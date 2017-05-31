Dim arr_bases
Dim arr_platform
Dim prefix
Dim nodename
Dim ServiceNode()
Dim counter

arr_bases = Array("buh", "erp", "uh")
arr_platform = Array("V83")
prefix = ""

setnodename()
For Each Platform In arr_platform
  counter = 0
  Set Connector = CreateObject(Platform & ".COMConnector")
  Set Connection = Connector.ConnectAgent("tcp://localhost")
  Clasters = Connection.GetClusters()
  Set Cluster = Clasters (0)
  Connection.Authenticate Cluster,"",""
  Bases = Connection.GetInfoBases (Cluster)
  get_service_node()
  For Each base1c In Bases
    If inArray(arr_bases, base1c.Name)  >= 0 Then
      If inArray(ServiceNode, base1c.Name) >= 0 Then
        ' база существует, проверяем статус
        ' check_and_set_status(base1c.Name, "pass") 
      Else:
        ' добавляем базу в качестве сервиса
        ' add_service(base1c.Name, "pass")
      End If
    End If
  Next
Next

WScript.Quit 0

public Function add_service(name, status)
  On Error Resume Next
  ' отправляем запрос в consul
  Err.Clear()
End Function


public Function get_service_node()
  On Error Resume Next
  url = "http://localhost:8500/v1/catalog/node/" + nodename + "?pretty" 
  set json = CreateObject("Chilkat_9_5_0.JsonObject")
  success = json.Load(GetHTTPResponse(url))
  Set Services  = json.ObjectOf("Services")
  numServices = Services.Size
  For i = 0 To numServices - 1
    Set srv = Services.ObjectAt(i)
    If inArray(srv.ArrayOf("Tags"), "1c") Then
      ReDim preserve ServiceNode(counter)
      ServiceNode(counter) = srv.StringOf("ID")
      WScript.echo "!"+ ServiceNode(counter) +"!"
      counter = counter + 1
    End If
  Next
  Err.Clear()  
End Function

public Function inArray(arr, obj)
  On Error Resume Next
  Dim x: x = -1
  If isArray(arr) Then
    For i = 0 To UBound(arr)
      If arr(i) = obj Then
        x = i
        Exit For
      End If
    Next
  End If
	
  Err.Clear()
  inArray = x
End Function

public Function setnodename()
  nodename = CreateObject("wscript.network").ComputerName
End Function

Private Function GetHTTPResponse(sURL)
    Dim oXMLHTTP
    On Error Resume Next
    Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    With oXMLHTTP
        .Open "GET", sURL, False
        .SetRequestHeader "Cache-Control", "max-age=0"
        .SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.41 Safari/537.36 OPR/35.0.2066.10 (Edition beta)"
        .SetRequestHeader "Accept-Encoding", "deflate"
        .SetRequestHeader "Accept-Language", "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4"
        .send
        GetHTTPResponse = .responseText
    End With
    Set oXMLHTTP = Nothing
End Function
