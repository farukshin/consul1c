Dim arrCheckBases
Dim checkAllBases: checkAllBases = False
Dim arrCheckPlatform
Dim prefix: prefix = ""
Dim nodename
Dim arrServiceNode()
Dim counter
Dim Host: Host = "http://localhost:8500"
Dim arrBaseCluster()

arrCheckBases = Array("buh", "erp", "uh")
arrCheckPlatform = Array("V83")

setnodename()
For Each Platform In arrCheckPlatform
  Set Connector = CreateObject(Platform & ".COMConnector")
  Set Connection = Connector.ConnectAgent("tcp://localhost")
  Clasters = Connection.GetClusters()
  Set Cluster = Clasters (0)
  Connection.Authenticate Cluster,"",""
  Bases = Connection.GetInfoBases(Cluster)
  count = 0
  For Each base1c In Bases
    ReDim preserve arrBaseCluster(count)
    arrBaseCluster(count) = base1c.Name
    count = count + 1
  Next
  
  If count > 0 Then
    counter = 0
    get_service_node()

    For Each base1c In arrBaseCluster
      If (inArray(arrCheckBases, base1c)  >= 0) OR (checkAllBases = True) Then
        If inArray(arrServiceNode, base1c) = -1 Then
          r = setService(base1c, "pass")
        End If
      End If
    Next
    
    For Each srvBase1c In arrServiceNode
      If inArray(arrBaseCluster, srvBase1c) = -1 Then
        r = setService(srvBase1c, "del")
      End If
    Next
 
  End If
Next

WScript.Quit 0

public Function setService(name, status)
  On Error Resume Next
  If status = "pass" Then
    body = "{""ID"": """+name+""", ""Name"": """+name+""", ""Tags"": [""1c""]}"
    resp = GetHTTPResponse(Host + "/v1/agent/service/register", "POST", body)
  ElseIf status = "del" Then
    resp = GetHTTPResponse(Host + "/v1/agent/service/deregister/" + name, "GET", "")
  End If
  Err.Clear()
End Function


public Function get_service_node()
  On Error Resume Next
  url = Host + "/v1/catalog/node/" + nodename + "?pretty" 
  set json = CreateObject("Chilkat_9_5_0.JsonObject")
  success = json.Load(GetHTTPResponse(url, "GET", ""))
  Set Services  = json.ObjectOf("Services")
  numServices = Services.Size
  For i = 0 To numServices - 1
    Set srv = Services.ObjectAt(i)
    If inArray(srv.ArrayOf("Tags"), "1c") Then
      ReDim preserve arrServiceNode(counter)
      arrServiceNode(counter) = srv.StringOf("ID")
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

Private Function GetHTTPResponse(sURL, method, body)
    Dim oXMLHTTP
    On Error Resume Next
    Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    With oXMLHTTP
        .Open method, sURL, False
        .SetRequestHeader "Cache-Control", "max-age=0"
        .SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.41 Safari/537.36 OPR/35.0.2066.10 (Edition beta)"
        .SetRequestHeader "Accept-Encoding", "deflate"
        .SetRequestHeader "Accept-Language", "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4"
        .send body
        GetHTTPResponse = .responseText
    End With
    Set oXMLHTTP = Nothing
End Function
