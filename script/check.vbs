' массив анализируемых информационных баз 1С
Dim arr_bases
Dim arr_platform
Dim prefix

arr_bases = Array("buh", "erp", "ut")
arr_platform = Array("V83", "V82")
prefix = ""

For Each Platform In arr_platform
  Set Connector = CreateObject(Platform & ".COMConnector")
  Set Connection = Connector.ConnectAgent("tcp://localhost")
  Clasters = Connection.GetClusters()
  Set Cluster = Clasters (0)
  Connection.Authenticate Cluster,"",""
  Bases = Connection.GetInfoBases (Cluster)
  For Each base1c In Bases
    If inArray(arr_bases, base1c.Name)  >= 0 Then
      ' проверка существования и статуса в consul
    End If
  Next
Next

WScript.Quit 0

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