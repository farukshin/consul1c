' массив анализируемых информационных баз 1С
arr_bases = ["buh", "erp", "ut"]
arr_platform = ["V83", "V82"]

For Each Platform In arr_platform
  Set Connector = CreateObject(Platform & ".COMConnector")
  Set Connection = Connector.ConnectAgent("tcp://localhost")
  Clasters = Connection.GetClusters()
  Set Cluster = Clasters (0)
  Connection.Authenticate Cluster,"",""
  Bases = Connection.GetInfoBases (Cluster)
  For Each base1c In Bases
    If base1c.Name = "erp" Then ' существует в массиве arr_bases
      ' проверка существования и статуса в consul
    End If
  Next
Next

WScript.Quit 0
