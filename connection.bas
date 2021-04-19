Attribute VB_Name = "connection"
Public Sub fileconnection(ByVal strpath As String, ByVal list As TreeView, ByVal flag As Boolean)
    Dim fso As New FileSystemObject
    Dim folder As folder
    Dim node As node
    Set folder = fso.GetFolder(strpath)
    Dim f As file
    Dim strkey As String
    strkey = strpath
    If flag Then
    Else
        Set node = list.Nodes.Add(, , strkey, folder.Path)
    End If
    Dim sf As folder
    If folder.Files.Count > 0 Then
        For Each f In folder.Files
            If f.Name Like "*.ta" Or f.Name Like "*.la" Or f.Name Like "*.pa" Or f.Name Like "*.db" Then
            list.Nodes.Add strkey, tvwChild, f.Path, f.Name
            End If
        Next f
    End If
    If folder.SubFolders.Count > 0 Then
        For Each sf In folder.SubFolders
            list.Nodes.Add strkey, tvwChild, sf.Path, sf.Name
            Call fileconnection(sf.Path, list, True)
        Next sf
    End If
    Set f = Nothing
    Set sf = Nothing
End Sub
