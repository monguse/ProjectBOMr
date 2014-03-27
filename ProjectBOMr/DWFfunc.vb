Module DWFfunc

    Public Sub ProcessDWFFolder(ByVal dwfFolderPath As String)
        Dim dwfTempPath, dwfHREF As String
        Dim projBOM As New System.Data.DataTable

        projBOM.Columns.Add("NUMBER", GetType(String))
        projBOM.Columns.Add("PARENT", GetType(String))
        projBOM.Columns.Add("PROCESS CODE", GetType(String))
        projBOM.Columns.Add("DESCRIPTION", GetType(String))
        projBOM.Columns.Add("MATERIAL", GetType(String))
        projBOM.Columns.Add("QTY", GetType(String))
        projBOM.Columns.Add("LG", GetType(String))
        projBOM.Columns.Add("WD", GetType(String))

        For Each dwfFile As String In My.Computer.FileSystem.GetFiles(dwfFolderPath, FileIO.SearchOption.SearchAllSubDirectories, "*.dwf")
            dwfTempPath = UnpackDWF(dwfFile)
            dwfHREF = GetXMLTable(dwfTempPath + "\manifest.xml")
            ParseXMLTables(projBOM, dwfTempPath + "\" + dwfHREF)
            My.Computer.FileSystem.DeleteDirectory(dwfTempPath, FileIO.DeleteDirectoryOption.DeleteAllContents)
        Next
    End Sub

    Private Function UnpackDWF(ByVal dwfPath As String) As String
        Dim tempPath As String

        tempPath = My.Computer.FileSystem.SpecialDirectories.Temp + "\" + "_ProjectBOMr"
        Using dwfFile As Ionic.Zip.ZipFile = Ionic.Zip.ZipFile.Read(dwfPath)
            Dim e As Ionic.Zip.ZipEntry
            For Each e In dwfFile
                e.Extract(tempPath, Ionic.Zip.ExtractExistingFileAction.OverwriteSilently)
            Next
        End Using

        Return tempPath
    End Function

    Private Sub ParseXMLTables(ByRef dt As System.Data.DataTable, ByVal dwfXMLPath As String)
        Dim partListContents As String = System.IO.File.ReadAllText(dwfXMLPath)

        Using reader As System.Xml.XmlReader = System.Xml.XmlReader.Create(New System.IO.StringReader(partListContents))
            While reader.Read()
                If reader.Name = "dwf:Properties" Then
                    While reader.NodeType <> Xml.XmlNodeType.EndElement
                        reader.Read()
                        If reader.Name = "dwf:Property" Then
                            reader.MoveToAttribute("name")
                            Debug.Print(reader.ReadContentAsString)
                            reader.MoveToAttribute("value")
                            Debug.Print(reader.ReadContentAsString)
                        End If
                    End While
                End If
            End While
        End Using

    End Sub

    Private Sub Get

    Private Function GetXMLTable(ByVal manifestPath As String) As String
        Dim manifestContents As String = System.IO.File.ReadAllText(manifestPath)

        Using reader As System.Xml.XmlReader = System.Xml.XmlReader.Create(New System.IO.StringReader(manifestContents))
            reader.ReadToFollowing("dwf:Content")
            reader.MoveToAttribute("href")
            Return reader.ReadContentAsString()
        End Using

    End Function
End Module
