Module DWFfunc

    Private Class DWFInfo
        Private cDateCreated As String
        Private cPartNumber As String
        Private cRev As String
        Private cTitle As String
        Private cProductCode As String
        Private cDrawingDocumentType As String
        Private cProcessCode As String
        Private cProductCategory As String

        Public Sub New()
            cDateCreated = ""
            cTitle = ""
            cProductCode = ""
            cDrawingDocumentType = ""
            cProcessCode = ""
            cProductCategory = ""
        End Sub

        Property DateCreated() As String
            Get
                Return cDateCreated
            End Get
            Set(value As String)
                cDateCreated = value
            End Set
        End Property

        Property PartNumber() As String
            Get
                Return cPartNumber
            End Get
            Set(value As String)
                cPartNumber = value
            End Set
        End Property

        Property Rev() As String
            Get
                Return cRev
            End Get
            Set(value As String)
                cRev = value
            End Set
        End Property

        Property Title() As String
            Get
                Return cTitle
            End Get
            Set(value As String)
                cTitle = value
            End Set
        End Property

        Property ProductCode() As String
            Get
                Return cProductCode
            End Get
            Set(value As String)
                cProductCode = value
            End Set
        End Property

        Property DrawingDocumentType() As String
            Get
                Return cDrawingDocumentType
            End Get
            Set(value As String)
                cDrawingDocumentType = value
            End Set
        End Property

        Property ProcessCode() As String
            Get
                Return cProcessCode
            End Get
            Set(value As String)
                cProcessCode = value
            End Set
        End Property

        Property ProductCategory() As String
            Get
                Return cProductCategory
            End Get
            Set(value As String)
                cProductCategory = value
            End Set
        End Property

    End Class

    Private revLevel As Integer

    Public Sub ProcessDWFFolder(ByRef dt As System.Data.DataTable, ByVal dwfFolderPath As String)
        Dim dwfTempPath, dwfTableContent, dwfSheetContent As String
        Dim dwfInfo As DWFInfo

        revLevel = 0
        dwfTableContent = ""
        dwfSheetContent = ""

        dt.Columns.Add("NUMBER", GetType(String))
        dt.Columns.Add("PRODCAT", GetType(String))
        dt.Columns.Add("PRODCODE", GetType(String))
        dt.Columns.Add("PARENT", GetType(String))
        dt.Columns.Add("REV", GetType(String))
        dt.Columns.Add("PROCESSCODE", GetType(String))
        dt.Columns.Add("DOCTYPE", GetType(String))
        dt.Columns.Add("DESCRIPTION", GetType(String))
        dt.Columns.Add("MATERIAL", GetType(String))
        dt.Columns.Add("QTY", GetType(String))
        dt.Columns.Add("LG", GetType(String))
        dt.Columns.Add("WD", GetType(String))
        dt.Columns.Add("REV" + CStr(revLevel), GetType(String))
        dt.Columns.Add("REVNOTE" + CStr(revLevel), GetType(String))
        dt.Columns.Add("REVDATE" + CStr(revLevel), GetType(String))

        For Each dwfFile As String In My.Computer.FileSystem.GetFiles(dwfFolderPath, FileIO.SearchOption.SearchAllSubDirectories, "*.dwf")
            dwfTableContent = ""
            dwfSheetContent = ""
            dwfTempPath = UnpackDWF(dwfFile)
            GetXML(dwfTempPath, dwfTableContent, dwfSheetContent)
            dwfInfo = ParseXMLSheet(dwfSheetContent)
            ParseXMLPartList(dt, dwfTableContent, dwfInfo)
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

    Private Sub GetXML(ByVal dwfPath As String, ByRef partsListContent As String, ByRef sheetPathContent As String)
        Dim manifestContents As String = System.IO.File.ReadAllText(dwfPath + "\manifest.xml")

        Using reader As System.Xml.XmlReader = System.Xml.XmlReader.Create(New System.IO.StringReader(manifestContents))
            reader.MoveToContent()
            While reader.Read()
                If reader.Name = "dwf:Contents" And reader.NodeType <> Xml.XmlNodeType.EndElement Then
                    reader.ReadToDescendant("dwf:Content")
                    reader.MoveToAttribute("href")
                    Dim partpath As String = reader.ReadContentAsString
                    partsListContent = System.IO.File.ReadAllText(dwfPath + "\" + partpath)
                End If
                If reader.Name = "dwf:Toc" And reader.NodeType <> Xml.XmlNodeType.EndElement Then
                    reader.ReadToDescendant("dwf:Resource")
                    Do
                        If reader.MoveToAttribute("role") Then
                            If reader.ReadContentAsString = "2d streaming graphics" Then
                                While reader.ReadToNextSibling("dwf:Resource")
                                    reader.MoveToAttribute("role")
                                    If reader.ReadContentAsString = "descriptor" Then
                                        reader.MoveToAttribute("href")
                                        sheetPathContent = System.IO.File.ReadAllText(dwfPath + "\" + reader.ReadContentAsString)
                                    End If
                                End While
                            End If
                        End If
                        If Not reader.ReadToNextSibling("dwf:Resource") Then
                            Exit Do
                        End If
                    Loop
                End If
            End While
        End Using
    End Sub

    Private Function ParseXMLSheet(ByVal sheetContent As String) As DWFInfo
        Dim dwfInfo As New DWFInfo

        Using reader As System.Xml.XmlReader = System.Xml.XmlReader.Create(New System.IO.StringReader(sheetContent))
            reader.MoveToContent()
            While reader.Read()
                If reader.Name = "ePlot:Properties" And reader.NodeType <> Xml.XmlNodeType.EndElement Then
                    reader.ReadToDescendant("ePlot:Property")
                    Do
                        reader.MoveToAttribute("name")
                        Select Case reader.ReadContentAsString
                            Case "Date Created"
                                reader.MoveToAttribute("value")
                                dwfInfo.DateCreated = reader.ReadContentAsString
                            Case "Part Number"
                                reader.MoveToAttribute("value")
                                dwfInfo.PartNumber = reader.ReadContentAsString
                            Case "Revision Number"
                                reader.MoveToAttribute("value")
                                dwfInfo.Rev = reader.ReadContentAsString
                            Case "Title"
                                reader.MoveToAttribute("value")
                                dwfInfo.Title = reader.ReadContentAsString
                            Case "SL02_Product_Code"
                                reader.MoveToAttribute("value")
                                dwfInfo.ProductCode = reader.ReadContentAsString
                            Case "SL04_Drawing_Document_Type"
                                reader.MoveToAttribute("value")
                                dwfInfo.DrawingDocumentType = reader.ReadContentAsString
                            Case "SL12_Process Code"
                                reader.MoveToAttribute("value")
                                dwfInfo.ProcessCode = reader.ReadContentAsString
                            Case "SL13_Product Category"
                                reader.MoveToAttribute("value")
                                dwfInfo.ProductCategory = reader.ReadContentAsString
                        End Select
                        If Not reader.ReadToNextSibling("ePlot:Property") Then
                            Exit Do
                        End If
                    Loop
                End If
            End While
        End Using

        Return dwfInfo
    End Function

    Private Sub ParseXMLPartList(dt As System.Data.DataTable, ByVal tableContent As String, ByVal dwfInfo As DWFInfo)
        Dim dataRow, revRow As System.Data.DataRow
        Dim readerBuf As String
        Dim sheetRevLevel, pos As Integer
        Dim isFirstRev As Boolean = True

        Using reader As System.Xml.XmlReader = System.Xml.XmlReader.Create(New System.IO.StringReader(tableContent))
            reader.MoveToContent()
            While reader.Read()
                If reader.Name = "dwf:Properties" And reader.NodeType <> Xml.XmlNodeType.EndElement Then
                    Debug.Print(dt.Rows.Count)
                    dataRow = dt.NewRow()
                    dt.Rows.Add(dataRow)
                    dataRow("PRODCAT") = dwfInfo.ProductCategory
                    dataRow("PRODCODE") = dwfInfo.ProductCode
                    dataRow("PARENT") = dwfInfo.PartNumber
                    dataRow("REV") = dwfInfo.Rev
                    dataRow("PROCESSCODE") = dwfInfo.ProcessCode
                    dataRow("DOCTYPE") = dwfInfo.DrawingDocumentType

                    reader.ReadToDescendant("dwf:Property")
                    Do
                        reader.MoveToAttribute("name")
                        Select Case reader.ReadContentAsString
                            Case "REV"
                                If isFirstRev Then
                                    revRow = dataRow
                                    isFirstRev = False
                                Else
                                    dt.Rows.Remove(dataRow)
                                End If
                                reader.MoveToAttribute("value")
                                sheetRevLevel = CInt(reader.ReadContentAsString)
                                If sheetRevLevel > revLevel Then
                                    For i As Integer = revLevel + 1 To sheetRevLevel
                                        dt.Columns.Add("REV" & CStr(i))
                                        dt.Columns.Add("REVNOTE" & CStr(i))
                                        dt.Columns.Add("REVDATE" & CStr(i))
                                    Next
                                    revLevel += 1
                                End If
                                revRow("REV" + CStr(sheetRevLevel)) = reader.ReadContentAsString
                                revRow("DESCRIPTION") = dwfInfo.Title
                                revRow("NUMBER") = dwfInfo.PartNumber
                                revRow("QTY") = "1"
                                While reader.ReadToNextSibling("dwf:Property")
                                    reader.MoveToAttribute("name")
                                    Select Case reader.ReadContentAsString
                                        Case "DESCRIPTION"
                                            reader.MoveToAttribute("value")
                                            revRow("REVNOTE" & CStr(sheetRevLevel)) = reader.ReadContentAsString
                                        Case "DATE"
                                            If reader.MoveToAttribute("value") Then
                                                revRow("REVDATE" & CStr(sheetRevLevel)) = reader.ReadContentAsString
                                            Else
                                                revRow("REVDATE" & CStr(sheetRevLevel)) = dwfInfo.DateCreated
                                            End If
                                    End Select
                                End While
                                Exit Do
                            Case "PART No."
                                reader.MoveToAttribute("value")
                                If Not IsDBNull(dataRow("NUMBER")) Then
                                    If dataRow("NUMBER") = "-" Or dataRow("NUMBER") = "" Then
                                        dataRow("NUMBER") = reader.ReadContentAsString
                                    End If
                                Else
                                    dataRow("NUMBER") = reader.ReadContentAsString
                                End If
                            Case "STOCK No."
                                reader.MoveToAttribute("value")
                                If Not IsDBNull(dataRow("NUMBER")) Then
                                    If dataRow("NUMBER") = "-" Or dataRow("NUMBER") = "" Then
                                        dataRow("NUMBER") = reader.ReadContentAsString
                                    End If
                                Else
                                    dataRow("NUMBER") = reader.ReadContentAsString
                                End If
                            Case "STOCK NUMBER"
                                reader.MoveToAttribute("value")
                                If Not IsDBNull(dataRow("NUMBER")) Then
                                    If dataRow("NUMBER") = "-" Or dataRow("NUMBER") = "" Then
                                        dataRow("NUMBER") = reader.ReadContentAsString
                                    End If
                                Else
                                    dataRow("NUMBER") = reader.ReadContentAsString
                                End If
                            Case "DESCRIPTION"
                                reader.MoveToAttribute("value")
                                dataRow("DESCRIPTION") = reader.ReadContentAsString
                            Case "MATERIAL"
                                reader.MoveToAttribute("value")
                                dataRow("MATERIAL") = reader.ReadContentAsString
                            Case "LENGTH"
                                reader.MoveToAttribute("value")
                                readerBuf = LCase(reader.ReadContentAsString)
                                pos = InStr(readerBuf, "x")
                                If pos > 0 Then
                                    dataRow("LG") = Trim(Left(readerBuf, pos - 1))
                                    dataRow("WD") = Trim(Right(readerBuf, readerBuf.Length - pos))
                                Else
                                    dataRow("LG") = readerBuf
                                End If
                            Case "DIAMETER"
                                reader.MoveToAttribute("value")
                                dataRow("LG") = reader.ReadContentAsString
                            Case "QTY"
                                reader.MoveToAttribute("value")
                                dataRow("QTY") = reader.ReadContentAsString
                            Case Else
                        End Select

                        If Not reader.ReadToNextSibling("dwf:Property") Then
                            Exit Do
                        End If
                    Loop
                    Try
                        If IsDBNull(dataRow("NUMBER")) Then
                            dt.Rows.Remove(dataRow)
                        End If
                    Catch ex As Exception
                        Debug.Print(ex.Message)
                    End Try
                End If
            End While
        End Using
    End Sub
End Module
