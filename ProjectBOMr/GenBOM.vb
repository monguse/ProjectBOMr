Module GenBOM
    Private Const BaseFilePath = "C:\projects\"

    ' Gets the user's history with the specified project number, history is in AppData
    Private Sub GetHistory(iProjNum As Long)

    End Sub

    ' After generating the combined BOM from the project folder, store BOM details for future deltas
    Private Sub SetHistory()

    End Sub

    Public Sub Cout(iStr As String)
        UserForm.rtb_Console.AppendText(iStr + vbNewLine)
        'UserForm.rtb_Console.ScrollToCaret()
    End Sub

    Public Function ProcessFolder(folderPath As String, sProjNum As String) As String
        Dim rawBOM As New System.Data.DataTable
        Dim parentBOM As New System.Data.DataTable
        Dim pType1BOM As New System.Data.DataTable
        Dim pType2BOM As New System.Data.DataTable
        Dim pType3BOM As New System.Data.DataTable
        Dim unknownBOM As New System.Data.DataTable

        ProcessFolder = ""

        'CSVFolderToTable(dt:=rawBOM, folderPath:=folderPath)
        ProcessDWFFolder(rawBOM, folderPath)
        'SortByParent(dt:=parentBOM, st:=rawBOM)
        'sortByType(dt:=pType1BOM, st:=rawBOM, processCode:=1)
        'sortByType(dt:=pType2BOM, st:=rawBOM, processCode:=2)
        'sortByType(dt:=pType3BOM, st:=rawBOM, processCode:=3)
        'sortByType(dt:=unknownBOM, st:=rawBOM, processCode:=-1)

        If DumpTablesToExcel(bomSavePath:=folderPath + "\BOM " + sProjNum + ".xlsx", _
                          stRaw:=rawBOM, _
                          stParent:=parentBOM, _
                          stType1:=pType1BOM, _
                          stType2:=pType2BOM, _
                          stType3:=pType3BOM, _
                          stUnknown:=unknownBOM) Then
            ProcessFolder = folderPath + "\BOM " + sProjNum + ".xlsx"
        End If
    End Function

    Public Function FindPath(sPath As String) As String
        FindPath = ""

        For Each foundDir As String In My.Computer.FileSystem.GetDirectories(BaseFilePath)
            If InStr(foundDir, sPath) > 0 Then
                FindPath = foundDir
            End If
        Next
    End Function

    Private Function TableToCSV(ByVal sourceTable As System.Data.DataTable, ByVal filePathName As String) As Boolean
        'Writes a datatable back into a csv 
        Try
            Dim sb As New System.Text.StringBuilder
            Dim nameArray(200) As Object
            Dim i As Integer = 0
            For Each col As DataColumn In sourceTable.Columns
                nameArray(i) = CType(col.ColumnName, Object)
                i += 1
            Next col
            ReDim Preserve nameArray(i - 1)
            sb.AppendLine(String.Join(",", Array.ConvertAll(Of Object, String)(nameArray, _
                            Function(o As Object) If(o.ToString.Contains(","), _
                            ControlChars.Quote & o.ToString & ControlChars.Quote, o))))
            For Each dr As DataRow In sourceTable.Rows
                sb.AppendLine(String.Join(",", Array.ConvertAll(Of Object, String)(dr.ItemArray, _
                                Function(o As Object) If(o.ToString.Contains(","), _
                                ControlChars.Quote & o.ToString & ControlChars.Quote, o.ToString))))
            Next
            System.IO.File.WriteAllText(filePathName, sb.ToString)
            Return True
        Catch ex As Exception
            Debug.Print(ex.ToString())
            Return False
        End Try
    End Function

    Private Function GetExcelColumnName(xIndex As Integer) As String
        GetExcelColumnName = ""
        Dim modulo, dividend As Integer
        dividend = xIndex

        While dividend > 0
            modulo = (dividend - 1) Mod 26
            GetExcelColumnName = Chr(65 + modulo).ToString + GetExcelColumnName
            dividend = (dividend - modulo) / 26
        End While

    End Function

    Private Function Frac2Num(ByVal X As String) As Double
        Dim P As Integer, N As Double, Num As Double, Den As Double
        X = Trim$(X)
        P = InStr(X, "/")
        If P = 0 Then
            N = Val(X)
        Else
            Den = Val(Mid$(X, P + 1))
            If Den = 0 Then Error 11 ' Divide by zero
            X = Trim$(Left$(X, P - 1))
            P = InStr(X, " ")
            If P = 0 Then
                Num = Val(X)
            Else
                Num = Val(Mid$(X, P + 1))
                N = Val(Left$(X, P - 1))
            End If
        End If
        If Den <> 0 Then
            N = N + Num / Den
        End If
        Frac2Num = N
    End Function

    Private Sub CSVFolderToTable(ByRef dt As System.Data.DataTable, ByVal folderPath As String)
        Dim parser As Microsoft.VisualBasic.FileIO.TextFieldParser

        dt.Columns.Add(columnName:="NUMBER", type:=GetType(String))
        dt.Columns.Add(columnName:="PARENT", type:=GetType(String))
        dt.Columns.Add(columnName:="PROCESS CODE", type:=GetType(String))
        dt.Columns.Add(columnName:="DESCRIPTION", type:=GetType(String))
        dt.Columns.Add(columnName:="MATERIAL", type:=GetType(String))
        dt.Columns.Add(columnName:="QTY", type:=GetType(String))
        dt.Columns.Add(columnName:="UNITQTY", type:=GetType(String))
        dt.Columns.Add(columnName:="LG", type:=GetType(String))
        dt.Columns.Add(columnName:="WD", type:=GetType(String))

        For Each foundFile As String In My.Computer.FileSystem.GetFiles(folderPath)
            If Not InStr(foundFile, "BOM") Then
                If System.IO.Path.GetExtension(foundFile) = ".csv" Then
                    parser = New Microsoft.VisualBasic.FileIO.TextFieldParser(foundFile)
                    parser.Delimiters = New String() {","}
                    parser.HasFieldsEnclosedInQuotes = True
                    parser.TextFieldType = FileIO.FieldType.Delimited
                    parser.TrimWhiteSpace = True
                    parser.ReadFields()
                    While Not parser.EndOfData
                        dt.Rows.Add(parser.ReadFields())
                    End While
                    parser.Close()
                End If
            End If
        Next

    End Sub

    Private Sub SortByParent(ByRef dt As System.Data.DataTable, ByRef st As System.Data.DataTable)
        Dim foundParent, isFirstRow, isTheEnd As Boolean
        Dim queryNumber As String
        Dim numColumns, dtrow, dtcolumns As Integer
        Dim dtDataRow As System.Data.DataRow

        dt.Columns.Add(columnName:="DWG QTY", type:=GetType(String))
        dt.Columns.Add(columnName:="DESCRIPTION", type:=GetType(String))
        dt.Columns.Add(columnName:="1", type:=GetType(String))

        queryNumber = ""

        For stRowParent As Integer = 0 To st.Rows.Count - 1
            foundParent = False
            queryNumber = st.Rows(stRowParent)(1)

            If queryNumber = "" Then
                Continue For
            End If

            For stRowChild As Integer = 0 To st.Rows.Count - 1
                If stRowParent = stRowChild Then
                    Continue For
                ElseIf queryNumber = st.Rows(stRowChild)(0) Then
                    foundParent = True
                End If
            Next

            If Not foundParent Then
                dt.Rows.Add({st.Rows(stRowParent)(5), st.Rows(stRowParent)(3), queryNumber})
            End If
        Next


        numColumns = 1
        dtcolumns = numColumns
        dtrow = 0
        Dim runcount As Integer = 0
        Do
            Debug.Print("runcount: " & runcount & " " & dtrow & " " & dt.Rows.Count)
            isFirstRow = True
            isTheEnd = True
            queryNumber = ""
            Do While dtrow < dt.Rows.Count
                If Not IsDBNull(dt.Rows(dtrow)(CStr(dtcolumns))) Then
                    queryNumber = dt.Rows(dtrow)(CStr(dtcolumns))
                Else
                    dtrow += 1
                    Continue Do
                End If
                Debug.Print("meep: " & queryNumber & " " & dtrow)
                For stRow As Integer = 0 To st.Rows.Count - 1
                    If st.Rows(stRow)(1) = queryNumber And st.Rows(stRow)(0) <> queryNumber And st.Rows(stRow)(0) <> "" Then
                        If isFirstRow Then
                            numColumns += 1
                            dt.Columns.Add(columnName:=CStr(numColumns), type:=GetType(String))
                            isFirstRow = False
                            isTheEnd = False
                        End If
                        If st.Rows(stRow)(0)(0) = "4"c Then
                            Debug.Print(dtrow & " " & dtcolumns & " " & dt.Rows.Count & " " & stRow & " " & st.Rows(stRow)(0) & ":" & st.Rows(stRow)(1) & ":" & queryNumber)
                            dtDataRow = dt.NewRow()
                            dtDataRow(columnName:=CStr(numColumns)) = st.Rows(stRow)(0)
                            dtDataRow(columnName:="DWG QTY") = st.Rows(stRow)(5)
                            dtDataRow(columnName:="DESCRIPTION") = st.Rows(stRow)(3)
                            dt.Rows.InsertAt(row:=dtDataRow, pos:=dtrow + 1)
                        End If
                    End If
                Next
                dtrow += 1
            Loop
            If isTheEnd Then
                Exit Do
            End If
            runcount += 1
            dtrow = 0
            dtcolumns = numColumns
        Loop
    End Sub

    Private Sub sortByType(ByRef dt As System.Data.DataTable, ByRef st As System.Data.DataTable, processCode As Integer)

        Dim queryNumber, queryParent, queryDescription, queryMaterial As String
        Dim queryLG, queryWD, queryQty As Double
        Dim foundChild As Boolean
        Dim dtRow, queryCode As Integer

        dt.Columns.Add(columnName:="NUMBER", type:=GetType(String))
        dt.Columns.Add(columnName:="DESCRIPTION", type:=GetType(String))
        dt.Columns.Add(columnName:="MATERIAL", type:=GetType(String))
        dt.Columns.Add(columnName:="QTY", type:=GetType(Double))
        dt.Columns.Add(columnName:="QTY UNIT", type:=GetType(String))
        dt.Columns.Add(columnName:="PARENTS", type:=GetType(String))
        dt.Columns.Add(columnName:="PROCESS CODE", type:=GetType(Integer))

        For stRow As Integer = 0 To st.Rows.Count - 1
            queryNumber = st.Rows(stRow)(0)
            queryParent = st.Rows(stRow)(1)
            queryDescription = st.Rows(stRow)(3)
            queryMaterial = st.Rows(stRow)(4)
            queryCode = st.Rows(stRow)(2)

            queryLG = Frac2Num(st.Rows(stRow)(7))
            queryWD = Frac2Num(st.Rows(stRow)(8))
            queryQty = CDbl(st.Rows(stRow)(5))
            Debug.Print("moop: " & queryLG & " " & queryWD)

            foundChild = False

            If processCode = -1 Then
                If queryCode = 1 Or queryCode = 2 Or queryCode = 3 Or queryCode = 0 Then
                    Continue For
                End If
            ElseIf queryCode <> processCode Or queryCode = 0 Then
                Continue For
            End If

            For dtRow = 0 To dt.Rows.Count - 1
                If dt.Rows(dtRow)(0) = queryNumber And _
                    dt.Rows(dtRow)(1) = queryDescription And _
                    dt.Rows(dtRow)(2) = queryMaterial Then
                    foundChild = True
                    Exit For
                End If
            Next

            If st.Rows(stRow)(0) <> st.Rows(stRow)(1) Then
                If foundChild Then
                    If queryLG = 0 Then
                        dt.Rows(dtRow)("QTY") = dt.Rows(dtRow)("QTY") + queryQty
                    ElseIf queryWD = 0 Then
                        dt.Rows(dtRow)("QTY") = dt.Rows(dtRow)("QTY") + (queryQty * queryLG)
                    Else
                        dt.Rows(dtRow)("QTY") = dt.Rows(dtRow)("QTY") + (queryQty * queryLG * queryWD)
                    End If
                    If InStr(dt.Rows(dtRow)("PARENTS"), queryParent) = 0 Then
                        dt.Rows(dtRow)("PARENTS") = dt.Rows(dtRow)("PARENTS") & ", " & queryParent
                    End If
                Else
                    dt.Rows.Add({queryNumber, queryDescription, queryMaterial, 0, "", queryParent, queryCode})
                    If queryLG = 0 Then
                        dt.Rows(dt.Rows.Count - 1)("QTY") = queryQty
                        dt.Rows(dt.Rows.Count - 1)("QTY UNIT") = "EACH"
                    ElseIf queryWD = 0 Then
                        dt.Rows(dt.Rows.Count - 1)("QTY") = queryQty * queryLG
                        dt.Rows(dt.Rows.Count - 1)("QTY UNIT") = "INCH"
                    Else
                        dt.Rows(dt.Rows.Count - 1)("QTY") = queryQty * queryLG * queryWD
                        dt.Rows(dt.Rows.Count - 1)("QTY UNIT") = "SQ. INCH"
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub TableToWorksheet(ByRef st As System.Data.DataTable, ByRef ws As Excel.Worksheet)
        For col As Integer = 0 To st.Columns.Count - 1
            ws.Cells(1, col + 1) = st.Columns(col).ColumnName
            For row As Integer = 0 To st.Rows.Count - 1
                ws.Cells(row + 2, col + 1) = st.Rows(row)(col)
            Next
        Next
    End Sub

    Private Function DumpTablesToExcel(ByVal bomSavePath As String, _
                                  ByRef stRaw As System.Data.DataTable, _
                                  ByRef stParent As System.Data.DataTable, _
                                  ByRef stType1 As System.Data.DataTable, _
                                  ByRef stType2 As System.Data.DataTable, _
                                  ByRef stType3 As System.Data.DataTable, _
                                  ByRef stUnknown As System.Data.DataTable) As Boolean

        Dim xlApp As Excel.Application = Nothing
        Dim xlWB As Excel.Workbook = Nothing
        Dim xlSh As Excel.Worksheet = Nothing
        DumpTablesToExcel = False

        Try
            xlApp = New Excel.Application()
            xlApp.AlertBeforeOverwriting = False
            xlApp.DisplayAlerts = False

            xlWB = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet)
            xlWB.SaveAs(Filename:=bomSavePath, FileFormat:=XlFileFormat.xlWorkbookDefault)

            xlSh = DirectCast(xlWB.Sheets(1), Excel.Worksheet)
            xlSh.Name = "RAW DATA"
            TableToWorksheet(st:=stRaw, ws:=xlSh)
            xlSh.Columns.AutoFit()
            xlSh.Rows(1).font.bold = True
            xlSh.Range(xlSh.Cells(1, 1), xlSh.Cells(1, 9)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Plum)
            xlSh.Columns(1).horizontalalignment = Excel.XlHAlign.xlHAlignLeft

            xlSh = DirectCast(xlWB.Worksheets.Add(), Excel.Worksheet)
            xlSh.Name = "Process Code Unknown"
            TableToWorksheet(st:=stUnknown, ws:=xlSh)
            xlSh.Columns.AutoFit()
            xlSh.Rows(1).font.bold = True
            xlSh.Range(xlSh.Cells(1, 1), xlSh.Cells(1, 7)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Plum)
            xlSh.Columns(1).horizontalalignment = Excel.XlHAlign.xlHAlignLeft
            xlSh.Columns(6).horizontalalignment = Excel.XlHAlign.xlHAlignLeft

            xlSh = DirectCast(xlWB.Worksheets.Add(), Excel.Worksheet)
            xlSh.Name = "Process Code 3"
            TableToWorksheet(st:=stType3, ws:=xlSh)
            xlSh.Columns.AutoFit()
            xlSh.Rows(1).font.bold = True
            xlSh.Range(xlSh.Cells(1, 1), xlSh.Cells(1, 7)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Plum)
            xlSh.Columns(1).horizontalalignment = Excel.XlHAlign.xlHAlignLeft
            xlSh.Columns(6).horizontalalignment = Excel.XlHAlign.xlHAlignLeft

            xlSh = DirectCast(xlWB.Worksheets.Add(), Excel.Worksheet)
            xlSh.Name = "Process Code 2"
            TableToWorksheet(st:=stType2, ws:=xlSh)
            xlSh.Columns.AutoFit()
            xlSh.Rows(1).font.bold = True
            xlSh.Range(xlSh.Cells(1, 1), xlSh.Cells(1, 7)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Plum)
            xlSh.Columns(1).horizontalalignment = Excel.XlHAlign.xlHAlignLeft
            xlSh.Columns(6).horizontalalignment = Excel.XlHAlign.xlHAlignLeft

            xlSh = DirectCast(xlWB.Worksheets.Add(), Excel.Worksheet)
            xlSh.Name = "Process Code 1"
            TableToWorksheet(st:=stType1, ws:=xlSh)
            xlSh.Columns.AutoFit()
            xlSh.Rows(1).font.bold = True
            xlSh.Range(xlSh.Cells(1, 1), xlSh.Cells(1, 7)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Plum)
            xlSh.Columns(1).horizontalalignment = Excel.XlHAlign.xlHAlignLeft
            xlSh.Columns(6).horizontalalignment = Excel.XlHAlign.xlHAlignLeft

            xlSh = DirectCast(xlWB.Worksheets.Add(), Excel.Worksheet)
            xlSh.Name = "Project Hierarchy"
            TableToWorksheet(st:=stParent, ws:=xlSh)
            xlSh.Columns.AutoFit()
            xlSh.Rows(1).font.bold = True
            xlSh.Range(xlSh.Cells(1, 1), xlSh.Cells(1, stParent.Columns.Count)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Plum)

        Catch ex As Exception
            Debug.Print("DumpTablesToExcel1: " & ex.Message)
        End Try

        Try
            If xlWB IsNot Nothing Then
                xlWB.Save()
                xlWB.Close(Nothing, Nothing, Nothing)
            End If

            xlApp.Quit()

            If xlSh IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSh)
            End If

            If xlWB IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWB)
            End If

            If xlApp IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
            End If
        Catch ex As Exception
            Debug.Print("DumpTablesToExcel2: " & ex.Message)
        Finally
            DumpTablesToExcel = True
        End Try

        xlWB = Nothing
        xlApp = Nothing
        ' force final cleanup!
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Function

End Module
