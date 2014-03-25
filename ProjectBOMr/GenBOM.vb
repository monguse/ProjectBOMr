Module GenBOM
    Private Const BaseFilePath = "C:\projects\"
    Private xlApp As Excel.Application = Nothing
    Private xlWB As Excel.Workbook = Nothing
    Private a_sErrorMessage As String = String.Empty

    ' Gets the user's history with the specified project number
    ' The history files are stored in the user's AppData
    Private Sub GetHistory(iProjNum As Long)

    End Sub

    ' After generating the combined BOM from the project folder, store BOM details for future deltas
    Private Sub SetHistory()

    End Sub

    Public Sub Cout(iStr As String)
        UserForm.rtb_Console.AppendText(iStr + vbNewLine)
        'UserForm.rtb_Console.ScrollToCaret()
    End Sub

    Public Function ProcessFolder(sPath As String, sProjNum As String) As String
        Dim combinedBOM As New System.Data.DataTable
        ProcessFolder = ""

        Dim isFirst As Boolean = True
        For Each foundFile As String In My.Computer.FileSystem.GetFiles(sPath)
            If Not InStr(foundFile, "BOM") Then
                If System.IO.Path.GetExtension(foundFile) = ".csv" Then
                    'Cout(foundFile)
                    AddCsvToTable(foundFile, combinedBOM, isFirst)
                    isFirst = False
                End If
            End If
        Next
        'Dim fPath As String = sPath + "\BOM " + sProjNum + ".csv"
        Dim fPath As String = sPath + "\BOM " + sProjNum + ".xlsx"
        'If TableToCSV(combinedBOM, fPath) Then
        OpenExcel()
        If TableToExcel(combinedBOM, fPath) Then
            SortByParent(fPath)
            SortByType(fPath)
            ProcessFolder = fPath
        End If
        CloseExcel()


    End Function

    Public Function FindPath(sPath As String) As String
        FindPath = ""

        For Each foundDir As String In My.Computer.FileSystem.GetDirectories(BaseFilePath)
            If InStr(foundDir, sPath) > 0 Then
                FindPath = foundDir
            End If
        Next
    End Function

    Private Sub OpenExcel()
        Try
            xlApp = New Excel.Application()
            xlApp.AlertBeforeOverwriting = False
            xlApp.DisplayAlerts = False
        Catch ex As Exception
            a_sErrorMessage = "Error in export: " & ex.Message
        End Try
    End Sub

    Private Sub CloseExcel()
        Try
            If xlWB IsNot Nothing Then
                xlWB.Save()
                xlWB.Close(Nothing, Nothing, Nothing)
            End If
            xlApp.Workbooks.Close()
            xlApp.Quit()
            If xlWB IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWB)
            End If
            If xlApp IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
            End If
        Catch
        End Try
        xlWB = Nothing
        xlApp = Nothing
        ' force final cleanup!
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub AddCsvToTable(ByVal sPath As String, ByRef dt As System.Data.DataTable, isFirst As Boolean)
        Dim sBuf As Object
        Dim bBool As Boolean = isFirst
        Dim parser As New Microsoft.VisualBasic.FileIO.TextFieldParser(sPath)
        parser.Delimiters = New String() {","}
        parser.HasFieldsEnclosedInQuotes = True 'use if data may contain delimiters 
        parser.TextFieldType = FileIO.FieldType.Delimited
        parser.TrimWhiteSpace = True
        If Not bBool Then
            sBuf = parser.ReadFields()
        End If
        While Not parser.EndOfData
            AddValuesToTable(parser.ReadFields, dt, bBool)
            bBool = False
        End While
        parser.Close()
    End Sub

    Private Sub AddValuesToTable(ByRef source() As String, ByRef destination As System.Data.DataTable, Optional ByVal HeaderFlag As Boolean = False)
        Dim existing As Integer = destination.Columns.Count
        If HeaderFlag Then
            For i As Integer = 0 To source.Length - existing - 1
                destination.Columns.Add(source(i).ToString, GetType(String))
            Next i
            Exit Sub
        End If
        For i As Integer = 0 To source.Length - existing - 1
            destination.Columns.Add("Column" & (existing + 1 + i).ToString, GetType(String))
        Next
        destination.Rows.Add(source)
    End Sub

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
            Console.WriteLine(ex.ToString())
            Return False
        End Try
    End Function

    Private Function TableToExcel(ByVal sourceTable As System.Data.DataTable, ByVal filePathName As String) As Boolean
        TableToExcel = False
        Dim xlSh As Excel.Worksheet = Nothing
        Try
            xlWB = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet)
            xlWB.SaveAs(Filename:=filePathName, FileFormat:=XlFileFormat.xlWorkbookDefault)

            xlSh = DirectCast(xlApp.ActiveWorkbook.ActiveSheet, Excel.Worksheet)
            xlSh.Rows(1).font.bold = True
            xlSh.Rows(1).interior.color = 17
            xlSh.Name = "RAW DATA"

            Dim i, j As Integer
            i = 0

            For Each col As DataColumn In sourceTable.Columns
                xlSh.Cells(1, i + 1) = col.ColumnName
                i += 1
            Next col

            j = 1
            For Each dr As DataRow In sourceTable.Rows
                i = 0
                For Each col As DataColumn In sourceTable.Columns
                    xlSh.Cells(j + 1, i + 1) = dr.Item(col, DataRowVersion.Current)
                    i += 1
                Next
                j += 1
            Next

            xlSh.Columns.AutoFit()

            If String.IsNullOrEmpty(filePathName) Then
                xlApp.Caption = "Untitled"
            Else
                xlApp.Caption = filePathName
            End If

            xlWB.Save()
            TableToExcel = True
        Catch ex As System.Runtime.InteropServices.COMException
            If ex.ErrorCode = -2147221164 Then
                a_sErrorMessage = "Error in export 1: Please install Microsoft Office (Excel) to use the Export to Excel feature."
            ElseIf ex.ErrorCode = -2146827284 Then
                a_sErrorMessage = "Error in export 1: Excel allows only 65,536 maximum rows in a sheet."
            Else
                a_sErrorMessage = (("Error in export 1: " & ex.Message) + Environment.NewLine & " Error: ") + ex.ErrorCode
            End If
        Catch ex As Exception
            a_sErrorMessage = "Error in export 1: " & ex.Message
        Finally
            If xlSh IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSh)
            End If
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

    Private Sub SortByParent(sPath As String)
        Try
            Dim xlRawSh As Excel.Worksheet = xlWB.Sheets("RAW DATA")
            Dim xlHierarchySh As Excel.Worksheet = CType(xlWB.Worksheets.Add(), Excel.Worksheet)

            Dim tableX, tableY, zeroX, zeroY, lastColumn, k As Integer
            Dim foundParent, isFirstRow, isTheEnd As Boolean
            Dim sQueryNumber As String = ""

            xlHierarchySh.Name = "Project Hierarchy"
            xlHierarchySh.Cells(1, 1) = "Project Number"
            xlHierarchySh.Cells(1, 2) = UserForm.tb_Project.Text
            xlHierarchySh.Range(xlHierarchySh.Cells(1, 2), xlHierarchySh.Cells(1, 3)).Merge()
            xlHierarchySh.Cells(2, 1) = "Project Folder Path"
            xlHierarchySh.Cells(2, 2) = sPath
            xlHierarchySh.Range(xlHierarchySh.Cells(2, 2), xlHierarchySh.Cells(2, 3)).Merge()
            xlHierarchySh.Rows(1).font.bold = True
            xlHierarchySh.Rows(2).font.bold = True

            xlHierarchySh.Cells(4, 1) = "DWG QTY"
            xlHierarchySh.Cells(4, 2) = "1"
            xlHierarchySh.Cells(4, 3) = "DESCRIPTION"
            xlHierarchySh.Cells(4, 4) = "REV"
            xlHierarchySh.Rows(4).font.bold = True
            xlHierarchySh.Rows(4).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Plum)
            'xlHierarchySh.Cells(3, 4) = "DATE RELEASE"

            zeroX = 2
            zeroY = 5
            tableX = 0
            tableY = 0

            ' find top level or orphaned items and dump them into hierarchical view
            k = 0
            For i As Integer = 2 To xlRawSh.UsedRange.Rows.Count + 100
                foundParent = False
                sQueryNumber = xlRawSh.Cells(i, 2).value

                If sQueryNumber = "" Then
                    Continue For
                End If

                For j As Integer = 2 To xlRawSh.UsedRange.Rows.Count
                    If i = j Then
                        k = j
                        Continue For
                    ElseIf sQueryNumber = xlRawSh.Cells(j, 1).value Then
                        foundParent = True
                    End If
                Next
                If foundParent = False Then
                    xlHierarchySh.Cells(zeroY + tableY, zeroX + tableX).value = sQueryNumber
                    xlHierarchySh.Cells(zeroY + tableY, zeroX + tableX).offset(0, -1).value = xlRawSh.Cells(k, 6).value
                    xlHierarchySh.Cells(zeroY + tableY, zeroX + tableX).offset(0, 1).value = xlRawSh.Cells(k, 4).value
                    tableY += 1
                End If
            Next

            ' add all children, working backwards from known values in hierarchical view
            lastColumn = 2
            Do
                isFirstRow = True
                isTheEnd = True
                For i As Integer = 5 To xlHierarchySh.UsedRange.Rows.Count + 100
                    sQueryNumber = CStr(xlHierarchySh.Cells(i, lastColumn).value)
                    If sQueryNumber = "" Then
                        Continue For
                    Else
                        'Debug.Print(sQueryNumber & "(" & lastColumn & ")")
                    End If
                    For j As Integer = 2 To xlRawSh.UsedRange.Rows.Count + 100
                        If CStr(xlRawSh.Cells(j, 2).value) = sQueryNumber And CStr(xlRawSh.Cells(j, 1).value) <> sQueryNumber Then
                            'Debug.Print("(" & i & "," & lastColumn & ")" & ">>>" & sQueryNumber & "|||" & CStr(xlRawSh.Cells(j, 2).value) & "||" & CStr(xlRawSh.Cells(j, 1).value))
                            If isFirstRow Then
                                xlHierarchySh.Cells(i, lastColumn).entirecolumn.offset(0, 1).insert()
                                isFirstRow = False
                                isTheEnd = False
                            End If
                            If CStr(xlRawSh.Cells(j, 1).value)(0) = "4"c Then
                                xlHierarchySh.Cells(i, lastColumn).Offset(1).EntireRow.Insert()
                                xlHierarchySh.Cells(i, lastColumn).offset(1, 1).value = CStr(xlRawSh.Cells(j, 1).value)
                                xlHierarchySh.Cells(i, lastColumn).offset(1, 2).value = CStr(xlRawSh.Cells(j, 4).value)
                                xlHierarchySh.Cells(i, 1).offset(1, 0).value = CStr(xlRawSh.Cells(j, 6).value)
                            End If
                        Else
                            'Debug.Print("(" & i & "," & lastColumn & ")" & "<<<" & sQueryNumber & "|||" & CStr(xlRawSh.Cells(j, 2).value) & "||" & CStr(xlRawSh.Cells(j, 1).value))
                        End If
                    Next
                Next
                If isTheEnd Then
                    Exit Do
                End If
                lastColumn += 1
            Loop

            xlHierarchySh.Columns.AutoFit()

            If xlRawSh IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRawSh)
            End If

            If xlHierarchySh IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlHierarchySh)
            End If

        Catch ex As System.Runtime.InteropServices.COMException
            If ex.ErrorCode = -2147221164 Then
                a_sErrorMessage = "Error in export 2: Please install Microsoft Office (Excel) to use the Export to Excel feature."
            ElseIf ex.ErrorCode = -2146827284 Then
                a_sErrorMessage = "Error in export 2: Excel allows only 65,536 maximum rows in a sheet."
            Else
                a_sErrorMessage = (("Error in export 2: " & ex.Message) + Environment.NewLine & " Error: ") + ex.ErrorCode
            End If
        Catch ex As Exception
            a_sErrorMessage = "Error in export 2: " & ex.Message
        Finally
            'xlWB.Save()
            If a_sErrorMessage <> "" Then
                Debug.Print(a_sErrorMessage)
            End If
        End Try

    End Sub

    Private Sub SortByType(sPath As String)
        Try
            Dim xlRawSh As Excel.Worksheet = xlWB.Sheets("RAW DATA")
            Dim xlTypeUnknownSh As Excel.Worksheet = CType(xlWB.Worksheets.Add(), Excel.Worksheet)
            Dim xlTypeAssemblySh As Excel.Worksheet = CType(xlWB.Worksheets.Add(), Excel.Worksheet)
            Dim xlTypeWeldmentSh As Excel.Worksheet = CType(xlWB.Worksheets.Add(), Excel.Worksheet)
            Dim xlTypeFabricationSh As Excel.Worksheet = CType(xlWB.Worksheets.Add(), Excel.Worksheet)

            Dim lastRowA, lastRowW, lastRowF, lastRowU As Integer
            Dim sQueryNumber, sQueryDescription, sQueryMaterial, sQueryType, sQueryParent As String
            Dim iQueryQty, iQueryLG, iQueryWD As Double
            Dim iFoundIndex As Integer
            Dim bFoundItem As Boolean

            xlTypeUnknownSh.Name = "Unknown Items"
            xlTypeUnknownSh.Cells(1, 1) = "Project Number"
            xlTypeUnknownSh.Cells(1, 2) = UserForm.tb_Project.Text
            xlTypeUnknownSh.Cells(2, 1) = "Project Folder Path"
            xlTypeUnknownSh.Cells(2, 2) = sPath
            xlTypeUnknownSh.Cells(4, 1) = "NUMBER"
            xlTypeUnknownSh.Cells(4, 2) = "DESCRIPTION"
            xlTypeUnknownSh.Cells(4, 3) = "MATERIAL"
            xlTypeUnknownSh.Cells(4, 4) = "QTY"
            xlTypeUnknownSh.Cells(4, 5) = "PARENTS"
            xlTypeUnknownSh.Cells(4, 6) = "TYPE"
            xlTypeUnknownSh.Rows(1).font.bold = True
            xlTypeUnknownSh.Rows(2).font.bold = True
            xlTypeUnknownSh.Rows(4).font.bold = True
            xlTypeUnknownSh.Rows(4).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Plum)
            xlTypeUnknownSh.Range(xlTypeUnknownSh.Cells(1, 2), xlTypeUnknownSh.Cells(1, 3)).Merge()
            xlTypeUnknownSh.Range(xlTypeUnknownSh.Cells(2, 2), xlTypeUnknownSh.Cells(2, 3)).Merge()
            lastRowU = 5

            xlTypeAssemblySh.Name = "Assembly Totals"
            xlTypeAssemblySh.Cells(1, 1) = "Project Number"
            xlTypeAssemblySh.Cells(1, 2) = UserForm.tb_Project.Text
            xlTypeAssemblySh.Cells(2, 1) = "Project Folder Path"
            xlTypeAssemblySh.Cells(2, 2) = sPath
            xlTypeAssemblySh.Cells(4, 1) = "NUMBER"
            xlTypeAssemblySh.Cells(4, 2) = "DESCRIPTION"
            xlTypeAssemblySh.Cells(4, 3) = "MATERIAL"
            xlTypeAssemblySh.Cells(4, 4) = "QTY"
            xlTypeAssemblySh.Cells(4, 5) = "PARENTS"
            xlTypeAssemblySh.Cells(4, 6) = "TYPE"
            xlTypeAssemblySh.Rows(1).font.bold = True
            xlTypeAssemblySh.Rows(2).font.bold = True
            xlTypeAssemblySh.Rows(4).font.bold = True
            xlTypeAssemblySh.Rows(4).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Plum)
            xlTypeAssemblySh.Range(xlTypeAssemblySh.Cells(1, 2), xlTypeAssemblySh.Cells(1, 3)).Merge()
            xlTypeAssemblySh.Range(xlTypeAssemblySh.Cells(2, 2), xlTypeAssemblySh.Cells(2, 3)).Merge()
            lastRowA = 5

            xlTypeWeldmentSh.Name = "Weldment Totals"
            xlTypeWeldmentSh.Cells(1, 1) = "Project Number"
            xlTypeWeldmentSh.Cells(1, 2) = UserForm.tb_Project.Text
            xlTypeWeldmentSh.Cells(2, 1) = "Project Folder Path"
            xlTypeWeldmentSh.Cells(2, 2) = sPath
            xlTypeWeldmentSh.Cells(4, 1) = "NUMBER"
            xlTypeWeldmentSh.Cells(4, 2) = "DESCRIPTION"
            xlTypeWeldmentSh.Cells(4, 3) = "MATERIAL"
            xlTypeWeldmentSh.Cells(4, 4) = "QTY"
            xlTypeWeldmentSh.Cells(4, 5) = "PARENTS"
            xlTypeWeldmentSh.Cells(4, 6) = "TYPE"
            xlTypeWeldmentSh.Rows(1).font.bold = True
            xlTypeWeldmentSh.Rows(2).font.bold = True
            xlTypeWeldmentSh.Rows(4).font.bold = True
            xlTypeWeldmentSh.Rows(4).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Plum)
            xlTypeWeldmentSh.Range(xlTypeWeldmentSh.Cells(1, 2), xlTypeWeldmentSh.Cells(1, 3)).Merge()
            xlTypeWeldmentSh.Range(xlTypeWeldmentSh.Cells(2, 2), xlTypeWeldmentSh.Cells(2, 3)).Merge()
            lastRowW = 5

            xlTypeFabricationSh.Name = "Fabrication Totals"
            xlTypeFabricationSh.Cells(1, 1) = "Project Number"
            xlTypeFabricationSh.Cells(1, 2) = UserForm.tb_Project.Text
            xlTypeFabricationSh.Cells(2, 1) = "Project Folder Path"
            xlTypeFabricationSh.Cells(2, 2) = sPath
            xlTypeFabricationSh.Cells(4, 1) = "NUMBER"
            xlTypeFabricationSh.Cells(4, 2) = "DESCRIPTION"
            xlTypeFabricationSh.Cells(4, 3) = "MATERIAL"
            xlTypeFabricationSh.Cells(4, 4) = "QTY"
            xlTypeFabricationSh.Cells(4, 5) = "PARENTS"
            xlTypeFabricationSh.Cells(4, 6) = "TYPE"
            xlTypeFabricationSh.Rows(4).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Plum)
            xlTypeFabricationSh.Rows(1).font.bold = True
            xlTypeFabricationSh.Rows(2).font.bold = True
            xlTypeFabricationSh.Rows(4).font.bold = True
            xlTypeFabricationSh.Range(xlTypeFabricationSh.Cells(1, 2), xlTypeFabricationSh.Cells(1, 3)).Merge()
            xlTypeFabricationSh.Range(xlTypeFabricationSh.Cells(2, 2), xlTypeFabricationSh.Cells(2, 3)).Merge()
            lastRowF = 5


            For i As Integer = 2 To xlRawSh.usedrange.rows.count + 100
                sQueryNumber = CStr(xlRawSh.cells(i, 1).value)
                sQueryDescription = CStr(xlRawSh.cells(i, 4).value)
                sQueryMaterial = CStr(xlRawSh.cells(i, 5).value)
                sQueryType = CStr(xlRawSh.cells(i, 3).value)
                sQueryParent = CStr(xlRawSh.cells(i, 2).value)

                iQueryQty = Frac2Num(CStr(xlRawSh.cells(i, 6).value))
                iQueryLG = Frac2Num(CStr(xlRawSh.cells(i, 8).value))
                iQueryWD = Frac2Num(CStr(xlRawSh.cells(i, 9).value))



                If sQueryNumber = "" Then
                    Continue For
                End If

                If sQueryNumber(0) = "4"c Then
                    Continue For
                End If

                Select Case sQueryType
                    Case "Weldment"
                        bFoundItem = False
                        For iFoundIndex = 5 To xlTypeWeldmentSh.UsedRange.Rows.Count + 100
                            If CStr(xlTypeWeldmentSh.Cells(iFoundIndex, 1).value) = sQueryNumber And _
                                CStr(xlTypeWeldmentSh.Cells(iFoundIndex, 2).value) = sQueryDescription And _
                                CStr(xlTypeWeldmentSh.Cells(iFoundIndex, 3).value) = sQueryMaterial Then
                                bFoundItem = True
                                Exit For
                            End If
                        Next

                        If bFoundItem Then
                            xlTypeWeldmentSh.Cells(iFoundIndex, 4) = CType(xlTypeWeldmentSh.Cells(iFoundIndex, 4).value, Double) + (Frac2Num(iQueryLG) * iQueryQty)
                            If InStr(CStr(xlTypeWeldmentSh.Cells(iFoundIndex, 5).value), sQueryParent) = 0 Then
                                xlTypeWeldmentSh.Cells(iFoundIndex, 5) = CStr(xlTypeWeldmentSh.Cells(iFoundIndex, 5).value) & ", " & sQueryParent
                            End If
                        Else
                            xlTypeWeldmentSh.Cells(lastRowW, 4) = Frac2Num(iQueryLG) * iQueryQty
                            xlTypeWeldmentSh.Cells(lastRowW, 1) = sQueryNumber
                            xlTypeWeldmentSh.Cells(lastRowW, 2) = sQueryDescription
                            xlTypeWeldmentSh.Cells(lastRowW, 3) = sQueryMaterial
                            xlTypeWeldmentSh.Cells(lastRowW, 5) = CStr(sQueryParent)
                            xlTypeWeldmentSh.Cells(lastRowW, 6) = sQueryType
                            lastRowW += 1
                        End If
                    Case "Assembly"
                        bFoundItem = False
                        For iFoundIndex = 5 To xlTypeAssemblySh.UsedRange.Rows.Count + 100
                            If CStr(xlTypeAssemblySh.Cells(iFoundIndex, 1).value) = sQueryNumber And _
                                CStr(xlTypeAssemblySh.Cells(iFoundIndex, 2).value) = sQueryDescription And _
                                CStr(xlTypeAssemblySh.Cells(iFoundIndex, 3).value) = sQueryMaterial Then
                                bFoundItem = True
                                Exit For
                            End If
                        Next

                        If bFoundItem Then
                            xlTypeAssemblySh.Cells(iFoundIndex, 4) = CType(xlTypeAssemblySh.Cells(iFoundIndex, 4).value, Double) + (Frac2Num(iQueryLG) * iQueryQty)
                            If InStr(CStr(xlTypeAssemblySh.Cells(iFoundIndex, 5).value), sQueryParent) = 0 Then
                                xlTypeAssemblySh.Cells(iFoundIndex, 5) = CStr(xlTypeAssemblySh.Cells(iFoundIndex, 5).value) & ", " & sQueryParent
                            End If
                        Else
                            xlTypeAssemblySh.Cells(lastRowA, 4) = Frac2Num(iQueryLG) * iQueryQty
                            xlTypeAssemblySh.Cells(lastRowA, 1) = sQueryNumber
                            xlTypeAssemblySh.Cells(lastRowA, 2) = sQueryDescription
                            xlTypeAssemblySh.Cells(lastRowA, 3) = sQueryMaterial
                            xlTypeAssemblySh.Cells(lastRowA, 5) = CStr(sQueryParent)
                            xlTypeAssemblySh.Cells(lastRowA, 6) = sQueryType
                            lastRowA += 1
                        End If
                    Case "Single Part"
                        bFoundItem = False
                        For iFoundIndex = 5 To xlTypeFabricationSh.UsedRange.Rows.Count + 100
                            If CStr(xlTypeFabricationSh.Cells(iFoundIndex, 1).value) = sQueryNumber And _
                                CStr(xlTypeFabricationSh.Cells(iFoundIndex, 2).value) = sQueryDescription And _
                                CStr(xlTypeFabricationSh.Cells(iFoundIndex, 3).value) = sQueryMaterial Then
                                bFoundItem = True
                                Exit For
                            End If
                        Next

                        If bFoundItem Then
                            xlTypeFabricationSh.Cells(iFoundIndex, 4) = CType(xlTypeFabricationSh.Cells(iFoundIndex, 4).value, Double) + (Frac2Num(iQueryLG) * iQueryQty)
                            If InStr(CStr(xlTypeFabricationSh.Cells(iFoundIndex, 5).value), sQueryParent) = 0 Then
                                xlTypeFabricationSh.Cells(iFoundIndex, 5) = CStr(xlTypeFabricationSh.Cells(iFoundIndex, 5).value) & ", " & sQueryParent
                            End If
                        Else
                            xlTypeFabricationSh.Cells(lastRowF, 4) = Frac2Num(iQueryLG) * iQueryQty
                            xlTypeFabricationSh.Cells(lastRowF, 1) = sQueryNumber
                            xlTypeFabricationSh.Cells(lastRowF, 2) = sQueryDescription
                            xlTypeFabricationSh.Cells(lastRowF, 3) = sQueryMaterial
                            xlTypeFabricationSh.Cells(lastRowF, 5) = CStr(sQueryParent)
                            xlTypeFabricationSh.Cells(lastRowF, 6) = sQueryType
                            lastRowF += 1
                        End If
                    Case Else
                            bFoundItem = False
                            For iFoundIndex = 5 To xlTypeUnknownSh.UsedRange.Rows.Count + 100
                                If CStr(xlTypeUnknownSh.Cells(iFoundIndex, 1).value) = sQueryNumber And _
                                    CStr(xlTypeUnknownSh.Cells(iFoundIndex, 2).value) = sQueryDescription And _
                                    CStr(xlTypeUnknownSh.Cells(iFoundIndex, 3).value) = sQueryMaterial Then
                                    bFoundItem = True
                                    Exit For
                                End If
                            Next

                            If bFoundItem Then
                                xlTypeUnknownSh.Cells(iFoundIndex, 4) = CType(xlTypeUnknownSh.Cells(iFoundIndex, 4).value, Double) + (Frac2Num(iQueryLG) * iQueryQty)
                                xlTypeUnknownSh.Cells(iFoundIndex, 5) = CStr(xlTypeUnknownSh.Cells(iFoundIndex, 5).value) & "," & sQueryParent
                            Else
                                xlTypeUnknownSh.Cells(lastRowU, 4) = Frac2Num(iQueryLG) * iQueryQty
                                xlTypeUnknownSh.Cells(lastRowU, 1) = sQueryNumber
                                xlTypeUnknownSh.Cells(lastRowU, 2) = sQueryDescription
                                xlTypeUnknownSh.Cells(lastRowU, 3) = sQueryMaterial
                                xlTypeUnknownSh.Cells(lastRowU, 5) = CStr(sQueryParent)
                                xlTypeUnknownSh.Cells(lastRowU, 6) = sQueryType
                                lastRowU += 1
                            End If
                End Select
            Next

            xlTypeAssemblySh.Columns.AutoFit()
            xlTypeWeldmentSh.Columns.AutoFit()
            xlTypeFabricationSh.Columns.AutoFit()
            xlTypeUnknownSh.Columns.AutoFit()

        Catch ex As System.Runtime.InteropServices.COMException
            If ex.ErrorCode = -2147221164 Then
                a_sErrorMessage = "Error in export 2: Please install Microsoft Office (Excel) to use the Export to Excel feature."
            ElseIf ex.ErrorCode = -2146827284 Then
                a_sErrorMessage = "Error in export 2: Excel allows only 65,536 maximum rows in a sheet."
            Else
                a_sErrorMessage = (("Error in export 2: " & ex.Message) + Environment.NewLine & " Error: ") + ex.ErrorCode
            End If
        Catch ex As Exception
            a_sErrorMessage = "Error in export 2: " & ex.Message
        Finally
            'xlWB.Save()
            If a_sErrorMessage <> "" Then
                Debug.Print(a_sErrorMessage)
            End If
        End Try
    End Sub

End Module
