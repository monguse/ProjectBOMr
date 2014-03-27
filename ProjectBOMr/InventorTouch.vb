Module InventorTouch
    Private ivApp As New Inventor.ApprenticeServerComponent
    Private Const BaseVaultPath = "C:\VaultWorkspace"

    Public Function ProcessDrawing(ByVal dwgNum As String) As System.Data.DataTable
        Dim ivDoc As Inventor.ApprenticeServerDrawingDocument
        Dim foundDrawingPath As String
        Dim projBOM As New System.Data.DataTable

        projBOM.Columns.Add(columnName:="DESCRIPTION", type:=GetType(String))
        projBOM.Columns.Add(columnName:="MATERIAL", type:=GetType(String))
        projBOM.Columns.Add(columnName:="NUMBER", type:=GetType(String))
        projBOM.Columns.Add(columnName:="QTY", type:=GetType(String))
        projBOM.Columns.Add(columnName:="UNITQTY", type:=GetType(String))
        projBOM.Columns.Add(columnName:="LG", type:=GetType(String))
        projBOM.Columns.Add(columnName:="WD", type:=GetType(String))
        projBOM.Columns.Add(columnName:="PARENT", type:=GetType(String))
        projBOM.Columns.Add(columnName:="PROCESSCODE", type:=GetType(String))

        ivApp = New Inventor.ApprenticeServerComponent
        ProcessDrawing = Nothing

        foundDrawingPath = RecursiveFind(BaseVaultPath, dwgNum)
        If foundDrawingPath <> "" Then
            ivDoc = ivApp.Open(foundDrawingPath)
            GetDrawingBOM(projBOM, ivDoc)
            ProcessDrawing = projBOM
        End If

    End Function

    Private Function RecursiveFind(ByVal folderPath As String, ByVal fileName As String) As String
        Dim filePath As String

        filePath = folderPath + "\" + fileName + ".idw"
        RecursiveFind = ""

        If My.Computer.FileSystem.FileExists(filePath) Then
            RecursiveFind = filePath
        Else
            For Each foundDir As String In My.Computer.FileSystem.GetDirectories(folderPath)
                RecursiveFind = RecursiveFind + RecursiveFind(foundDir, fileName)
            Next
        End If
    End Function

    Private Sub GetDrawingBOM(ByRef dt As System.Data.DataTable, ByRef sd As Inventor.ApprenticeServerDrawingDocument)
        Dim dtDataRow As System.Data.DataRow

        dtDataRow = dt.NewRow()

        For Each dSheet As Inventor.Sheet In sd.Sheets
            For Each pList As Inventor.PartsList In dSheet.PartsLists
                For Each pRow As Inventor.PartsListRow In pList.PartsListRows
                    If pRow.Visible Then
                        GetCells(dtDataRow, pRow, pList.PartsListColumns)
                    End If
                Next
            Next
        Next
    End Sub

    Private Sub GetCells(ByRef dt As System.Data.DataRow, ByRef sRow As Inventor.PartsListRow, ByRef sCols As Inventor.PartsListColumns)
        Dim propSetID As String
        Dim propID As Long

        propSetID = ""
        propID = -1

        'If sCol.PropertyType = Inventor.PropertyTypeEnum.kFileProperty Then
        '    sCol.GetFilePropertyId(propSetID, propID)
        'End If

        'Select Case sCol.PropertyType
        '    Case Inventor.PropertyTypeEnum.kFileProperty
        '    Case Inventor.PropertyTypeEnum.kMaterialPartsListProperty
        '    Case Inventor.PropertyTypeEnum.kCustomProperty
        '    Case Inventor.PropertyTypeEnum.kQuantityPartsListProperty
        '    Case Inventor.PropertyTypeEnum.kQuantityPartsListProperty
        '    Case Else
        '        Debug.Print("Unknown Property: ")
        'End Select



    End Sub
End Module
