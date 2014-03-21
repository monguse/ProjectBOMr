Module GenBOM
    Public Const BaseFilePath = "C:\Projects\"
    Public fPath As String
    Public fExists As Boolean

    ' Gets the user's history with the specified project number
    ' The history files are stored in the user's AppData
    Private Sub GetHistory(iProjNum As Long)

    End Sub

    ' After generating the combined BOM from the project folder, store BOM details for future deltas
    Private Sub SetHistory()

    End Sub

    Public Sub Cout(iStr As String)
        UserForm.rtb_Console.AppendText(iStr + vbNewLine)
    End Sub

    Public Function ProcessFolder() As Collection
        Dim fso, f
        Dim cDrawingBOMs As New Collection

        fso = CreateObject("Scripting.FileSystemObject")
        For Each f In fso.GetFolder(fPath)
            pos()
        Next

        For 
    End Function

    Public Function ReadCSV(sCSVFilePath As String) As cDrawingBOM
        ReadCSV = New cDrawingBOM
    End Function

    Public Function FindPath(sPath As String) As Boolean
        Dim fso, f
        Dim pos As Integer

        FindPath = False
        fExists = False

        fso = CreateObject("Scripting.FileSystemObject")
        For Each f In fso.GetFolder(BaseFilePath).subfolders
            pos = InStr(f.name, sPath)
            If pos > 0 Then
                FindPath = True
                fExists = True
                fPath = BaseFilePath + f.name + "\"
                Exit For
            End If
        Next
        fso = Nothing
    End Function


End Module
