Public Class UserForm

    Private isThinking As Boolean

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        isThinking = False
        tb_Project.Text = ""
        GenBOM.fExists = False
        GenBOM.fPath = ""
    End Sub

    Private Sub tb_Project_TextChanged(sender As Object, e As EventArgs) Handles tb_Project.TextChanged
        If Not isThinking Then
            Cout("Thinking...")
            isThinking = True
        End If
    End Sub

    Private Sub tb_Project_Leave(sender As Object, e As System.EventArgs) Handles tb_Project.Leave
        ' Run project locating sub here
        If tb_Project.Text = "" Then
            Exit Sub
        End If

        Cout("Searching...")
        If FindPath(tb_Project.Text) Then
            Cout("Project Located!")
            Cout("  " + GenBOM.fPath)
        Else
            Cout("Project " + tb_Project.Text + " not found!")
        End If
        isThinking = False
    End Sub
End Class
