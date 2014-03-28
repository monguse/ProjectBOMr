Public Class UserForm

    Private isThinking As Boolean

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        isThinking = False
        tb_Project.Text = ""
    End Sub

    Private Sub tb_Project_TextChanged(sender As Object, e As EventArgs) Handles tb_Project.TextChanged
        If Not isThinking Then
            'Cout("Thinking...")
            isThinking = True
        End If
    End Sub

    Private Sub tb_Project_Leave(sender As Object, e As System.EventArgs) Handles tb_Project.Leave
        If tb_Project.Text <> "" Then
            Dim sProjPath As String = FindPath(tb_Project.Text)

            Cout("Searching...")
            If sProjPath <> "" Then
                Cout("Project Located!")
                Cout("  " + sProjPath)
            Else
                Cout("Project " + tb_Project.Text + " not found!")
            End If
        End If

        isThinking = False
    End Sub

    Private Sub tb_Project_Leave(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tb_Project.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.SelectNextControl(DirectCast(sender, System.Windows.Forms.TextBox), True, True, False, True)
        End If
    End Sub

    Private Sub b_OK_Click(sender As Object, e As EventArgs) Handles b_OK.Click
        Dim sProjPath As String = FindPath(tb_Project.Text)
        If sProjPath = "" Then
            Exit Sub
        End If
        b_OK.Enabled = False
        Cout("Writing BOM...")

        If sProjPath <> "" Then
            Dim sBomPath = GenBOM.ProcessFolder(sProjPath, tb_Project.Text)
            If sBomPath <> "" Then
                Cout("Wrote BOM! ")
                Cout("  " & sBomPath)
            Else
                Cout("Failed to write BOM")
            End If
        End If
        b_OK.Enabled = True
    End Sub
End Class
