<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UserForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tb_Project = New System.Windows.Forms.TextBox()
        Me.b_OK = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.rtb_Console = New System.Windows.Forms.RichTextBox()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Project Number"
        '
        'tb_Project
        '
        Me.tb_Project.Location = New System.Drawing.Point(15, 25)
        Me.tb_Project.Name = "tb_Project"
        Me.tb_Project.Size = New System.Drawing.Size(296, 20)
        Me.tb_Project.TabIndex = 1
        '
        'b_OK
        '
        Me.b_OK.Location = New System.Drawing.Point(317, 23)
        Me.b_OK.Name = "b_OK"
        Me.b_OK.Size = New System.Drawing.Size(75, 23)
        Me.b_OK.TabIndex = 2
        Me.b_OK.Text = "Get BOM"
        Me.b_OK.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.rtb_Console)
        Me.GroupBox1.Location = New System.Drawing.Point(15, 51)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(377, 235)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Messages"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(6, 199)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "Clear"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(258, 199)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(113, 23)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Copy to Clipboard"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'rtb_Console
        '
        Me.rtb_Console.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.rtb_Console.Enabled = False
        Me.rtb_Console.Location = New System.Drawing.Point(6, 19)
        Me.rtb_Console.Name = "rtb_Console"
        Me.rtb_Console.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.rtb_Console.Size = New System.Drawing.Size(365, 174)
        Me.rtb_Console.TabIndex = 0
        Me.rtb_Console.Text = ""
        '
        'UserForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(404, 294)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.b_OK)
        Me.Controls.Add(Me.tb_Project)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "UserForm"
        Me.Text = "ProjectBOMr"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tb_Project As System.Windows.Forms.TextBox
    Friend WithEvents b_OK As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rtb_Console As System.Windows.Forms.RichTextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button

End Class
