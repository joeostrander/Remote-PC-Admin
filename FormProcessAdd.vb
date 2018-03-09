Public Class FormProcessAdd
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents LabelWorkingDirectory As System.Windows.Forms.Label
    Friend WithEvents lblCommand As System.Windows.Forms.Label
    Friend WithEvents ButtonExecute As System.Windows.Forms.Button
    Friend WithEvents TextBoxWorkingDir As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxCommand As System.Windows.Forms.TextBox
    Friend WithEvents ButtonCancel As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormProcessAdd))
        Me.TextBoxWorkingDir = New System.Windows.Forms.TextBox
        Me.TextBoxCommand = New System.Windows.Forms.TextBox
        Me.LabelWorkingDirectory = New System.Windows.Forms.Label
        Me.lblCommand = New System.Windows.Forms.Label
        Me.ButtonExecute = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.ButtonCancel = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'TextBoxWorkingDir
        '
        Me.TextBoxWorkingDir.Location = New System.Drawing.Point(12, 105)
        Me.TextBoxWorkingDir.Name = "TextBoxWorkingDir"
        Me.TextBoxWorkingDir.Size = New System.Drawing.Size(368, 20)
        Me.TextBoxWorkingDir.TabIndex = 4
        '
        'TextBoxCommand
        '
        Me.TextBoxCommand.Location = New System.Drawing.Point(12, 33)
        Me.TextBoxCommand.Name = "TextBoxCommand"
        Me.TextBoxCommand.Size = New System.Drawing.Size(368, 20)
        Me.TextBoxCommand.TabIndex = 1
        '
        'LabelWorkingDirectory
        '
        Me.LabelWorkingDirectory.Location = New System.Drawing.Point(9, 86)
        Me.LabelWorkingDirectory.Name = "LabelWorkingDirectory"
        Me.LabelWorkingDirectory.Size = New System.Drawing.Size(148, 16)
        Me.LabelWorkingDirectory.TabIndex = 3
        Me.LabelWorkingDirectory.Text = "Working Directory (optional):"
        '
        'lblCommand
        '
        Me.lblCommand.Location = New System.Drawing.Point(9, 14)
        Me.lblCommand.Name = "lblCommand"
        Me.lblCommand.Size = New System.Drawing.Size(65, 16)
        Me.lblCommand.TabIndex = 0
        Me.lblCommand.Text = "Command:"
        '
        'ButtonExecute
        '
        Me.ButtonExecute.Location = New System.Drawing.Point(92, 153)
        Me.ButtonExecute.Name = "ButtonExecute"
        Me.ButtonExecute.Size = New System.Drawing.Size(75, 23)
        Me.ButtonExecute.TabIndex = 5
        Me.ButtonExecute.Text = "&Execute"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(12, 56)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(352, 16)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "Ex:  C:\WINDOWS\System32\wscript.exe ""c:\temp\test.vbs"""
        '
        'ButtonCancel
        '
        Me.ButtonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ButtonCancel.Location = New System.Drawing.Point(224, 153)
        Me.ButtonCancel.Name = "ButtonCancel"
        Me.ButtonCancel.Size = New System.Drawing.Size(75, 23)
        Me.ButtonCancel.TabIndex = 6
        Me.ButtonCancel.Text = "&Cancel"
        '
        'FormProcessAdd
        '
        Me.AcceptButton = Me.ButtonExecute
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.ButtonCancel
        Me.ClientSize = New System.Drawing.Size(392, 199)
        Me.ControlBox = False
        Me.Controls.Add(Me.ButtonCancel)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.TextBoxCommand)
        Me.Controls.Add(Me.TextBoxWorkingDir)
        Me.Controls.Add(Me.ButtonExecute)
        Me.Controls.Add(Me.lblCommand)
        Me.Controls.Add(Me.LabelWorkingDirectory)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FormProcessAdd"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Execute Process"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region


    Private strTitle As String = "Execute Process"



    Private Sub FormProcessAdd_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.Text = FormRemotePCAdmin.TextBoxComputer.Text & " - " & strTitle

    End Sub

    Private Sub ButtonExecute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExecute.Click

        If TextBoxCommand.Text = "" Then
            MsgBox("Please enter a command.", MsgBoxStyle.Exclamation)
            TextBoxCommand.Focus()
            Exit Sub
        End If

        FormRemotePCAdmin.strProcessCommand = TextBoxCommand.Text

        Dim strWorkDir As String = TextBoxWorkingDir.Text

        If strWorkDir = "" Then strWorkDir = "C:\Windows\System32"

        FormRemotePCAdmin.strProcessWorkingDir = strWorkDir

        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()


    End Sub
End Class
