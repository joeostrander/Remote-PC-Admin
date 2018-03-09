<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FindText
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
        Me.ButtonFindNext = New System.Windows.Forms.Button
        Me.ButtonClose = New System.Windows.Forms.Button
        Me.ButtonFindPrevious = New System.Windows.Forms.Button
        Me.ComboBoxSearchString = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.CheckBoxMatchCase = New System.Windows.Forms.CheckBox
        Me.CheckBoxRegEx = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'ButtonFindNext
        '
        Me.ButtonFindNext.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonFindNext.Location = New System.Drawing.Point(336, 12)
        Me.ButtonFindNext.Name = "ButtonFindNext"
        Me.ButtonFindNext.Size = New System.Drawing.Size(84, 23)
        Me.ButtonFindNext.TabIndex = 2
        Me.ButtonFindNext.Text = "Find &Next"
        '
        'ButtonClose
        '
        Me.ButtonClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ButtonClose.Location = New System.Drawing.Point(336, 70)
        Me.ButtonClose.Name = "ButtonClose"
        Me.ButtonClose.Size = New System.Drawing.Size(84, 23)
        Me.ButtonClose.TabIndex = 4
        Me.ButtonClose.Text = "&Close"
        '
        'ButtonFindPrevious
        '
        Me.ButtonFindPrevious.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonFindPrevious.Location = New System.Drawing.Point(336, 41)
        Me.ButtonFindPrevious.Name = "ButtonFindPrevious"
        Me.ButtonFindPrevious.Size = New System.Drawing.Size(84, 23)
        Me.ButtonFindPrevious.TabIndex = 3
        Me.ButtonFindPrevious.Text = "Find &Previous"
        '
        'ComboBoxSearchString
        '
        Me.ComboBoxSearchString.FormattingEnabled = True
        Me.ComboBoxSearchString.Location = New System.Drawing.Point(12, 28)
        Me.ComboBoxSearchString.Name = "ComboBoxSearchString"
        Me.ComboBoxSearchString.Size = New System.Drawing.Size(302, 21)
        Me.ComboBoxSearchString.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(74, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "&Search String:"
        '
        'CheckBoxMatchCase
        '
        Me.CheckBoxMatchCase.AutoSize = True
        Me.CheckBoxMatchCase.Location = New System.Drawing.Point(15, 70)
        Me.CheckBoxMatchCase.Name = "CheckBoxMatchCase"
        Me.CheckBoxMatchCase.Size = New System.Drawing.Size(82, 17)
        Me.CheckBoxMatchCase.TabIndex = 5
        Me.CheckBoxMatchCase.Text = "&Match case"
        Me.CheckBoxMatchCase.UseVisualStyleBackColor = True
        '
        'CheckBoxRegEx
        '
        Me.CheckBoxRegEx.AutoSize = True
        Me.CheckBoxRegEx.Location = New System.Drawing.Point(115, 70)
        Me.CheckBoxRegEx.Name = "CheckBoxRegEx"
        Me.CheckBoxRegEx.Size = New System.Drawing.Size(117, 17)
        Me.CheckBoxRegEx.TabIndex = 6
        Me.CheckBoxRegEx.Text = "&Regular Expression"
        Me.CheckBoxRegEx.UseVisualStyleBackColor = True
        '
        'FindText
        '
        Me.AcceptButton = Me.ButtonFindNext
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.ButtonClose
        Me.ClientSize = New System.Drawing.Size(435, 119)
        Me.Controls.Add(Me.CheckBoxRegEx)
        Me.Controls.Add(Me.CheckBoxMatchCase)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ComboBoxSearchString)
        Me.Controls.Add(Me.ButtonClose)
        Me.Controls.Add(Me.ButtonFindPrevious)
        Me.Controls.Add(Me.ButtonFindNext)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FindText"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Find Text"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ButtonFindNext As System.Windows.Forms.Button
    Friend WithEvents ButtonClose As System.Windows.Forms.Button
    Friend WithEvents ButtonFindPrevious As System.Windows.Forms.Button
    Friend WithEvents ComboBoxSearchString As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CheckBoxMatchCase As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxRegEx As System.Windows.Forms.CheckBox

End Class
