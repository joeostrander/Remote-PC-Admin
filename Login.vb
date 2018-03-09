Public Class Login

    Private Sub ButtonCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCancel.Click
        Me.Close()
    End Sub

    Private Sub ButtonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSave.Click
        SaveCredentials()
    End Sub

    Private Sub SaveCredentials()
        If TextBoxUsername.Text = "" Then
            MsgBox("Please enter a username.", MsgBoxStyle.Exclamation, "Invalid Entry")
            TextBoxUsername.Focus()
            Exit Sub
        End If

        If TextBoxPassword.Text = "" Then
            MsgBox("Please enter a password.", MsgBoxStyle.Exclamation, "Invalid Entry")
            TextBoxPassword.Focus()
            Exit Sub
        End If

        FormRemotePCAdmin.strUsername = TextBoxUsername.Text
        FormRemotePCAdmin.strPassword = TextBoxPassword.Text

        Me.Close()
    End Sub

    Private Sub Login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBoxUsername.Text = FormRemotePCAdmin.strUsername
        TextBoxPassword.Text = FormRemotePCAdmin.strPassword
    End Sub

    Private Sub TextBoxPassword_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBoxPassword.KeyPress
        If Asc(e.KeyChar) = Keys.Enter Then SaveCredentials()
    End Sub


    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
        MsgBox("You can enter your name like:" & vbCrLf & vbCrLf & _
        "youruser@yourdomain.com" & vbCrLf & "or" & vbCrLf & "YOURDOMAIN\youruser", MsgBoxStyle.Information, "Setting the Username")
    End Sub
End Class