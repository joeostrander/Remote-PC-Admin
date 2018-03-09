Imports System.Management

Public Class FormServiceAdd
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblDisplayName As System.Windows.Forms.Label
    Friend WithEvents lblCommand As System.Windows.Forms.Label
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents ButtonAddService As System.Windows.Forms.Button
    Friend WithEvents txtServiceName As System.Windows.Forms.TextBox
    Friend WithEvents txtServiceDisplayName As System.Windows.Forms.TextBox
    Friend WithEvents txtServiceCmd As System.Windows.Forms.TextBox
    Friend WithEvents ComboBoxServiceType As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ButtonCancel As System.Windows.Forms.Button
    Friend WithEvents CheckBoxInteractive As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LabelUserID As System.Windows.Forms.Label
    Friend WithEvents TextBoxUserID As System.Windows.Forms.TextBox
    Friend WithEvents RadioButtonSystemAccount As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonOtherAccount As System.Windows.Forms.RadioButton
    Friend WithEvents TextBoxPassword As System.Windows.Forms.TextBox
    Friend WithEvents LabelPassword As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Label5 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormServiceAdd))
        Me.txtServiceName = New System.Windows.Forms.TextBox
        Me.txtServiceDisplayName = New System.Windows.Forms.TextBox
        Me.txtServiceCmd = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblDisplayName = New System.Windows.Forms.Label
        Me.lblCommand = New System.Windows.Forms.Label
        Me.ComboBoxServiceType = New System.Windows.Forms.ComboBox
        Me.lblType = New System.Windows.Forms.Label
        Me.ButtonAddService = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.ButtonCancel = New System.Windows.Forms.Button
        Me.CheckBoxInteractive = New System.Windows.Forms.CheckBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.LabelUserID = New System.Windows.Forms.Label
        Me.TextBoxUserID = New System.Windows.Forms.TextBox
        Me.RadioButtonSystemAccount = New System.Windows.Forms.RadioButton
        Me.RadioButtonOtherAccount = New System.Windows.Forms.RadioButton
        Me.TextBoxPassword = New System.Windows.Forms.TextBox
        Me.LabelPassword = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SuspendLayout()
        '
        'txtServiceName
        '
        Me.txtServiceName.Location = New System.Drawing.Point(14, 31)
        Me.txtServiceName.Name = "txtServiceName"
        Me.txtServiceName.Size = New System.Drawing.Size(200, 20)
        Me.txtServiceName.TabIndex = 1
        '
        'txtServiceDisplayName
        '
        Me.txtServiceDisplayName.Location = New System.Drawing.Point(14, 84)
        Me.txtServiceDisplayName.Name = "txtServiceDisplayName"
        Me.txtServiceDisplayName.Size = New System.Drawing.Size(200, 20)
        Me.txtServiceDisplayName.TabIndex = 4
        '
        'txtServiceCmd
        '
        Me.txtServiceCmd.Location = New System.Drawing.Point(14, 139)
        Me.txtServiceCmd.Name = "txtServiceCmd"
        Me.txtServiceCmd.Size = New System.Drawing.Size(341, 20)
        Me.txtServiceCmd.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(11, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "&Name:"
        '
        'lblDisplayName
        '
        Me.lblDisplayName.Location = New System.Drawing.Point(11, 65)
        Me.lblDisplayName.Name = "lblDisplayName"
        Me.lblDisplayName.Size = New System.Drawing.Size(80, 16)
        Me.lblDisplayName.TabIndex = 3
        Me.lblDisplayName.Text = "&Display Name:"
        '
        'lblCommand
        '
        Me.lblCommand.Location = New System.Drawing.Point(11, 120)
        Me.lblCommand.Name = "lblCommand"
        Me.lblCommand.Size = New System.Drawing.Size(65, 16)
        Me.lblCommand.TabIndex = 6
        Me.lblCommand.Text = "&Command:"
        '
        'ComboBoxServiceType
        '
        Me.ComboBoxServiceType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxServiceType.Items.AddRange(New Object() {"Automatic", "Manual"})
        Me.ComboBoxServiceType.Location = New System.Drawing.Point(15, 192)
        Me.ComboBoxServiceType.Name = "ComboBoxServiceType"
        Me.ComboBoxServiceType.Size = New System.Drawing.Size(156, 21)
        Me.ComboBoxServiceType.TabIndex = 10
        '
        'lblType
        '
        Me.lblType.Location = New System.Drawing.Point(12, 173)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(56, 16)
        Me.lblType.TabIndex = 9
        Me.lblType.Text = "&Type:"
        '
        'ButtonAddService
        '
        Me.ButtonAddService.Location = New System.Drawing.Point(139, 376)
        Me.ButtonAddService.Name = "ButtonAddService"
        Me.ButtonAddService.Size = New System.Drawing.Size(75, 23)
        Me.ButtonAddService.TabIndex = 19
        Me.ButtonAddService.Text = "&Add Service"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(220, 34)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(144, 16)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Ex:  WUAUSERV"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(219, 87)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(136, 16)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Ex:  Automatic Updates"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(361, 142)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(157, 16)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Ex:  C:\My Folder\service.exe"
        '
        'ButtonCancel
        '
        Me.ButtonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ButtonCancel.Location = New System.Drawing.Point(305, 376)
        Me.ButtonCancel.Name = "ButtonCancel"
        Me.ButtonCancel.Size = New System.Drawing.Size(75, 23)
        Me.ButtonCancel.TabIndex = 20
        Me.ButtonCancel.Text = "&Cancel"
        '
        'CheckBoxInteractive
        '
        Me.CheckBoxInteractive.AutoSize = True
        Me.CheckBoxInteractive.Location = New System.Drawing.Point(152, 251)
        Me.CheckBoxInteractive.Name = "CheckBoxInteractive"
        Me.CheckBoxInteractive.Size = New System.Drawing.Size(76, 17)
        Me.CheckBoxInteractive.TabIndex = 13
        Me.CheckBoxInteractive.Text = "&Interactive"
        Me.CheckBoxInteractive.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(12, 232)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(79, 16)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "&Log on as:"
        '
        'LabelUserID
        '
        Me.LabelUserID.Enabled = False
        Me.LabelUserID.Location = New System.Drawing.Point(125, 277)
        Me.LabelUserID.Name = "LabelUserID"
        Me.LabelUserID.Size = New System.Drawing.Size(56, 16)
        Me.LabelUserID.TabIndex = 15
        Me.LabelUserID.Text = "User ID:"
        '
        'TextBoxUserID
        '
        Me.TextBoxUserID.Enabled = False
        Me.TextBoxUserID.Location = New System.Drawing.Point(187, 274)
        Me.TextBoxUserID.Name = "TextBoxUserID"
        Me.TextBoxUserID.Size = New System.Drawing.Size(193, 20)
        Me.TextBoxUserID.TabIndex = 16
        Me.TextBoxUserID.Text = "NT AUTHORITY\LocalService"
        '
        'RadioButtonSystemAccount
        '
        Me.RadioButtonSystemAccount.AutoSize = True
        Me.RadioButtonSystemAccount.Checked = True
        Me.RadioButtonSystemAccount.Location = New System.Drawing.Point(15, 251)
        Me.RadioButtonSystemAccount.Name = "RadioButtonSystemAccount"
        Me.RadioButtonSystemAccount.Size = New System.Drawing.Size(131, 17)
        Me.RadioButtonSystemAccount.TabIndex = 12
        Me.RadioButtonSystemAccount.TabStop = True
        Me.RadioButtonSystemAccount.Text = "Local System Account"
        Me.RadioButtonSystemAccount.UseVisualStyleBackColor = True
        '
        'RadioButtonOtherAccount
        '
        Me.RadioButtonOtherAccount.AutoSize = True
        Me.RadioButtonOtherAccount.Location = New System.Drawing.Point(15, 274)
        Me.RadioButtonOtherAccount.Name = "RadioButtonOtherAccount"
        Me.RadioButtonOtherAccount.Size = New System.Drawing.Size(90, 17)
        Me.RadioButtonOtherAccount.TabIndex = 14
        Me.RadioButtonOtherAccount.Text = "This account:"
        Me.RadioButtonOtherAccount.UseVisualStyleBackColor = True
        '
        'TextBoxPassword
        '
        Me.TextBoxPassword.Enabled = False
        Me.TextBoxPassword.Location = New System.Drawing.Point(187, 300)
        Me.TextBoxPassword.Name = "TextBoxPassword"
        Me.TextBoxPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TextBoxPassword.Size = New System.Drawing.Size(193, 20)
        Me.TextBoxPassword.TabIndex = 18
        '
        'LabelPassword
        '
        Me.LabelPassword.Enabled = False
        Me.LabelPassword.Location = New System.Drawing.Point(125, 303)
        Me.LabelPassword.Name = "LabelPassword"
        Me.LabelPassword.Size = New System.Drawing.Size(56, 16)
        Me.LabelPassword.TabIndex = 17
        Me.LabelPassword.Text = "Password:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Help
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.CornflowerBlue
        Me.Label6.Location = New System.Drawing.Point(386, 277)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(14, 13)
        Me.Label6.TabIndex = 21
        Me.Label6.Text = "?"
        Me.ToolTip1.SetToolTip(Me.Label6, "Make sure this account has been granted the permission ""Log on as a service""")
        '
        'FormServiceAdd
        '
        Me.AcceptButton = Me.ButtonAddService
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.ButtonCancel
        Me.ClientSize = New System.Drawing.Size(530, 452)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.LabelPassword)
        Me.Controls.Add(Me.TextBoxPassword)
        Me.Controls.Add(Me.RadioButtonOtherAccount)
        Me.Controls.Add(Me.RadioButtonSystemAccount)
        Me.Controls.Add(Me.TextBoxUserID)
        Me.Controls.Add(Me.LabelUserID)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.CheckBoxInteractive)
        Me.Controls.Add(Me.ButtonCancel)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtServiceCmd)
        Me.Controls.Add(Me.txtServiceDisplayName)
        Me.Controls.Add(Me.txtServiceName)
        Me.Controls.Add(Me.ButtonAddService)
        Me.Controls.Add(Me.lblType)
        Me.Controls.Add(Me.ComboBoxServiceType)
        Me.Controls.Add(Me.lblCommand)
        Me.Controls.Add(Me.lblDisplayName)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FormServiceAdd"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Service Add"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region


    Private strTitle As String = "Service Add"

    Private strServiceCmd As String
    Private strServiceDisplayName As String
    Private strServiceName As String
    Private strServiceType As String


    Private Sub FormServiceAdd_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = FormRemotePCAdmin.TextBoxComputer.Text & " - " & strTitle
        ComboBoxServiceType.SelectedItem = "Automatic"
    End Sub

    Private Sub ButtonAddService_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonAddService.Click

        If txtServiceName.Text = "" Then
            MsgBox("Please enter a Service Name.", MsgBoxStyle.Information, strTitle)
            txtServiceName.Focus()
            Exit Sub
        End If

        If txtServiceDisplayName.Text = "" Then
            MsgBox("Please enter a Service Display Name.", MsgBoxStyle.Information, strTitle)
            txtServiceDisplayName.Focus()
            Exit Sub
        End If

        If txtServiceCmd.Text = "" Then
            MsgBox("Please enter a Service Command.", MsgBoxStyle.Information, strTitle)
            txtServiceCmd.Focus()
            Exit Sub
        End If

        CreateService()

    End Sub

    Private Sub ButtonCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub


    Private Sub CreateService()


        Const OWN_PROCESS As Integer = 16
        'Const NOT_INTERACTIVE As Boolean = False
        Const NORMAL_ERROR_CONTROL As Integer = 2

        Dim strComputer As String = FormRemotePCAdmin.TextBoxComputer.Text
        If strComputer = "" Then strComputer = "."

        strServiceCmd = txtServiceCmd.Text
        strServiceDisplayName = txtServiceDisplayName.Text
        strServiceName = txtServiceName.Text
        strServiceType = ComboBoxServiceType.SelectedItem.ToString

        Dim strServiceUser As String = TextBoxUserID.Text
        Dim strServicePassword As String = TextBoxPassword.Text
        Dim boolInteractive As Boolean = CheckBoxInteractive.Checked




        Dim oConn As ConnectionOptions = New ConnectionOptions
        If FormRemotePCAdmin.CheckBoxAlternateCredentials.Checked And strComputer <> "." Then
            If FormRemotePCAdmin.strUsername = "" Or FormRemotePCAdmin.strPassword = "" Then
                MsgBox("Please set the credentials.", MsgBoxStyle.Information, "Credentials not set.")
                Exit Sub
            End If
            oConn.Username = FormRemotePCAdmin.strUsername
            oConn.Password = FormRemotePCAdmin.strPassword
        End If


        Try
            Dim mp As ManagementPath = New ManagementPath("\\" & strComputer & "\root\cimv2")
            Dim oMs As ManagementScope = New ManagementScope(mp, oConn)

            Dim path As ManagementPath = New ManagementPath("Win32_BaseService")
            Dim serviceClass As ManagementClass = New ManagementClass(oMs, path, Nothing)

            Dim createParams As ManagementBaseObject = serviceClass.GetMethodParameters("Create")
            createParams.SetPropertyValue("Name", strServiceName)
            createParams.SetPropertyValue("DisplayName", strServiceDisplayName)
            createParams.SetPropertyValue("PathName", strServiceCmd)
            createParams.SetPropertyValue("ServiceType", OWN_PROCESS)
            createParams.SetPropertyValue("ErrorControl", NORMAL_ERROR_CONTROL)
            createParams.SetPropertyValue("StartMode", strServiceType)
            If RadioButtonSystemAccount.Checked Then
                strServiceUser = "LocalSystem"
                createParams.SetPropertyValue("DesktopInteract", boolInteractive)       'TO DO: add optional interactive
            End If

            createParams.SetPropertyValue("StartName", strServiceUser) 'username
            If RadioButtonOtherAccount.Checked Then createParams.SetPropertyValue("StartPassword", strServicePassword) 'password
            'createParams.SetPropertyValue("LoadOrderGroup", "")
            'createParams.SetPropertyValue("LoadOrderGroupDependencies", New String() {})
            'createParams.SetPropertyValue("ServiceDependencies", New String() {})

            Dim ret As ManagementBaseObject = serviceClass.InvokeMethod("Create", createParams, Nothing)
            Dim msg As String
            Dim intReturn As Integer = DirectCast(ret("ReturnValue"), UInteger)

            Dim style As MsgBoxStyle = MsgBoxStyle.Exclamation


            Select Case intReturn
                Case 0
                    msg = "The request was accepted."
                    style = MsgBoxStyle.Information
                Case 1
                    msg = "The request is not supported."
                Case 2
                    msg = "The user did not have the necessary access."
                Case 3
                    msg = "The service cannot be stopped because other services that are running are dependent on it."
                Case 4
                    msg = "The requested control code is not valid, or it is unacceptable to the service."
                Case 5
                    msg = "The requested control code cannot be sent to the service because the state of the service (Win32_BaseServiceState property) is equal to 0, 1, or 2."
                Case 6
                    msg = "The service has not been started."
                Case 7
                    msg = "The service did not respond to the start request in a timely fashion."
                Case 8
                    msg = "Interactive process."
                Case 9
                    msg = "The directory path to the service executable file was not found."
                Case 10
                    msg = "The service is already running."
                Case 11
                    msg = "The database to add a new service is locked."
                Case 12
                    msg = "A dependency on which this service relies has been removed from the system."
                Case 13
                    msg = "The service failed to find the service needed from a dependent service."
                Case 14
                    msg = "The service has been disabled from the system."
                Case 15
                    msg = "The service does not have the correct authentication to run on the system."
                Case 16
                    msg = "This service is being removed from the system."
                Case 17
                    msg = "There is no execution thread for the service."
                Case 18
                    msg = "There are circular dependencies when starting the service."
                Case 19
                    msg = "There is a service running under the same name."
                Case 20
                    msg = "There are invalid characters in the name of the service."
                Case 21
                    msg = "Invalid parameters have been passed to the service."
                Case 22
                    msg = "The account which this service is to run under is either invalid or lacks the permissions to run the service."
                Case 23
                    msg = "The service exists in the database of services available from the system."
                Case 24
                    msg = "The service is currently paused in the system."
                Case Else
                    msg = "Unable to create service."
            End Select

            MsgBox(msg, style)

            If intReturn = 0 Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)
        End Try

    End Sub



    Private Sub RadioButtons_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonSystemAccount.CheckedChanged, RadioButtonOtherAccount.CheckedChanged
        CheckBoxInteractive.Enabled = RadioButtonSystemAccount.Checked
        LabelUserID.Enabled = RadioButtonOtherAccount.Checked
        LabelPassword.Enabled = RadioButtonOtherAccount.Checked
        TextBoxUserID.Enabled = RadioButtonOtherAccount.Checked
        TextBoxPassword.Enabled = RadioButtonOtherAccount.Checked


    End Sub

    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label6.Click
        MsgBox("Make sure this account has been granted the permission ""Log on as a service""", MsgBoxStyle.Information)
    End Sub
End Class
