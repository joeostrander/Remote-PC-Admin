Imports System.Management
Imports Microsoft.Win32
Imports System.Collections.Generic
Imports System.Text.RegularExpressions

Public Class FormRemotePCAdmin
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
    Friend WithEvents lblHost As System.Windows.Forms.Label
    Friend WithEvents TextBoxComputer As System.Windows.Forms.TextBox
    Friend WithEvents CheckBoxAlternateCredentials As System.Windows.Forms.CheckBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents ContextMenuIcon As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItemIcon3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemIcon1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItemIcon2 As System.Windows.Forms.MenuItem
    Friend WithEvents TabSoftware As System.Windows.Forms.TabPage
    Friend WithEvents lstSoftware As System.Windows.Forms.ListView
    Friend WithEvents TabSystemRestore As System.Windows.Forms.TabPage
    Friend WithEvents btnSystemRestore As System.Windows.Forms.Button
    Friend WithEvents LabelSystemRestore As System.Windows.Forms.Label
    Friend WithEvents GroupBoxSystemRestore As System.Windows.Forms.GroupBox
    Friend WithEvents ListViewSystemRestore As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents TabPerformance As System.Windows.Forms.TabPage
    Friend WithEvents lblUptime As System.Windows.Forms.Label
    Friend WithEvents GroupBoxDisk3 As System.Windows.Forms.GroupBox
    Friend WithEvents lblDisk3Max As System.Windows.Forms.Label
    Friend WithEvents ProgBarDisk3 As System.Windows.Forms.ProgressBar
    Friend WithEvents GroupBoxDisk1 As System.Windows.Forms.GroupBox
    Friend WithEvents ProgBarDisk1 As System.Windows.Forms.ProgressBar
    Friend WithEvents lblDisk1Max As System.Windows.Forms.Label
    Friend WithEvents GroupBoxCPU As System.Windows.Forms.GroupBox
    Friend WithEvents lblCpuMax As System.Windows.Forms.Label
    Friend WithEvents ProgBarCpu As System.Windows.Forms.ProgressBar
    Friend WithEvents GroupBoxMem As System.Windows.Forms.GroupBox
    Friend WithEvents lblMemMax As System.Windows.Forms.Label
    Friend WithEvents ProgBarMem As System.Windows.Forms.ProgressBar
    Friend WithEvents GroupBoxDisk2 As System.Windows.Forms.GroupBox
    Friend WithEvents lblDisk2Max As System.Windows.Forms.Label
    Friend WithEvents ProgBarDisk2 As System.Windows.Forms.ProgressBar
    Friend WithEvents TabServices As System.Windows.Forms.TabPage
    Friend WithEvents lstServices As System.Windows.Forms.ListView
    Friend WithEvents TabProcesses As System.Windows.Forms.TabPage
    Friend WithEvents lstProcesses As System.Windows.Forms.ListView
    Friend WithEvents TabEventLogs As System.Windows.Forms.TabPage
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents RadioButtonApplication As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonSecurity As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonSystem As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBoxEventLogs As System.Windows.Forms.GroupBox
    Friend WithEvents RichTextBoxEventLogs As System.Windows.Forms.RichTextBox
    Friend WithEvents MaxResults As System.Windows.Forms.ComboBox
    Friend WithEvents TabUpdates As System.Windows.Forms.TabPage
    Friend WithEvents GroupBoxUpdates As System.Windows.Forms.GroupBox
    Friend WithEvents RichTextBoxUpdates As System.Windows.Forms.RichTextBox
    Friend WithEvents TabSystemInfo As System.Windows.Forms.TabPage
    Friend WithEvents GroupBoxSystemInfo As System.Windows.Forms.GroupBox
    Friend WithEvents RichTextBoxSystemInfo As System.Windows.Forms.RichTextBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents ContextMenuStripProcesses As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents EndProcessToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ViewPathToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ShowInExplorerToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ContextMenuStripSoftware As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ContextMenuStripCopy As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents CopyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ButtonSetCredentials As System.Windows.Forms.Button
    Friend WithEvents ButtonGo As System.Windows.Forms.Button
    Friend WithEvents ColumnHeaderName As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeaderPID As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeaderUsername As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeaderMemUsage As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeaderCPU As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeaderCommandLine As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeaderExecutablePath As System.Windows.Forms.ColumnHeader
    Friend WithEvents CopyAllToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ColumnHeaderDisplayName As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeaderStatus As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeaderStartupType As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeaderDescription As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader_Name As System.Windows.Forms.ColumnHeader
    Friend WithEvents ContextMenuStripServices As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ViewInfoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RestartToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StopToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StartToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RemoveToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ContextMenuStripSystemRestore As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents RestoreToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ActionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ViewToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NewToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ProcessToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ServiceToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RebootToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ShutdownToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LogOffUserToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExploreToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RegeditToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RemoteDesktopToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RemoteAssistToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ManageToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RefreshToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CheckForUpdatesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CopyToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CopyToolStripMenuItem2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CopyAsTableToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CopyAsTableToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FindToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ColumnHeaderCreated As System.Windows.Forms.ColumnHeader
    Friend WithEvents ShowCPUToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UninstallToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormRemotePCAdmin))
        Me.TextBoxComputer = New System.Windows.Forms.TextBox()
        Me.lblHost = New System.Windows.Forms.Label()
        Me.CheckBoxAlternateCredentials = New System.Windows.Forms.CheckBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ListViewSystemRestore = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ContextMenuStripSystemRestore = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.RestoreToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.btnSystemRestore = New System.Windows.Forms.Button()
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenuIcon = New System.Windows.Forms.ContextMenu()
        Me.MenuItemIcon1 = New System.Windows.Forms.MenuItem()
        Me.MenuItemIcon2 = New System.Windows.Forms.MenuItem()
        Me.MenuItemIcon3 = New System.Windows.Forms.MenuItem()
        Me.TabSoftware = New System.Windows.Forms.TabPage()
        Me.lstSoftware = New System.Windows.Forms.ListView()
        Me.ContextMenuStripSoftware = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.UninstallToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CopyToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.CopyToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.TabSystemRestore = New System.Windows.Forms.TabPage()
        Me.LabelSystemRestore = New System.Windows.Forms.Label()
        Me.GroupBoxSystemRestore = New System.Windows.Forms.GroupBox()
        Me.TabPerformance = New System.Windows.Forms.TabPage()
        Me.lblUptime = New System.Windows.Forms.Label()
        Me.GroupBoxDisk3 = New System.Windows.Forms.GroupBox()
        Me.lblDisk3Max = New System.Windows.Forms.Label()
        Me.ProgBarDisk3 = New System.Windows.Forms.ProgressBar()
        Me.GroupBoxDisk1 = New System.Windows.Forms.GroupBox()
        Me.ProgBarDisk1 = New System.Windows.Forms.ProgressBar()
        Me.lblDisk1Max = New System.Windows.Forms.Label()
        Me.GroupBoxCPU = New System.Windows.Forms.GroupBox()
        Me.lblCpuMax = New System.Windows.Forms.Label()
        Me.ProgBarCpu = New System.Windows.Forms.ProgressBar()
        Me.GroupBoxMem = New System.Windows.Forms.GroupBox()
        Me.lblMemMax = New System.Windows.Forms.Label()
        Me.ProgBarMem = New System.Windows.Forms.ProgressBar()
        Me.GroupBoxDisk2 = New System.Windows.Forms.GroupBox()
        Me.lblDisk2Max = New System.Windows.Forms.Label()
        Me.ProgBarDisk2 = New System.Windows.Forms.ProgressBar()
        Me.TabServices = New System.Windows.Forms.TabPage()
        Me.lstServices = New System.Windows.Forms.ListView()
        Me.ColumnHeaderDisplayName = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeaderStatus = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeaderStartupType = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeaderDescription = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader_Name = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ContextMenuStripServices = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ViewInfoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RestartToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StopToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StartToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RemoveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CopyAsTableToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.TabProcesses = New System.Windows.Forms.TabPage()
        Me.lstProcesses = New System.Windows.Forms.ListView()
        Me.ColumnHeaderName = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeaderPID = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeaderUsername = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeaderCreated = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeaderMemUsage = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeaderCPU = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeaderCommandLine = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeaderExecutablePath = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ContextMenuStripProcesses = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.EndProcessToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewPathToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ShowInExplorerToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CopyAsTableToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ShowCPUToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TabEventLogs = New System.Windows.Forms.TabPage()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.RadioButtonApplication = New System.Windows.Forms.RadioButton()
        Me.RadioButtonSecurity = New System.Windows.Forms.RadioButton()
        Me.RadioButtonSystem = New System.Windows.Forms.RadioButton()
        Me.GroupBoxEventLogs = New System.Windows.Forms.GroupBox()
        Me.RichTextBoxEventLogs = New System.Windows.Forms.RichTextBox()
        Me.ContextMenuStripCopy = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.CopyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CopyAllToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FindToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MaxResults = New System.Windows.Forms.ComboBox()
        Me.TabUpdates = New System.Windows.Forms.TabPage()
        Me.GroupBoxUpdates = New System.Windows.Forms.GroupBox()
        Me.RichTextBoxUpdates = New System.Windows.Forms.RichTextBox()
        Me.TabSystemInfo = New System.Windows.Forms.TabPage()
        Me.GroupBoxSystemInfo = New System.Windows.Forms.GroupBox()
        Me.RichTextBoxSystemInfo = New System.Windows.Forms.RichTextBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.ButtonSetCredentials = New System.Windows.Forms.Button()
        Me.ButtonGo = New System.Windows.Forms.Button()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ActionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProcessToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ServiceToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RebootToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ShutdownToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LogOffUserToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExploreToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RegeditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RemoteDesktopToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RemoteAssistToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ManageToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RefreshToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CheckForUpdatesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContextMenuStripSystemRestore.SuspendLayout()
        Me.TabSoftware.SuspendLayout()
        Me.ContextMenuStripSoftware.SuspendLayout()
        Me.TabSystemRestore.SuspendLayout()
        Me.GroupBoxSystemRestore.SuspendLayout()
        Me.TabPerformance.SuspendLayout()
        Me.GroupBoxDisk3.SuspendLayout()
        Me.GroupBoxDisk1.SuspendLayout()
        Me.GroupBoxCPU.SuspendLayout()
        Me.GroupBoxMem.SuspendLayout()
        Me.GroupBoxDisk2.SuspendLayout()
        Me.TabServices.SuspendLayout()
        Me.ContextMenuStripServices.SuspendLayout()
        Me.TabProcesses.SuspendLayout()
        Me.ContextMenuStripProcesses.SuspendLayout()
        Me.TabEventLogs.SuspendLayout()
        Me.GroupBoxEventLogs.SuspendLayout()
        Me.ContextMenuStripCopy.SuspendLayout()
        Me.TabUpdates.SuspendLayout()
        Me.GroupBoxUpdates.SuspendLayout()
        Me.TabSystemInfo.SuspendLayout()
        Me.GroupBoxSystemInfo.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBoxComputer
        '
        Me.TextBoxComputer.Location = New System.Drawing.Point(46, 38)
        Me.TextBoxComputer.Name = "TextBoxComputer"
        Me.TextBoxComputer.Size = New System.Drawing.Size(197, 20)
        Me.TextBoxComputer.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.TextBoxComputer, "Enter Hostname.  If left blank, local computer is assumed.")
        '
        'lblHost
        '
        Me.lblHost.Location = New System.Drawing.Point(8, 35)
        Me.lblHost.Name = "lblHost"
        Me.lblHost.Size = New System.Drawing.Size(32, 24)
        Me.lblHost.TabIndex = 0
        Me.lblHost.Text = "H&ost:"
        Me.lblHost.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'CheckBoxAlternateCredentials
        '
        Me.CheckBoxAlternateCredentials.Location = New System.Drawing.Point(295, 36)
        Me.CheckBoxAlternateCredentials.Name = "CheckBoxAlternateCredentials"
        Me.CheckBoxAlternateCredentials.Size = New System.Drawing.Size(159, 24)
        Me.CheckBoxAlternateCredentials.TabIndex = 3
        Me.CheckBoxAlternateCredentials.Text = "&Use Alternate Credentials"
        Me.ToolTip1.SetToolTip(Me.CheckBoxAlternateCredentials, "Click to connect as a different user.")
        '
        'ListViewSystemRestore
        '
        Me.ListViewSystemRestore.BackColor = System.Drawing.Color.MidnightBlue
        Me.ListViewSystemRestore.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ListViewSystemRestore.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4})
        Me.ListViewSystemRestore.ContextMenuStrip = Me.ContextMenuStripSystemRestore
        Me.ListViewSystemRestore.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListViewSystemRestore.ForeColor = System.Drawing.Color.CornflowerBlue
        Me.ListViewSystemRestore.FullRowSelect = True
        Me.ListViewSystemRestore.Location = New System.Drawing.Point(3, 16)
        Me.ListViewSystemRestore.MultiSelect = False
        Me.ListViewSystemRestore.Name = "ListViewSystemRestore"
        Me.ListViewSystemRestore.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ListViewSystemRestore.Size = New System.Drawing.Size(672, 405)
        Me.ListViewSystemRestore.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.ListViewSystemRestore, "Use Right-click to restore system to the selected point.")
        Me.ListViewSystemRestore.UseCompatibleStateImageBehavior = False
        Me.ListViewSystemRestore.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "#"
        Me.ColumnHeader1.Width = 38
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Description"
        Me.ColumnHeader2.Width = 148
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Type"
        Me.ColumnHeader3.Width = 133
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Date"
        Me.ColumnHeader4.Width = 176
        '
        'ContextMenuStripSystemRestore
        '
        Me.ContextMenuStripSystemRestore.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RestoreToolStripMenuItem})
        Me.ContextMenuStripSystemRestore.Name = "ContextMenuStripSystemRestore"
        Me.ContextMenuStripSystemRestore.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.ContextMenuStripSystemRestore.ShowImageMargin = False
        Me.ContextMenuStripSystemRestore.Size = New System.Drawing.Size(89, 26)
        '
        'RestoreToolStripMenuItem
        '
        Me.RestoreToolStripMenuItem.Name = "RestoreToolStripMenuItem"
        Me.RestoreToolStripMenuItem.Size = New System.Drawing.Size(88, 22)
        Me.RestoreToolStripMenuItem.Text = "&Restore"
        '
        'btnSystemRestore
        '
        Me.btnSystemRestore.Location = New System.Drawing.Point(232, 16)
        Me.btnSystemRestore.Name = "btnSystemRestore"
        Me.btnSystemRestore.Size = New System.Drawing.Size(75, 23)
        Me.btnSystemRestore.TabIndex = 2
        Me.btnSystemRestore.Text = "Enable / Disable"
        Me.ToolTip1.SetToolTip(Me.btnSystemRestore, "Click to change status.")
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "Remote PC Admin"
        '
        'ContextMenuIcon
        '
        Me.ContextMenuIcon.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItemIcon1, Me.MenuItemIcon2, Me.MenuItemIcon3})
        '
        'MenuItemIcon1
        '
        Me.MenuItemIcon1.Index = 0
        Me.MenuItemIcon1.Text = "&Open Remote PC Admin"
        '
        'MenuItemIcon2
        '
        Me.MenuItemIcon2.Index = 1
        Me.MenuItemIcon2.Text = "&About"
        '
        'MenuItemIcon3
        '
        Me.MenuItemIcon3.Index = 2
        Me.MenuItemIcon3.Text = "E&xit"
        '
        'TabSoftware
        '
        Me.TabSoftware.Controls.Add(Me.lstSoftware)
        Me.TabSoftware.Location = New System.Drawing.Point(4, 22)
        Me.TabSoftware.Name = "TabSoftware"
        Me.TabSoftware.Padding = New System.Windows.Forms.Padding(3)
        Me.TabSoftware.Size = New System.Drawing.Size(684, 475)
        Me.TabSoftware.TabIndex = 11
        Me.TabSoftware.Text = "Software"
        Me.TabSoftware.UseVisualStyleBackColor = True
        '
        'lstSoftware
        '
        Me.lstSoftware.BackColor = System.Drawing.Color.MidnightBlue
        Me.lstSoftware.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstSoftware.ContextMenuStrip = Me.ContextMenuStripSoftware
        Me.lstSoftware.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstSoftware.ForeColor = System.Drawing.Color.CornflowerBlue
        Me.lstSoftware.FullRowSelect = True
        Me.lstSoftware.Location = New System.Drawing.Point(3, 3)
        Me.lstSoftware.Name = "lstSoftware"
        Me.lstSoftware.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstSoftware.Size = New System.Drawing.Size(678, 469)
        Me.lstSoftware.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.lstSoftware.TabIndex = 2
        Me.lstSoftware.UseCompatibleStateImageBehavior = False
        Me.lstSoftware.View = System.Windows.Forms.View.Details
        '
        'ContextMenuStripSoftware
        '
        Me.ContextMenuStripSoftware.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UninstallToolStripMenuItem, Me.CopyToolStripMenuItem1, Me.CopyToolStripMenuItem2})
        Me.ContextMenuStripSoftware.Name = "ContextMenuStripSoftware"
        Me.ContextMenuStripSoftware.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.ContextMenuStripSoftware.ShowImageMargin = False
        Me.ContextMenuStripSoftware.Size = New System.Drawing.Size(148, 70)
        '
        'UninstallToolStripMenuItem
        '
        Me.UninstallToolStripMenuItem.Name = "UninstallToolStripMenuItem"
        Me.UninstallToolStripMenuItem.Size = New System.Drawing.Size(147, 22)
        Me.UninstallToolStripMenuItem.Text = "&Uninstall"
        '
        'CopyToolStripMenuItem1
        '
        Me.CopyToolStripMenuItem1.Name = "CopyToolStripMenuItem1"
        Me.CopyToolStripMenuItem1.Size = New System.Drawing.Size(147, 22)
        Me.CopyToolStripMenuItem1.Text = "&Copy Software List"
        '
        'CopyToolStripMenuItem2
        '
        Me.CopyToolStripMenuItem2.Name = "CopyToolStripMenuItem2"
        Me.CopyToolStripMenuItem2.Size = New System.Drawing.Size(147, 22)
        Me.CopyToolStripMenuItem2.Text = "Copy as &Table"
        '
        'TabSystemRestore
        '
        Me.TabSystemRestore.Controls.Add(Me.btnSystemRestore)
        Me.TabSystemRestore.Controls.Add(Me.LabelSystemRestore)
        Me.TabSystemRestore.Controls.Add(Me.GroupBoxSystemRestore)
        Me.TabSystemRestore.Location = New System.Drawing.Point(4, 22)
        Me.TabSystemRestore.Name = "TabSystemRestore"
        Me.TabSystemRestore.Size = New System.Drawing.Size(684, 475)
        Me.TabSystemRestore.TabIndex = 8
        Me.TabSystemRestore.Text = "System Restore"
        Me.TabSystemRestore.UseVisualStyleBackColor = True
        '
        'LabelSystemRestore
        '
        Me.LabelSystemRestore.Location = New System.Drawing.Point(24, 16)
        Me.LabelSystemRestore.Name = "LabelSystemRestore"
        Me.LabelSystemRestore.Size = New System.Drawing.Size(208, 23)
        Me.LabelSystemRestore.TabIndex = 1
        Me.LabelSystemRestore.Text = "System restore is currently..."
        Me.LabelSystemRestore.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GroupBoxSystemRestore
        '
        Me.GroupBoxSystemRestore.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBoxSystemRestore.Controls.Add(Me.ListViewSystemRestore)
        Me.GroupBoxSystemRestore.Location = New System.Drawing.Point(3, 48)
        Me.GroupBoxSystemRestore.Name = "GroupBoxSystemRestore"
        Me.GroupBoxSystemRestore.Size = New System.Drawing.Size(678, 424)
        Me.GroupBoxSystemRestore.TabIndex = 0
        Me.GroupBoxSystemRestore.TabStop = False
        '
        'TabPerformance
        '
        Me.TabPerformance.Controls.Add(Me.lblUptime)
        Me.TabPerformance.Controls.Add(Me.GroupBoxDisk3)
        Me.TabPerformance.Controls.Add(Me.GroupBoxDisk1)
        Me.TabPerformance.Controls.Add(Me.GroupBoxCPU)
        Me.TabPerformance.Controls.Add(Me.GroupBoxMem)
        Me.TabPerformance.Controls.Add(Me.GroupBoxDisk2)
        Me.TabPerformance.Location = New System.Drawing.Point(4, 22)
        Me.TabPerformance.Name = "TabPerformance"
        Me.TabPerformance.Size = New System.Drawing.Size(684, 475)
        Me.TabPerformance.TabIndex = 6
        Me.TabPerformance.Text = "Performance"
        Me.TabPerformance.UseVisualStyleBackColor = True
        '
        'lblUptime
        '
        Me.lblUptime.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUptime.Location = New System.Drawing.Point(8, 8)
        Me.lblUptime.Name = "lblUptime"
        Me.lblUptime.Size = New System.Drawing.Size(304, 16)
        Me.lblUptime.TabIndex = 15
        Me.lblUptime.Text = "Computer has been up for ..."
        '
        'GroupBoxDisk3
        '
        Me.GroupBoxDisk3.Controls.Add(Me.lblDisk3Max)
        Me.GroupBoxDisk3.Controls.Add(Me.ProgBarDisk3)
        Me.GroupBoxDisk3.Location = New System.Drawing.Point(8, 256)
        Me.GroupBoxDisk3.Name = "GroupBoxDisk3"
        Me.GroupBoxDisk3.Size = New System.Drawing.Size(496, 48)
        Me.GroupBoxDisk3.TabIndex = 14
        Me.GroupBoxDisk3.TabStop = False
        Me.GroupBoxDisk3.Text = "Disk 3"
        '
        'lblDisk3Max
        '
        Me.lblDisk3Max.Location = New System.Drawing.Point(296, 16)
        Me.lblDisk3Max.Name = "lblDisk3Max"
        Me.lblDisk3Max.Size = New System.Drawing.Size(192, 23)
        Me.lblDisk3Max.TabIndex = 9
        Me.lblDisk3Max.Text = "Disk Space"
        Me.lblDisk3Max.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ProgBarDisk3
        '
        Me.ProgBarDisk3.Location = New System.Drawing.Point(8, 16)
        Me.ProgBarDisk3.Name = "ProgBarDisk3"
        Me.ProgBarDisk3.Size = New System.Drawing.Size(272, 24)
        Me.ProgBarDisk3.TabIndex = 7
        '
        'GroupBoxDisk1
        '
        Me.GroupBoxDisk1.Controls.Add(Me.ProgBarDisk1)
        Me.GroupBoxDisk1.Controls.Add(Me.lblDisk1Max)
        Me.GroupBoxDisk1.Location = New System.Drawing.Point(8, 144)
        Me.GroupBoxDisk1.Name = "GroupBoxDisk1"
        Me.GroupBoxDisk1.Size = New System.Drawing.Size(496, 48)
        Me.GroupBoxDisk1.TabIndex = 12
        Me.GroupBoxDisk1.TabStop = False
        Me.GroupBoxDisk1.Text = "Disk 1"
        '
        'ProgBarDisk1
        '
        Me.ProgBarDisk1.Location = New System.Drawing.Point(8, 16)
        Me.ProgBarDisk1.Name = "ProgBarDisk1"
        Me.ProgBarDisk1.Size = New System.Drawing.Size(272, 24)
        Me.ProgBarDisk1.TabIndex = 4
        '
        'lblDisk1Max
        '
        Me.lblDisk1Max.Location = New System.Drawing.Point(296, 16)
        Me.lblDisk1Max.Name = "lblDisk1Max"
        Me.lblDisk1Max.Size = New System.Drawing.Size(192, 23)
        Me.lblDisk1Max.TabIndex = 6
        Me.lblDisk1Max.Text = "Disk Space"
        Me.lblDisk1Max.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GroupBoxCPU
        '
        Me.GroupBoxCPU.Controls.Add(Me.lblCpuMax)
        Me.GroupBoxCPU.Controls.Add(Me.ProgBarCpu)
        Me.GroupBoxCPU.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBoxCPU.Location = New System.Drawing.Point(8, 88)
        Me.GroupBoxCPU.Name = "GroupBoxCPU"
        Me.GroupBoxCPU.Size = New System.Drawing.Size(496, 48)
        Me.GroupBoxCPU.TabIndex = 11
        Me.GroupBoxCPU.TabStop = False
        Me.GroupBoxCPU.Text = "CPU Usage"
        '
        'lblCpuMax
        '
        Me.lblCpuMax.Location = New System.Drawing.Point(296, 16)
        Me.lblCpuMax.Name = "lblCpuMax"
        Me.lblCpuMax.Size = New System.Drawing.Size(184, 23)
        Me.lblCpuMax.TabIndex = 6
        Me.lblCpuMax.Text = "CPU Time"
        Me.lblCpuMax.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ProgBarCpu
        '
        Me.ProgBarCpu.Location = New System.Drawing.Point(8, 16)
        Me.ProgBarCpu.Name = "ProgBarCpu"
        Me.ProgBarCpu.Size = New System.Drawing.Size(272, 24)
        Me.ProgBarCpu.TabIndex = 4
        '
        'GroupBoxMem
        '
        Me.GroupBoxMem.Controls.Add(Me.lblMemMax)
        Me.GroupBoxMem.Controls.Add(Me.ProgBarMem)
        Me.GroupBoxMem.Location = New System.Drawing.Point(8, 32)
        Me.GroupBoxMem.Name = "GroupBoxMem"
        Me.GroupBoxMem.Size = New System.Drawing.Size(496, 48)
        Me.GroupBoxMem.TabIndex = 10
        Me.GroupBoxMem.TabStop = False
        Me.GroupBoxMem.Text = "Memory"
        '
        'lblMemMax
        '
        Me.lblMemMax.Location = New System.Drawing.Point(296, 16)
        Me.lblMemMax.Name = "lblMemMax"
        Me.lblMemMax.Size = New System.Drawing.Size(184, 23)
        Me.lblMemMax.TabIndex = 3
        Me.lblMemMax.Text = "Free"
        Me.lblMemMax.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ProgBarMem
        '
        Me.ProgBarMem.Location = New System.Drawing.Point(8, 16)
        Me.ProgBarMem.Name = "ProgBarMem"
        Me.ProgBarMem.Size = New System.Drawing.Size(272, 24)
        Me.ProgBarMem.TabIndex = 1
        '
        'GroupBoxDisk2
        '
        Me.GroupBoxDisk2.Controls.Add(Me.lblDisk2Max)
        Me.GroupBoxDisk2.Controls.Add(Me.ProgBarDisk2)
        Me.GroupBoxDisk2.Location = New System.Drawing.Point(8, 200)
        Me.GroupBoxDisk2.Name = "GroupBoxDisk2"
        Me.GroupBoxDisk2.Size = New System.Drawing.Size(496, 48)
        Me.GroupBoxDisk2.TabIndex = 13
        Me.GroupBoxDisk2.TabStop = False
        Me.GroupBoxDisk2.Text = "Disk 2"
        '
        'lblDisk2Max
        '
        Me.lblDisk2Max.Location = New System.Drawing.Point(296, 16)
        Me.lblDisk2Max.Name = "lblDisk2Max"
        Me.lblDisk2Max.Size = New System.Drawing.Size(192, 23)
        Me.lblDisk2Max.TabIndex = 9
        Me.lblDisk2Max.Text = "Disk Space"
        Me.lblDisk2Max.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ProgBarDisk2
        '
        Me.ProgBarDisk2.Location = New System.Drawing.Point(8, 16)
        Me.ProgBarDisk2.Name = "ProgBarDisk2"
        Me.ProgBarDisk2.Size = New System.Drawing.Size(272, 24)
        Me.ProgBarDisk2.TabIndex = 7
        '
        'TabServices
        '
        Me.TabServices.Controls.Add(Me.lstServices)
        Me.TabServices.Location = New System.Drawing.Point(4, 22)
        Me.TabServices.Name = "TabServices"
        Me.TabServices.Size = New System.Drawing.Size(684, 475)
        Me.TabServices.TabIndex = 0
        Me.TabServices.Text = "Services"
        Me.TabServices.UseVisualStyleBackColor = True
        '
        'lstServices
        '
        Me.lstServices.BackColor = System.Drawing.Color.MidnightBlue
        Me.lstServices.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstServices.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeaderDisplayName, Me.ColumnHeaderStatus, Me.ColumnHeaderStartupType, Me.ColumnHeaderDescription, Me.ColumnHeader_Name})
        Me.lstServices.ContextMenuStrip = Me.ContextMenuStripServices
        Me.lstServices.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstServices.ForeColor = System.Drawing.Color.CornflowerBlue
        Me.lstServices.FullRowSelect = True
        Me.lstServices.Location = New System.Drawing.Point(0, 0)
        Me.lstServices.Name = "lstServices"
        Me.lstServices.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstServices.Size = New System.Drawing.Size(684, 475)
        Me.lstServices.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.lstServices.TabIndex = 1
        Me.lstServices.UseCompatibleStateImageBehavior = False
        Me.lstServices.View = System.Windows.Forms.View.Details
        '
        'ColumnHeaderDisplayName
        '
        Me.ColumnHeaderDisplayName.Tag = "String"
        Me.ColumnHeaderDisplayName.Text = "Display Name"
        Me.ColumnHeaderDisplayName.Width = 320
        '
        'ColumnHeaderStatus
        '
        Me.ColumnHeaderStatus.Tag = "String"
        Me.ColumnHeaderStatus.Text = "Status"
        Me.ColumnHeaderStatus.Width = 81
        '
        'ColumnHeaderStartupType
        '
        Me.ColumnHeaderStartupType.Tag = "String"
        Me.ColumnHeaderStartupType.Text = "Startup Type"
        Me.ColumnHeaderStartupType.Width = 85
        '
        'ColumnHeaderDescription
        '
        Me.ColumnHeaderDescription.Tag = "String"
        Me.ColumnHeaderDescription.Text = "Description"
        Me.ColumnHeaderDescription.Width = 0
        '
        'ColumnHeader_Name
        '
        Me.ColumnHeader_Name.Tag = "String"
        Me.ColumnHeader_Name.Text = "Name"
        Me.ColumnHeader_Name.Width = 0
        '
        'ContextMenuStripServices
        '
        Me.ContextMenuStripServices.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ViewInfoToolStripMenuItem, Me.RestartToolStripMenuItem, Me.StopToolStripMenuItem, Me.StartToolStripMenuItem, Me.RemoveToolStripMenuItem, Me.CopyAsTableToolStripMenuItem1})
        Me.ContextMenuStripServices.Name = "ContextMenuStripServices"
        Me.ContextMenuStripServices.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.ContextMenuStripServices.ShowImageMargin = False
        Me.ContextMenuStripServices.Size = New System.Drawing.Size(123, 136)
        '
        'ViewInfoToolStripMenuItem
        '
        Me.ViewInfoToolStripMenuItem.Name = "ViewInfoToolStripMenuItem"
        Me.ViewInfoToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.ViewInfoToolStripMenuItem.Text = "&View Info"
        '
        'RestartToolStripMenuItem
        '
        Me.RestartToolStripMenuItem.Name = "RestartToolStripMenuItem"
        Me.RestartToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.RestartToolStripMenuItem.Text = "&Restart"
        Me.RestartToolStripMenuItem.Visible = False
        '
        'StopToolStripMenuItem
        '
        Me.StopToolStripMenuItem.Name = "StopToolStripMenuItem"
        Me.StopToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.StopToolStripMenuItem.Text = "Sto&p"
        '
        'StartToolStripMenuItem
        '
        Me.StartToolStripMenuItem.Name = "StartToolStripMenuItem"
        Me.StartToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.StartToolStripMenuItem.Text = "&Start"
        '
        'RemoveToolStripMenuItem
        '
        Me.RemoveToolStripMenuItem.Name = "RemoveToolStripMenuItem"
        Me.RemoveToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.RemoveToolStripMenuItem.Text = "Remo&ve"
        '
        'CopyAsTableToolStripMenuItem1
        '
        Me.CopyAsTableToolStripMenuItem1.Name = "CopyAsTableToolStripMenuItem1"
        Me.CopyAsTableToolStripMenuItem1.Size = New System.Drawing.Size(122, 22)
        Me.CopyAsTableToolStripMenuItem1.Text = "&Copy as Table"
        '
        'TabProcesses
        '
        Me.TabProcesses.Controls.Add(Me.lstProcesses)
        Me.TabProcesses.Location = New System.Drawing.Point(4, 22)
        Me.TabProcesses.Name = "TabProcesses"
        Me.TabProcesses.Size = New System.Drawing.Size(684, 475)
        Me.TabProcesses.TabIndex = 2
        Me.TabProcesses.Text = "Processes"
        Me.TabProcesses.UseVisualStyleBackColor = True
        '
        'lstProcesses
        '
        Me.lstProcesses.AllowColumnReorder = True
        Me.lstProcesses.BackColor = System.Drawing.Color.MidnightBlue
        Me.lstProcesses.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstProcesses.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeaderName, Me.ColumnHeaderPID, Me.ColumnHeaderUsername, Me.ColumnHeaderCreated, Me.ColumnHeaderMemUsage, Me.ColumnHeaderCPU, Me.ColumnHeaderCommandLine, Me.ColumnHeaderExecutablePath})
        Me.lstProcesses.ContextMenuStrip = Me.ContextMenuStripProcesses
        Me.lstProcesses.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstProcesses.ForeColor = System.Drawing.Color.CornflowerBlue
        Me.lstProcesses.FullRowSelect = True
        Me.lstProcesses.Location = New System.Drawing.Point(0, 0)
        Me.lstProcesses.Name = "lstProcesses"
        Me.lstProcesses.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstProcesses.Size = New System.Drawing.Size(684, 475)
        Me.lstProcesses.TabIndex = 0
        Me.lstProcesses.UseCompatibleStateImageBehavior = False
        Me.lstProcesses.View = System.Windows.Forms.View.Details
        '
        'ColumnHeaderName
        '
        Me.ColumnHeaderName.Tag = "String"
        Me.ColumnHeaderName.Text = "Name"
        Me.ColumnHeaderName.Width = 140
        '
        'ColumnHeaderPID
        '
        Me.ColumnHeaderPID.Tag = "Numeric"
        Me.ColumnHeaderPID.Text = "PID"
        Me.ColumnHeaderPID.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeaderPID.Width = 50
        '
        'ColumnHeaderUsername
        '
        Me.ColumnHeaderUsername.Text = "User Name"
        Me.ColumnHeaderUsername.Width = 180
        '
        'ColumnHeaderCreated
        '
        Me.ColumnHeaderCreated.Tag = "Date"
        Me.ColumnHeaderCreated.Text = "Created"
        Me.ColumnHeaderCreated.Width = 140
        '
        'ColumnHeaderMemUsage
        '
        Me.ColumnHeaderMemUsage.Tag = "Numeric"
        Me.ColumnHeaderMemUsage.Text = "Mem Usage"
        Me.ColumnHeaderMemUsage.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ColumnHeaderMemUsage.Width = 75
        '
        'ColumnHeaderCPU
        '
        Me.ColumnHeaderCPU.Tag = "Numeric"
        Me.ColumnHeaderCPU.Text = "CPU"
        Me.ColumnHeaderCPU.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'ColumnHeaderCommandLine
        '
        Me.ColumnHeaderCommandLine.Tag = "String"
        Me.ColumnHeaderCommandLine.Text = "Command Line"
        Me.ColumnHeaderCommandLine.Width = 500
        '
        'ColumnHeaderExecutablePath
        '
        Me.ColumnHeaderExecutablePath.Tag = "String"
        Me.ColumnHeaderExecutablePath.Text = "Executable Path"
        '
        'ContextMenuStripProcesses
        '
        Me.ContextMenuStripProcesses.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EndProcessToolStripMenuItem, Me.ViewPathToolStripMenuItem, Me.ShowInExplorerToolStripMenuItem, Me.CopyAsTableToolStripMenuItem, Me.ShowCPUToolStripMenuItem})
        Me.ContextMenuStripProcesses.Name = "ContextMenuStripProcesses"
        Me.ContextMenuStripProcesses.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.ContextMenuStripProcesses.ShowCheckMargin = True
        Me.ContextMenuStripProcesses.ShowImageMargin = False
        Me.ContextMenuStripProcesses.Size = New System.Drawing.Size(185, 114)
        '
        'EndProcessToolStripMenuItem
        '
        Me.EndProcessToolStripMenuItem.Name = "EndProcessToolStripMenuItem"
        Me.EndProcessToolStripMenuItem.Size = New System.Drawing.Size(184, 22)
        Me.EndProcessToolStripMenuItem.Text = "&End Process"
        '
        'ViewPathToolStripMenuItem
        '
        Me.ViewPathToolStripMenuItem.Name = "ViewPathToolStripMenuItem"
        Me.ViewPathToolStripMenuItem.Size = New System.Drawing.Size(184, 22)
        Me.ViewPathToolStripMenuItem.Text = "&View Command Line"
        '
        'ShowInExplorerToolStripMenuItem
        '
        Me.ShowInExplorerToolStripMenuItem.Name = "ShowInExplorerToolStripMenuItem"
        Me.ShowInExplorerToolStripMenuItem.Size = New System.Drawing.Size(184, 22)
        Me.ShowInExplorerToolStripMenuItem.Text = "&Show in Explorer"
        '
        'CopyAsTableToolStripMenuItem
        '
        Me.CopyAsTableToolStripMenuItem.Name = "CopyAsTableToolStripMenuItem"
        Me.CopyAsTableToolStripMenuItem.Size = New System.Drawing.Size(184, 22)
        Me.CopyAsTableToolStripMenuItem.Text = "&Copy as Table"
        '
        'ShowCPUToolStripMenuItem
        '
        Me.ShowCPUToolStripMenuItem.CheckOnClick = True
        Me.ShowCPUToolStripMenuItem.Name = "ShowCPUToolStripMenuItem"
        Me.ShowCPUToolStripMenuItem.Size = New System.Drawing.Size(184, 22)
        Me.ShowCPUToolStripMenuItem.Text = "Show CPU (slower)"
        '
        'TabEventLogs
        '
        Me.TabEventLogs.Controls.Add(Me.Label2)
        Me.TabEventLogs.Controls.Add(Me.RadioButtonApplication)
        Me.TabEventLogs.Controls.Add(Me.RadioButtonSecurity)
        Me.TabEventLogs.Controls.Add(Me.RadioButtonSystem)
        Me.TabEventLogs.Controls.Add(Me.GroupBoxEventLogs)
        Me.TabEventLogs.Controls.Add(Me.MaxResults)
        Me.TabEventLogs.Location = New System.Drawing.Point(4, 22)
        Me.TabEventLogs.Name = "TabEventLogs"
        Me.TabEventLogs.Size = New System.Drawing.Size(684, 475)
        Me.TabEventLogs.TabIndex = 7
        Me.TabEventLogs.Text = "Event Logs"
        Me.TabEventLogs.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(336, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 23)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Max. Results:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'RadioButtonApplication
        '
        Me.RadioButtonApplication.Checked = True
        Me.RadioButtonApplication.Location = New System.Drawing.Point(16, 8)
        Me.RadioButtonApplication.Name = "RadioButtonApplication"
        Me.RadioButtonApplication.Size = New System.Drawing.Size(104, 24)
        Me.RadioButtonApplication.TabIndex = 4
        Me.RadioButtonApplication.TabStop = True
        Me.RadioButtonApplication.Text = "Application"
        '
        'RadioButtonSecurity
        '
        Me.RadioButtonSecurity.Location = New System.Drawing.Point(128, 8)
        Me.RadioButtonSecurity.Name = "RadioButtonSecurity"
        Me.RadioButtonSecurity.Size = New System.Drawing.Size(104, 24)
        Me.RadioButtonSecurity.TabIndex = 3
        Me.RadioButtonSecurity.Text = "Security"
        '
        'RadioButtonSystem
        '
        Me.RadioButtonSystem.Location = New System.Drawing.Point(240, 8)
        Me.RadioButtonSystem.Name = "RadioButtonSystem"
        Me.RadioButtonSystem.Size = New System.Drawing.Size(72, 24)
        Me.RadioButtonSystem.TabIndex = 2
        Me.RadioButtonSystem.Text = "System"
        '
        'GroupBoxEventLogs
        '
        Me.GroupBoxEventLogs.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBoxEventLogs.Controls.Add(Me.RichTextBoxEventLogs)
        Me.GroupBoxEventLogs.Location = New System.Drawing.Point(3, 40)
        Me.GroupBoxEventLogs.Name = "GroupBoxEventLogs"
        Me.GroupBoxEventLogs.Size = New System.Drawing.Size(678, 432)
        Me.GroupBoxEventLogs.TabIndex = 1
        Me.GroupBoxEventLogs.TabStop = False
        '
        'RichTextBoxEventLogs
        '
        Me.RichTextBoxEventLogs.BackColor = System.Drawing.Color.MidnightBlue
        Me.RichTextBoxEventLogs.ContextMenuStrip = Me.ContextMenuStripCopy
        Me.RichTextBoxEventLogs.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RichTextBoxEventLogs.ForeColor = System.Drawing.Color.CornflowerBlue
        Me.RichTextBoxEventLogs.Location = New System.Drawing.Point(3, 16)
        Me.RichTextBoxEventLogs.Name = "RichTextBoxEventLogs"
        Me.RichTextBoxEventLogs.ReadOnly = True
        Me.RichTextBoxEventLogs.Size = New System.Drawing.Size(672, 413)
        Me.RichTextBoxEventLogs.TabIndex = 0
        Me.RichTextBoxEventLogs.Text = "EventLogs..."
        '
        'ContextMenuStripCopy
        '
        Me.ContextMenuStripCopy.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CopyToolStripMenuItem, Me.CopyAllToolStripMenuItem, Me.FindToolStripMenuItem})
        Me.ContextMenuStripCopy.Name = "ContextMenuStripCopy"
        Me.ContextMenuStripCopy.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.ContextMenuStripCopy.ShowImageMargin = False
        Me.ContextMenuStripCopy.Size = New System.Drawing.Size(125, 70)
        '
        'CopyToolStripMenuItem
        '
        Me.CopyToolStripMenuItem.Name = "CopyToolStripMenuItem"
        Me.CopyToolStripMenuItem.Size = New System.Drawing.Size(124, 22)
        Me.CopyToolStripMenuItem.Text = "&Copy Selected"
        '
        'CopyAllToolStripMenuItem
        '
        Me.CopyAllToolStripMenuItem.Name = "CopyAllToolStripMenuItem"
        Me.CopyAllToolStripMenuItem.Size = New System.Drawing.Size(124, 22)
        Me.CopyAllToolStripMenuItem.Text = "Copy &All"
        '
        'FindToolStripMenuItem
        '
        Me.FindToolStripMenuItem.Name = "FindToolStripMenuItem"
        Me.FindToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F), System.Windows.Forms.Keys)
        Me.FindToolStripMenuItem.Size = New System.Drawing.Size(124, 22)
        Me.FindToolStripMenuItem.Text = "&Find"
        '
        'MaxResults
        '
        Me.MaxResults.Items.AddRange(New Object() {"25", "50", "100", "200"})
        Me.MaxResults.Location = New System.Drawing.Point(432, 8)
        Me.MaxResults.Name = "MaxResults"
        Me.MaxResults.Size = New System.Drawing.Size(72, 21)
        Me.MaxResults.TabIndex = 11
        Me.MaxResults.Text = "25"
        '
        'TabUpdates
        '
        Me.TabUpdates.Controls.Add(Me.GroupBoxUpdates)
        Me.TabUpdates.Location = New System.Drawing.Point(4, 22)
        Me.TabUpdates.Name = "TabUpdates"
        Me.TabUpdates.Size = New System.Drawing.Size(684, 475)
        Me.TabUpdates.TabIndex = 10
        Me.TabUpdates.Text = "Updates"
        Me.TabUpdates.UseVisualStyleBackColor = True
        '
        'GroupBoxUpdates
        '
        Me.GroupBoxUpdates.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBoxUpdates.Controls.Add(Me.RichTextBoxUpdates)
        Me.GroupBoxUpdates.Location = New System.Drawing.Point(3, 6)
        Me.GroupBoxUpdates.Name = "GroupBoxUpdates"
        Me.GroupBoxUpdates.Size = New System.Drawing.Size(678, 466)
        Me.GroupBoxUpdates.TabIndex = 5
        Me.GroupBoxUpdates.TabStop = False
        '
        'RichTextBoxUpdates
        '
        Me.RichTextBoxUpdates.BackColor = System.Drawing.Color.MidnightBlue
        Me.RichTextBoxUpdates.ContextMenuStrip = Me.ContextMenuStripCopy
        Me.RichTextBoxUpdates.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RichTextBoxUpdates.ForeColor = System.Drawing.Color.CornflowerBlue
        Me.RichTextBoxUpdates.Location = New System.Drawing.Point(3, 16)
        Me.RichTextBoxUpdates.Name = "RichTextBoxUpdates"
        Me.RichTextBoxUpdates.ReadOnly = True
        Me.RichTextBoxUpdates.Size = New System.Drawing.Size(672, 447)
        Me.RichTextBoxUpdates.TabIndex = 1
        Me.RichTextBoxUpdates.Text = "Updates..."
        '
        'TabSystemInfo
        '
        Me.TabSystemInfo.Controls.Add(Me.GroupBoxSystemInfo)
        Me.TabSystemInfo.Location = New System.Drawing.Point(4, 22)
        Me.TabSystemInfo.Name = "TabSystemInfo"
        Me.TabSystemInfo.Size = New System.Drawing.Size(684, 475)
        Me.TabSystemInfo.TabIndex = 3
        Me.TabSystemInfo.Text = "System Info"
        Me.TabSystemInfo.UseVisualStyleBackColor = True
        '
        'GroupBoxSystemInfo
        '
        Me.GroupBoxSystemInfo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBoxSystemInfo.Controls.Add(Me.RichTextBoxSystemInfo)
        Me.GroupBoxSystemInfo.Location = New System.Drawing.Point(3, 8)
        Me.GroupBoxSystemInfo.Name = "GroupBoxSystemInfo"
        Me.GroupBoxSystemInfo.Size = New System.Drawing.Size(678, 463)
        Me.GroupBoxSystemInfo.TabIndex = 3
        Me.GroupBoxSystemInfo.TabStop = False
        '
        'RichTextBoxSystemInfo
        '
        Me.RichTextBoxSystemInfo.BackColor = System.Drawing.Color.MidnightBlue
        Me.RichTextBoxSystemInfo.ContextMenuStrip = Me.ContextMenuStripCopy
        Me.RichTextBoxSystemInfo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RichTextBoxSystemInfo.ForeColor = System.Drawing.Color.CornflowerBlue
        Me.RichTextBoxSystemInfo.Location = New System.Drawing.Point(3, 16)
        Me.RichTextBoxSystemInfo.Name = "RichTextBoxSystemInfo"
        Me.RichTextBoxSystemInfo.ReadOnly = True
        Me.RichTextBoxSystemInfo.Size = New System.Drawing.Size(672, 444)
        Me.RichTextBoxSystemInfo.TabIndex = 5
        Me.RichTextBoxSystemInfo.Text = "System Info..."
        '
        'TabControl1
        '
        Me.TabControl1.AllowDrop = True
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabSystemInfo)
        Me.TabControl1.Controls.Add(Me.TabSoftware)
        Me.TabControl1.Controls.Add(Me.TabUpdates)
        Me.TabControl1.Controls.Add(Me.TabEventLogs)
        Me.TabControl1.Controls.Add(Me.TabProcesses)
        Me.TabControl1.Controls.Add(Me.TabServices)
        Me.TabControl1.Controls.Add(Me.TabPerformance)
        Me.TabControl1.Controls.Add(Me.TabSystemRestore)
        Me.TabControl1.ItemSize = New System.Drawing.Size(61, 18)
        Me.TabControl1.Location = New System.Drawing.Point(9, 75)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(692, 501)
        Me.TabControl1.TabIndex = 4
        '
        'ButtonSetCredentials
        '
        Me.ButtonSetCredentials.FlatAppearance.BorderColor = System.Drawing.SystemColors.Control
        Me.ButtonSetCredentials.FlatAppearance.BorderSize = 0
        Me.ButtonSetCredentials.FlatAppearance.MouseDownBackColor = System.Drawing.SystemColors.Control
        Me.ButtonSetCredentials.FlatAppearance.MouseOverBackColor = System.Drawing.SystemColors.Control
        Me.ButtonSetCredentials.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonSetCredentials.ForeColor = System.Drawing.Color.RoyalBlue
        Me.ButtonSetCredentials.Location = New System.Drawing.Point(448, 36)
        Me.ButtonSetCredentials.Name = "ButtonSetCredentials"
        Me.ButtonSetCredentials.Size = New System.Drawing.Size(90, 23)
        Me.ButtonSetCredentials.TabIndex = 4
        Me.ButtonSetCredentials.Text = "Set &Credentials"
        Me.ButtonSetCredentials.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ButtonSetCredentials.UseVisualStyleBackColor = True
        Me.ButtonSetCredentials.Visible = False
        '
        'ButtonGo
        '
        Me.ButtonGo.Location = New System.Drawing.Point(249, 36)
        Me.ButtonGo.Name = "ButtonGo"
        Me.ButtonGo.Size = New System.Drawing.Size(34, 23)
        Me.ButtonGo.TabIndex = 2
        Me.ButtonGo.Text = "&Go"
        Me.ButtonGo.UseVisualStyleBackColor = True
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ActionsToolStripMenuItem, Me.ViewToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.MenuStrip1.Size = New System.Drawing.Size(713, 24)
        Me.MenuStrip1.TabIndex = 5
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ActionsToolStripMenuItem
        '
        Me.ActionsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewToolStripMenuItem, Me.RebootToolStripMenuItem, Me.ShutdownToolStripMenuItem, Me.LogOffUserToolStripMenuItem, Me.ExploreToolStripMenuItem, Me.RegeditToolStripMenuItem, Me.RemoteDesktopToolStripMenuItem, Me.RemoteAssistToolStripMenuItem, Me.ManageToolStripMenuItem})
        Me.ActionsToolStripMenuItem.Name = "ActionsToolStripMenuItem"
        Me.ActionsToolStripMenuItem.Size = New System.Drawing.Size(59, 20)
        Me.ActionsToolStripMenuItem.Text = "&Actions"
        '
        'NewToolStripMenuItem
        '
        Me.NewToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ProcessToolStripMenuItem, Me.ServiceToolStripMenuItem})
        Me.NewToolStripMenuItem.Name = "NewToolStripMenuItem"
        Me.NewToolStripMenuItem.Size = New System.Drawing.Size(173, 22)
        Me.NewToolStripMenuItem.Text = "&New"
        '
        'ProcessToolStripMenuItem
        '
        Me.ProcessToolStripMenuItem.Name = "ProcessToolStripMenuItem"
        Me.ProcessToolStripMenuItem.Size = New System.Drawing.Size(114, 22)
        Me.ProcessToolStripMenuItem.Text = "&Process"
        '
        'ServiceToolStripMenuItem
        '
        Me.ServiceToolStripMenuItem.Name = "ServiceToolStripMenuItem"
        Me.ServiceToolStripMenuItem.Size = New System.Drawing.Size(114, 22)
        Me.ServiceToolStripMenuItem.Text = "&Service"
        '
        'RebootToolStripMenuItem
        '
        Me.RebootToolStripMenuItem.Name = "RebootToolStripMenuItem"
        Me.RebootToolStripMenuItem.Size = New System.Drawing.Size(173, 22)
        Me.RebootToolStripMenuItem.Text = "Re&boot"
        '
        'ShutdownToolStripMenuItem
        '
        Me.ShutdownToolStripMenuItem.Name = "ShutdownToolStripMenuItem"
        Me.ShutdownToolStripMenuItem.Size = New System.Drawing.Size(173, 22)
        Me.ShutdownToolStripMenuItem.Text = "&Shutdown"
        '
        'LogOffUserToolStripMenuItem
        '
        Me.LogOffUserToolStripMenuItem.Name = "LogOffUserToolStripMenuItem"
        Me.LogOffUserToolStripMenuItem.Size = New System.Drawing.Size(173, 22)
        Me.LogOffUserToolStripMenuItem.Text = "&Log off User"
        '
        'ExploreToolStripMenuItem
        '
        Me.ExploreToolStripMenuItem.Name = "ExploreToolStripMenuItem"
        Me.ExploreToolStripMenuItem.Size = New System.Drawing.Size(173, 22)
        Me.ExploreToolStripMenuItem.Text = "&Explore"
        '
        'RegeditToolStripMenuItem
        '
        Me.RegeditToolStripMenuItem.Name = "RegeditToolStripMenuItem"
        Me.RegeditToolStripMenuItem.Size = New System.Drawing.Size(173, 22)
        Me.RegeditToolStripMenuItem.Text = "&Regedit"
        '
        'RemoteDesktopToolStripMenuItem
        '
        Me.RemoteDesktopToolStripMenuItem.Name = "RemoteDesktopToolStripMenuItem"
        Me.RemoteDesktopToolStripMenuItem.Size = New System.Drawing.Size(173, 22)
        Me.RemoteDesktopToolStripMenuItem.Text = "Re&mote Desktop"
        '
        'RemoteAssistToolStripMenuItem
        '
        Me.RemoteAssistToolStripMenuItem.Name = "RemoteAssistToolStripMenuItem"
        Me.RemoteAssistToolStripMenuItem.Size = New System.Drawing.Size(173, 22)
        Me.RemoteAssistToolStripMenuItem.Text = "Remote &Assistance"
        '
        'ManageToolStripMenuItem
        '
        Me.ManageToolStripMenuItem.Name = "ManageToolStripMenuItem"
        Me.ManageToolStripMenuItem.Size = New System.Drawing.Size(173, 22)
        Me.ManageToolStripMenuItem.Text = "Mana&ge"
        '
        'ViewToolStripMenuItem
        '
        Me.ViewToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RefreshToolStripMenuItem})
        Me.ViewToolStripMenuItem.Name = "ViewToolStripMenuItem"
        Me.ViewToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.ViewToolStripMenuItem.Text = "&View"
        '
        'RefreshToolStripMenuItem
        '
        Me.RefreshToolStripMenuItem.Name = "RefreshToolStripMenuItem"
        Me.RefreshToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F5
        Me.RefreshToolStripMenuItem.Size = New System.Drawing.Size(132, 22)
        Me.RefreshToolStripMenuItem.Text = "&Refresh"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutToolStripMenuItem, Me.CheckForUpdatesToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "&Help"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(171, 22)
        Me.AboutToolStripMenuItem.Text = "&About"
        '
        'CheckForUpdatesToolStripMenuItem
        '
        Me.CheckForUpdatesToolStripMenuItem.Name = "CheckForUpdatesToolStripMenuItem"
        Me.CheckForUpdatesToolStripMenuItem.Size = New System.Drawing.Size(171, 22)
        Me.CheckForUpdatesToolStripMenuItem.Text = "&Check for Updates"
        '
        'FormRemotePCAdmin
        '
        Me.AcceptButton = Me.ButtonGo
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(713, 588)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.ButtonGo)
        Me.Controls.Add(Me.TextBoxComputer)
        Me.Controls.Add(Me.lblHost)
        Me.Controls.Add(Me.ButtonSetCredentials)
        Me.Controls.Add(Me.CheckBoxAlternateCredentials)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MinimumSize = New System.Drawing.Size(550, 560)
        Me.Name = "FormRemotePCAdmin"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Remote PC Admin"
        Me.ContextMenuStripSystemRestore.ResumeLayout(False)
        Me.TabSoftware.ResumeLayout(False)
        Me.ContextMenuStripSoftware.ResumeLayout(False)
        Me.TabSystemRestore.ResumeLayout(False)
        Me.GroupBoxSystemRestore.ResumeLayout(False)
        Me.TabPerformance.ResumeLayout(False)
        Me.GroupBoxDisk3.ResumeLayout(False)
        Me.GroupBoxDisk1.ResumeLayout(False)
        Me.GroupBoxCPU.ResumeLayout(False)
        Me.GroupBoxMem.ResumeLayout(False)
        Me.GroupBoxDisk2.ResumeLayout(False)
        Me.TabServices.ResumeLayout(False)
        Me.ContextMenuStripServices.ResumeLayout(False)
        Me.TabProcesses.ResumeLayout(False)
        Me.ContextMenuStripProcesses.ResumeLayout(False)
        Me.TabEventLogs.ResumeLayout(False)
        Me.GroupBoxEventLogs.ResumeLayout(False)
        Me.ContextMenuStripCopy.ResumeLayout(False)
        Me.TabUpdates.ResumeLayout(False)
        Me.GroupBoxUpdates.ResumeLayout(False)
        Me.TabSystemInfo.ResumeLayout(False)
        Me.GroupBoxSystemInfo.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    '???
    'Me.Menu = Me.MainMenu1

    'TO DO add Service "Restart"?
    'TO DO - regedit alternate creds w/o runas?
    'TO DO - check return values from service start/stop/etc....events
    'TO DO - return status for service start/stop/etc
    'TO DO - Event log time... option to show remote, not local or gmt... find bias??

    Public strUsername As String = ""
    Public strPassword As String = ""

    Private strUpdateCheckURL As String = "https://raw.githubusercontent.com/joeostrander/Remote-PC-Admin/master/AssemblyInfo.vb"
    Private strUpdateDownloadURL As String = "https://github.com/joeostrander/Remote-PC-Admin/raw/master/bin/Release/Remote%20PC%20Admin.exe"


    Private PreviousTab As String = "TabSystemInfo"

    Public strProcessCommand As String
    Public strProcessWorkingDir As String


    Friend ListRecentSearches As New List(Of String)

    Structure MyProcess
        Public ID As String
        Public Name As String
        Public Owner As String
        Public ExecutablePath As String
        Public CommandLine As String
        Public CreationDate As String
        Public MemUsage As String
        Public CPU As String
    End Structure

    Private Sub menuItemEndProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim c As Integer
        c = lstProcesses.SelectedIndices(0)

        Dim strProcess As String
        Dim intProcessID As Integer
        Dim objProcess As Object
        Dim strProcMsg As String
        Dim AskYesNo As MsgBoxResult
        strProcess = lstProcesses.Items(c).SubItems(0).Text
        intProcessID = CInt(lstProcesses.Items(c).SubItems(1).Text)


        Dim strComputer As String = TextBoxComputer.Text
        Dim oConn As ConnectionOptions = New ConnectionOptions
        If CheckBoxAlternateCredentials.Checked Then
            oConn.Username = strUsername
            oConn.Password = strPassword
        End If
        Dim oMs As ManagementScope = New ManagementScope(strComputer, oConn)

        Dim oQuery As SelectQuery = New SelectQuery(String.Format("select * from Win32_Process where ProcessId = {0}", intProcessID))
        Dim oSearcher As ManagementObjectSearcher = New ManagementObjectSearcher(oMs, oQuery)
        Dim colProcesses As ManagementObjectCollection = oSearcher.Get()


        strProcMsg = vbCrLf & vbCrLf & "Name:  " & strProcess & vbCrLf & "PID:  " & intProcessID

        For Each objProcess In colProcesses
            AskYesNo = MsgBox("Are you sure you wish to end this process?" & strProcMsg, MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation, Application.ProductName)
            Select Case AskYesNo
                Case MsgBoxResult.Yes
                    objProcess.Terminate(intProcessID)
                    MsgBox("Termination request sent for:  " & strProcMsg, MsgBoxStyle.Information, Application.ProductName)
                    ViewProcesses()
                Case Else
                    MsgBox("Action Cancelled.", MsgBoxStyle.Information, Application.ProductName)
            End Select

        Next

    End Sub

    Private Function GetLoggedOnUserByProcess() As String


        Dim strComputer As String = TextBoxComputer.Text

        Dim colProcesses As ManagementObjectCollection = GetCollection("select * from Win32_Process Where Name = 'explorer.exe'")
        If IsNothing(colProcesses) Then
            GetLoggedOnUserByProcess = Nothing
            Exit Function
        End If

        Dim strNameOfUser As String = ""
        Dim strDomain As String = ""
        Dim strUser As String = "Nobody"
        Try
            For Each objProcess As ManagementObject In colProcesses
                Dim argList() As String = {"", ""}
                Dim retVal As Integer = objProcess.InvokeMethod("GetOwner", argList)
                If retVal = 0 Then
                    strNameOfUser = argList(0)
                    strDomain = argList(1)
                    If InStr(strDomain, "NT AUTHORITY") = False Then
                        If strDomain = "" Then
                            strUser = strNameOfUser
                        Else
                            If strNameOfUser <> "" Then
                                strUser = strDomain & "\" & strNameOfUser
                            End If
                        End If
                        Exit For
                    End If
                End If

            Next
        Catch ex As Exception
            'Do Nothing
            strUser = "???"
        End Try

        GetLoggedOnUserByProcess = strUser
    End Function

    Private Sub CheckBoxAlternateCredentials_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxAlternateCredentials.CheckedChanged
        If CheckBoxAlternateCredentials.CheckState = CheckState.Checked Then
            ButtonSetCredentials.Visible = True
        End If

        If CheckBoxAlternateCredentials.CheckState = CheckState.Unchecked Then
            ButtonSetCredentials.Visible = False
        End If
    End Sub


    Private Sub ViewProcesses()
        lstProcesses.Items.Clear()

        Dim listProcs As List(Of MyProcess) = New List(Of MyProcess)

        Dim PCap As String
        Dim PID As Integer
        Dim ExecutablePath As String
        Dim CommandLine As String

        Dim strOwner As String = ""


        Dim MemUsage As String
        Dim TotalTime As Integer = 0
        Dim PercentTime As Integer = 0

        Dim colProcesses As ManagementObjectCollection = GetCollection("select * from Win32_Process")
        If IsNothing(colProcesses) Then Exit Sub


        Dim colEnum As ManagementObjectCollection = Nothing

        If ShowCPUToolStripMenuItem.Checked Then
            colEnum = GetCollection("Select * From Win32_PerfFormattedData_PerfProc_Process")
            'If IsNothing(colEnum) Then Exit Sub
        End If



        'Create Columns

        'Loop through processes
        For Each objProcess As ManagementObject In colProcesses
            Dim strCreationDate As String = ""

            Dim dmtfDate As String = objProcess.GetPropertyValue("CreationDate")
            If dmtfDate <> "" Then
                Dim dtmCreationDate As DateTime = ManagementDateTimeConverter.ToDateTime(dmtfDate)
                strCreationDate = dtmCreationDate.ToString()
            End If



            'Image Name (process name)
            PCap = objProcess.GetPropertyValue("Name")


            'Process ID
            PID = objProcess.GetPropertyValue("ProcessID")

            'User Name
            Dim methodArgs() As String = {"", ""}

            Dim strNameOfUser As String = ""
            Dim strDomain As String = ""

            Try
                objProcess.InvokeMethod("GetOwner", methodArgs)
                strNameOfUser = methodArgs(0)
                strDomain = methodArgs(1)
            Catch ex As Exception
                'Do Nothing
            End Try


            If strNameOfUser <> "" Then
                If strDomain = "" Then
                    strOwner = strNameOfUser
                Else
                    If strNameOfUser <> "" Then
                        strOwner = strDomain & "\" & strNameOfUser
                    End If

                End If
            Else
                strOwner = ""
            End If

            If ShowCPUToolStripMenuItem.Checked Then
                If Not IsNothing(colEnum) Then
                    For Each objInstance As ManagementObject In colEnum
                        If CLng(objInstance.GetPropertyValue("IdProcess")) = CLng(objProcess.GetPropertyValue("Handle")) Then

                            PercentTime = objInstance.GetPropertyValue("PercentProcessorTime")
                            Exit For
                        End If


                    Next
                End If


                'CPU Usage

            End If


            'Memory Usage
            MemUsage = FormatNumber(objProcess.GetPropertyValue("WorkingSetSize") / 1024 / 1024, 0) & " MB"

            ''Executable Path
            ExecutablePath = objProcess.GetPropertyValue("ExecutablePath")
            CommandLine = objProcess.GetPropertyValue("CommandLine")


            If IsNumeric(PID) = False Then PID = " "
            If strOwner = "" Then strOwner = " "
            If MemUsage = "" Then MemUsage = " "
            If CommandLine = "" Then CommandLine = " "
            If ExecutablePath = "" Then ExecutablePath = " "

            Dim strCPU As String
            Dim idx As Integer = GetColumnIndex(lstProcesses, "CPU")
            If ShowCPUToolStripMenuItem.Checked AndAlso (Not IsNothing(colEnum)) Then
                strCPU = PercentTime.ToString
                lstProcesses.Columns(idx).Width = 60
            Else
                strCPU = ""
                lstProcesses.Columns(idx).Width = 0
            End If


            If PCap <> "System Idle Process" Then
                Dim proc As New MyProcess
                proc.ID = PID
                proc.Name = PCap
                proc.Owner = strOwner
                proc.CreationDate = strCreationDate
                proc.MemUsage = MemUsage
                proc.CPU = strCPU
                proc.CommandLine = CommandLine
                proc.ExecutablePath = ExecutablePath

                listProcs.Add(proc)
            End If

            strNameOfUser = ""
            strDomain = ""

        Next

        For Each myProc As MyProcess In listProcs
            Dim objListItem As ListViewItem
            objListItem = lstProcesses.Items.Add(myProc.Name, 0)
            objListItem.SubItems.Add(myProc.ID)
            objListItem.SubItems.Add(myProc.Owner)
            objListItem.SubItems.Add(myProc.CreationDate)
            objListItem.SubItems.Add(myProc.MemUsage)
            objListItem.SubItems.Add(myProc.CPU)
            objListItem.SubItems.Add(myProc.CommandLine)
            objListItem.SubItems.Add(myProc.ExecutablePath)
        Next




    End Sub

    Private Function GetColumnIndex(ByVal lv As ListView, ByVal strColumnTitle As String) As Integer
        For Each col As ColumnHeader In lv.Columns
            If col.Text = strColumnTitle Then
                Return col.Index
            End If
        Next
    End Function


    Private Sub ViewServices()

        Dim objListItem As ListViewItem
        Dim strSrvName As String = ""
        Dim strSrvDisplayName As String = ""
        Dim srvStatus As String = ""
        Dim srvStartupType As String = ""
        Dim strSrvDescription As String = ""


        lstServices.Items.Clear()


        Dim colServices As ManagementObjectCollection = GetCollection("select * from Win32_Service")
        If IsNothing(colServices) Then Exit Sub

        'Loop through services
        For Each objService As ManagementObject In colServices
            'Display Name
            If objService("DisplayName") = "" Then
                strSrvDisplayName = "Unknown"
            Else
                strSrvDisplayName = objService.GetPropertyValue("DisplayName")
            End If

            'Name
            strSrvName = objService.GetPropertyValue("Name")
            If strSrvName = "" Then strSrvName = "Unknown"

            'Status
            srvStatus = objService.GetPropertyValue("State")
            If srvStatus = "" Then srvStatus = "Unknown"

            'Startup Type
            srvStartupType = objService.GetPropertyValue("StartMode")
            If srvStartupType = "" Then srvStartupType = "Unknown"


            'Description
            strSrvDescription = objService.GetPropertyValue("Description")
            If strSrvDescription = "" Then strSrvDescription = "Unknown"



            objListItem = lstServices.Items.Add(strSrvDisplayName, 0)
            objListItem.SubItems.Add(srvStatus)
            objListItem.SubItems.Add(srvStartupType)
            objListItem.SubItems.Add(strSrvDescription)
            objListItem.SubItems.Add(strSrvName)
        Next
    End Sub




    Private Sub TabCheck(ByVal strTab As String)
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Select Case strTab
            Case "TabSystemInfo"
                ViewSystemInfo()
            Case "TabSoftware"
                ViewSoftware()
            Case "TabUpdates"
                ViewHotfixes()
            Case "TabProcesses"
                ViewProcesses()
            Case "TabServices"
                ViewServices()
            Case "TabPerformance"
                Uptime()
                MemUsage()
                CpuUsage()
                DiskUsage()
            Case "TabEventLogs"
                If RadioButtonApplication.Checked = True Then
                    GetLog("Application")
                ElseIf RadioButtonSecurity.Checked = True Then
                    GetLog("Security")
                ElseIf RadioButtonSystem.Checked = True Then
                    GetLog("System")
                End If
            Case "TabSystemRestore"
                ListViewSystemRestore.Items.Clear()
                SystemRestoreStatus()
                ListRestorePoints()
            Case Else
        End Select
    End Sub

    Private Sub menuItemProcViewPath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error Resume Next
        Dim strProcPath As String
        Dim i As Integer
        i = lstProcesses.SelectedIndices(0)
        strProcPath = lstProcesses.Items(i).SubItems(5).Text
        MsgBox(strProcPath, MsgBoxStyle.Information)
    End Sub


    Private Sub ServiceChange(ByVal strStatus As String)

        Dim i As Integer
        i = lstServices.SelectedIndices(0)


        Dim strSrvDisplayName As String
        Dim srvStatus As String
        Dim strSrvName As String
        Dim srvMsg As String
        Dim AskYesNo As MsgBoxResult
        strSrvDisplayName = lstServices.Items(i).SubItems(0).Text
        srvStatus = lstServices.Items(i).SubItems(1).Text
        strSrvName = lstServices.Items(i).SubItems(4).Text

        'Dim strComputer As String = TextBoxComputer.Text
        'Dim oConn As ConnectionOptions = New ConnectionOptions
        'If CheckBoxAlternateCredentials.Checked Then
        '    oConn.Username = strUsername
        '    oConn.Password = strPassword
        'End If
        'Dim oMs As ManagementScope = New ManagementScope(strComputer, oConn)

        'Dim oQuery As SelectQuery = New SelectQuery(String.Format("Select * from Win32_Service where Name='" & strSrvName & "'"))
        'Dim oSearcher As ManagementObjectSearcher = New ManagementObjectSearcher(oMs, oQuery)
        Dim colServices As ManagementObjectCollection = GetObjectCollection("Select * from Win32_Service where Name='" & strSrvName & "'")

        Select Case srvStatus
            Case "Running"
                If strStatus = "Start" Then
                    MsgBox(strSrvDisplayName & " is already started.", MsgBoxStyle.Information)
                    Exit Sub
                End If
            Case "Stopped"
                If strStatus = "Stop" Then
                    MsgBox(strSrvDisplayName & " is already stopped.", MsgBoxStyle.Information)
                    Exit Sub
                End If
        End Select



        srvMsg = vbCrLf & vbCrLf & strSrvDisplayName & " (" & strSrvName & ")"
        Dim intResult As Integer
        Dim msg As String = ""
        Dim style As MsgBoxStyle = MsgBoxStyle.Exclamation
        For Each objService As ManagementObject In colServices
            AskYesNo = MsgBox("Are you sure you wish to " & strStatus & " this service?" & srvMsg, MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation)
            Select Case AskYesNo
                Case MsgBoxResult.Yes
                    Select Case strStatus
                        Case "Stop"
                            intResult = DirectCast(objService.InvokeMethod("StopService", Nothing), UInteger)
                        Case "Start"
                            intResult = DirectCast(objService.InvokeMethod("StartService", Nothing), UInteger)
                        Case "Restart"
                            If srvStatus = "Running" Then
                                intResult = DirectCast(objService.InvokeMethod("StopService", Nothing), UInteger)
                            End If
                            intResult = DirectCast(objService.InvokeMethod("StartService", Nothing), UInteger)
                        Case "Remove"
                            intResult = DirectCast(objService.InvokeMethod("StopService", Nothing), UInteger)
                            intResult = DirectCast(objService.InvokeMethod("Delete", Nothing), UInteger)
                    End Select

                    Select Case intResult
                        Case 0
                            msg = "Operation submitted successfully."
                            style = MsgBoxStyle.Information
                        Case 1
                            msg = "Not Supported"
                        Case 2
                            msg = "Access Denied"
                        Case 3
                            msg = "Dependent Services Running"
                        Case 4
                            msg = "Invalid Service Control"
                        Case 5
                            msg = "Service Cannot Accept Control"
                        Case 6
                            msg = "Service Not Active"
                        Case 7
                            msg = "Service Request timeout"
                        Case 8
                            msg = "Unknown Failure"
                        Case 9
                            msg = "Path Not Found"
                        Case 10
                            msg = "Service Already Stopped"
                        Case 11
                            msg = "Service Database Locked"
                        Case 12
                            msg = "Service Dependency Deleted"
                        Case 13
                            msg = "Service Dependency Failure"
                        Case 14
                            msg = "Service Disabled"
                        Case 15
                            msg = "Service Logon Failed"
                        Case 16
                            msg = "Service Marked For Deletion"
                        Case 17
                            msg = "Service No Thread"
                        Case 18
                            msg = "Status Circular Dependency"
                        Case 19
                            msg = "Status Duplicate Name"
                        Case 20
                            msg = "Status - Invalid Name"
                        Case 21
                            msg = "Status - Invalid Parameter"
                        Case 22
                            msg = "Status - Invalid Service Account"
                        Case 23
                            msg = "Status - Service Exists"
                        Case 24
                            msg = "Service Already Paused"
                        Case Else
                            msg = "Unknown Failure"
                    End Select

                    MsgBox("Request sent to " & strStatus & ":" & srvMsg & vbCrLf & vbCrLf & msg, style)
                    ViewServices()
                Case Else
                    MsgBox("Action Canceled", MsgBoxStyle.Information)
            End Select
        Next
    End Sub


    Private Sub RebootUsingShutdownExe(ByVal boolRestart As Boolean)


        Dim strComputer As String = TextBoxComputer.Text
        If strComputer = "" Then strComputer = "."

        Dim strMsg As String
        Dim strCommand As String
        If boolRestart = True Then
            strMsg = "Are you sure you wish to send a reboot request to " & strComputer & "?"
            strCommand = "cmd /c shutdown -f -r -t 0"
        Else
            strMsg = "Are you sure you wish to send a shutdown request to " & strComputer & "?"
            strCommand = "cmd /c shutdown -f -s -t 0"
        End If


        Dim AskYesNo As MsgBoxResult
        AskYesNo = MsgBox(strMsg, MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation)
        If AskYesNo <> MsgBoxResult.Yes Then
            MsgBox("Action Canceled", MsgBoxStyle.Information)
            Exit Sub
        End If

        ExecuteProcess(strCommand, "")



    End Sub

    Private Sub CreateProcess()

        If FormProcessAdd.ShowDialog <> Windows.Forms.DialogResult.OK Then Exit Sub

        ExecuteProcess(strProcessCommand, strProcessWorkingDir)


    End Sub

    Private Sub ExecuteProcess(ByVal command As String, ByVal workingdir As String)

        Dim strComputer As String = TextBoxComputer.Text
        If strComputer = "" Then strComputer = "."

        Dim oConn As ConnectionOptions = New ConnectionOptions
        If CheckBoxAlternateCredentials.Checked And strComputer <> "." Then
            If strUsername = "" Or strPassword = "" Then
                MsgBox("Please set the credentials.", MsgBoxStyle.Information, "Credentials not set.")
                Exit Sub
            End If
            oConn.Username = strUsername
            oConn.Password = strPassword
        End If


        Try
            Dim mp As ManagementPath = New ManagementPath("\\" & strComputer & "\root\cimv2")
            Dim oMs As ManagementScope = New ManagementScope(mp, oConn)

            Dim path As ManagementPath = New ManagementPath("Win32_Process")
            Dim processClass As ManagementClass = New ManagementClass(oMs, path, Nothing)

            Dim createParams As ManagementBaseObject = processClass.GetMethodParameters("Create")
            'createParams("CommandLine") = strProcess
            createParams.SetPropertyValue("CommandLine", command)
            If workingdir <> "" Then
                createParams("CurrentDirectory") = workingdir
            End If

            'createParams("ProcessStartupInformation") = ...

            'Dim instance As New ManagementOperationObserver()
            'Dim completionHandlerObj As New completionHandler.MyHandler()
            ''Dim handler As CompletedEventHandler
            'AddHandler instance.ObjectReady, AddressOf ProcessCreated

            Dim ret As ManagementBaseObject = processClass.InvokeMethod("Create", createParams, Nothing)
            Dim msg As String
            Dim intReturn As Integer = DirectCast(ret("ReturnValue"), UInteger)

            Dim style As MsgBoxStyle = MsgBoxStyle.Exclamation

            Select Case intReturn
                Case 0
                    msg = "Process created successfully." & vbCrLf & vbCrLf & "Process ID:  " & ret("ProcessID")
                    style = MsgBoxStyle.Information
                Case 2
                    msg = "Access Denied"
                Case 3
                    msg = "Insufficient Privilege"
                Case 8
                    msg = "Unknown failure"
                Case 9
                    msg = "Path Not Found"
                Case 21
                    msg = "Invalid Parameter"
                Case Else
                    msg = "Failed to create process"
            End Select

            MsgBox(msg, style)

        Catch unauth As System.UnauthorizedAccessException
            MsgBox("Access is denied.", MsgBoxStyle.Exclamation, Application.ProductName)
            Return
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, Application.ProductName)
            Return
        End Try
    End Sub


    'Private Sub ProcessCreated(ByVal sender As Object, ByVal e As ObjectReadyEventArgs)
    '    MsgBox("Process created.  PID=" & e.NewObject("ProcessID").ToString & vbCrLf & "ReturnValue=  " & e.NewObject("ReturnValue").ToString, MsgBoxStyle.Information)

    'End Sub

    Private Sub ViewSystemInfo()

        RichTextBoxSystemInfo.Clear()

        Dim ChassisArray() As String = {"Unknown", "Other", "Unknown", "Desktop", "Low Profile Desktop", "Pizza Box", "Mini Tower", "Tower", "Portable", "Laptop", "Notebook", "Handheld", "Docking Station", "All in One", "Sub Notebook", "Space-Saving", "Lunch Box", "Main System Chassis", "Expansion Chassis", "SubChassis", "Bus Expansion Chassis", "Peripheral Chassis", "Storage Chassis", "Rack Mount Chassis", "Sealed-Case PC"}

        Dim colCompSysProd, colCompSys, colOpSys, colProc, colLogDisk, colAdapters, colChassis As ManagementObjectCollection

        colCompSysProd = GetCollection("Select * from Win32_ComputerSystemProduct")
        If IsNothing(colCompSysProd) Then Exit Sub

        colCompSys = GetCollection("Select * from Win32_ComputerSystem")
        If IsNothing(colCompSys) Then Exit Sub

        colOpSys = GetCollection("Select * from Win32_OperatingSystem")
        If IsNothing(colOpSys) Then Exit Sub

        colProc = GetCollection("Select * from Win32_Processor")
        If IsNothing(colProc) Then Exit Sub

        colLogDisk = GetCollection("Select * from Win32_LogicalDisk Where DriveType =3")
        If IsNothing(colLogDisk) Then Exit Sub

        colAdapters = GetCollection("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=TRUE")
        If IsNothing(colAdapters) Then Exit Sub

        colChassis = GetCollection("Select * from Win32_SystemEnclosure")
        If IsNothing(colChassis) Then Exit Sub

        Dim strLoggedOnUser As String
        Dim mem As Long
        '---------------COMPUTER SECTION---------------
        RichTextBoxSystemInfo.AppendText("-------------------------------------" & vbCrLf)
        RichTextBoxSystemInfo.SelectionFont = New Font(RichTextBoxSystemInfo.SelectionFont, FontStyle.Bold)
        RichTextBoxSystemInfo.AppendText("Computer:" & vbCrLf)
        RichTextBoxSystemInfo.AppendText("-------------------------------------" & vbCrLf)
        For Each objCompSys As ManagementObject In colCompSys
            If objCompSys("Username") = "" Then
                strLoggedOnUser = GetLoggedOnUserByProcess()
            Else
                strLoggedOnUser = objCompSys("Username")
            End If

            RichTextBoxSystemInfo.AppendText(vbTab & strLoggedOnUser & " is currently logged on." & vbCrLf)
            RichTextBoxSystemInfo.AppendText(vbTab & "System:  " & objCompSys("Name") & vbCrLf)
            RichTextBoxSystemInfo.AppendText(vbTab & "Manufacturer:  " & objCompSys("Manufacturer") & vbCrLf)
            RichTextBoxSystemInfo.AppendText(vbTab & "Model:  " & objCompSys("Model") & vbCrLf)
            mem = objCompSys("TotalPhysicalMemory")
        Next
        'Try
        For Each objChassis As ManagementObject In colChassis
            Try
                For Each objChassisItem As Object In objChassis("ChassisTypes")
                    RichTextBoxSystemInfo.AppendText(vbTab & "Type:  " & ChassisArray(objChassisItem) & vbCrLf)
                Next
            Catch ex As Exception
                Debug.WriteLine(ex.Message)
            End Try
        Next
        'Catch ex As Exception
        '    Debug.WriteLine(ex.Message)
        'End Try

        For Each objCompSysProd As ManagementObject In colCompSysProd
            RichTextBoxSystemInfo.AppendText(vbTab & "Identifying Number:  " & objCompSysProd("IdentifyingNumber") & vbCrLf)
            RichTextBoxSystemInfo.AppendText(vbTab & "UUID:  " & objCompSysProd("UUID") & vbCrLf)
        Next
        '---------------OPERATING SYSTEM SECTION---------------

        RichTextBoxSystemInfo.AppendText("-------------------------------------" & vbCrLf)
        RichTextBoxSystemInfo.SelectionFont = New Font(RichTextBoxSystemInfo.SelectionFont, FontStyle.Bold)
        RichTextBoxSystemInfo.AppendText("Operating System:" & vbCrLf)
        RichTextBoxSystemInfo.AppendText("-------------------------------------" & vbCrLf)
        Dim objSWbemDateTime As Object = CreateObject("WbemScripting.SWbemDateTime")
        For Each objOpSys As ManagementObject In colOpSys
            RichTextBoxSystemInfo.AppendText(vbTab & "Operating System:  " & objOpSys("Caption") & vbCrLf)
            RichTextBoxSystemInfo.AppendText(vbTab & "Version: " & objOpSys("Version") & vbCrLf)
            RichTextBoxSystemInfo.AppendText(vbTab & "Service Pack: " & objOpSys("ServicePackMajorVersion") & "." & objOpSys("ServicePackMinorVersion") & vbCrLf)
            objSWbemDateTime.value = objOpSys("LastBootUpTime")
            Dim dtmLastBootupTime As Date = WMIDateStringToDate(objOpSys("LastBootUpTime"))

            Dim tsSystemUptime As TimeSpan
            tsSystemUptime = Now.Subtract(dtmLastBootupTime)

            RichTextBoxSystemInfo.AppendText(vbTab & "Computer has been up for " & FormatNumber(tsSystemUptime.TotalHours / 24, 1) & " days" & vbCrLf)
        Next


        '---------------PHYSICAL MEMORY SECTION---------------
        RichTextBoxSystemInfo.AppendText("-------------------------------------" & vbCrLf)
        RichTextBoxSystemInfo.SelectionFont = New Font(RichTextBoxSystemInfo.SelectionFont, FontStyle.Bold)
        RichTextBoxSystemInfo.AppendText("Physical Memory:" & vbCrLf)
        RichTextBoxSystemInfo.AppendText("-------------------------------------" & vbCrLf)
        ' **"mem" is from the computer section
        'RichTextBoxSystemInfo.AppendText(vbTab & "Total RAM:  " & FormatNumber(mem / 1048576, 0) & "MB" & vbCrLf)
        RichTextBoxSystemInfo.AppendText(vbTab & "Total RAM:  " & StrFormatByteSize(mem) & vbCrLf)

        '---------------PROCESSOR SECTION---------------
        RichTextBoxSystemInfo.AppendText("-------------------------------------" & vbCrLf)
        RichTextBoxSystemInfo.SelectionFont = New Font(RichTextBoxSystemInfo.SelectionFont, FontStyle.Bold)
        RichTextBoxSystemInfo.AppendText("Processor:" & vbCrLf)
        RichTextBoxSystemInfo.AppendText("-------------------------------------" & vbCrLf)
        For Each objProc As ManagementObject In colProc
            RichTextBoxSystemInfo.AppendText(vbTab & "Manufacturer:  " & objProc("Manufacturer") & vbCrLf)
            RichTextBoxSystemInfo.AppendText(vbTab & "Maximum Clock Speed:  " & objProc("MaxClockSpeed") & vbCrLf)
            RichTextBoxSystemInfo.AppendText(vbTab & "Name:  " & objProc("Name") & vbCrLf)
        Next
        '---------------LOGICAL DISK SECTION---------------
        RichTextBoxSystemInfo.AppendText("-------------------------------------" & vbCrLf)
        RichTextBoxSystemInfo.SelectionFont = New Font(RichTextBoxSystemInfo.SelectionFont, FontStyle.Bold)
        RichTextBoxSystemInfo.AppendText("Logical Disks:" & vbCrLf)
        RichTextBoxSystemInfo.AppendText("-------------------------------------" & vbCrLf)
        For Each objLogDisk As ManagementObject In colLogDisk
            RichTextBoxSystemInfo.AppendText(vbTab & objLogDisk("DeviceID") & vbCrLf)
            'RichTextBoxSystemInfo.AppendText(vbTab & vbTab & FormatNumber(objLogDisk.Size / 1073741824, 3) & " GB Total" & vbCrLf)
            RichTextBoxSystemInfo.AppendText(vbTab & vbTab & StrFormatByteSize(objLogDisk("Size")) & " Total" & vbCrLf)
            'RichTextBoxSystemInfo.AppendText(vbTab & vbTab & FormatNumber(objLogDisk.FreeSpace / 1073741824, 3) & " GB Remaining" & vbCrLf & vbCrLf)
            RichTextBoxSystemInfo.AppendText(vbTab & vbTab & StrFormatByteSize(objLogDisk("FreeSpace")) & " Remaining" & vbCrLf & vbCrLf)
        Next

        '---------------NETWORK SECTION---------------
        RichTextBoxSystemInfo.AppendText("-------------------------------------" & vbCrLf)
        RichTextBoxSystemInfo.SelectionFont = New Font(RichTextBoxSystemInfo.SelectionFont, FontStyle.Bold)
        RichTextBoxSystemInfo.AppendText("Network:" & vbCrLf)
        RichTextBoxSystemInfo.AppendText("-------------------------------------" & vbCrLf)
        Dim i As Integer
        For Each objAdapter As ManagementObject In colAdapters
            If Not IsNothing(objAdapter("IPAddress")) Then
                For i = LBound(objAdapter("IPAddress")) To UBound(objAdapter("IPAddress"))
                    RichTextBoxSystemInfo.AppendText(vbTab & objAdapter("Description") & vbCrLf)
                    RichTextBoxSystemInfo.AppendText(vbTab & "Mac Address:  " & objAdapter("MACAddress") & vbCrLf)
                    RichTextBoxSystemInfo.AppendText(vbTab & "IP Address:  " & objAdapter("IPAddress")(i) & vbCrLf)
                Next
            End If
        Next

    End Sub



    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub


    Private Sub Win32Shutdown(ByVal action As String)
        On Error Resume Next

        Dim strAction As String = ""

        Select Case strAction
            Case "Reboot"
                strAction = "2"
            Case "Shutdown"
                strAction = "1"
            Case "Logoff"
                strAction = "0"
        End Select


        Dim strComputer As String
        If TextBoxComputer.Text = "" Then
            strComputer = "LOCALHOST"
        Else
            strComputer = TextBoxComputer.Text
        End If

        Dim AskYesNo As Microsoft.VisualBasic.MsgBoxResult
        AskYesNo = MsgBox("Are you sure you wish to send a " & action & " request to " & strComputer & "?", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation)
        If AskYesNo <> MsgBoxResult.Yes Then
            MsgBox("Action Canceled", MsgBoxStyle.Information)
            Exit Sub
        End If

        Dim objWbemLocator As Object
        Dim objServices As Object
        Const aCall As Integer = 3
        Const iImpersonate As Integer = 3
        Dim objWbem As Object
        Dim objService As Object
        Dim strQuery As String
        Dim colOpSys As Object
        Dim objSystem As Object = Nothing



        ' Create the Locator object
        objWbemLocator = CreateObject("WbemScripting.SWbemLocator")

        If CheckBoxAlternateCredentials.Checked Then
            If strUsername = "" Or strPassword = "" Then
                MsgBox("Please set the credentials.", MsgBoxStyle.Information, "Credentials not set.")
                Exit Sub
            End If
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\CIMV2", strUsername, strPassword)
        Else
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\CIMV2")
        End If


        If Err.Number <> 0 Then
            Dim strMsg As String
            strMsg = "Could not connect to:  " & strComputer & vbCrLf & vbCrLf & "Error Number:  " & Err.Number & vbCrLf & "Description:  " & Err.Description
            MsgBox(strMsg, MsgBoxStyle.Exclamation)
            Return
        End If

        TextBoxComputer.Text = UCase(strComputer)
        objServices.Security_.AuthenticationLevel = aCall
        objServices.Security_.ImpersonationLevel = iImpersonate
        'objServices.Security_.Privileges.Add(18)

        objWbem = objServices

        strQuery = "Select Name From Win32_OperatingSystem"
        colOpSys = objWbem.ExecQuery(strQuery)

        For Each objSystem In colOpSys
            objSystem.Win32Shutdown(strAction)
        Next
        MsgBox(action & " request sent for " & TextBoxComputer.Text, MsgBoxStyle.Information)
        objSystem.Dispose()
    End Sub


    Private Function GetScope() As ManagementScope
        Dim strComputer As String = TextBoxComputer.Text
        If strComputer = "" Then strComputer = "."

        Dim oConn As ConnectionOptions = New ConnectionOptions

        'Testing!!!
        'oConn.Username = ".\administrator"
        'oConn.Password = "??????"
        'oConn.Authentication = AuthenticationLevel.PacketPrivacy
        'oConn.Impersonation = ImpersonationLevel.Impersonate
        'oConn.EnablePrivileges = True

        ''



        If CheckBoxAlternateCredentials.Checked And strComputer <> "." Then
            If strUsername = "" Or strPassword = "" Then
                MsgBox("Please set the credentials.", MsgBoxStyle.Information, "Credentials not set.")
                Return Nothing
            End If
            oConn.Username = strUsername
            oConn.Password = strPassword
        End If
        ''


        Dim oMs As New ManagementScope

        Try
            'Dim mp As ManagementPath = New ManagementPath("\\" & strComputer & "\root\cimv2")
            'Dim mp As ManagementPath = New ManagementPath("\\" & strComputer & "\root\StdRegProv")
            Dim mp As ManagementPath = New ManagementPath("\\" & strComputer & "\root\default")
            oMs = New ManagementScope(mp, oConn)
        Catch ex As Exception
            'Do Nothing
        End Try


        'oMs.Options.Authentication = AuthenticationLevel.Default
        'oMs.Options.Impersonation = ImpersonationLevel.Impersonate
        'oMs.Options.EnablePrivileges = True

        Return oMs

    End Function


    Private Sub ViewSoftware()
        lstSoftware.Clear()
        lstSoftware.Columns.Add("Display Name", 500, HorizontalAlignment.Left)
        lstSoftware.Columns.Add("Type", 40, HorizontalAlignment.Left)
        lstSoftware.Columns.Add("Uninstall String", 0, HorizontalAlignment.Left)

        'lstSoftware.Items.Clear()
        Dim bool64bit As Boolean = Is64BitRemote()
        Dim strPrefix As String = ""

        If bool64bit Then strPrefix = "64bit - "
        ListSofwareFromRegKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\", strPrefix)
        If bool64bit Then
            strPrefix = "32bit - "
            ListSofwareFromRegKey("SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\", strPrefix)
        End If

    End Sub

    Private Sub ListSofwareFromRegKey(ByVal strKeyPath As String, ByVal strPrefix As String)

        Dim objListItem As ListViewItem

        Dim strValue As String

        Const HKEY_LOCAL_MACHINE As UInteger = &H80000002UI

        Dim strDisplayName As String = "DisplayName"
        Dim strUninstallString As String = "UninstallString"
        Dim strUninstallCMD As String



        Dim subkey As String
        Dim AllSoftware As String = ""


        Dim oMs As ManagementScope = GetScope()
        If oMs Is Nothing Then Exit Sub
        'Dim oMs As ManagementScope = GetScope("cimv2")
        'Dim oMs As ManagementScope = GetScope("StdRegProv")

        Dim path As ManagementPath = New ManagementPath("StdRegProv")
        Dim mc As ManagementClass = New ManagementClass(oMs, path, Nothing)



        Dim inParams As ManagementBaseObject = mc.GetMethodParameters("EnumKey")
        inParams("hDefKey") = HKEY_LOCAL_MACHINE
        'inParams.SetPropertyValue("hDefKey", HKEY_LOCAL_MACHINE)
        inParams.SetPropertyValue("sSubKeyName", strKeyPath)

        ''''
        Dim outParams As ManagementBaseObject = mc.InvokeMethod("EnumKey", inParams, Nothing)

        inParams = mc.GetMethodParameters("GetStringValue")
        inParams.SetPropertyValue("hDefKey", HKEY_LOCAL_MACHINE)

        Dim arrSubKeys As String() = DirectCast(outParams("sNames"), String())


        'Dim Total As Integer

        For Each subkey In arrSubKeys
            inParams.SetPropertyValue("sSubKeyName", strKeyPath & subkey)
            inParams.SetPropertyValue("sValueName", strDisplayName)
            outParams = mc.InvokeMethod("GetStringValue", inParams, Nothing)

            strValue = DirectCast(outParams("sValue"), String)

            If Not String.IsNullOrEmpty(strValue) Then
                If InStr(strValue, " (KB") = False Then
                    If InStr(AllSoftware, strValue & vbCrLf) = False Then
                        'Total = Total + 1
                        inParams.SetPropertyValue("sSubKeyName", strKeyPath & subkey)
                        inParams.SetPropertyValue("sValueName", strUninstallString)
                        outParams = mc.InvokeMethod("GetStringValue", inParams, Nothing)

                        strUninstallCMD = DirectCast(outParams("sValue"), String)

                        objListItem = lstSoftware.Items.Add(strPrefix & strValue)
                        If InStr(LCase(strUninstallCMD), "msiexec") Then
                            objListItem.SubItems.Add("MSI")
                        Else
                            objListItem.SubItems.Add("")
                        End If
                        objListItem.SubItems.Add(strUninstallCMD)
                    End If

                End If
            End If

        Next


    End Sub

    Private Sub FormRemotePCAdmin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        RefreshTab()

        Dim UPD As New Updater
        UPD.DeleteOldFiles()
        UPD.CheckForUpdates(strUpdateCheckURL, strUpdateDownloadURL, True)

    End Sub

    Private Sub MemUsage()
        Dim MemAvailable As Integer
        Dim colMemory As Object
        Dim colCompSys As Object
        Dim objItem As Object
        Dim strComputer As String



        Dim objWbemLocator As Object
        Dim objWbem As Object
        Dim objServices As Object
        Const aCall As Integer = 3
        Const iImpersonate As Integer = 3
        Dim objMemItem As Object
        Dim MemTotal As Long

        If TextBoxComputer.Text = "" Then
            strComputer = "LOCALHOST"
        Else
            strComputer = TextBoxComputer.Text
        End If


        On Error Resume Next
        ' Create the Locator object
        objWbemLocator = CreateObject("WbemScripting.SWbemLocator")

        If CheckBoxAlternateCredentials.Checked Then
            If strUsername = "" Or strPassword = "" Then
                MsgBox("Please set the credentials.", MsgBoxStyle.Information, "Credentials not set.")
                Exit Sub
            End If
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\CIMV2", strUsername, strPassword)
        Else
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\CIMV2")
        End If

        If Err.Number <> 0 Then
            Dim strMsg As String
            strMsg = "Could not connect to:  " & strComputer & vbCrLf & vbCrLf & "Error Number:  " & Err.Number & vbCrLf & "Description:  " & Err.Description
            MsgBox(strMsg, MsgBoxStyle.Exclamation)
            Return
        End If

        TextBoxComputer.Text = UCase(strComputer)

        objServices.Security_.AuthenticationLevel = aCall
        objServices.Security_.ImpersonationLevel = iImpersonate

        objWbem = objServices

        colMemory = objWbem.ExecQuery("Select * from Win32_PerfFormattedData_PerfOS_Memory")
        colCompSys = objWbem.ExecQuery("Select * from Win32_ComputerSystem")

        'Loop through processes to find total mem usage
        For Each objMemItem In colMemory
            MemAvailable = objMemItem.AvailableMBytes
        Next

        'Find Total Memory
        For Each objItem In colCompSys
            MemTotal = objItem.TotalPhysicalMemory
        Next

        MemTotal = FormatNumber(MemTotal / 1048576, 0)

        ProgBarMem.Maximum = MemTotal
        ProgBarMem.Value = MemTotal - MemAvailable
        GroupBoxMem.Text = "Physical Memory (" & (MemTotal - MemAvailable) & " of " & MemTotal & " MB Used)"
        lblMemMax.Text = FormatPercent(MemAvailable / MemTotal) & " Available"
    End Sub

    Private Sub CpuUsage()
        Dim colProc As Object
        Dim objItem As Object
        Dim strComputer As String



        Dim objWbemLocator As Object
        Dim objWbem As Object
        Dim objServices As Object
        Dim objInstance1 As Object
        Dim perf_instance2 As Object
        Dim PercentProcessorTime As Object
        Dim N1 As Object
        Dim D1 As Object
        Dim N2 As Object
        Dim D2 As Object
        Const aCall As Integer = 3
        Const iImpersonate As Integer = 3

        If TextBoxComputer.Text = "" Then
            strComputer = "LOCALHOST"
        Else
            strComputer = TextBoxComputer.Text
        End If

        On Error Resume Next
        ' Create the Locator object
        objWbemLocator = CreateObject("WbemScripting.SWbemLocator")

        If CheckBoxAlternateCredentials.Checked Then
            If strUsername = "" Or strPassword = "" Then
                MsgBox("Please set the credentials.", MsgBoxStyle.Information, "Credentials not set.")
                Exit Sub
            End If
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\CIMV2", strUsername, strPassword)
        Else
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\CIMV2")
        End If

        If Err.Number <> 0 Then
            Dim strMsg As String
            strMsg = "Could not connect to:  " & strComputer & vbCrLf & vbCrLf & "Error Number:  " & Err.Number & vbCrLf & "Description:  " & Err.Description
            MsgBox(strMsg, MsgBoxStyle.Exclamation)
            Return
        End If

        TextBoxComputer.Text = UCase(strComputer)

        objServices.Security_.AuthenticationLevel = aCall
        objServices.Security_.ImpersonationLevel = iImpersonate

        objWbem = objServices
        objInstance1 = objWbem.Get("Win32_PerfRawData_PerfOS_Processor.Name='_Total'")
        N1 = objInstance1.PercentProcessorTime
        D1 = objInstance1.TimeStamp_Sys100NS

        System.Threading.Thread.Sleep(2000)

        perf_instance2 = objWbem.get("Win32_PerfRawData_PerfOS_Processor.Name='_Total'")
        N2 = perf_instance2.PercentProcessorTime
        D2 = perf_instance2.TimeStamp_Sys100NS

        ' CounterType - PERF_100NSEC_TIMER_INV
        ' Formula - (1- ((N2 - N1) / (D2 - D1))) x 100
        PercentProcessorTime = FormatNumber((1 - ((N2 - N1) / (D2 - D1))) * 100, 2)
        ProgBarCpu.Value = PercentProcessorTime
        GroupBoxCPU.Text = "CPU (" & PercentProcessorTime & "% Used)"
        lblCpuMax.Text = (100 - PercentProcessorTime) & "% Available"
        PercentProcessorTime = Nothing
    End Sub
    Private Sub DiskUsage()
        GroupBoxDisk1.Visible = True
        GroupBoxDisk2.Visible = True
        GroupBoxDisk3.Visible = True

        Dim colDisks As Object
        Dim objDisk As Object
        Dim strComputer As String



        Dim objWbemLocator As Object
        Dim objWbem As Object
        Dim objServices As Object
        Const aCall As Integer = 3
        Const iImpersonate As Integer = 3
        Const HARD_DISK As Integer = 3

        If TextBoxComputer.Text = "" Then
            strComputer = "LOCALHOST"
        Else
            strComputer = TextBoxComputer.Text
        End If


        On Error Resume Next
        ' Create the Locator object
        objWbemLocator = CreateObject("WbemScripting.SWbemLocator")

        If CheckBoxAlternateCredentials.Checked Then
            If strUsername = "" Or strPassword = "" Then
                MsgBox("Please set the credentials.", MsgBoxStyle.Information, "Credentials not set.")
                Exit Sub
            End If
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\CIMV2", strUsername, strPassword)
        Else
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\CIMV2")
        End If

        If Err.Number <> 0 Then
            Dim strMsg As String
            strMsg = "Could not connect to:  " & strComputer & vbCrLf & vbCrLf & "Error Number:  " & Err.Number & vbCrLf & "Description:  " & Err.Description
            MsgBox(strMsg, MsgBoxStyle.Exclamation)
            Return
        End If

        TextBoxComputer.Text = UCase(strComputer)

        objServices.Security_.AuthenticationLevel = aCall
        objServices.Security_.ImpersonationLevel = iImpersonate

        objWbem = objServices
        colDisks = objWbem.ExecQuery("Select * from Win32_LogicalDisk Where DriveType = " & HARD_DISK & "")

        Dim Drivecount As Integer = 0
        Dim Max As Long
        Dim Free As Long
        Dim Used As Long
        For Each objDisk In colDisks
            Drivecount += 1
            Max = objDisk.Size
            Free = objDisk.FreeSpace
            Used = objDisk.Size - objDisk.FreeSpace
            Select Case Drivecount
                Case "1"
                    ProgBarDisk1.Maximum = Max / 1073741824 * 100
                    ProgBarDisk1.Value = Used / 1073741824 * 100
                    'GroupBoxDisk1.Text = "Drive " & objDisk.DeviceID & " (" & FormatNumber(Used / 1073741824, 3) & " of " & FormatNumber(Max / 1073741824, 3) & " GB Used)"
                    'lblDisk1Max.Text = FormatPercent(Free / Max, 2) & " Available (" & FormatNumber(Free / 1073741824, 3) & " GB)"
                    GroupBoxDisk1.Text = "Drive " & objDisk.DeviceID & " (" & StrFormatByteSize(Used) & " of " & StrFormatByteSize(Max) & " Used)"
                    lblDisk1Max.Text = FormatPercent(Free / Max, 2) & " Available (" & StrFormatByteSize(Free) & ")"
                Case "2"
                    ProgBarDisk2.Maximum = Max / 1073741824 * 100
                    ProgBarDisk2.Value = Used / 1073741824 * 100
                    'GroupBoxDisk2.Text = "Drive " & objDisk.DeviceID & " (" & FormatNumber(Used / 1073741824, 3) & " of " & FormatNumber(Max / 1073741824, 3) & " GB Used)"
                    'lblDisk2Max.Text = FormatPercent(Free / Max, 2) & " Available (" & FormatNumber(Free / 1073741824, 3) & " GB)"
                    GroupBoxDisk2.Text = "Drive " & objDisk.DeviceID & " (" & StrFormatByteSize(Used) & " of " & StrFormatByteSize(Max) & " Used)"
                    lblDisk2Max.Text = FormatPercent(Free / Max, 2) & " Available (" & StrFormatByteSize(Free) & ")"
                Case "3"
                    ProgBarDisk3.Maximum = Max / 1073741824 * 100
                    ProgBarDisk3.Value = Used / 1073741824 * 100
                    'GroupBoxDisk3.Text = "Drive " & objDisk.DeviceID & " (" & FormatNumber(Used / 1073741824, 3) & " of " & FormatNumber(Max / 1073741824, 3) & " GB Used)"
                    'lblDisk3Max.Text = FormatPercent(Free / Max, 2) & " Available (" & FormatNumber(Free / 1073741824, 3) & " GB)"
                    GroupBoxDisk3.Text = "Drive " & objDisk.DeviceID & " (" & StrFormatByteSize(Used) & " of " & StrFormatByteSize(Max) & " Used)"
                    lblDisk3Max.Text = FormatPercent(Free / Max, 2) & " Available (" & StrFormatByteSize(Free) & ")"
                Case Else
                    MsgBox("There are more than 3 drives available.  Please use the MMC to view all drives", MsgBoxStyle.Exclamation)
            End Select
            Max = Nothing
            Free = Nothing
            Used = Nothing
        Next
        If Drivecount < 3 Then
            GroupBoxDisk3.Visible = False
        End If
        If Drivecount < 2 Then
            GroupBoxDisk2.Visible = False
        End If

    End Sub
    Private Sub Uptime()
        Dim colOpSys As Object
        Dim objOS As Object
        Dim strComputer As String



        Dim objWbemLocator As Object
        Dim objWbem As Object
        Dim objServices As Object
        Const aCall As Integer = 3
        Const iImpersonate As Integer = 3

        If TextBoxComputer.Text = "" Then
            strComputer = "LOCALHOST"
        Else
            strComputer = TextBoxComputer.Text
        End If


        On Error Resume Next
        ' Create the Locator object
        objWbemLocator = CreateObject("WbemScripting.SWbemLocator")

        If CheckBoxAlternateCredentials.Checked Then
            If strUsername = "" Or strPassword = "" Then
                MsgBox("Please set the credentials.", MsgBoxStyle.Information, "Credentials not set.")
                Exit Sub
            End If
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\CIMV2", strUsername, strPassword)
        Else
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\CIMV2")
        End If

        If Err.Number <> 0 Then
            Dim strMsg As String
            strMsg = "Could not connect to:  " & strComputer & vbCrLf & vbCrLf & "Error Number:  " & Err.Number & vbCrLf & "Description:  " & Err.Description
            MsgBox(strMsg, MsgBoxStyle.Exclamation)
            Return
        End If

        TextBoxComputer.Text = UCase(strComputer)

        objServices.Security_.AuthenticationLevel = aCall
        objServices.Security_.ImpersonationLevel = iImpersonate

        objWbem = objServices
        colOpSys = objWbem.ExecQuery("Select * from Win32_OperatingSystem")
        Dim dtmBootup As String
        Dim dtmLastBootupTime As String
        Dim dtmSystemUptime As String = ""
        For Each objOS In colOpSys
            dtmBootup = objOS.LastBootUpTime
            dtmLastBootupTime = WMIDateStringToDate(dtmBootup)
            dtmSystemUptime = DateDiff("h", dtmLastBootupTime, Now)
        Next

        lblUptime.Text = "Computer has been up for " & Microsoft.VisualBasic.Left(dtmSystemUptime / 24, 3) & " Days"
    End Sub

    Function WMIDateStringToDate(ByVal strDateString As String) As Date
        Dim objSWbemDateTime As Object = CreateObject("WbemScripting.SWbemDateTime")
        objSWbemDateTime.value = strDateString
        Return CDate(objSWbemDateTime.GetVarDate(True))
    End Function


    Private Sub GetLog(ByVal myLog As String)
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        GroupBoxEventLogs.Text = myLog & " Log"
        Dim strComputer As String



        Dim objWbemLocator As Object
        Dim objWbem As Object
        Dim objServices As Object
        Dim colItems As Object
        Dim objEvent As Object
        'Dim strDate As String
        'Dim strTime As String
        Dim strMsg As String
        Dim strType As String
        Dim color As Color
        Dim strEvent As String

        Dim strEventType As String

        Const aCall As Integer = 3
        Const iImpersonate As Integer = 3

        If TextBoxComputer.Text = "" Then
            strComputer = "LOCALHOST"
        Else
            strComputer = TextBoxComputer.Text
        End If

        Dim intMax As Integer = 25
        If MaxResults.Text = "" Then MaxResults.Text = intMax

        If Not IsNumeric(MaxResults.Text) Then
            MsgBox("Enter a valid number for Max. Results!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        intMax = CInt(MaxResults.Text)


        On Error Resume Next
        ' Create the Locator object
        objWbemLocator = CreateObject("WbemScripting.SWbemLocator")

        If CheckBoxAlternateCredentials.Checked Then
            If strUsername = "" Or strPassword = "" Then
                MsgBox("Please set the credentials.", MsgBoxStyle.Information, "Credentials not set.")
                Exit Sub
            End If
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\CIMV2", strUsername, strPassword)
        Else
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\CIMV2")
        End If

        If Err.Number <> 0 Then
            strMsg = "Could not connect to:  " & strComputer & vbCrLf & vbCrLf & "Error Number:  " & Err.Number & vbCrLf & "Description:  " & Err.Description
            MsgBox(strMsg, MsgBoxStyle.Exclamation)
            Return
        End If

        TextBoxComputer.Text = UCase(strComputer)

        objServices.Security_.AuthenticationLevel = aCall
        objServices.Security_.ImpersonationLevel = iImpersonate

        objWbem = objServices



        colItems = objWbem.ExecQuery("Select * from Win32_NTLogEvent Where Logfile='" & myLog & "'")

        RichTextBoxEventLogs.Clear()


        Dim count As Integer

        For Each objEvent In colItems
            count = count + 1
            If count > intMax Then Exit For
            strEventType = objEvent.Type
            Select Case LCase(strEventType)
                Case "information"
                    color = System.Drawing.Color.White
                    strType = "INFORMATION: "
                Case "warning"
                    color = System.Drawing.Color.Yellow
                    strType = "WARNING: "
                Case "error"
                    color = System.Drawing.Color.Red
                    strType = "ERROR: "
                Case "audit success"
                    color = System.Drawing.Color.Green
                    strType = "SUCCESS: "
                Case "audit failure"
                    color = System.Drawing.Color.Red
                    strType = "FAILURE: "
                Case Else
                    color = System.Drawing.Color.White
                    strType = "INFORMATION: "
            End Select
            RichTextBoxEventLogs.SelectionFont = New Font(RichTextBoxEventLogs.SelectionFont, FontStyle.Bold)
            RichTextBoxEventLogs.SelectionColor = color
            RichTextBoxEventLogs.AppendText(strType)
            RichTextBoxEventLogs.SelectionColor = RichTextBoxEventLogs.ForeColor
            RichTextBoxEventLogs.SelectionFont = New Font(RichTextBoxEventLogs.SelectionFont, FontStyle.Regular)


            Dim dtmTimeWritten As Date = WMIDateStringToDate(objEvent.TimeWritten)

            strEvent = "(" & objEvent.SourceName & ") " & objEvent.Message

            If objEvent.EventCode = "6009" Then
                RichTextBoxEventLogs.SelectionFont = New Font(RichTextBoxEventLogs.SelectionFont, FontStyle.Bold)
                RichTextBoxEventLogs.AppendText(dtmTimeWritten.ToString & " - ")
                RichTextBoxEventLogs.AppendText("****Computer Restarted**** -- " & strEvent)
            Else
                RichTextBoxEventLogs.AppendText(dtmTimeWritten.ToString & " - ")
                RichTextBoxEventLogs.AppendText(Replace(strEvent, vbCrLf & vbCrLf, vbCrLf))
            End If
            RichTextBoxEventLogs.AppendText(vbCrLf)

            strEventType = ""
            strType = ""
        Next

    End Sub
    Private Sub RadioButtonApplication_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonApplication.CheckedChanged
        GetLog("Application")
    End Sub
    Private Sub RadioButtonSecurity_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonSecurity.CheckedChanged
        GetLog("Security")
    End Sub
    Private Sub RadioButtonSystem_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonSystem.CheckedChanged
        GetLog("System")
    End Sub
    Private Sub ViewHotfixes()
        RichTextBoxUpdates.Clear()
        Dim objWbemLocator As Object
        Const aCall As Integer = 3
        Const iImpersonate As Integer = 3
        Dim objWbem As Object


        Dim strComputer As String

        Dim objServices As Object
        Dim colQuickFixes As Object
        Dim objQuickFix As Object
        Dim InstalledBy As String
        Dim Description As String
        Dim HotFixID As String
        Dim InstalledOn As String

        On Error Resume Next
        If TextBoxComputer.Text = "" Then
            strComputer = "LOCALHOST"
        Else
            strComputer = TextBoxComputer.Text
        End If

        ' Create the Locator object
        objWbemLocator = CreateObject("WbemScripting.SWbemLocator")

        If CheckBoxAlternateCredentials.Checked Then
            If strUsername = "" Or strPassword = "" Then
                MsgBox("Please set the credentials.", MsgBoxStyle.Information, "Credentials not set.")
                Exit Sub
            End If
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\CIMV2", strUsername, strPassword)
        Else
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\CIMV2")
        End If

        If Err.Number <> 0 Then
            Dim strMsg As String
            strMsg = "Could not connect to:  " & strComputer & vbCrLf & vbCrLf & "Error Number:  " & Err.Number & vbCrLf & "Description:  " & Err.Description
            MsgBox(strMsg, MsgBoxStyle.Exclamation)
            Return
        End If

        TextBoxComputer.Text = UCase(strComputer)

        objServices.Security_.AuthenticationLevel = aCall
        objServices.Security_.ImpersonationLevel = iImpersonate

        objWbem = objServices

        colQuickFixes = objWbem.ExecQuery("Select * from Win32_QuickFixEngineering")

        RichTextBoxUpdates.AppendText("----------------------------------------------" & vbCrLf)
        RichTextBoxUpdates.SelectionFont = New Font(RichTextBoxUpdates.SelectionFont, FontStyle.Bold)
        RichTextBoxUpdates.AppendText(UCase(strComputer) & " - Hotfixes" & vbCrLf)
        RichTextBoxUpdates.AppendText("----------------------------------------------" & vbCrLf & vbCrLf)

        For Each objQuickFix In colQuickFixes

            InstalledOn = objQuickFix.InstalledOn
            HotFixID = objQuickFix.HotFixID
            Description = objQuickFix.Description
            If InStr(Description, " \n") Then
                Description = Replace(Description, "\n", " ")
            End If


            If objQuickFix.InstalledBy = "" Then
                InstalledBy = ""
            Else
                InstalledBy = ", Installed by " & objQuickFix.InstalledBy
            End If
            If Description <> "" Then
                RichTextBoxUpdates.AppendText(InstalledOn & " - " & HotFixID & " " & Description & InstalledBy & vbCrLf)
            End If
        Next
    End Sub

    Private Sub SystemRestoreStatus()
        Dim strUser As String

        Dim objWbemLocator As Object
        Const aCall As Integer = 3
        Const iImpersonate As Integer = 3
        Dim objWbem As Object


        Dim strComputer As String
        Dim objServices As Object
        Dim objWMIClass As Object
        Dim SREnabled As Boolean

        Dim vRegistry As Object
        Const HKEY_LOCAL_MACHINE As Integer = &H80000002
        Dim strKeyPath As String = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore\"
        Dim dwValueName As String = "DisableSR"

        On Error Resume Next
        If TextBoxComputer.Text = "" Then
            strComputer = "LOCALHOST"
        Else
            strComputer = TextBoxComputer.Text
        End If


        ' Create the Locator object
        objWbemLocator = CreateObject("WbemScripting.SWbemLocator")

        If CheckBoxAlternateCredentials.Checked Then
            If strUsername = "" Or strPassword = "" Then
                MsgBox("Please set the credentials.", MsgBoxStyle.Information, "Credentials not set.")
                Exit Sub
            End If
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\CIMV2", strUsername, strPassword)
        Else
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\CIMV2")
        End If

        If Err.Number <> 0 Then
            Dim strMsg As String
            strMsg = "Could not connect to:  " & strComputer & vbCrLf & vbCrLf & "Error Number:  " & Err.Number & vbCrLf & "Description:  " & Err.Description
            MsgBox(strMsg, MsgBoxStyle.Exclamation)
            Return
        End If

        TextBoxComputer.Text = UCase(strComputer)

        objServices.Security_.AuthenticationLevel = aCall
        objServices.Security_.ImpersonationLevel = iImpersonate

        objWbem = objServices

        vRegistry = objWbem.Get("StdRegProv")

        Dim dwValue As Object = Nothing
        Dim strStatus As String
        vRegistry.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, dwValueName, dwValue)

        'Looks like Vista & Higher don't have the DisableSR reg value

        If IsDBNull(dwValue) Then
            Dim objSR As Object
            objSR = objWbem.Get("SystemRestoreConfig='SR'")
            Dim RPSessionInterval As Boolean
            RPSessionInterval = objSR.RPSessionInterval
            If RPSessionInterval = True Then
                SREnabled = True
            Else
                SREnabled = False
            End If
        Else
            'MsgBox(dwValue)
            If dwValue = 0 Then
                SREnabled = True
            Else
                SREnabled = False
            End If

        End If


        If SREnabled = True Then
            LabelSystemRestore.Text = "System Restore is currently ENABLED."
            btnSystemRestore.Text = "Disable"
            strStatus = "Disable"
        Else
            LabelSystemRestore.Text = "System Restore is currently DISABLED."
            btnSystemRestore.Text = "Enable"
            strStatus = "Enable"
        End If
    End Sub

    Private Sub btnSystemRestore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSystemRestore.Click
        Dim AskYesNo As MsgBoxResult
        AskYesNo = MsgBox("Are you sure you want to " & btnSystemRestore.Text & " System Restore?", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation)
        If AskYesNo <> MsgBoxResult.Yes Then
            MsgBox("Action Canceled", MsgBoxStyle.Information)
            Exit Sub
        End If


        Dim strUser As String

        Dim objWbemLocator As Object
        Const aCall As Integer = 3
        Const iImpersonate As Integer = 3
        Dim objWbem As Object


        Dim strComputer As String
        Dim objServices As Object
        Dim objWMI As Object
        Dim strStatus As String = btnSystemRestore.Text
        On Error Resume Next
        If TextBoxComputer.Text = "" Then
            strComputer = "LOCALHOST"
        Else
            strComputer = TextBoxComputer.Text
        End If


        ' Create the Locator object
        objWbemLocator = CreateObject("WbemScripting.SWbemLocator")

        If CheckBoxAlternateCredentials.Checked Then
            If strUsername = "" Or strPassword = "" Then
                MsgBox("Please set the credentials.", MsgBoxStyle.Information, "Credentials not set.")
                Exit Sub
            End If
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\CIMV2", strUsername, strPassword)
        Else
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\CIMV2")
        End If

        If Err.Number <> 0 Then
            Dim strMsg As String
            strMsg = "Could not connect to:  " & strComputer & vbCrLf & vbCrLf & "Error Number:  " & Err.Number & vbCrLf & "Description:  " & Err.Description
            MsgBox(strMsg, MsgBoxStyle.Exclamation)
            Return
        End If

        TextBoxComputer.Text = UCase(strComputer)

        objServices.Security_.AuthenticationLevel = aCall
        objServices.Security_.ImpersonationLevel = iImpersonate

        objWbem = objServices

        objWMI = objWbem.Get("SystemRestore")
        Dim errResults As Integer
        Select Case strStatus
            Case "Disable"
                errResults = objWMI.Disable("")
                If errResults = 0 Then
                    btnSystemRestore.Text = "Enable"
                    LabelSystemRestore.Text = "System Restore is currently DISABLED."
                End If
            Case "Enable"
                errResults = objWMI.Enable("")
                If errResults = 0 Then
                    btnSystemRestore.Text = "Disable"
                    LabelSystemRestore.Text = "System Restore is currently ENABLED."
                End If
        End Select
        Select Case errResults
            Case 0
                MsgBox("Successfully " & strStatus & "d System Restore", MsgBoxStyle.Information)
            Case 1717
                MsgBox("System Restore was already Disabled", MsgBoxStyle.Exclamation)
                LabelSystemRestore.Text = "System Restore is currently DISABLED."
                btnSystemRestore.Text = "Disable"
            Case 1056
                MsgBox("System Restore was already enabled", MsgBoxStyle.Exclamation)
                LabelSystemRestore.Text = "System Restore is currently ENABLED."
                btnSystemRestore.Text = "Enable"
            Case Else
                MsgBox("Error code:  " & errResults, MsgBoxStyle.Exclamation)
        End Select

    End Sub
    Private Sub ListRestorePoints()
        Dim strUser As String

        Dim objWbemLocator As Object
        Const aCall As Integer = 3
        Const iImpersonate As Integer = 3
        Dim objWbem As Object


        Dim strComputer As String
        Dim objServices As Object
        Dim colItems As Object
        Dim objItem As Object
        Dim strRPDescription As String
        Dim strRPNumber As String
        Dim strRPType As String
        Dim strRPDate As String
        Dim arrDate As Array
        Dim objListItem As Object

        On Error Resume Next
        If TextBoxComputer.Text = "" Then
            strComputer = "LOCALHOST"
        Else
            strComputer = TextBoxComputer.Text
        End If


        ' Create the Locator object
        objWbemLocator = CreateObject("WbemScripting.SWbemLocator")

        If CheckBoxAlternateCredentials.Checked Then
            If strUsername = "" Or strPassword = "" Then
                MsgBox("Please set the credentials.", MsgBoxStyle.Information, "Credentials not set.")
                Exit Sub
            End If
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\default", strUsername, strPassword)
        Else
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\default")
        End If

        If Err.Number <> 0 Then
            Dim strMsg As String
            strMsg = "Could not connect to:  " & strComputer & vbCrLf & vbCrLf & "Error Number:  " & Err.Number & vbCrLf & "Description:  " & Err.Description
            MsgBox(strMsg, MsgBoxStyle.Exclamation)
            Return
        End If

        TextBoxComputer.Text = UCase(strComputer)

        objServices.Security_.AuthenticationLevel = aCall
        objServices.Security_.ImpersonationLevel = iImpersonate

        objWbem = objServices

        colItems = objWbem.ExecQuery("Select * from SystemRestore")

        If colItems.Count = 0 Then
            objListItem = ListViewSystemRestore.Items.Add("", 0)
            objListItem.SubItems.Add("No restore points found.")
            Exit Sub
        End If

        For Each objItem In colItems
            strRPDescription = objItem.Description
            strRPNumber = objItem.SequenceNumber
            strRPDate = WMIDateStringToDate(objItem.CreationTime)
            strRPType = objItem.RestorePointType
            Select Case strRPType
                Case 0
                    strRPType = "Application installation"
                Case 1
                    strRPType = "Application uninstall"
                Case 6
                    strRPType = "Restore"
                Case 7
                    strRPType = "Checkpoint"
                Case 10
                    strRPType = "Device drive installation"
                Case 11
                    strRPType = "First run"
                Case 12
                    strRPType = "Modify settings"
                Case 13
                    strRPType = "Cancelled operation"
                Case 14
                    strRPType = "Backup recovery"
                Case Else
                    strRPType = "Unknown"
            End Select
            objListItem = ListViewSystemRestore.Items.Add(strRPNumber, 0)
            objListItem.SubItems.Add(strRPDescription)
            objListItem.SubItems.Add(strRPType)
            objListItem.SubItems.Add(strRPDate)
        Next



    End Sub

    Private Sub RestoreSystem(ByVal RestorePoint As Integer)
        Dim strUser As String

        Dim objWbemLocator As Object
        Const aCall As Integer = 3
        Const iImpersonate As Integer = 3
        Dim objWbem As Object


        Dim strComputer As String
        Dim objServices As Object
        Dim objWMI As Object
        On Error Resume Next
        If TextBoxComputer.Text = "" Then
            strComputer = "LOCALHOST"
        Else
            strComputer = TextBoxComputer.Text
        End If

        ' Create the Locator object
        objWbemLocator = CreateObject("WbemScripting.SWbemLocator")

        If CheckBoxAlternateCredentials.Checked Then
            If strUsername = "" Or strPassword = "" Then
                MsgBox("Please set the credentials.", MsgBoxStyle.Information, "Credentials not set.")
                Exit Sub
            End If
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\CIMV2", strUsername, strPassword)
        Else
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\CIMV2")
        End If

        If Err.Number <> 0 Then
            Dim strMsg As String
            strMsg = "Could not connect to:  " & strComputer & vbCrLf & vbCrLf & "Error Number:  " & Err.Number & vbCrLf & "Description:  " & Err.Description
            MsgBox(strMsg, MsgBoxStyle.Exclamation)
            Return
        End If

        TextBoxComputer.Text = UCase(strComputer)

        objServices.Security_.AuthenticationLevel = aCall
        objServices.Security_.ImpersonationLevel = iImpersonate

        objWbem = objServices

        objWMI = objWbem.Get("SystemRestore")

        Dim errResults As Integer
        errResults = objWMI.Restore(RestorePoint)
        If errResults <> 0 Then
            MsgBox("System Restore failed:  " & errResults, MsgBoxStyle.Critical)
        Else
            MsgBox("System Restore request submitted successfully." & vbCrLf & vbCrLf &
            "System Restore should occur at next shutdown.", MsgBoxStyle.Information)
        End If

    End Sub




    Private Sub RefreshTab()
        Dim strTab As String = TabControl1.SelectedTab.Name
        TabCheck(strTab)
    End Sub






    Private Sub NotifyIcon1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDown
        If e.Button = MouseButtons.Right Then
            NotifyIcon1.ContextMenu = ContextMenuIcon
        End If
        If e.Button = MouseButtons.Left Then
            'Me.Activate()
            Me.Visible = True
            If Me.WindowState = FormWindowState.Minimized Then
                Me.WindowState = FormWindowState.Normal
            End If
            NotifyIcon1.Visible = False
        End If

    End Sub

    'Private Sub mainForm_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
    '    If Me.WindowState = FormWindowState.Minimized Then
    '        NotifyIcon1.Visible = True
    '        Me.Visible = False
    '        NotifyIcon1.BalloonTipTitle = "Remote PC Admin"
    '        NotifyIcon1.BalloonTipText = "Click icon to reactivate."
    '        NotifyIcon1.ShowBalloonTip(600)
    '    End If
    'End Sub


    Private Sub MenuItemIcon3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemIcon3.Click
        Application.Exit()
    End Sub

    Private Sub MenuItemIcon2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemIcon2.Click
        AboutBox1.ShowDialog()
    End Sub

    Private Sub MenuItemIcon1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemIcon1.Click
        Me.Visible = True
        If Me.WindowState = FormWindowState.Minimized Then
            Me.WindowState = FormWindowState.Normal
        End If
        NotifyIcon1.Visible = False
    End Sub





    Private Sub rewriteHTM()
        Dim NewText As String
        Dim strFile As String
        Dim arrInputData As Array
        Dim Text As String
        Dim line As String

        strFile = "C:\WINDOWS\pchealth\helpctr\Vendors\CN=Microsoft Corporation,L=Redmond,S=Washington,C=US\Remote Assistance\Escalation\Unsolicited\UnsolicitedRCUI.htm"


        NewText = "function onLoad() " & vbCrLf &
        "	{ " & vbCrLf &
        "		KickArse(); " & vbCrLf &
        "	}  " & vbCrLf &
        "function KickArse() " & vbCrLf &
        "	{ " & vbCrLf &
        "		try " & vbCrLf &
        "		{ " & vbCrLf &
        "			setTimeout(""idComputerName.focus()"",250); " & vbCrLf &
        "			g_oSAFRemoteDesktopConnection = oSAFClassFactory.CreateObject_RemoteDesktopConnection(); " & vbCrLf &
        "			if (window.location.search.substring(1)!= """") " & vbCrLf &
        "			{ " & vbCrLf &
        "			idComputerName.value = window.location.search.substring(1); // load the Machine Name " & vbCrLf &
        "			document.getElementById('btnConnect').click(); // clicks the connect button " & vbCrLf &
        "			document.getElementById('btnStart').click(); // clicks the Start Remote Assistance button " & vbCrLf &
        "			} " & vbCrLf &
        "		} " & vbCrLf &
        "		catch(error) " & vbCrLf &
        "		{ " & vbCrLf &
        "			FatalError( L_RCCTL_Text, error ); " & vbCrLf &
        "		} " & vbCrLf &
        "		return; " & vbCrLf &
        "	} " & vbCrLf &
        "function onLoadDISABLED() "

        If System.IO.File.Exists(strFile) = False Then
            Exit Sub
        End If

        Dim objReader As New System.IO.StreamReader(strFile)
        Text = objReader.ReadToEnd
        objReader.Close()
        arrInputData = Text.Split(vbNewLine)

        If InStr(Text, "KickArse") = False Then
            Dim objWriter As New System.IO.StreamWriter(strFile)
            For Each line In arrInputData
                If InStr(line, "function onLoad()") Then
                    objWriter.WriteLine(NewText)
                Else
                    objWriter.WriteLine(line)
                End If
            Next
            objWriter.Close()
        End If

    End Sub




    Private Sub UninstallSoftware(ByVal strSoftwareName As String, ByVal strUninstallString As String)

        Dim AskYesNo As MsgBoxResult

        If strSoftwareName = "" Then Exit Sub
        Dim boolMsi As Boolean = False

        'Modify uinstallstring for MSI
        If InStr(strUninstallString, "msiexec", CompareMethod.Text) > 0 Then
            boolMsi = True
            Dim out As String = strUninstallString
            Dim strPattern As String = "msiexec /i|msiexec.exe /i"
            Dim options As RegexOptions
            options = RegexOptions.IgnoreCase

            Dim regex As Regex = New Regex(strPattern, options)

            strUninstallString = regex.Replace(strUninstallString, "msiexec.exe /x")

            If InStr(strUninstallString, "/qn", CompareMethod.Text) = 0 Then
                strUninstallString = strUninstallString & " /qn"
            End If

        End If


        'if not local and not MSI... throw a warning
        Dim strComputer As String = TextBoxComputer.Text

        AskYesNo = MsgBox("Are you sure you wish to uninstall:" & vbCrLf & vbCrLf & strSoftwareName, MsgBoxStyle.YesNo + MsgBoxStyle.Question)
        Select Case AskYesNo
            Case MsgBoxResult.Yes
                'If boolMsi = False Then
                If Not IsLocalComputer(strComputer) Then
                    MsgBox("WARNING!!" & vbCrLf & vbCrLf & "You'll be prompted to enter a command to execute on the computer.  Make sure to add any switches/parameters that the uninstaller needs to run silent/unattended and not automatically reboot!", MsgBoxStyle.Exclamation)
                End If
                'End If

                Dim frmProcAdd As FormProcessAdd = New FormProcessAdd
                frmProcAdd.TextBoxCommand.Text = strUninstallString

                frmProcAdd.ShowDialog()

                'MsgBox(strComputer & vbCrLf & vbCrLf & "Uninstall request submitted... check event log for results.", MsgBoxStyle.Information)
            Case Else
                MsgBox("Action Cancelled.", MsgBoxStyle.Information)
        End Select

    End Sub

    Private Sub EndProcessToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EndProcessToolStripMenuItem.Click
        Dim intPIDColumnIndex As Integer = 0
        Dim intNameColumnIndex As Integer = 0
        For Each col As ColumnHeader In lstProcesses.Columns
            If col.Text = "PID" Then
                intPIDColumnIndex = col.Index
            ElseIf col.Text = "Name" Then
                intNameColumnIndex = col.Index
            End If
        Next

        Dim c As Integer
        c = lstProcesses.SelectedIndices(0)
        Dim objWbemLocator As Object
        'Dim colProcesses As Object
        Dim objWbem As Object



        Dim objServices As Object
        Const aCall As Integer = 3
        Const iImpersonate As Integer = 3
        Dim strComputer As String
        strComputer = TextBoxComputer.Text
        objWbemLocator = CreateObject("WbemScripting.SWbemLocator")

        If CheckBoxAlternateCredentials.Checked Then
            If strUsername = "" Or strPassword = "" Then
                MsgBox("Please set the credentials.", MsgBoxStyle.Information, "Credentials not set.")
                Exit Sub
            End If
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\CIMV2", strUsername, strPassword)
        Else
            objServices = objWbemLocator.ConnectServer(strComputer, "Root\CIMV2")
        End If

        objServices.Security_.AuthenticationLevel = aCall
        objServices.Security_.ImpersonationLevel = iImpersonate

        objWbem = objServices

        Dim strProcess As String
        Dim intProcessID As Integer
        Dim objProcess As Object
        Dim strProcMsg As String
        Dim AskYesNo As MsgBoxResult
        strProcess = lstProcesses.Items(c).SubItems(intNameColumnIndex).Text
        intProcessID = CInt(lstProcesses.Items(c).SubItems(intPIDColumnIndex).Text)
        strProcMsg = vbCrLf & vbCrLf & "Name:  " & strProcess & vbCrLf & "PID:  " & intProcessID

        For Each objProcess In objWbem.InstancesOf("Win32_Process where ProcessID='" & intProcessID & "'")
            AskYesNo = MsgBox("Are you sure you wish to end this process?" & strProcMsg, MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation)
            Select Case AskYesNo
                Case MsgBoxResult.Yes
                    objProcess.Terminate(intProcessID)
                    MsgBox("Termination request sent for:  " & strProcMsg, MsgBoxStyle.Information)

                    ViewProcesses()
                Case Else
                    MsgBox("Action Canceled", MsgBoxStyle.Information)
            End Select
        Next
    End Sub

    Private Sub ViewPathToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewPathToolStripMenuItem.Click
        On Error Resume Next
        Dim intCmdColumnIndex As Integer = 0
        For Each col As ColumnHeader In lstProcesses.Columns
            If col.Text = "Command Line" Then
                intCmdColumnIndex = col.Index
                Exit For
            End If
        Next

        Dim strProcPath As String
        Dim i As Integer
        i = lstProcesses.SelectedIndices(0)
        strProcPath = lstProcesses.Items(i).SubItems(intCmdColumnIndex).Text
        MsgBox(strProcPath, MsgBoxStyle.Information)
    End Sub


    Private Sub UninstallToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UninstallToolStripMenuItem.Click
        On Error Resume Next

        Dim strUninstallString As String
        Dim strInstallType As String
        Dim strSoftwareName As String
        Dim strComputer As String
        Dim c As Integer
        strComputer = TextBoxComputer.Text
        c = lstSoftware.SelectedIndices(0)
        strSoftwareName = lstSoftware.Items(c).SubItems(0).Text
        strInstallType = lstSoftware.Items(c).SubItems(1).Text
        strUninstallString = lstSoftware.Items(c).SubItems(2).Text
        'If this is the local computer, run the uninstall string
        If IsLocalComputer(strComputer) Then
            If strUninstallString <> "" Then
                Shell(strUninstallString, AppWinStyle.NormalFocus)
            End If
        Else
            'must be a remote computer, try triggering if it's an MSI
            UninstallSoftware(strSoftwareName, strUninstallString)
        End If
    End Sub

    Private Function IsLocalComputer(ByVal strComputer As String) As Boolean
        If strComputer = "" Then Return True
        If strComputer = "." Then Return True
        If strComputer.ToLower() = "localhost" Then Return True
        If strComputer.ToLower() = System.Net.Dns.GetHostName().ToLower() Then Return True
        If strComputer.ToLower() = My.Computer.Name Then Return True
        Return False
    End Function

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripMenuItem.Click

        Dim theBox As RichTextBox = ContextMenuStripCopy.SourceControl

        If (Not theBox Is Nothing) Then
            ' Put selected text on Clipboard.
            theBox.Copy()
        End If

    End Sub


    Private Sub ShowInExplorerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowInExplorerToolStripMenuItem.Click
        On Error Resume Next
        Dim intExeColumnIndex As Integer = 0
        For Each col As ColumnHeader In lstProcesses.Columns
            If col.Text = "Executable Path" Then
                intExeColumnIndex = col.Index
                Exit For
            End If
        Next

        Dim strProcPath As String
        Dim i As Integer
        Dim strComputer As String
        strComputer = TextBoxComputer.Text
        i = lstProcesses.SelectedIndices(0)
        strProcPath = lstProcesses.Items(i).SubItems(intExeColumnIndex).Text

        If Not IsLocalComputer(strComputer) Then
            strProcPath = "\\" & strComputer & "\" & Replace(strProcPath, ":", "$")
        End If

        Shell("explorer.exe /e,/select," & strProcPath, AppWinStyle.NormalFocus)

    End Sub

    Private Sub ContextMenuStripSoftware_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStripSoftware.Opening
        On Error Resume Next
        Dim strInstallType As String
        Dim c As Integer
        Dim strComputer As String
        Dim strUninstallString As String

        strComputer = TextBoxComputer.Text
        c = lstSoftware.SelectedIndices(0)
        strInstallType = lstSoftware.Items(c).SubItems(1).Text
        strUninstallString = lstSoftware.Items(c).SubItems(2).Text
        'If it's a local host, make sure there's an uninstall string
        If IsLocalComputer(strComputer) Then
            If strUninstallString = "" Then e.Cancel = True
            UninstallToolStripMenuItem.Visible = False
        Else
            UninstallToolStripMenuItem.Visible = True
            'If strInstallType <> "MSI" Then e.Cancel = True
        End If


    End Sub




    Private Sub ButtonSetCredentials_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSetCredentials.Click
        Login.ShowDialog()
    End Sub

    Private Function StrFormatByteSize(ByVal lngBytes As Long) As String
        Debug.WriteLine("bytes:  " & lngBytes)
        If lngBytes < 1024 Then Return lngBytes & " bytes"

        Dim suf As String() = {"B", "KB", "MB", "GB", "TB", "PB"}
        Dim place As Integer = Convert.ToInt32(Math.Floor(Math.Log(lngBytes, 1024)))

        Dim num As Double = Math.Round(lngBytes / Math.Pow(1024, place), 1)
        Dim readable As String = num.ToString() & " " & suf(place)

        StrFormatByteSize = readable
    End Function


    Private Function GetCollection(ByVal strQuery As String) As ManagementObjectCollection
        Dim strComputer As String = TextBoxComputer.Text
        If strComputer = "" Then strComputer = "."


        Dim oConn As ConnectionOptions = New ConnectionOptions
        If CheckBoxAlternateCredentials.Checked And strComputer <> "." Then
            If strUsername = "" Or strPassword = "" Then
                MsgBox("Please set the credentials.", MsgBoxStyle.Information, "Credentials not set.")
                Return Nothing
            End If
            oConn.Username = strUsername
            oConn.Password = strPassword
        End If
        'oConn.EnablePrivileges = True
        'oConn.Impersonation = ImpersonationLevel.Impersonate
        'oConn.Timeout = New TimeSpan(0, 5, 0)

        Try
            Dim mp As ManagementPath = New ManagementPath("\\" & strComputer & "\root\cimv2")
            Dim oMs As ManagementScope = New ManagementScope(mp, oConn)

            Dim oQuery As SelectQuery = New SelectQuery(String.Format(strQuery))
            Dim oSearcher As ManagementObjectSearcher = New ManagementObjectSearcher(oMs, oQuery)
            Dim colTemp As ManagementObjectCollection = oSearcher.Get()
            Dim intCount As Integer = colTemp.Count ' if there was a problem, this may force the exception
            Return colTemp
        Catch unauth As System.UnauthorizedAccessException
            MsgBox("Access is denied.", MsgBoxStyle.Exclamation, Application.ProductName)
            Return Nothing
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & vbCrLf & strQuery, MsgBoxStyle.Exclamation, Application.ProductName)
            Return Nothing
        End Try

    End Function


    Private Function GetObjectCollection(ByVal strQuery As String) As ManagementObjectCollection
        Dim strComputer As String = TextBoxComputer.Text
        If strComputer = "" Then strComputer = "."


        Dim oConn As ConnectionOptions = New ConnectionOptions
        If CheckBoxAlternateCredentials.Checked And strComputer <> "." Then
            If strUsername = "" Or strPassword = "" Then
                MsgBox("Please set the credentials.", MsgBoxStyle.Information, "Credentials not set.")
                Return Nothing
            End If
            oConn.Username = strUsername
            oConn.Password = strPassword
        End If
        oConn.EnablePrivileges = True
        oConn.Impersonation = ImpersonationLevel.Impersonate
        oConn.Timeout = New TimeSpan(0, 5, 0)

        Try
            Dim mp As ManagementPath = New ManagementPath("\\" & strComputer & "\root\cimv2")
            Dim oMs As ManagementScope = New ManagementScope(mp, oConn)
            oMs.Connect()
            If oMs.IsConnected Then
                Dim oQuery As ObjectQuery = New ObjectQuery(String.Format(strQuery))
                Dim oSearcher As ManagementObjectSearcher = New ManagementObjectSearcher(oMs, oQuery)
                Dim colTemp As ManagementObjectCollection = oSearcher.Get()

                Return colTemp
            Else
                Return Nothing
            End If

        Catch unauth As System.UnauthorizedAccessException
            MsgBox("Access is denied.", MsgBoxStyle.Exclamation, Application.ProductName)
            Return Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, Application.ProductName)
            Return Nothing
        End Try

    End Function



    Private Sub TabControl1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.Click

        If TabControl1.SelectedTab.Name = PreviousTab Then
            Debug.WriteLine(Now & vbTab & "REFRESH!")
            RefreshTab()
        Else
            PreviousTab = TabControl1.SelectedTab.Name
        End If


    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        RefreshTab()
    End Sub

    Private Sub ButtonGo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonGo.Click
        RefreshTab()
    End Sub

    Private Function Is64BitRemote() As Boolean
        Dim colProc As ManagementObjectCollection
        colProc = GetCollection("Select * from Win32_Processor")
        Dim bits As Integer
        For Each proc As ManagementObject In colProc
            bits = proc.GetPropertyValue("AddressWidth")
        Next
        If bits = 64 Then
            Is64BitRemote = True
        Else
            Is64BitRemote = False
        End If
    End Function

    Private Sub ListView_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lstProcesses.ColumnClick, lstServices.ColumnClick, lstSoftware.ColumnClick, ListViewSystemRestore.ColumnClick
        SortMyListView(sender, e.Column, , True)
    End Sub

    Friend Sub SortMyListView(ByVal ListViewToSort As ListView, ByVal ColumnNumber As Integer, Optional ByVal Resort As Boolean = False, Optional ByVal ForceSort As Boolean = False)
        ' Sorts a list view column by string, number, or date.  The three types of sorts must be specified within the listview columns "tag" property unless the ForceSort option is enabled.
        ' ListViewToSort - The listview to sort
        ' ColumnNumber - The column number to sort by
        ' Resort - Resorts a listview in the same direction as the last sort
        ' ForceSort - Forces a sort even if no listview tag data exists and assumes a string sort.  If this is false (default) and no tag is specified the procedure will exit
        Dim SortOrder As SortOrder
        Static LastSortColumn As Integer = -1
        Static LastSortOrder As SortOrder = SortOrder.Ascending
        If Resort = True Then
            SortOrder = LastSortOrder
        Else
            If LastSortColumn = ColumnNumber Then
                If LastSortOrder = SortOrder.Ascending Then
                    SortOrder = SortOrder.Descending
                Else
                    SortOrder = SortOrder.Ascending
                End If
            Else
                SortOrder = SortOrder.Ascending
            End If
        End If

        ' In case a tag wasn't specified check if we should force a string sort
        If String.IsNullOrEmpty(CStr(ListViewToSort.Columns(ColumnNumber).Tag)) Then
            If ForceSort = True Then
                ListViewToSort.Columns(ColumnNumber).Tag = "String"
            Else
                ' don't sort since no tag was specified.
                Exit Sub
            End If
        End If

        Select Case ListViewToSort.Columns(ColumnNumber).Tag.ToString
            Case "Numeric"
                ListViewToSort.ListViewItemSorter = New ListViewNumericSort(ColumnNumber, SortOrder)
            Case "Date"
                ListViewToSort.ListViewItemSorter = New ListViewDateSort(ColumnNumber, SortOrder)
            Case "String"
                ListViewToSort.ListViewItemSorter = New ListViewStringSort(ColumnNumber, SortOrder)
            Case "IP"
                ListViewToSort.ListViewItemSorter = New ListViewIPSort(ColumnNumber, SortOrder)
        End Select
        LastSortColumn = ColumnNumber
        LastSortOrder = SortOrder
        ListViewToSort.ListViewItemSorter = Nothing
    End Sub



    Private Sub CopyAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyAllToolStripMenuItem.Click

        Dim theBox As RichTextBox = ContextMenuStripCopy.SourceControl

        If (Not theBox Is Nothing) Then
            ' Put selected text on Clipboard.
            theBox.SelectAll()
            theBox.Copy()
            theBox.DeselectAll()
        End If
    End Sub

    Private Sub ViewInfoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewInfoToolStripMenuItem.Click
        On Error Resume Next
        Dim strSrvDisplayName As String
        Dim strSrvName As String
        Dim strSrvDescription As String
        Dim i As Integer
        i = lstServices.SelectedIndices(0)
        strSrvDisplayName = lstServices.Items(i).SubItems(0).Text
        strSrvName = lstServices.Items(i).SubItems(4).Text
        strSrvDescription = lstServices.Items(i).SubItems(3).Text
        MsgBox(strSrvDisplayName & " (" & strSrvName & "):" & vbCrLf & vbCrLf & strSrvDescription, MsgBoxStyle.Information)
    End Sub

    Private Sub RestartToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RestartToolStripMenuItem.Click
        ServiceChange("Restart")
    End Sub

    Private Sub StopToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StopToolStripMenuItem.Click
        ServiceChange("Stop")
    End Sub

    Private Sub StartToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StartToolStripMenuItem.Click
        ServiceChange("Start")
    End Sub

    Private Sub RemoveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveToolStripMenuItem.Click
        ServiceChange("Remove")
    End Sub

    Private Sub RestoreToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RestoreToolStripMenuItem.Click
        On Error Resume Next
        Dim RestorePoint As Integer = ListViewSystemRestore.SelectedIndices(0)
        If IsNumeric(RestorePoint) = False Then
            MsgBox("Please select a Restore Point.", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Dim strRPNumber As String = ListViewSystemRestore.Items(RestorePoint).SubItems(0).Text
        Dim strRPDescription As String
        strRPDescription = ListViewSystemRestore.Items(RestorePoint).SubItems(1).Text
        Dim AskYesNo As MsgBoxResult
        AskYesNo = MsgBox("Are you sure you wish to restore the system to point " & strRPNumber & "?" &
        vbCrLf & vbCrLf & "(" & strRPDescription & ")", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation)
        If AskYesNo <> MsgBoxResult.Yes Then
            MsgBox("Action Canceled", MsgBoxStyle.Information)
            Exit Sub
        End If

        RestoreSystem(strRPNumber)
    End Sub

    Private Sub TextBoxComputer_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxComputer.TextChanged
        Dim strTitle As String = Application.ProductName
        Dim strComputer As String = TextBoxComputer.Text
        If strComputer <> "" Then strTitle += " - " & strComputer
        Me.Text = strTitle
    End Sub



    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        AboutBox1.ShowDialog()
    End Sub



    Private Sub CheckForUpdatesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckForUpdatesToolStripMenuItem.Click
        Dim UPD As New Updater
        UPD.CheckForUpdates(strUpdateCheckURL, strUpdateDownloadURL, False)
    End Sub

    Private Sub ProcessToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProcessToolStripMenuItem.Click
        CreateProcess()
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        RefreshTab()
    End Sub

    Private Sub ManageToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ManageToolStripMenuItem.Click
        Dim strComputer As String = TextBoxComputer.Text
        Dim strCommand As String
        Dim strWinDir As String
        strWinDir = Environment.GetEnvironmentVariable("SystemRoot")

        If IsLocalComputer(strComputer) Then
            strCommand = "mmc " & strWinDir & "\system32\compmgmt.msc"
        Else
            strCommand = "mmc " & strWinDir & "\system32\compmgmt.msc /computer=" & strComputer
        End If


        If CheckBoxAlternateCredentials.Checked Then
            If strUsername = "" Or strPassword = "" Then
                MsgBox("Please set the credentials.", MsgBoxStyle.Information, "Credentials not set.")
                Exit Sub
            End If

            Shell("runas /user:" & strUsername & " " & Chr(34) & strCommand & Chr(34), AppWinStyle.NormalFocus)

        Else
            Shell(strCommand, AppWinStyle.NormalFocus)
        End If
    End Sub

    Private Sub RemoteAssistToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoteAssistToolStripMenuItem.Click
        Dim strWindir As String
        strWindir = Environment.GetEnvironmentVariable("SystemRoot")

        Dim strComputer As String = TextBoxComputer.Text
        If IsLocalComputer(strComputer) Then
            MsgBox("Please enter a REMOTE host!", MsgBoxStyle.Exclamation)
        Else
            If System.IO.File.Exists(strWindir & "\System32\msra.exe") Then
                'If Vista or Higher...

                Try
                    'Fails with UAC!  msra.exe must run in the non-elevated token
                    'Shell("cmd /s /k ""msra.exe /offerRA " & strComputer & Chr(34), AppWinStyle.NormalFocus)

                    Dim proc As Process = New Process
                    proc.StartInfo.UseShellExecute = True
                    proc.StartInfo.WindowStyle = ProcessWindowStyle.Normal
                    proc.StartInfo.WorkingDirectory = strWindir & "\System32"

                    ' launch non-elevated
                    proc.StartInfo.FileName = "msra.exe"
                    proc.StartInfo.Verb = ""                ' important!
                    proc.StartInfo.Arguments = "/offerRA " & strComputer

                    proc.Start()

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation)
                End Try

            Else
                'If XP...
                rewriteHTM()
                Shell("reg add ""\\" & strComputer & "\HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services"" /v fAllowUnsolicited /t REG_DWORD /d 1", AppWinStyle.Hide, True)
                Shell("reg add ""\\" & strComputer & "\HKLM\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services"" /v fAllowUnsolicitedFullControl /t REG_DWORD /d 1", AppWinStyle.Hide, True)
                Shell(strWindir & "\pchealth\helpctr\binaries\helpctr.exe -FromStartHelp -url hcp://CN=Microsoft%20Corporation,L=Redmond,S=Washington,C=US/Remote%20Assistance/Escalation/Unsolicited/unsolicitedrcui.htm?" & strComputer, AppWinStyle.NormalFocus)
            End If


        End If
    End Sub

    Private Sub RemoteDesktopToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoteDesktopToolStripMenuItem.Click
        Dim strComputer As String = TextBoxComputer.Text
        If IsLocalComputer(strComputer) Then
            MsgBox("Please enter a REMOTE host!", MsgBoxStyle.Exclamation)
        Else
            Shell("mstsc.exe /v:" & strComputer & " /f", AppWinStyle.NormalFocus)
        End If
    End Sub

    Private Sub RegeditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegeditToolStripMenuItem.Click
        Dim strComputer As String = TextBoxComputer.Text
        Dim strCommand As String
        'Dim strWindir
        'strWindir = Environment.GetEnvironmentVariable("SystemRoot")
        strCommand = "regedit"

        If CheckBoxAlternateCredentials.CheckState = CheckState.Checked Then
            'MsgBox("Sorry, cannot use RunAs option", MsgBoxStyle.Information)

            Shell("runas /user:" & strUsername & " " & Chr(34) & strCommand & Chr(34), AppWinStyle.NormalFocus)
        Else
            Shell(strCommand, AppWinStyle.NormalFocus)
        End If

        If Not (strComputer = "LOCALHOST" Or strComputer = "" Or strComputer = "127.0.0.1") Then
            Dim reg As New RemoteReg
            reg.SelectComputer(strComputer)
        End If
    End Sub

    Private Sub ExploreToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExploreToolStripMenuItem.Click
        Dim strComputer As String = TextBoxComputer.Text
        If IsLocalComputer(strComputer) Then
            Shell("explorer.exe /e,c:\", AppWinStyle.NormalFocus)
        Else
            Shell("explorer.exe /e,\\" & TextBoxComputer.Text & "\c$", AppWinStyle.NormalFocus)
        End If
    End Sub


    Private Sub LogOffUserToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogOffUserToolStripMenuItem.Click
        Win32Shutdown("Logoff")
    End Sub

    Private Sub ShutdownToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShutdownToolStripMenuItem.Click
        Dim result As MsgBoxResult
        result = MsgBox("Shutdown using forceful shutdown (potential data loss)?", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "Forceful shutdown?")
        Select Case result
            Case MsgBoxResult.Yes
                RebootUsingShutdownExe(False)
            Case MsgBoxResult.No
                Win32Shutdown("Shutdown")
            Case Else
                MsgBox("Action cancelled.", MsgBoxStyle.Information)
        End Select
    End Sub

    Private Sub RebootToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RebootToolStripMenuItem.Click
        Dim result As MsgBoxResult
        result = MsgBox("Reboot using forceful shutdown (potential data loss)?", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "Forceful reboot?")
        Select Case result
            Case MsgBoxResult.Yes
                RebootUsingShutdownExe(True)
            Case MsgBoxResult.No
                Win32Shutdown("Reboot")
            Case Else
                MsgBox("Action cancelled.", MsgBoxStyle.Information)
        End Select
    End Sub

    Private Sub ServiceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ServiceToolStripMenuItem.Click
        FormServiceAdd.ShowDialog()
    End Sub



    Private Sub CopyToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripMenuItem1.Click
        Dim strAllSoftware As String = ""

        For Each itm As ListViewItem In lstSoftware.Items
            strAllSoftware += itm.Text & vbCrLf
        Next
        Clipboard.SetText(strAllSoftware)

    End Sub



    Private Sub CopyAsTable(ByRef sender As System.Object)
        Dim strip As ContextMenuStrip = DirectCast(sender, ToolStripMenuItem).Owner
        Dim theListView As ListView = strip.SourceControl

        Dim strAllText As String = ""

        'For Each colTitle As ColumnHeader In ListViewUsers.Columns

        For Each col As ColumnHeader In theListView.Columns
            strAllText += col.Text & vbTab
        Next
        If strAllText <> "" Then strAllText += vbCrLf

        For Each itm As ListViewItem In theListView.Items
            'strAllSoftware += itm.Text & vbCrLf
            Dim row As String = ""
            For Each subitm As ListViewItem.ListViewSubItem In itm.SubItems
                row += subitm.Text & vbTab
            Next
            If row <> "" Then strAllText += row & vbCrLf

        Next
        Clipboard.SetText(strAllText)
    End Sub

    Private Sub CopyAsTableToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyAsTableToolStripMenuItem.Click
        CopyAsTable(sender)
    End Sub

    Private Sub CopyAsTableToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyAsTableToolStripMenuItem1.Click
        CopyAsTable(sender)
    End Sub

    Private Sub FindToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindToolStripMenuItem.Click
        FindText.ShowDialog()
    End Sub


    Private Sub ShowCPUToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowCPUToolStripMenuItem.Click
        Dim idx As Integer = GetColumnIndex(lstProcesses, "CPU")
        If ShowCPUToolStripMenuItem.Checked Then
            lstProcesses.Columns(idx).Width = 60
        Else
            lstProcesses.Columns(idx).Width = 0
        End If
    End Sub
End Class

