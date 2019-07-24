Imports System.Configuration
Imports Microsoft.Win32.Registry
Imports System
Imports System.Reflection

Public Class Logon
    '-------------------------------------------------------------------------------
    ' MUSTER.MUSTER.MusterContainer
    '   Provides the logon form for all operations in the application.
    '
    ' Copyright (C) 2004 CIBER, Inc.
    ' All rights reserved.
    '
    ' Release   Initials    Date        Description
    '  1.0        ??      ??/??/??    Original class definition.
    '  1.1        JC      12/31/04    Changed over to Muster.BussinessLogic.pUser
    '  1.2        JVC2    01/19/05    Changed check for usr.name during logon to
    '                                   also check for string.empty as a failure
    '                                   indicator.
    '  1.3        AN      02/10/05    Integrated AppFlags new object model
    '  1.4        JVC2    03/07/05    Changed SSPI check to true on load
    '  1.5        JVC2    03/14/05    Removed login name and password from logon window
    '                                   and added use of System.Security.Principal to
    '                                   determine which user ID to accept.
    '  1.6        JVC     06/02/05    Added features to allow logon to DB Server if
    '                                   Integrated Security is not checked.  Only works
    '                                   when app is used in DEBUG mode.
    '  1.7        JVC     06/07/05    Added check for user deleted when processing logon.
    '                                   If user is deleted (inactive) then display 
    '                                   invalid logon message
    '-------------------------------------------------------------------------------
    '
    Inherits System.Windows.Forms.Form

    Dim App As ConfigurationSettings
    Dim MyFrm As MusterContainer
    Dim LocalUserSettings As Microsoft.Win32.Registry
    Dim MyConnectionString As String
    Dim Usr As MUSTER.BusinessLogic.pUser
    Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long
    Private WithEvents winChangePassword As ChangePassword
    Dim bPasswordCancelFlag As Boolean = False
    Dim returnVal As String = String.Empty

#Region " Windows Form Designer generated code "

    Public Sub New(ByRef frmCaller As MusterContainer)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        MyFrm = frmCaller

        '
        '  If the connection string is not found in the registry, then add it...
        '
        MyConnectionString = LocalUserSettings.CurrentUser.GetValue("MusterSQLConnection", "Not Available")
        If MyConnectionString = "Not Available" Then
            LocalUserSettings.CurrentUser.SetValue("MusterSQLConnection", App.AppSettings("SQLConnectionString").ToString)
        End If

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
    Friend WithEvents txtLogonName As System.Windows.Forms.TextBox
    Friend WithEvents lblLogonID As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnLogon As System.Windows.Forms.Button
    Friend WithEvents txtUserID As System.Windows.Forms.TextBox
    Friend WithEvents txtInitialCatalog As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents chkIntegratedSecurity As System.Windows.Forms.CheckBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents grpUserSpecs As System.Windows.Forms.GroupBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblPassword As System.Windows.Forms.Label
    Friend WithEvents txtLogonPwd As System.Windows.Forms.TextBox
    Friend WithEvents cboServers As System.Windows.Forms.ComboBox
    Friend WithEvents ChkChangePsw As System.Windows.Forms.CheckBox
    Friend WithEvents lblDBLogin As System.Windows.Forms.Label
    Friend WithEvents lblDBPwd As System.Windows.Forms.Label
    Friend WithEvents txtDBLogin As System.Windows.Forms.TextBox
    Friend WithEvents txtDBPwd As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtLogonName = New System.Windows.Forms.TextBox
        Me.lblLogonID = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnLogon = New System.Windows.Forms.Button
        Me.grpUserSpecs = New System.Windows.Forms.GroupBox
        Me.txtDBPwd = New System.Windows.Forms.TextBox
        Me.lblDBPwd = New System.Windows.Forms.Label
        Me.txtDBLogin = New System.Windows.Forms.TextBox
        Me.cboServers = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.chkIntegratedSecurity = New System.Windows.Forms.CheckBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtInitialCatalog = New System.Windows.Forms.TextBox
        Me.txtUserID = New System.Windows.Forms.TextBox
        Me.lblDBLogin = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.lblPassword = New System.Windows.Forms.Label
        Me.txtLogonPwd = New System.Windows.Forms.TextBox
        Me.ChkChangePsw = New System.Windows.Forms.CheckBox
        Me.grpUserSpecs.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtLogonName
        '
        Me.txtLogonName.Location = New System.Drawing.Point(88, 48)
        Me.txtLogonName.Name = "txtLogonName"
        Me.txtLogonName.Size = New System.Drawing.Size(160, 20)
        Me.txtLogonName.TabIndex = 8
        Me.txtLogonName.Text = ""
        '
        'lblLogonID
        '
        Me.lblLogonID.Location = New System.Drawing.Point(12, 50)
        Me.lblLogonID.Name = "lblLogonID"
        Me.lblLogonID.Size = New System.Drawing.Size(56, 16)
        Me.lblLogonID.TabIndex = 0
        Me.lblLogonID.Text = "User ID :"
        Me.lblLogonID.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(12, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(248, 32)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Enter your user id in the text box below, select your data server connection and " & _
        "then click OK"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnLogon
        '
        Me.btnLogon.Location = New System.Drawing.Point(40, 240)
        Me.btnLogon.Name = "btnLogon"
        Me.btnLogon.Size = New System.Drawing.Size(80, 24)
        Me.btnLogon.TabIndex = 5
        Me.btnLogon.Text = "&OK"
        '
        'grpUserSpecs
        '
        Me.grpUserSpecs.Controls.Add(Me.txtDBPwd)
        Me.grpUserSpecs.Controls.Add(Me.lblDBPwd)
        Me.grpUserSpecs.Controls.Add(Me.txtDBLogin)
        Me.grpUserSpecs.Controls.Add(Me.cboServers)
        Me.grpUserSpecs.Controls.Add(Me.Label4)
        Me.grpUserSpecs.Controls.Add(Me.chkIntegratedSecurity)
        Me.grpUserSpecs.Controls.Add(Me.Label3)
        Me.grpUserSpecs.Controls.Add(Me.Label2)
        Me.grpUserSpecs.Controls.Add(Me.txtInitialCatalog)
        Me.grpUserSpecs.Controls.Add(Me.txtUserID)
        Me.grpUserSpecs.Controls.Add(Me.lblDBLogin)
        Me.grpUserSpecs.Controls.Add(Me.Label5)
        Me.grpUserSpecs.Location = New System.Drawing.Point(8, 136)
        Me.grpUserSpecs.Name = "grpUserSpecs"
        Me.grpUserSpecs.Size = New System.Drawing.Size(256, 96)
        Me.grpUserSpecs.TabIndex = 3
        Me.grpUserSpecs.TabStop = False
        Me.grpUserSpecs.Text = "Server Connection"
        '
        'txtDBPwd
        '
        Me.txtDBPwd.Location = New System.Drawing.Point(114, 148)
        Me.txtDBPwd.Name = "txtDBPwd"
        Me.txtDBPwd.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtDBPwd.Size = New System.Drawing.Size(144, 20)
        Me.txtDBPwd.TabIndex = 7
        Me.txtDBPwd.Text = "password"
        '
        'lblDBPwd
        '
        Me.lblDBPwd.Location = New System.Drawing.Point(1, 127)
        Me.lblDBPwd.Name = "lblDBPwd"
        Me.lblDBPwd.Size = New System.Drawing.Size(104, 16)
        Me.lblDBPwd.TabIndex = 0
        Me.lblDBPwd.Text = "DB Password"
        '
        'txtDBLogin
        '
        Me.txtDBLogin.Location = New System.Drawing.Point(104, 96)
        Me.txtDBLogin.Name = "txtDBLogin"
        Me.txtDBLogin.Size = New System.Drawing.Size(144, 20)
        Me.txtDBLogin.TabIndex = 6
        Me.txtDBLogin.Text = "sa"
        '
        'cboServers
        '
        Me.cboServers.Items.AddRange(New Object() {"OPCGW", "muster.deq.state.ms.us", "GARD-PROD"})
        Me.cboServers.Location = New System.Drawing.Point(106, 24)
        Me.cboServers.Name = "cboServers"
        Me.cboServers.Size = New System.Drawing.Size(144, 21)
        Me.cboServers.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(50, 28)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 16)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Server"
        '
        'chkIntegratedSecurity
        '
        Me.chkIntegratedSecurity.Location = New System.Drawing.Point(24, 72)
        Me.chkIntegratedSecurity.Name = "chkIntegratedSecurity"
        Me.chkIntegratedSecurity.Size = New System.Drawing.Size(192, 16)
        Me.chkIntegratedSecurity.TabIndex = 4
        Me.chkIntegratedSecurity.Text = "Use Integrated Security"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(50, 28)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 16)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Password"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(50, 27)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 16)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "User ID"
        Me.Label2.Visible = False
        '
        'txtInitialCatalog
        '
        Me.txtInitialCatalog.Location = New System.Drawing.Point(106, 48)
        Me.txtInitialCatalog.Name = "txtInitialCatalog"
        Me.txtInitialCatalog.Size = New System.Drawing.Size(144, 20)
        Me.txtInitialCatalog.TabIndex = 3
        Me.txtInitialCatalog.Text = "Muster_Prd"
        '
        'txtUserID
        '
        Me.txtUserID.Location = New System.Drawing.Point(106, 24)
        Me.txtUserID.Name = "txtUserID"
        Me.txtUserID.Size = New System.Drawing.Size(144, 20)
        Me.txtUserID.TabIndex = 0
        Me.txtUserID.Text = "CIBER"
        Me.txtUserID.Visible = False
        '
        'lblDBLogin
        '
        Me.lblDBLogin.Location = New System.Drawing.Point(3, 99)
        Me.lblDBLogin.Name = "lblDBLogin"
        Me.lblDBLogin.Size = New System.Drawing.Size(104, 16)
        Me.lblDBLogin.TabIndex = 0
        Me.lblDBLogin.Text = "DB Login"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Location = New System.Drawing.Point(-3, 52)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(102, 16)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Database (Catalog)"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(152, 240)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 24)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "&Cancel"
        '
        'lblPassword
        '
        Me.lblPassword.Location = New System.Drawing.Point(16, 80)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(56, 16)
        Me.lblPassword.TabIndex = 0
        Me.lblPassword.Text = "Password"
        Me.lblPassword.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtLogonPwd
        '
        Me.txtLogonPwd.Location = New System.Drawing.Point(88, 80)
        Me.txtLogonPwd.Name = "txtLogonPwd"
        Me.txtLogonPwd.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtLogonPwd.Size = New System.Drawing.Size(160, 20)
        Me.txtLogonPwd.TabIndex = 0
        Me.txtLogonPwd.Text = ""
        '
        'ChkChangePsw
        '
        Me.ChkChangePsw.Location = New System.Drawing.Point(32, 104)
        Me.ChkChangePsw.Name = "ChkChangePsw"
        Me.ChkChangePsw.Size = New System.Drawing.Size(184, 24)
        Me.ChkChangePsw.TabIndex = 1
        Me.ChkChangePsw.Text = "Change password after login."
        '
        'Logon
        '
        Me.AcceptButton = Me.btnLogon
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(272, 271)
        Me.ControlBox = False
        Me.Controls.Add(Me.grpUserSpecs)
        Me.Controls.Add(Me.ChkChangePsw)
        Me.Controls.Add(Me.lblPassword)
        Me.Controls.Add(Me.txtLogonPwd)
        Me.Controls.Add(Me.txtLogonName)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblLogonID)
        Me.Controls.Add(Me.btnLogon)
        Me.Name = "Logon"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Logon"
        Me.TopMost = True
        Me.grpUserSpecs.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnLogon_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogon.Click
        Dim SqlConn As New SqlClient.SqlConnection
        Dim strErr As String
        Dim myAssemblyName As AssemblyName

        Try
            MyFrm.Update()
            LockWindowUpdate(CLng(MyFrm.Handle.ToInt64))
            '
            ' Relocated so that new connection string is written to registry before
            '  attempt is made to establish connection.
            '
            Dim WorkConnStr As String
            'Dim dbServer, dbCatalog As String
            'dbServer = IIf(cboServers.SelectedText Is DBNull.Value, String.Empty, cboServers.SelectedText)
            'If cboServers.SelectedItem Is Nothing Then
            '    dbServer = String.Empty
            'Else
            '    dbServer = cboServers.GetItemText(cboServers.SelectedItem)
            'End If
            'If cboDatabaseCatalog.SelectedItem Is Nothing Then
            '    dbCatalog = String.Empty
            'Else
            '    dbCatalog = cboDatabaseCatalog.GetItemText(cboDatabaseCatalog.SelectedItem)
            'End If
            'dbCatalog = IIf(cboDatabaseCatalog.SelectedText Is DBNull.Value, String.Empty, cboDatabaseCatalog.SelectedText)
            'WorkConnStr = IIf(cboServers.Text = String.Empty, "", "Data Source=" & cboServers.Text & ";") & _
            '    IIf(txtInitialCatalog.Text = String.Empty, "", "Initial Catalog=" & txtInitialCatalog.Text & ";")

            If txtLogonName.Text = String.Empty Then
                strErr = "User ID, "
                'MsgBox("Please Enter the User ID.")
                'txtLogonName.Focus()
                'Exit Sub
            End If
            If txtLogonPwd.Text = String.Empty Then
                strErr = "Password, "
                'MsgBox("Please Enter the Password.")
                'txtLogonPwd.Focus()
                'Exit Sub
            End If
            If cboServers.Text = String.Empty Then
                strErr = "Server, "
            End If
            If txtInitialCatalog.Text = String.Empty Then
                strErr += "Database (Catolog), "
            End If
            If strErr <> String.Empty Then
                MsgBox("The following fields are required" + vbCrLf + strErr.Trim.TrimEnd(","), MsgBoxStyle.Exclamation)
                Exit Sub
            End If
            WorkConnStr = IIf(cboServers.Text = String.Empty, "", "Data Source=" & cboServers.Text & ";") & _
                            IIf(txtInitialCatalog.Text = String.Empty, "", "Initial Catalog=" & txtInitialCatalog.Text & ";")
            'WorkConnStr = "Data Source=" + dbServer + ";" + "Initial Catalog=" + dbCatalog + ";"
            'IIf(txtUserID.Text = String.Empty, "", "User ID=" & txtUserID.Text & ";") '& _
            '"Password=" & txtPassword.Text & ";"
            If chkIntegratedSecurity.Checked Then
                WorkConnStr += "Integrated Security=SSPI;"
            Else
                If cboServers.Text = "GARD-PROD" Then  'DBLogin = MusterApp
                    WorkConnStr += "user='MusterApp';password='8f1-4c9A';"
                    Me.txtDBLogin.Text = "MusterApp"
                Else
                    If cboServers.Text = "OPCGW" Then
                        WorkConnStr += "user='" + txtDBLogin.Text + "';password='password';"
                    Else
                        If cboServers.Text = "muster.deq.state.ms.us" Then
                            WorkConnStr += "user='" + txtDBLogin.Text + "';password='4b3dD60w';"
                        Else 'for local DB login
                            WorkConnStr += "user='" + txtDBLogin.Text + "';password='password';"
                        End If
                    End If
                End If

            End If
            WorkConnStr = WorkConnStr.Substring(0, WorkConnStr.Length)

            'LocalUserSettings.CurrentUser.SetValue("MusterSQLConnection", WorkConnStr)
            '
            'Next, check the connection to make certain it's valid
            '
            Try
                SqlConn.ConnectionString = WorkConnStr
                'SqlConn.ConnectionString = "Data Source=MUSTER-DEV\SQL2008R2;Initial Catalog=Sus_muster_temp;user='susheel';password='susheel123';"
                SqlConn.Open()
                SqlConn.Close()
            Catch ex As Exception
                MsgBox("Either server " & cboServers.Text & " cannot be located or database " & txtInitialCatalog.Text & " does not exist on the server." & vbCrLf & "Please Try Again!", MsgBoxStyle.Information & MsgBoxStyle.OKOnly, "Server Not Found!")
                Exit Sub
            End Try

            LocalUserSettings.CurrentUser.SetValue("MusterSQLConnection", WorkConnStr)
            LocalUserSettings.CurrentUser.SetValue("MusterLastLoginUser", txtLogonName.Text)

            '
            'Now, instantiate the pUser object for the requested logon
            '
            Usr = New MUSTER.BusinessLogic.pUser(txtLogonName.Text)
            'Usr.Retrieve(txtLogonName.Text)
            If Usr.Name Is Nothing Then
                MsgBox("Invalid login. Please try again!", MsgBoxStyle.Critical & MsgBoxStyle.OKOnly, "Invalid Login")
                Exit Sub
            Else
                If Usr.Name = String.Empty Or Usr.Deleted Then
                    MsgBox("Invalid login. Please try again!", MsgBoxStyle.Critical & MsgBoxStyle.OKOnly, "Invalid Login")
                    Exit Sub
                End If
            End If

            Usr.Password = txtLogonPwd.Text
            If Usr.VerifyPassword = False Then
                MsgBox("Password is Invalid.Please Enter Correct Password.")
                txtLogonPwd.Focus()
                Usr = Nothing
                Exit Sub
            End If

            If txtLogonPwd.Text.ToLower.IndexOf("password") >= 0 Then
                MsgBox("Your password is a temporary password.  Please supply your permanent password in the next window that appears.", MsgBoxStyle.Information & MsgBoxStyle.OKOnly, "Password Change Required")
                winChangePassword = New ChangePassword
                winChangePassword.SetUser(Usr)
                Me.Hide()
                winChangePassword.ShowDialog()
                'Added to Close the Logon Form.
                If bPasswordCancelFlag = True Then
                    Me.Dispose()
                    Me.Close()
                    Exit Sub
                End If

                Usr.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), Usr.UserKey, returnVal, Usr.ID, True)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
                Me.Show()
            End If

            If ChkChangePsw.Checked Then
                winChangePassword = New ChangePassword
                winChangePassword.SetUser(Usr)
                Me.Hide()
                winChangePassword.ShowDialog()
                'Added to Close the Logon Form.
                'If bPasswordCancelFlag = True Then
                '    Me.Dispose()
                '    Me.Close()
                '    Exit Sub
                'End If
                Usr.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), Usr.UserKey, returnVal, Usr.ID, True)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
                Me.Show()
            End If
            Dim dsVersion As New DataSet
            Dim FileNames As String()
            Dim FileName As String
            Dim FileList As System.IO.Directory
            dsVersion = Usr.RunSQLQuery("SELECT PROPERTY_NAME AS Version FROM dbo.tblSYS_PROPERTY_MASTER " & _
                                            "WHERE dbo.tblSYS_PROPERTY_MASTER.PROPERTY_TYPE_ID IN (" & _
                                            "SELECT PROPERTY_TYPE_ID FROM dbo.tblSYS_PROPERTY_TYPE " & _
                                            "WHERE dbo.tblSYS_PROPERTY_TYPE.PROPERTY_TYPE_NAME = 'DBVersion')")
            If (dsVersion.Tables(0).Rows.Count > 0) Then
                FileNames = FileList.GetFiles(Application.StartupPath, "muster.exe")
                For Each FileName In FileNames
                    myAssemblyName = AssemblyName.GetAssemblyName(FileName)
                    If myAssemblyName.Version.ToString <> Convert.ToString(dsVersion.Tables(0).Rows(0)(0)) Then
                        MsgBox("This version (" & myAssemblyName.Version.ToString & ") does not match the current system release version (" & Convert.ToString(dsVersion.Tables(0).Rows(0)(0)) & ")." & vbCrLf & "Please upgrade or contact your system administrator for help.", MsgBoxStyle.Information & MsgBoxStyle.OKOnly, "Version Conflict Notice")
                    End If
                Next
            End If

            ' Now, test to see if we can obtain a connection
            ' Declaring the data manager as New forces it to connect to the database
            MyFrm.AppUser = New MUSTER.BusinessLogic.pUser(txtLogonName.Text)

            Me.Dispose()
        Catch ex As Exception
            Dim frmErr As New ErrorReport(ex)
            frmErr.ShowDialog()
            Me.txtLogonName.Focus()
        Finally
            LockWindowUpdate(CLng(0))
        End Try
    End Sub
    'Function to Close the Application.
    Private Sub PasswordFormClosed() Handles winChangePassword.PasswordClosed
        bPasswordCancelFlag = True
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click

        Me.Dispose()

    End Sub

    Private Sub Logon_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim SqlConn As New SqlClient.SqlConnection
        'Dim index As Int16
        'Dim ServerKey As String
        ' adding local so inspectors can see the name in drop down list
        Dim testSwitchFile As String = String.Format("{0}\TestMuster.txt", Application.StartupPath)
        Dim setupIniFile As String = String.Format("{0}\MusterSetup.ini", Application.StartupPath)
        Dim flag As String
        cboServers.Items.Add(System.Net.Dns.GetHostName)
        Me.grpUserSpecs.Enabled = False

        If Not System.IO.File.Exists(testSwitchFile) Then
            If Not System.IO.File.Exists(setupIniFile) Then
                'For the first time that Muster is installed
                Dim file As New System.IO.StreamWriter(setupIniFile)
                file.WriteLine("0")
                flag = 0
                Threading.Thread.Sleep(500)
                file.Close()
                file = Nothing
            Else
                Dim iniFile As New System.IO.StreamReader(setupIniFile)
                flag = iniFile.ReadLine
                Threading.Thread.Sleep(500)
                iniFile.Close()
                iniFile = Nothing
            End If
        Else
            'Rename TestMuster.txt to MusterSetup.ini for only once
            If Not System.IO.File.Exists(setupIniFile) Then
                System.IO.File.Copy(testSwitchFile, setupIniFile)
            End If
            Dim iniFile As New System.IO.StreamReader(setupIniFile)
            flag = iniFile.ReadLine
            Threading.Thread.Sleep(500)
            iniFile.Close()
            iniFile = Nothing
            System.IO.File.Delete(testSwitchFile)
        End If

#If DEBUG Then
        Me.grpUserSpecs.Enabled = False

        txtLogonName.Text = "admin"
        txtLogonPwd.Text = "admin"
        'cboServers.Items.Add(System.Net.Dns.GetHostByName("localhost").HostName)
        cboServers.SelectedIndex = 0
        If chkIntegratedSecurity.Checked Then
            chkIntegratedSecurity.Checked = False
        Else
            chkIntegratedSecurity.Checked = True
            chkIntegratedSecurity.Checked = False
        End If
        cboServers.Text = "OPCGW"
        txtInitialCatalog.Text = "MUSTER_SIT"

        ' canTest = True

#Else

        cboServers.Text = "OPCGW"
        txtInitialCatalog.Text = "MUSTER_SIT"
#End If

        Dim getLastData As Boolean = False
        'Dim strSQL As String
        'Dim strConnStr As String = "Data Source=" + System.Net.Dns.GetHostName + ";Initial Catalog=master;user='sa';password='password';"

        If flag = "2" Then
            grpUserSpecs.Enabled = True
            getLastData = True

        ElseIf MyFrm.Inspector Then
            txtInitialCatalog.Text = "MUSTER_PRD"
            cboServers.Text = System.Net.Dns.GetHostName

        Else
            Me.grpUserSpecs.Enabled = False

            txtInitialCatalog.Text = "MUSTER_PRD"
            If flag = "0" Then
                cboServers.Text = "GARD-PROD"
            Else 'if flag = "1" local read only DB
                cboServers.Text = System.Net.Dns.GetHostName
                'strConnStr = strConnStr.Substring(0, strConnStr.Length)
                'LocalUserSettings.CurrentUser.SetValue("MusterSQLConnection", strConnStr)
                'strSQL = "USE master;EXEC sp_dboption 'Muster_Prd', 'read only', 'TRUE'"
                'SqlHelper.ExecuteNonQuery(strConnStr, CommandType.Text, strSQL)

            End If
            getLastData = True
        End If


        Me.Text += " " + MusterContainer.DomainUser.Identity.Name

        Dim lastLogin As Object = LocalUserSettings.CurrentUser.GetValue("MusterSQLConnection")
        Dim lastLoginUser As String = LocalUserSettings.CurrentUser.GetValue("MusterLastLoginUser", "")
        Dim lastLoginServer, lastLoginDB As String

        If flag = "2" And getLastData Then

            For Each str As String In lastLogin.ToString.Split(";")
                If str.StartsWith("Data Source") Then
                    lastLoginServer = str.Split("=")(1)
                ElseIf str.StartsWith("Initial Catalog") Then
                    lastLoginDB = str.Split("=")(1)
                End If
                If lastLoginServer <> String.Empty And lastLoginDB <> String.Empty Then Exit For
            Next

            cboServers.Text = IIf(lastLoginServer = String.Empty, cboServers.Text, lastLoginServer)
            ' cboServers.Text = System.Net.Dns.GetHostName
            txtInitialCatalog.Text = IIf(lastLoginDB = String.Empty, txtInitialCatalog.Text, lastLoginDB)

        End If

        txtLogonName.Text = IIf(lastLoginUser = String.Empty, txtLogonName.Text, lastLoginUser)


        If txtLogonName.Text <> String.Empty Then

            Me.BringToFront()

            txtLogonPwd.Focus()
        End If


        '        cboServers.Items.Clear()
        '        Me.Cursor = Windows.Forms.Cursors.AppStarting
        '        Try
        '            bolLoading = True
        '            Dim strSQL As String = "SELECT PROPERTY_ID, PROPERTY_NAME, PROPERTY_POSITION FROM tblsys_property_master WHERE PROPERTY_TYPE_ID IN (SELECT PROPERTY_TYPE_ID FROM tblSYS_PROPERTY_TYPE WHERE PROPERTY_TYPE_NAME LIKE '%DATABASESERVERS%') AND PROPERTY_ACTIVE = 'YES'"
        '            Dim ds As DataSet = oPropType.GetDS(strSQL)
        '            Dim dv As DataView
        '            dv = ds.Tables(0).DefaultView
        '            dv.Sort = "PROPERTY_POSITION"
        '            cboServers.DataSource = dv
        '            cboServers.ValueMember = "PROPERTY_ID"
        '            cboServers.DisplayMember = "PROPERTY_NAME"
        '            cboServers.SelectedIndex = -1
        '            cboDatabaseCatalog.DataSource = Nothing
        '            bolLoading = False
        '        Catch ex As Exception
        '            Dim frmErr As New ErrorReport(ex)
        '            frmErr.ShowDialog()
        '            Me.txtLogonName.Focus()
        '            Exit Sub
        '        End Try
        'For index = 1 To 4
        '    ServerKey = "Server" & index.ToString
        '    If Not App.AppSettings.Get(ServerKey) Is Nothing Then
        '        SqlConn.ConnectionString = "Data Source=" & App.AppSettings.Get(ServerKey) & ";Integrated Security=SSPI;Timeout=5;"
        '        Try
        '            SqlConn.Open()
        '            cboServers.Items.Add(App.AppSettings.Get(ServerKey))
        '            SqlConn.Close()
        '        Catch ex As Exception
        '        Finally
        '        End Try

        '    End If
        'Next


    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkChangePsw.CheckedChanged

    End Sub

    Private Sub chkIntegratedSecurity_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkIntegratedSecurity.CheckedChanged
#If DEBUG Then
        If chkIntegratedSecurity.Checked Then
            grpUserSpecs.Height = 96
            Me.Height = 304
            btnCancel.Top = 240
            btnLogon.Top = 240
            txtDBLogin.Enabled = False
            txtDBPwd.Enabled = False
        Else
            Me.Height = 384
            btnCancel.Top = 296
            btnLogon.Top = 296
            grpUserSpecs.Height = 152
            txtDBLogin.Enabled = True
            txtDBPwd.Enabled = True
        End If
        Me.Refresh()

        Me.txtLogonPwd.Focus()
#End If
    End Sub

    'Private Sub cboServers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboServers.SelectedIndexChanged
    '    Try
    '        If bolLoading Then Exit Sub
    '        If cboServers.SelectedIndex = -1 Then
    '            cboDatabaseCatalog.DataSource = Nothing
    '        Else
    '            Dim strSQL As String = "SELECT PROPERTY_ID, PROPERTY_NAME, PROPERTY_POSITION FROM tblsys_property_master WHERE PROPERTY_ID IN (SELECT PROPERTY_ID_CHILD FROM tblSYS_PROPERTY_RELATION WHERE PROPERTY_ID_PARENT = " + cboServers.SelectedValue.ToString + ")  AND PROPERTY_ACTIVE = 'YES'"
    '            Dim ds As DataSet = oPropType.GetDS(strSQL)
    '            Dim dv As DataView
    '            If ds.Tables.Count > 0 Then
    '                dv = ds.Tables(0).DefaultView
    '                dv.Sort = "PROPERTY_POSITION"
    '                cboDatabaseCatalog.DataSource = dv
    '                cboDatabaseCatalog.ValueMember = "PROPERTY_ID"
    '                cboDatabaseCatalog.DisplayMember = "PROPERTY_NAME"
    '                cboDatabaseCatalog.SelectedIndex = 0
    '            Else
    '                cboDatabaseCatalog.DataSource = Nothing
    '            End If
    '        End If
    '    Catch ex As Exception
    '        Dim frmErr As New ErrorReport(ex)
    '        frmErr.ShowDialog()
    '        Me.cboServers.Focus()
    '    End Try
    'End Sub
End Class
