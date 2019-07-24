Imports System
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic

'  1.0        AN      02/10/05    Integrated AppFlags new object model

Public Class MusterAbout
    Inherits System.Windows.Forms.Form
    Dim MyFrm As MusterContainer
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub
    Public Sub New(ByRef frmCaller As MusterContainer)
        MyBase.New()
        InitializeComponent()
        MyFrm = frmCaller
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
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lvwAppVersions As System.Windows.Forms.ListView
    Friend WithEvents Resource As System.Windows.Forms.ColumnHeader
    Friend WithEvents Version As System.Windows.Forms.ColumnHeader
    Friend WithEvents Compatibility As System.Windows.Forms.ColumnHeader
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.lvwAppVersions = New System.Windows.Forms.ListView
        Me.Resource = New System.Windows.Forms.ColumnHeader
        Me.Version = New System.Windows.Forms.ColumnHeader
        Me.Compatibility = New System.Windows.Forms.ColumnHeader
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(256, 24)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "MUSTER Application Resources"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(256, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Currently installed services versions :"
        '
        'lvwAppVersions
        '
        Me.lvwAppVersions.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Resource, Me.Version, Me.Compatibility})
        Me.lvwAppVersions.Location = New System.Drawing.Point(16, 72)
        Me.lvwAppVersions.Name = "lvwAppVersions"
        Me.lvwAppVersions.Size = New System.Drawing.Size(392, 176)
        Me.lvwAppVersions.TabIndex = 3
        Me.lvwAppVersions.View = System.Windows.Forms.View.Details
        '
        'Resource
        '
        Me.Resource.Text = "Resource"
        Me.Resource.Width = 150
        '
        'Version
        '
        Me.Version.Text = "Version"
        Me.Version.Width = 144
        '
        'Compatibility
        '
        Me.Compatibility.Text = "Compatibility"
        Me.Compatibility.Width = 86
        '
        'MusterAbout
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(424, 266)
        Me.Controls.Add(Me.lvwAppVersions)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "MusterAbout"
        Me.Text = "About Muster"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub MusterAbout_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim lvwItem As New ListViewItem
        Dim FileList As System.IO.Directory
        Dim FileNames As String()
        Dim FileName As String
        Dim myAssemblyName As AssemblyName
        Try
            FileNames = FileList.GetFiles(Application.StartupPath, "muster.exe")
            For Each FileName In FileNames
                lvwItem = New ListViewItem
                myAssemblyName = AssemblyName.GetAssemblyName(FileName)
                lvwItem.Text = myAssemblyName.Name
                lvwItem.SubItems.Add(myAssemblyName.Version.ToString)
                lvwItem.SubItems.Add(myAssemblyName.VersionCompatibility)
                lvwAppVersions.Items.Add(lvwItem)
            Next

            Try
                lvwItem = New ListViewItem
                lvwItem.Text = "User ID"
                'lvwItem.SubItems.Add(MyFrm.AppSemaphores.Retrieve("User ID", "").Value.ToString)
                'lvwItem.SubItems.Add(MyFrm.AppSemaphores.GetValuePair("User ID", ""))
                lvwItem.SubItems.Add(MusterContainer.AppUser.ID)
                lvwAppVersions.Items.Add(lvwItem)
            Catch ex As Exception
                lvwItem.SubItems.Add("Default")
                lvwAppVersions.Items.Add(lvwItem)
            End Try

            Try
                lvwItem = New ListViewItem
                lvwItem.Text = "DataBase"
                'lvwItem.SubItems.Add(MyFrm.AppSemaphores.Retrieve("Initial Catalog", "").Value.ToString)
                lvwItem.SubItems.Add(MyFrm.AppSemaphores.GetValuePair("Initial Catalog", ""))
                lvwAppVersions.Items.Add(lvwItem)
            Catch ex As Exception
                lvwItem.SubItems.Add("Default")
                lvwAppVersions.Items.Add(lvwItem)
            End Try

            Try
                lvwItem = New ListViewItem
                lvwItem.Text = "Data Source"
                'lvwItem.SubItems.Add(MyFrm.AppSemaphores.Retrieve("Data Source", "").Value.ToString)
                lvwItem.SubItems.Add(MyFrm.AppSemaphores.GetValuePair("Data Source", ""))
                lvwAppVersions.Items.Add(lvwItem)
            Catch ex As Exception
                lvwItem.SubItems.Add("Default")
                lvwAppVersions.Items.Add(lvwItem)
            End Try

            'Try
            '    lvwItem = New ListViewItem
            '    lvwItem.Text = "Report Location"
            '    'lvwItem.SubItems.Add(MyFrm.AppSemaphores.Retrieve("Report Location", "").Value.ToString)
            '    lvwItem.SubItems.Add(MyFrm.AppSemaphores.GetValuePair("Report Location", ""))
            '    lvwAppVersions.Items.Add(lvwItem)
            'Catch ex As Exception
            '    lvwItem.SubItems.Add("Default")
            '    lvwAppVersions.Items.Add(lvwItem)
            'End Try

            'FileNames = FileList.GetFiles(Application.StartupPath, "*.dll")
            'For Each FileName In FileNames
            '    lvwItem = New ListViewItem
            '    myAssemblyName = AssemblyName.GetAssemblyName(FileName)
            '    lvwItem.Text = myAssemblyName.Name
            '    lvwItem.SubItems.Add(myAssemblyName.Version.ToString)
            '    lvwItem.SubItems.Add(myAssemblyName.VersionCompatibility)
            '    lvwAppVersions.Items.Add(lvwItem)
            'Next
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub MusterAbout_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        lvwAppVersions.Width = Me.Width - 36
    End Sub
End Class
