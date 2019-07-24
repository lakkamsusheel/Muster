Imports Infragistics.Shared
Imports Infragistics.Win.UltraWinGrid
Imports System.Data.SqlClient

Public Class RedTagProhibitionChangeManager
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    Public Sub New(ByVal _data As Data.DataSet)
        Me.New()
        _RedTagOCE = _data
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
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents BtnUnprohibitAll As System.Windows.Forms.Button
    Friend WithEvents btnProhibitAll As System.Windows.Forms.Button
    Friend WithEvents ugBeingRedTagged As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugBeingRemovedFromRedTag As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.BtnUnprohibitAll = New System.Windows.Forms.Button
        Me.btnProhibitAll = New System.Windows.Forms.Button
        Me.ugBeingRedTagged = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.ugBeingRemovedFromRedTag = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.ugBeingRedTagged, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ugBeingRemovedFromRedTag, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.ugBeingRedTagged)
        Me.Panel1.Location = New System.Drawing.Point(8, 40)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(872, 112)
        Me.Panel1.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.ugBeingRemovedFromRedTag)
        Me.Panel2.Location = New System.Drawing.Point(8, 200)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(872, 112)
        Me.Panel2.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(232, 16)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Tanks for C& E Sites Just Red Tagged"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 184)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(280, 16)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Tanks for C& E Sites Recently off red Tag Status"
        '
        'BtnUnprohibitAll
        '
        Me.BtnUnprohibitAll.Location = New System.Drawing.Point(256, 176)
        Me.BtnUnprohibitAll.Name = "BtnUnprohibitAll"
        Me.BtnUnprohibitAll.Size = New System.Drawing.Size(112, 23)
        Me.BtnUnprohibitAll.TabIndex = 4
        Me.BtnUnprohibitAll.Text = "Unprohibit All"
        '
        'btnProhibitAll
        '
        Me.btnProhibitAll.Location = New System.Drawing.Point(256, 16)
        Me.btnProhibitAll.Name = "btnProhibitAll"
        Me.btnProhibitAll.Size = New System.Drawing.Size(112, 23)
        Me.btnProhibitAll.TabIndex = 5
        Me.btnProhibitAll.Text = "Prohibit All"
        '
        'ugBeingRedTagged
        '
        Me.ugBeingRedTagged.Cursor = System.Windows.Forms.Cursors.Default
        Appearance1.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugBeingRedTagged.DisplayLayout.Override.CellAppearance = Appearance1
        Me.ugBeingRedTagged.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Appearance2.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugBeingRedTagged.DisplayLayout.Override.RowAppearance = Appearance2
        Me.ugBeingRedTagged.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugBeingRedTagged.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugBeingRedTagged.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugBeingRedTagged.Location = New System.Drawing.Point(0, 0)
        Me.ugBeingRedTagged.Name = "ugBeingRedTagged"
        Me.ugBeingRedTagged.Size = New System.Drawing.Size(868, 108)
        Me.ugBeingRedTagged.TabIndex = 1
        '
        'ugBeingRemovedFromRedTag
        '
        Me.ugBeingRemovedFromRedTag.Cursor = System.Windows.Forms.Cursors.Default
        Appearance3.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugBeingRemovedFromRedTag.DisplayLayout.Override.CellAppearance = Appearance3
        Me.ugBeingRemovedFromRedTag.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Appearance4.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugBeingRemovedFromRedTag.DisplayLayout.Override.RowAppearance = Appearance4
        Me.ugBeingRemovedFromRedTag.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugBeingRemovedFromRedTag.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugBeingRemovedFromRedTag.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugBeingRemovedFromRedTag.Location = New System.Drawing.Point(0, 0)
        Me.ugBeingRemovedFromRedTag.Name = "ugBeingRemovedFromRedTag"
        Me.ugBeingRemovedFromRedTag.Size = New System.Drawing.Size(868, 108)
        Me.ugBeingRemovedFromRedTag.TabIndex = 1
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(8, 336)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "Cancel"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(776, 336)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(104, 23)
        Me.btnSave.TabIndex = 7
        Me.btnSave.Text = "Save Changes"
        '
        'RedTagProhibitionChangeManager
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(896, 365)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnProhibitAll)
        Me.Controls.Add(Me.BtnUnprohibitAll)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "RedTagProhibitionChangeManager"
        Me.Text = "RedTagProhibitionChangeManager"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        CType(Me.ugBeingRedTagged, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ugBeingRemovedFromRedTag, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "private members"
    Private _RedTagOCE As Data.DataSet
#End Region

#Region "Form Events"

    Sub LoadForm(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

        If Not _RedTagOCE Is Nothing AndAlso _RedTagOCE.Tables.Count > 0 Then
            ugBeingRedTagged.DataSource = _RedTagOCE.Tables(0)

            If _RedTagOCE.Tables.Count > 1 Then
                Me.ugBeingRemovedFromRedTag.DataSource = _RedTagOCE.Tables(1)
            End If
        End If

    End Sub

#End Region

#Region "CheckChange Logic"
    Private Sub SetProhib(ByVal facid As Integer, ByVal tankid As Integer, ByVal value As Boolean)



        Dim LocalUserSettings As Microsoft.Win32.Registry
        Dim conSQLConnection As SqlConnection
        Dim cmdSQLCommand As SqlCommand
        Try
            conSQLConnection = New SqlConnection
            cmdSQLCommand = New SqlCommand

            conSQLConnection.ConnectionString = LocalUserSettings.CurrentUser.GetValue("MusterSQLConnection")

            conSQLConnection.Open()
            cmdSQLCommand.Connection = conSQLConnection

            If (value) Then

                ' cmdSQLCommand.CommandText = "Delete tblReg_Prohibition where Tank_id = " + tankid.ToString
                cmdSQLCommand.CommandText = "Update tblReg_tankplus set prohibition = 0, RevokeReason = NULL, RevokeDate = NULL where TankId = " + tankid.ToString
                cmdSQLCommand.ExecuteNonQuery()

                'cmdSQLCommand.CommandText = "Insert into tblReg_Prohibition values(" + facid.ToString + "," + tankid.ToString + ",'" + Date.Today.Now + "')"
                cmdSQLCommand.CommandText = "Update tblReg_tankplus set prohibition = 1, RevokeReason = NULL, RevokeDate = '" + Date.Today.Now + "' where TankID = " + tankid.ToString
                cmdSQLCommand.ExecuteNonQuery()

            Else
                'cmdSQLCommand.CommandText = "Delete tblReg_Prohibition where Facility_id = " + facid.ToString + " and Tank_id = " + tankid.ToString
                cmdSQLCommand.CommandText = "Update tblReg_tankplus set prohibition = 0, RevokeReason = NULL, RevokeDate = NULL where TankId = " + tankid.ToString
                cmdSQLCommand.ExecuteNonQuery()
            End If

        Catch ex As Exception
            Throw ex
        Finally
            If Not conSQLConnection Is Nothing Then

                If conSQLConnection.State = ConnectionState.Open Then
                    conSQLConnection.Close()
                End If

                conSQLConnection.Dispose()
            End If

            If Not cmdSQLCommand Is Nothing Then
                cmdSQLCommand.Dispose()
            End If
        End Try


    End Sub

#End Region

#Region "Grid Events"


    Sub RedTagCheckAll(ByVal sender As Object, ByVal e As EventArgs) Handles btnProhibitAll.Click

        For Each r As Infragistics.Win.UltraWinGrid.UltraGridRow In ugBeingRedTagged.Rows()
            r.Cells("PROHIBITED").Value = True
        Next


    End Sub

    Sub RedTagUnCheckAll(ByVal sender As Object, ByVal e As EventArgs) Handles BtnUnprohibitAll.Click

        For Each r As Infragistics.Win.UltraWinGrid.UltraGridRow In Me.ugBeingRemovedFromRedTag.Rows
            r.Cells("PROHIBITED").Value = False
        Next

    End Sub

    Sub Save(ByVal sender As Object, ByVal e As EventArgs) Handles btnSave.Click

        For Each r As Infragistics.Win.UltraWinGrid.UltraGridRow In ugBeingRedTagged.Rows()
            Me.Setprohib(r.Cells("ID").Value, r.Cells("TANK_ID").Value, r.Cells("PROHIBITED").Value)
        Next
        For Each r2 As Infragistics.Win.UltraWinGrid.UltraGridRow In Me.ugBeingRemovedFromRedTag.Rows
            Me.Setprohib(r2.Cells("ID").Value, r2.Cells("TANK_ID").Value, r2.Cells("PROHIBITED").Value)
        Next

        MsgBox("Prohibition has been saved")

        Me.Close()

    End Sub

    Sub CloseForm(ByVal sender As Object, ByVal e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub SetGrid(ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs)
        e.Layout.UseFixedHeaders = True
        e.Layout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        e.Layout.Bands(0).Columns("PROHIBITED").CellActivation = Activation.AllowEdit
        e.Layout.Bands(0).Columns("OCE OWNER").CellActivation = Activation.NoEdit
        e.Layout.Bands(0).Columns("ID").CellActivation = Activation.NoEdit
        e.Layout.Bands(0).Columns("NAME").CellActivation = Activation.NoEdit
        e.Layout.Bands(0).Columns("TANK #").CellActivation = Activation.NoEdit
        e.Layout.Bands(0).Columns("STATUS").CellActivation = Activation.NoEdit

        e.Layout.Override.FixedHeaderIndicator = Infragistics.Win.UltraWinGrid.FixedHeaderIndicator.None
        e.Layout.Bands(0).Columns("PROHIBITED").Header.Fixed = True
        e.Layout.Bands(0).Columns("NAME").Header.Fixed = True
        e.Layout.Bands(0).Columns("OCE OWNER").Header.Fixed = True


        e.Layout.Override.RowSizing = RowSizing.Fixed


    End Sub

    Private Sub ugBeingRedTagged_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugBeingRedTagged.InitializeLayout
        SetGrid(e)
    End Sub

    Private Sub ugBeingUnRedTagged_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugBeingRemovedFromRedTag.InitializeLayout
        SetGrid(e)
    End Sub

#End Region


End Class
