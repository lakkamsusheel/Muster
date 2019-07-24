Public Class FavSearchDialog
    '-------------------------------------------------------------------------------
    ' MUSTER.MUSTER.ShowComment.vb
    '   Provides the mechanism for managing Favorite Searches.
    '
    ' Copyright (C) 2004 CIBER, Inc.
    ' All rights reserved.
    '
    ' Release   Initials    Date        Description
    '  1.0        ??      8/??/04    Original class definition.
    '  1.1        JC      1/02/04    Changed AppUser.UserName to AppUser.ID to
    '                                  accomodate new use of pUser by application.
    '                                 
    '-------------------------------------------------------------------------------
    Inherits System.Windows.Forms.Form
    Friend frmParent As Form

    Public Event SaveFavoriteSearch(ByVal strFavSearchName As String)
    Private oAdvanceSearch As Muster.BusinessLogic.pAdvancedSearch


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub
    Public Sub New(ByRef oAdsearch As Muster.BusinessLogic.pAdvancedSearch)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        oAdvanceSearch = oAdsearch

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
    Friend WithEvents lblFav_Search_Name As System.Windows.Forms.Label
    Friend WithEvents txtFav_Search_Name As System.Windows.Forms.TextBox
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents chkPublic As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblFav_Search_Name = New System.Windows.Forms.Label
        Me.txtFav_Search_Name = New System.Windows.Forms.TextBox
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.chkPublic = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'lblFav_Search_Name
        '
        Me.lblFav_Search_Name.Location = New System.Drawing.Point(24, 24)
        Me.lblFav_Search_Name.Name = "lblFav_Search_Name"
        Me.lblFav_Search_Name.Size = New System.Drawing.Size(104, 16)
        Me.lblFav_Search_Name.TabIndex = 0
        Me.lblFav_Search_Name.Text = "Search Name"
        '
        'txtFav_Search_Name
        '
        Me.txtFav_Search_Name.Location = New System.Drawing.Point(128, 21)
        Me.txtFav_Search_Name.Name = "txtFav_Search_Name"
        Me.txtFav_Search_Name.Size = New System.Drawing.Size(304, 20)
        Me.txtFav_Search_Name.TabIndex = 1
        Me.txtFav_Search_Name.Text = ""
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(184, 64)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 2
        Me.btnSave.Text = "Save"
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(264, 64)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        '
        'chkPublic
        '
        Me.chkPublic.Location = New System.Drawing.Point(440, 24)
        Me.chkPublic.Name = "chkPublic"
        Me.chkPublic.Size = New System.Drawing.Size(104, 16)
        Me.chkPublic.TabIndex = 4
        Me.chkPublic.Text = "Public"
        '
        'FavSearchDialog
        '
        Me.AcceptButton = Me.btnSave
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(592, 102)
        Me.Controls.Add(Me.chkPublic)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.txtFav_Search_Name)
        Me.Controls.Add(Me.lblFav_Search_Name)
        Me.Name = "FavSearchDialog"
        Me.Text = "FavSearchDialog"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If txtFav_Search_Name.Text = String.Empty Then
                MessageBox.Show("Please Enter Valid Search Name.")
                txtFav_Search_Name.Focus()
                Exit Sub
            End If
            Dim oParentInfo As MUSTER.Info.FavSearchParentInfo
            If CType(frmParent, AdvancedSearch).cmbFavoriteSearches.SelectedIndex <> -1 Then
                oParentInfo = oAdvanceSearch.GetParentByID(CType(frmParent, AdvancedSearch).cmbFavoriteSearches.SelectedValue)
            Else
                oParentInfo = oAdvanceSearch.GetParentByID(CType(CType(frmParent, AdvancedSearch).oAdvanceSearch.ID, System.Int64))
            End If
            If IsNothing(oParentInfo) Then
                oParentInfo = New MUSTER.Info.FavSearchParentInfo
                oAdvanceSearch.AddParent(oParentInfo)
                'oAdvanceSearch.Retrieve(0)
            Else
                If oParentInfo.Name <> txtFav_Search_Name.Text Then
                    Dim drow As DataRow
                    'oParentInfo = New MUSTER.Info.FavSearchParentInfo
                    'oAdvanceSearch.AddParent(oParentInfo)
                    ''oAdvanceSearch.Retrieve(0)
                    For Each drow In CType(CType(frmParent, AdvancedSearch).ugSearchByLookFor.DataSource, DataTable).Rows
                        Dim oLocalChild As New MUSTER.Info.FavSearchChildInfo
                        oAdvanceSearch.AddChild(oLocalChild)
                        oLocalChild.CriterionName = drow("CRITERION_NAME")
                        oLocalChild.CriterionValue = drow("CRITERION_VALUE")
                        oLocalChild.CriterionDataType = drow("CRITERION_DATA_TYPE")
                        oLocalChild.Order = drow("CRITERION_ORDER")
                    Next
                End If
            End If
            oAdvanceSearch.User = MusterContainer.AppUser.ID
            oAdvanceSearch.IsPublic = Me.chkPublic.Checked
            oAdvanceSearch.Search_Type = CType(frmParent, AdvancedSearch).cmbSearchType.Text

            oAdvanceSearch.Name = txtFav_Search_Name.Text
            RaiseEvent SaveFavoriteSearch(txtFav_Search_Name.Text)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            Me.Close()
            Me.Dispose()
        End Try
    End Sub
    Private Sub FavSearchDialog_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Try
            Me.StartPosition = FormStartPosition.CenterParent

        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Dispose()
    End Sub

    Private Sub chkPublic_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPublic.CheckedChanged

    End Sub
End Class
