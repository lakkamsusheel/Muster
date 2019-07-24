' ========================================================================================
'
'   1.0         AN      02/10/05    Integrated AppFlags new object model
'   1.1         EN      02/11/05    Implemented the new object model 
'   1.2         AB      03/02/05    Corrected a number of bugs 	
'                                       - Address no longer shows non-printable characters
'	                                    - Now properly saving information when more than one pipe attached to a tank.
'	                                    - Will now refresh tank/pipe grid when selecting a facility cell other than cell(0)
'	                                    - 'Copy Current to Next' only updates necessary dates now
'	                                    - 'Apply to All' only updates necessary dates now
'   1.3         AB      03/15/05    Facility Grid now updates when the changes in Tank/Pipe are saved
'   2.17        HC      01/21/09    Integrated 10 new CAP fields
'   
' ========================================================================================
'
'
'

Imports Infragistics.Win
Imports Infragistics.Win.UltraWinGrid
Imports System.Data.SqlClient
Imports System.Text
Public Class CAPSignUp
    Inherits System.Windows.Forms.Form

#Region "User Defined Variables"
    Dim dsCapFac As New DataSet
    Dim dsCAPStatus As New DataSet
    Private WithEvents oOwner As MUSTER.BusinessLogic.pOwner

    
    Private controlCapEntry As CapEntry
    Friend CallingForm As Form

    Private bolFacCapStatusChanged As Boolean = False
    Dim returnVal As String = String.Empty

    Friend nOwnerID As Integer
    Friend nFacilityID As Integer
#End Region

#Region " Windows Form Designer generated code "
    Public Sub New(ByRef pOwn As MUSTER.BusinessLogic.pOwner)
        MyBase.New()

        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        oOwner = pOwn

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
    Friend WithEvents txtOwner As System.Windows.Forms.TextBox
    Friend WithEvents txtAddress As System.Windows.Forms.TextBox
    Friend WithEvents lblOwner As System.Windows.Forms.Label
    Friend WithEvents lblAddress As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents pnlTop As System.Windows.Forms.Panel
    Friend WithEvents ugFacilities As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlCenter As System.Windows.Forms.Panel
    Friend WithEvents pnlCenterRest As System.Windows.Forms.Panel
    Friend WithEvents pnlCenterTop As System.Windows.Forms.Panel
    Friend WithEvents pnlCenterRestRest As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.txtOwner = New System.Windows.Forms.TextBox
        Me.txtAddress = New System.Windows.Forms.TextBox
        Me.lblOwner = New System.Windows.Forms.Label
        Me.lblAddress = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.pnlTop = New System.Windows.Forms.Panel
        Me.pnlBottom = New System.Windows.Forms.Panel
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.pnlCenter = New System.Windows.Forms.Panel
        Me.pnlCenterRest = New System.Windows.Forms.Panel
        Me.pnlCenterRestRest = New System.Windows.Forms.Panel
        Me.pnlCenterTop = New System.Windows.Forms.Panel
        Me.ugFacilities = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlTop.SuspendLayout()
        Me.pnlBottom.SuspendLayout()
        Me.pnlCenter.SuspendLayout()
        Me.pnlCenterRest.SuspendLayout()
        Me.pnlCenterTop.SuspendLayout()
        CType(Me.ugFacilities, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtOwner
        '
        Me.txtOwner.Location = New System.Drawing.Point(64, 16)
        Me.txtOwner.Name = "txtOwner"
        Me.txtOwner.ReadOnly = True
        Me.txtOwner.Size = New System.Drawing.Size(216, 20)
        Me.txtOwner.TabIndex = 2
        Me.txtOwner.Text = ""
        '
        'txtAddress
        '
        Me.txtAddress.Location = New System.Drawing.Point(360, 16)
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.ReadOnly = True
        Me.txtAddress.Size = New System.Drawing.Size(368, 20)
        Me.txtAddress.TabIndex = 113
        Me.txtAddress.Text = ""
        '
        'lblOwner
        '
        Me.lblOwner.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwner.Location = New System.Drawing.Point(16, 16)
        Me.lblOwner.Name = "lblOwner"
        Me.lblOwner.Size = New System.Drawing.Size(40, 16)
        Me.lblOwner.TabIndex = 115
        Me.lblOwner.Text = "Owner"
        '
        'lblAddress
        '
        Me.lblAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAddress.Location = New System.Drawing.Point(296, 16)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(56, 16)
        Me.lblAddress.TabIndex = 116
        Me.lblAddress.Text = "Address"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(784, 8)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(96, 23)
        Me.btnClose.TabIndex = 119
        Me.btnClose.Text = "Close"
        '
        'pnlTop
        '
        Me.pnlTop.Controls.Add(Me.lblAddress)
        Me.pnlTop.Controls.Add(Me.txtOwner)
        Me.pnlTop.Controls.Add(Me.txtAddress)
        Me.pnlTop.Controls.Add(Me.lblOwner)
        Me.pnlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(888, 48)
        Me.pnlTop.TabIndex = 125
        '
        'pnlBottom
        '
        Me.pnlBottom.Controls.Add(Me.Label1)
        Me.pnlBottom.Controls.Add(Me.btnClose)
        Me.pnlBottom.Controls.Add(Me.Label2)
        Me.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlBottom.Location = New System.Drawing.Point(0, 573)
        Me.pnlBottom.Name = "pnlBottom"
        Me.pnlBottom.Size = New System.Drawing.Size(888, 40)
        Me.pnlBottom.TabIndex = 125
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Red
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 26)
        Me.Label1.TabIndex = 120
        Me.Label1.Text = "Invalid CAP Data"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Yellow
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(128, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 26)
        Me.Label2.TabIndex = 120
        Me.Label2.Text = "CAP Field"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlCenter
        '
        Me.pnlCenter.Controls.Add(Me.pnlCenterRest)
        Me.pnlCenter.Controls.Add(Me.pnlCenterTop)
        Me.pnlCenter.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlCenter.Location = New System.Drawing.Point(0, 48)
        Me.pnlCenter.Name = "pnlCenter"
        Me.pnlCenter.Size = New System.Drawing.Size(888, 525)
        Me.pnlCenter.TabIndex = 127
        '
        'pnlCenterRest
        '
        Me.pnlCenterRest.Controls.Add(Me.pnlCenterRestRest)
        Me.pnlCenterRest.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlCenterRest.Location = New System.Drawing.Point(0, 168)
        Me.pnlCenterRest.Name = "pnlCenterRest"
        Me.pnlCenterRest.Size = New System.Drawing.Size(888, 357)
        Me.pnlCenterRest.TabIndex = 129
        '
        'pnlCenterRestRest
        '
        Me.pnlCenterRestRest.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlCenterRestRest.Location = New System.Drawing.Point(0, 0)
        Me.pnlCenterRestRest.Name = "pnlCenterRestRest"
        Me.pnlCenterRestRest.Size = New System.Drawing.Size(888, 357)
        Me.pnlCenterRestRest.TabIndex = 126
        '
        'pnlCenterTop
        '
        Me.pnlCenterTop.Controls.Add(Me.ugFacilities)
        Me.pnlCenterTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCenterTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlCenterTop.Name = "pnlCenterTop"
        Me.pnlCenterTop.Size = New System.Drawing.Size(888, 168)
        Me.pnlCenterTop.TabIndex = 128
        '
        'ugFacilities
        '
        Me.ugFacilities.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugFacilities.DisplayLayout.AutoFitColumns = True
        Appearance1.FontData.BoldAsString = "True"
        Me.ugFacilities.DisplayLayout.CaptionAppearance = Appearance1
        Me.ugFacilities.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugFacilities.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        Me.ugFacilities.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugFacilities.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugFacilities.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ugFacilities.Location = New System.Drawing.Point(0, 0)
        Me.ugFacilities.Name = "ugFacilities"
        Me.ugFacilities.Size = New System.Drawing.Size(888, 168)
        Me.ugFacilities.TabIndex = 111
        Me.ugFacilities.Text = "Facilities"
        '
        'CAPSignUp
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(888, 613)
        Me.Controls.Add(Me.pnlCenter)
        Me.Controls.Add(Me.pnlBottom)
        Me.Controls.Add(Me.pnlTop)
        Me.Name = "CAPSignUp"
        Me.Text = "Registration - CAP Maintenance"
        Me.pnlTop.ResumeLayout(False)
        Me.pnlBottom.ResumeLayout(False)
        Me.pnlCenter.ResumeLayout(False)
        Me.pnlCenterRest.ResumeLayout(False)
        Me.pnlCenterTop.ResumeLayout(False)
        CType(Me.ugFacilities, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region

#Region "Form Page Events "

    Private Sub CAPSignUp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim strParticipationLevel As String
        Try
            Me.Height = MyBase.Height
            Dim strAddress As String = String.Empty
            ugFacilities.DataSource = Nothing
            If oOwner Is Nothing Then
                oOwner.Retrieve(nOwnerID, , False, True)
            End If

            'oAddressInfo = oOwner.Address()
            'strParticipationLevel = oOwner.CAPParticipationLevel
            strAddress = oOwner.Addresses.AddressLine1.ToString + IIf(oOwner.Addresses.AddressLine2.TrimEnd.Length = 0, ",", " " + oOwner.Addresses.AddressLine2.ToString + ",") + " " + oOwner.Addresses.City.Trim.ToString + ", " + oOwner.Addresses.State.Trim.ToString + " " + oOwner.Addresses.Zip.Trim.ToString
            Me.txtAddress.Text = strAddress
            FillUgFacililitesGrid(oOwner.ID, , , nFacilityID)
            Me.WindowState = FormWindowState.Maximized
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub

    'Private Sub CAPSignUp_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize


    '    'Me.Height = MyBase.Height

    '    ugTankandPipe.Width = Me.Width - 40
    '    ugTankandPipe.Height = Me.Height - lngugTankPipeBottom
    '    btnSave.Top = Me.Height - lngBtnActDist
    '    btnCancel.Top = Me.Height - lngBtnActDist
    '    btnClose.Top = Me.Height - lngBtnActDist
    '    ' Me.Height = 622

    'End Sub

    'Private Sub CAPSignUp_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
    '    Dim strOwnerInfo As String
    '    Try
    '        'MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGUID)
    '        'mContainer = Me.MdiParent
    '        strOwnerInfo = Me.txtOwner.Text & " : " & vbCrLf & txtOwner.Tag
    '        'mContainer.lblOwnerInfo.Text = strOwnerInfo.ToString
    '        'Me.Height = MyBase.Height
    '        Me.WindowState = FormWindowState.Maximized
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
#End Region

#Region " Facilities Grid Events & Processes "

    Private Sub FillUgFacililitesGrid(ByVal nOwnerId As Integer, Optional ByVal BolCheckedStatus As Boolean = False, Optional ByVal BolSetActive As Boolean = True, Optional ByVal facilityID As Integer = 0)
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow = Nothing
        Try
            dsCAPStatus = oOwner.Facilities.FacilityCAPTable(oOwner.ID)
            If dsCAPStatus.Tables(0).Rows.Count <= 0 Then
                MsgBox("No Records Found")
                Exit Sub
            Else
                If Not dsCAPStatus Is Nothing And dsCAPStatus.Tables(0).Rows.Count > 0 Then
                    ugFacilities.DataSource = dsCAPStatus
                    ugFacilities.DataBind()
                    ugFacilities.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                    ugFacilities.DisplayLayout.Bands(0).Columns("Facility ID").CellActivation = Activation.NoEdit

                    If facilityID <> 0 Then
                        For Each ugRow In ugFacilities.Rows
                            If ugRow.Cells("Facility ID").Value = nFacilityID Then
                                BolSetActive = True
                                Exit For
                            Else
                                ugRow = Nothing
                            End If
                        Next
                    End If

                    If BolSetActive Then
                        If ugRow Is Nothing Then
                            ugFacilities.ActiveRow = ugFacilities.Rows(0)
                        Else
                            ugFacilities.ActiveRow = ugRow
                        End If
                    End If

                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub ugFacilities_AfterRowActivate(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ugFacilities.AfterRowActivate

        If Not IsDBNull(ugFacilities.ActiveRow.Cells(0).Value) Then
            Try
                Me.Cursor = Cursors.WaitCursor

                ' will retrieve if tank / pipe is saved
                'oOwner.Facilities.Retrieve(oOwner.OwnerInfo, CInt(ugFacilities.ActiveRow.Cells(0).Value), , "FACILITY", False, True)
                ugFacilities.Tag = ugFacilities.ActiveRow.Cells(0).Value

                'LoadTankandPipe(ugFacilities.ActiveRow.Cells(0).Text)
                ReLoadCapEntry()


            Catch ex As Exception
                Me.Cursor = Cursors.Default
                Dim MyErr As New ErrorReport(ex)
                MyErr.ShowDialog()

            End Try

            Me.Cursor = Cursors.Default

        End If
    End Sub

#End Region


#Region "Control Cap Events"

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            If dsCAPStatus.Tables(0).Rows.Count > 0 Then
                controlCapEntry.Dispose()
                controlCapEntry = Nothing
            End If
            Me.Close()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub


#End Region

#Region "External Events"

    Private Sub FillGrid(ByVal BolVal As Boolean, ByVal facID As Integer) Handles oOwner.evtOwnFacilityCAPStatusChanged
        If BolVal Then
            Dim sender As Object
            Dim e As EventArgs
            For Each facInfo As MUSTER.Info.FacilityInfo In oOwner.OwnerInfo.facilityCollection.Values
                If facInfo.ID = facID Then
                    If oOwner.ID = facInfo.OwnerID Then
                        FillUgFacililitesGrid(facInfo.OwnerID, True)
                    Else
                        Exit For
                    End If
                End If
            Next
        End If
    End Sub


#End Region

    Private Sub CapLoadingNeedsfacilities(ByVal facID As Integer)
        FillUgFacililitesGrid(oOwner.ID, True, False, facID)

        For Each ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugFacilities.Rows
            If ugrow.Cells(0).Value = facID Then
                ugFacilities.ActiveRow = ugrow
                Exit For
            End If
        Next


    End Sub


    Private Sub ReLoadCapEntry()

        With pnlCenterRestRest
            .Visible = False

            If Not controlCapEntry Is Nothing Then

                RemoveHandler controlCapEntry.Fillfacilities, AddressOf CapLoadingNeedsfacilities
                controlCapEntry.Dispose()
                controlCapEntry = Nothing
            End If

            .Controls.Clear()

            If Not ugFacilities.ActiveRow Is Nothing Then
                controlCapEntry = New CapEntry(oOwner, ugFacilities.ActiveRow.Cells("FACILITY ID").Value, Me.CallingForm)
                controlCapEntry.Dock = DockStyle.Fill
                .Controls.Add(controlCapEntry)

                AddHandler controlCapEntry.Fillfacilities, AddressOf CapLoadingNeedsfacilities

                controlCapEntry.Show()
            End If
            .Visible = True

        End With

    End Sub



    'Private Sub ValidateCellsOld(ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs)
    '    Dim cellType As System.Type
    '    Dim strvalue As String
    '    Dim dttemp As Date


    '    cellType = e.Cell.Value.GetType
    '    If cellType.Equals(GetType(Date)) Or cellType.Equals(GetType(DBNull)) And IsDate(e.Cell.Text) Then
    '        Dim aDate As Date = CDate(e.Cell.Text)
    '        ' -+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
    '        ' -+-+-+ Basic Validations +-+-+-
    '        ' -+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
    '        If Not IsDBNull(e.Cell.OriginalValue) Then
    '            If DateDiff(DateInterval.Day, CDate(e.Cell.Text), CDate(e.Cell.OriginalValue)) = 0 Then
    '                Exit Sub
    '            Else
    '                If DateDiff(DateInterval.Day, aDate, Today()) < 0 Then
    '                    MsgBox(e.Cell.Column.ToString + " must not be greater than Today's date.")
    '                    e.Cell.Value = e.Cell.OriginalValue
    '                    Exit Sub
    '                End If
    '            End If
    '        Else
    '            If DateDiff(DateInterval.Day, aDate, Today()) < 0 Then
    '                MsgBox(e.Cell.Column.ToString + " must not be greater than Today's date.")
    '                If Not IsDBNull(e.Cell.OriginalValue) Then
    '                    e.Cell.Value = e.Cell.OriginalValue
    '                Else
    '                    e.Cell.Value = System.DBNull.Value
    '                End If
    '                Exit Sub
    '            End If
    '        End If

    '        ' AB - Not sure if these validations are required.
    '        ' -+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
    '        ' -+-+-+ Tank Validations  +-+-+-
    '        ' -+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
    '        If e.Cell.Band.Index = 0 Then
    '            If e.Cell.Row.Cells("STATUS").Text.IndexOf("Currently In Use") >= 0 Then

    '                If e.Cell.Column.ToString.IndexOf("TT DATE") >= 0 Then
    '                    If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 90, Today()), DateAdd(DateInterval.Year, 5, aDate)) <= 0 Then
    '                        dttemp = DateAdd(DateInterval.Day, DateDiff(DateInterval.Day, DateAdd(DateInterval.Year, 5, aDate), DateAdd(DateInterval.Day, 90, Today)), aDate)
    '                        MsgBox(e.Cell.Column.ToString + " must be greater than " + dttemp.ToShortDateString + " for Tank Site ID : " + e.Cell.Row.Cells("TANK SITE ID").Text)
    '                        If Not IsDBNull(e.Cell.OriginalValue) Then
    '                            e.Cell.Value = e.Cell.OriginalValue
    '                        Else
    '                            e.Cell.Value = System.DBNull.Value
    '                        End If
    '                        Exit Sub
    '                    End If
    '                End If

    '            End If


    '            If e.Cell.Row.Cells("STATUS").Text.IndexOf("Currently In Use") >= 0 Or e.Cell.Row.Cells("STATUS").Text.IndexOf("Temporarily Out of Service Indefinitely") >= 0 Then
    '                If e.Cell.Column.ToString.IndexOf("CP DATE") >= 0 Or e.Cell.Column.ToString.IndexOf("TERM CP TEST") >= 0 Then
    '                    If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 90, Today()), DateAdd(DateInterval.Year, 3, aDate)) <= 0 Then
    '                        dttemp = DateAdd(DateInterval.Day, DateDiff(DateInterval.Day, DateAdd(DateInterval.Year, 3, aDate), DateAdd(DateInterval.Day, 90, Today)), aDate)
    '                        MsgBox(e.Cell.Column.ToString & " must be greater than " & dttemp.ToShortDateString & " for Tank Site ID : " & e.Cell.Row.Cells("TANK SITE ID").Text)
    '                        If Not IsDBNull(e.Cell.OriginalValue) Then
    '                            e.Cell.Value = e.Cell.OriginalValue
    '                        Else
    '                            e.Cell.Value = System.DBNull.Value
    '                        End If
    '                        Exit Sub
    '                    End If
    '                End If

    '                ' Validation for IntInspectDate
    '                If e.Cell.Column.ToString.IndexOf("LI INSPECTED") >= 0 Then
    '                    If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 90, Today()), DateAdd(DateInterval.Year, 5, aDate)) <= 0 Then
    '                        dttemp = DateAdd(DateInterval.Day, DateDiff(DateInterval.Day, DateAdd(DateInterval.Year, 5, aDate), DateAdd(DateInterval.Day, 90, Today)), aDate)
    '                        MsgBox(e.Cell.Column.ToString + " must be greater than " + dttemp.ToShortDateString + " for Tank Site ID : " + e.Cell.Row.Cells("TANK SITE ID").Text)
    '                        If Not IsDBNull(e.Cell.OriginalValue) Then
    '                            e.Cell.Value = e.Cell.OriginalValue
    '                        Else
    '                            e.Cell.Value = System.DBNull.Value


    '                        End If
    '                        Exit Sub
    '                    End If
    '                End If

    '                ' Validation for IntInspectDate
    '                If e.Cell.Column.ToString.IndexOf("LI INSTALL") >= 0 Then
    '                    If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 90, Today()), DateAdd(DateInterval.Year, 10, aDate)) <= 0 Then
    '                        dttemp = DateAdd(DateInterval.Day, DateDiff(DateInterval.Day, DateAdd(DateInterval.Year, 10, aDate), DateAdd(DateInterval.Day, 90, Today)), aDate)
    '                        MsgBox(e.Cell.Column.ToString + " must be greater than " + dttemp.ToShortDateString + " for Tank Site ID : " + e.Cell.Row.Cells("TANK SITE ID").Text)
    '                        If Not IsDBNull(e.Cell.OriginalValue) Then
    '                            e.Cell.Value = e.Cell.OriginalValue
    '                        Else
    '                            e.Cell.Value = System.DBNull.Value
    '                        End If
    '                        Exit Sub
    '                    End If
    '                End If
    '            End If
    '        End If


    '        ' -+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
    '        ' -+-+-+ Pipe Validations  +-+-+-
    '        ' -+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
    '        If e.Cell.Band.Index = 1 Then
    '            If e.Cell.Row.Cells("STATUS").Text.IndexOf("Currently In Use") >= 0 Then

    '                If e.Cell.Column.ToString.IndexOf("TT DATE") >= 0 Then
    '                    If e.Cell.Row.Cells("PIPE_TYPE_DESC").Text.IndexOf("U.S.Suction") >= 0 Then
    '                        If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 90, Today()), DateAdd(DateInterval.Year, 3, aDate)) <= 0 Then
    '                            dttemp = DateAdd(DateInterval.Day, DateDiff(DateInterval.Day, DateAdd(DateInterval.Year, 3, aDate), DateAdd(DateInterval.Day, 90, Today)), aDate)
    '                            MsgBox(e.Cell.Column.ToString + " must be greater than " + dttemp.ToShortDateString + " for Pipe Site ID : " + e.Cell.Row.Cells("PIPE SITE ID").Text)
    '                            If Not IsDBNull(e.Cell.OriginalValue) Then
    '                                e.Cell.Value = e.Cell.OriginalValue
    '                            Else
    '                                e.Cell.Value = System.DBNull.Value
    '                            End If
    '                            Exit Sub
    '                        End If
    '                    End If
    '                    If e.Cell.Row.Cells("PIPE_TYPE_DESC").Text.IndexOf("Pressurized") >= 0 Then
    '                        If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 90, Today()), DateAdd(DateInterval.Year, 1, aDate)) <= 0 Then
    '                            dttemp = DateAdd(DateInterval.Day, DateDiff(DateInterval.Day, DateAdd(DateInterval.Year, 1, aDate), DateAdd(DateInterval.Day, 90, Today)), aDate)
    '                            MsgBox(e.Cell.Column.ToString + " must be greater than " + dttemp.ToShortDateString + " for Pipe Site ID : " + e.Cell.Row.Cells("PIPE SITE ID").Text)
    '                            If Not IsDBNull(e.Cell.OriginalValue) Then
    '                                e.Cell.Value = e.Cell.OriginalValue
    '                            Else
    '                                e.Cell.Value = System.DBNull.Value
    '                            End If
    '                            Exit Sub
    '                        End If

    '                    End If
    '                End If

    '                ' Validation for ALLDTestDate
    '                If e.Cell.Column.ToString.IndexOf("ALLD Test Date") >= 0 Then
    '                    If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 90, Today()), DateAdd(DateInterval.Year, 1, aDate)) <= 0 Then
    '                        dttemp = DateAdd(DateInterval.Day, DateDiff(DateInterval.Day, DateAdd(DateInterval.Year, 1, aDate), DateAdd(DateInterval.Day, 90, Today)), aDate)
    '                        MsgBox(e.Cell.Column.ToString + " must be greater than " + dttemp.ToShortDateString + " for Pipe Site ID : " + e.Cell.Row.Cells("PIPE SITE ID").Text)
    '                        If Not IsDBNull(e.Cell.OriginalValue) Then
    '                            e.Cell.Value = e.Cell.OriginalValue
    '                        Else
    '                            e.Cell.Value = System.DBNull.Value
    '                        End If
    '                        Exit Sub
    '                    End If

    '                End If

    '            End If


    '            If e.Cell.Row.Cells("STATUS").Text.IndexOf("Currently In Use") >= 0 Or e.Cell.Row.Cells("STATUS").Text.IndexOf("Temporarily Out of Service Indefinitely") >= 0 Then
    '                If e.Cell.Column.ToString.IndexOf("CP DATE") >= 0 Or e.Cell.Column.ToString.IndexOf("TERM CP TEST") >= 0 Then
    '                    If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 90, Today()), DateAdd(DateInterval.Year, 3, aDate)) <= 0 Then
    '                        dttemp = DateAdd(DateInterval.Day, DateDiff(DateInterval.Day, DateAdd(DateInterval.Year, 3, aDate), DateAdd(DateInterval.Day, 90, Today)), aDate)
    '                        MsgBox(e.Cell.Column.ToString & " must be greater than " & dttemp.ToShortDateString & " for Pipe Site ID : " & e.Cell.Row.Cells("PIPE SITE ID").Text)
    '                        If Not IsDBNull(e.Cell.OriginalValue) Then
    '                            e.Cell.Value = e.Cell.OriginalValue
    '                        Else
    '                            e.Cell.Value = System.DBNull.Value
    '                        End If
    '                        Exit Sub
    '                    End If
    '                End If

    '            End If
    '        End If

    '    End If

    'End Sub



End Class
