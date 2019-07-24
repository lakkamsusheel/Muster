Public Class DetachPipes
    Inherits System.Windows.Forms.Form

    Private nFacilityID, nTankID As Integer
    Private pTank As MUSTER.BusinessLogic.pTank
    Private parentFrm As Form
    Private nSelected As Integer = 0
    Private rp As New Remove_Pencil
    Private returnVal As String = String.Empty

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
    Friend WithEvents pnlBottom As System.Windows.Forms.Panel
    Friend WithEvents ugDetachPipe As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnDetach As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlBottom = New System.Windows.Forms.Panel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnDetach = New System.Windows.Forms.Button
        Me.ugDetachPipe = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlBottom.SuspendLayout()
        CType(Me.ugDetachPipe, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlBottom
        '
        Me.pnlBottom.Controls.Add(Me.btnCancel)
        Me.pnlBottom.Controls.Add(Me.btnDetach)
        Me.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlBottom.Location = New System.Drawing.Point(0, 285)
        Me.pnlBottom.Name = "pnlBottom"
        Me.pnlBottom.Size = New System.Drawing.Size(584, 40)
        Me.pnlBottom.TabIndex = 0
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(256, 8)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        '
        'btnDetach
        '
        Me.btnDetach.Location = New System.Drawing.Point(160, 8)
        Me.btnDetach.Name = "btnDetach"
        Me.btnDetach.TabIndex = 0
        Me.btnDetach.Text = "Detach"
        '
        'ugDetachPipe
        '
        Me.ugDetachPipe.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugDetachPipe.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugDetachPipe.Location = New System.Drawing.Point(0, 0)
        Me.ugDetachPipe.Name = "ugDetachPipe"
        Me.ugDetachPipe.Size = New System.Drawing.Size(584, 285)
        Me.ugDetachPipe.TabIndex = 0
        '
        'DetachPipes
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(584, 325)
        Me.Controls.Add(Me.ugDetachPipe)
        Me.Controls.Add(Me.pnlBottom)
        Me.Name = "DetachPipes"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Registration - Detach Pipes for Tank x - FacilityID (FacName)"
        Me.pnlBottom.ResumeLayout(False)
        CType(Me.ugDetachPipe, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Friend WriteOnly Property FacilityID() As Integer
        Set(ByVal Value As Integer)
            nFacilityID = Value
        End Set
    End Property
    Friend Property TankID() As Integer
        Get
            Return nTankID
        End Get
        Set(ByVal Value As Integer)
            nTankID = Value
        End Set
    End Property
    Friend WriteOnly Property TankObj() As MUSTER.BusinessLogic.pTank
        Set(ByVal Value As MUSTER.BusinessLogic.pTank)
            pTank = Value
        End Set
    End Property
    Friend Property CallingForm() As Form
        Get
            Return parentFrm
        End Get
        Set(ByVal Value As Form)
            parentFrm = Value
        End Set
    End Property

    Private Sub DetachPipes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If pTank Is Nothing Then
            pTank = New MUSTER.BusinessLogic.pTank
        End If
        LoadTanksPipes()
    End Sub

    Private Sub ShowError(ByVal ex As Exception)
        Dim MyErr As New ErrorReport(ex)
        MyErr.ShowDialog()
    End Sub
    Private Sub LoadTanksPipes()
        Try
            Dim strAttachedPipeIDs As String = pTank.GetAttachedPipeIDs(TankID)
            Dim dsPipes As DataSet = pTank.GetAttachedPipes(strAttachedPipeIDs)


            ugDetachPipe.DataSource = dsPipes
            ugDetachPipe.DrawFilter = rp
            btnDetach.Enabled = False
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub ugDetachPipe_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugDetachPipe.InitializeLayout
        Try
            e.Layout.Bands(0).Columns("FACILITY_ID").Hidden = True
            e.Layout.Bands(0).Columns("TANK_ID").Hidden = True
            e.Layout.Bands(0).Columns("PIPE_ID").Hidden = True
            e.Layout.Bands(0).Columns("POSITION").Hidden = True

            e.Layout.Bands(0).Columns("TANK SITE ID").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("COMPARTMENT NUMBER").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("PIPE SITE ID").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("PIPE STATUS").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            e.Layout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Free
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub ugDetachPipe_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugDetachPipe.CellChange
        Try
            If "SELECTED".Equals(e.Cell.Column.Key) Then
                If e.Cell.Text.ToUpper = "TRUE" Then
                    nSelected += 1
                Else
                    nSelected -= 1
                End If
                If nSelected = ugDetachPipe.Rows.Count And nSelected > 0 Then
                    MsgBox("Cannot detach all pipes")
                    e.Cell.CancelUpdate()
                    nSelected -= 1
                End If
                If nSelected = 0 Then
                    btnDetach.Enabled = False
                Else
                    btnDetach.Enabled = True
                End If
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub btnDetach_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDetach.Click
        Try
            Dim strDetachedPipes As String = String.Empty
            For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugDetachPipe.Rows
                If ugRow.Cells("SELECTED").Value Then

                    Dim nPipeID As Integer = ugRow.Cells("PIPE_ID").Value

                    pTank.DetachPipe(nPipeID, ugRow.Cells("TANK_ID").Value, ugRow.Cells("COMPARTMENT NUMBER").Value, UIUtilsGen.ModuleID.Registration, MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)

                    'Copies any extensions
                    If nPipeID > 0 Then

                        Dim LocalpPipe As New BusinessLogic.pPipe
                        Dim pipeExtensions As DataTable = LocalpPipe.GetPipeExtensions(nPipeID)


                        If Not pipeExtensions Is Nothing AndAlso pipeExtensions.Rows.Count > 0 Then

                            For Each children As DataRow In pipeExtensions.Rows

                                Dim nChildPipeID As Integer = children("PIPE_ID")

                                pTank.DetachPipe(nChildPipeID, ugRow.Cells("TANK_ID").Value, ugRow.Cells("COMPARTMENT NUMBER").Value, UIUtilsGen.ModuleID.Registration, MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)

                            Next
                        End If


                    End If

                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                    strDetachedPipes += "Tank: " + ugRow.Cells("TANK SITE ID").Text + "; Compartment: " + ugRow.Cells("COMPARTMENT NUMBER").Text + "; Pipe: " + ugRow.Cells("PIPE SITE ID").Text + vbCrLf
                End If
            Next
            If strDetachedPipes.Length > 0 Then
                MsgBox("Detached the foll pipe(s) successfully:" + vbCrLf + strDetachedPipes, MsgBoxStyle.OKOnly)
            End If
            If Not CallingForm Is Nothing Then
                CallingForm.Tag = "1"
            End If
            Me.Close()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

End Class
