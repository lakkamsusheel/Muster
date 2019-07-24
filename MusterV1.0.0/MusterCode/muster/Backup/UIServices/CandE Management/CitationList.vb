Public Class CitationList
    Inherits System.Windows.Forms.Form
#Region "Public Events"
    Public Event evtCitationSelected(ByVal CitationID As Integer, ByVal isDiscrep As Boolean, ByVal strDiscrepText As String)
    Public Event evtLCECitationSelected(ByVal drow As Infragistics.Win.UltraWinGrid.UltraGridRow)
#End Region
#Region "User defined variables"
    Private pFCE As MUSTER.BusinessLogic.pFacilityComplianceEvent
    Private pLCE As MUSTER.BusinessLogic.pLicenseeComplianceEvent
    Private nCitationID As Integer = 0
    Private strFrom As String
    Private bolisDiscrep As Boolean
    Private alCitationPenalty As ArrayList
#End Region
#Region " Windows Form Designer generated code "

    Public Sub New(ByVal from As String, Optional ByRef objFCE As MUSTER.BusinessLogic.pFacilityComplianceEvent = Nothing, Optional ByRef objLCE As MUSTER.BusinessLogic.pLicenseeComplianceEvent = Nothing, Optional ByVal alCitPenalty As ArrayList = Nothing, Optional ByVal isDiscrep As Boolean = False)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        strFrom = from
        If strFrom = "FCE" Then
            pFCE = objFCE
            If pFCE Is Nothing Then
                pFCE = New MUSTER.BusinessLogic.pFacilityComplianceEvent
            End If
            alCitationPenalty = alCitPenalty
            bolisDiscrep = isDiscrep
        Else
            pLCE = objLCE
        End If
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
    Friend WithEvents ugCitations As Infragistics.Win.UltraWinGrid.UltraGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ugCitations = New Infragistics.Win.UltraWinGrid.UltraGrid
        CType(Me.ugCitations, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ugCitations
        '
        Me.ugCitations.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugCitations.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugCitations.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugCitations.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugCitations.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugCitations.Location = New System.Drawing.Point(0, 0)
        Me.ugCitations.Name = "ugCitations"
        Me.ugCitations.Size = New System.Drawing.Size(600, 238)
        Me.ugCitations.TabIndex = 1
        '
        'CitationList
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(600, 238)
        Me.Controls.Add(Me.ugCitations)
        Me.Name = "CitationList"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CitationList"
        CType(Me.ugCitations, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub CitationList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim dsData As DataSet
        Try
            If strFrom = "FCE" Then
                Dim strWhere As String = ""
                If bolisDiscrep Then
                    Me.Text = "Discrep List"
                    For Each strText As String In alCitationPenalty
                        strWhere += " DISCREP_TEXT <> '" + strText + "' AND"
                    Next
                    If strWhere.Length > 0 Then
                        strWhere = " AND" + strWhere.Substring(0, strWhere.Length - 3)
                    End If
                    ugCitations.DataSource = pFCE.GetDiscrepText(strWhere)
                Else
                    For Each citID As Integer In alCitationPenalty
                        strWhere += " CITATION_ID <> " + citID.ToString + " AND"
                    Next
                    If strWhere.Length > 0 Then
                        strWhere = " WHERE" + strWhere.Substring(0, strWhere.Length - 3)
                    End If
                    ugCitations.DataSource = pFCE.GetCitations(strWhere)
                    'If pFCE.ID <= 0 Then
                    'Else
                    '    ugCitations.DataSource = pFCE.GetCitations(" WHERE C.DELETED = 0 AND P.DELETED = 0 ORDER BY CATEGORY")
                    'End If
                End If
            Else
                dsData = pLCE.PopulateCitationList()
                ugCitations.DataSource = dsData.Tables(0).DefaultView
                ugCitations.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugCitations_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugCitations.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            If Not (ugCitations.ActiveRow Is Nothing) Then
                If strFrom = "FCE" Then
                    If bolisDiscrep Then
                        nCitationID = ugCitations.ActiveRow.Cells("QUESTION_ID").Value
                        RaiseEvent evtCitationSelected(0, bolisDiscrep, ugCitations.ActiveRow.Cells("DISCREP TEXT").Text)
                    Else
                        nCitationID = ugCitations.ActiveRow.Cells("CITATION_ID").Value
                        RaiseEvent evtCitationSelected(nCitationID, bolisDiscrep, "")
                    End If
                Else
                    pLCE.LicenseeCitationID = ugCitations.ActiveRow.Cells("CITATION_ID").Value
                    pLCE.citationText = ugCitations.ActiveRow.Cells("CITATION_TEXT").Value
                    RaiseEvent evtLCECitationSelected(ugCitations.ActiveRow)
                    Me.Close()
                End If
                ugCitations.ActiveRow.Delete(False)
            Else
                Exit Sub
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugCitations_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugCitations.InitializeLayout
        ugCitations.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        ugCitations.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        ugCitations.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False

        ugCitations.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect

        If strFrom = "FCE" Then
            ugCitations.DisplayLayout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free
            If bolisDiscrep Then
                ugCitations.DisplayLayout.Bands(0).Columns("CITATION_ID").Hidden = True
                ugCitations.DisplayLayout.Bands(0).Columns("QUESTION_ID").Hidden = True
                'ugCitations.DisplayLayout.Bands(0).Columns("DISCREP TEXT").Header.Caption = "DISCREP TEXT"
                ugCitations.DisplayLayout.Bands(0).Columns("DISCREP TEXT").Width = Me.Width - 10
            Else
                ugCitations.DisplayLayout.Bands(0).Columns("CITATION_ID").Hidden = True
                'ugCitations.DisplayLayout.Bands(0).Columns("FederalCitation").Hidden = True
                ugCitations.DisplayLayout.Bands(0).Columns("Section").Hidden = True
                'ugCitations.DisplayLayout.Bands(0).Columns("Small").Hidden = True
                'ugCitations.DisplayLayout.Bands(0).Columns("Medium").Hidden = True
                'ugCitations.DisplayLayout.Bands(0).Columns("Large").Hidden = True
                'ugCitations.DisplayLayout.Bands(0).Columns("CorrectiveAction").Hidden = True
                'ugCitations.DisplayLayout.Bands(0).Columns("EPA").Hidden = True
                ugCitations.DisplayLayout.Bands(0).Columns("Created_By").Hidden = True
                ugCitations.DisplayLayout.Bands(0).Columns("DATE_Created").Hidden = True
                ugCitations.DisplayLayout.Bands(0).Columns("LAST_EDITED_By").Hidden = True
                ugCitations.DisplayLayout.Bands(0).Columns("DATE_LAST_EDITED").Hidden = True
                ugCitations.DisplayLayout.Bands(0).Columns("Deleted").Hidden = True

                ugCitations.DisplayLayout.Bands(0).Columns("StateCitation").Header.Caption = "CITATION"
                ugCitations.DisplayLayout.Bands(0).Columns("Category").Header.Caption = "CATEGORY"
                ugCitations.DisplayLayout.Bands(0).Columns("Description").Header.Caption = "CITATION TEXT"

                ugCitations.DisplayLayout.Bands(0).Columns("StateCitation").Header.VisiblePosition = 0
                ugCitations.DisplayLayout.Bands(0).Columns("Category").Header.VisiblePosition = 1
                ugCitations.DisplayLayout.Bands(0).Columns("Description").Header.VisiblePosition = 2
            End If
        Else
        End If
    End Sub
End Class
