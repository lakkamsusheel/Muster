'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Search
'   Provides the operations required to manipulate an Flag object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         MNR     12/08/04    Original class definition.
'   1.1         AN      01/03/05    Added Try catch and Exception Handling/Logging
'   1.2         EN      01/19/05    Added New GetResult and commented the Old GetResult.
'                                   Modified Event SearchErr.
'   1.3         EN      01/24/05    Commented the raiseevent in Keyword property. Added strSrc to SearchResultsevent.
'   1.4         MR      04/08/05    Commented Warning Msg in GetResult() Function.
'
' Function          Description
' GetResult()       Returns a resultant Dataset for the provided keyword in the
'                   selected Module, filtered by the selected Filter
'
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pSearch
#Region "Public Events"
        Public Event SearchErr(ByVal MsgStr As String, ByVal strColumnName As String, ByVal strSrc As String)
        Public Event SearchResults(ByVal nCount As Integer, ByVal strSrc As String)
#End Region
#Region "Private Member Variables"
        Private oSearchInfo As Muster.Info.SearchInfo
        Private oSearchDB As New Muster.DataAccess.SearchDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
#End Region
#Region "Constructors"
        Public Sub New()
            oSearchInfo = New Muster.Info.SearchInfo
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property Keyword() As String
            Get
                Return oSearchInfo.Keyword
            End Get
            Set(ByVal Value As String)
                oSearchInfo.Keyword = Value
            End Set
        End Property
        Public Property [Module]() As String
            Get
                Return oSearchInfo.Module
            End Get
            Set(ByVal Value As String)
                If Value = "Select Module" Or Value = String.Empty Then
                    RaiseEvent SearchErr("Select a Module", "Module", Me.ToString)
                    Exit Property
                Else
                    oSearchInfo.Module = Value
                End If
            End Set
        End Property
        Public Property Filter() As String
            Get
                Return oSearchInfo.Filter
            End Get
            Set(ByVal Value As String)
                If Value = "Filter By" Or Value = String.Empty Then
                    RaiseEvent SearchErr("Select a Filter By Input", "Filter", Me.ToString)
                    Exit Property
                Else
                    oSearchInfo.Filter = Value
                End If
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"

        Public Function GetResult() As DataSet

            Dim ds As DataSet
            Dim QueryOperator As String
            Dim strTempKeyWord As String
            Dim srchKeyword As String
            Dim bolDoSearch As Boolean = True

            Try
                If (oSearchInfo.Filter.IndexOf(" ID") > -1) And (oSearchInfo.Filter.IndexOf("BP2K") = -1) And IsNumeric(oSearchInfo.Keyword) = False Then
                    RaiseEvent SearchErr("Enter the Valid Numeric Input", "KEYWORD", Me.ToString)
                    bolDoSearch = False
                Else
                    srchKeyword = Trim(oSearchInfo.Keyword)
                    'If srchKeyword = String.Empty Then
                    '    Dim msgResult = MsgBox("WARNING!" & vbCrLf & "No search keyword specified may result is a LARGE result set." & _
                    '        vbCrLf & "Continue?", MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNo, "No Search Keyword Specified")
                    '    If msgResult = MsgBoxResult.No Then
                    '        bolDoSearch = False
                    '    End If
                    'End If

                    If bolDoSearch Then
                        Select Case Trim(oSearchInfo.Filter)
                            Case "Ensite AIID"
                                QueryOperator = "="
                                strTempKeyWord = srchKeyword
                            Case "Owner ID", "Facility ID", "Check #"
                                QueryOperator = "="
                                strTempKeyWord = srchKeyword
                            Case "Owner Name", "Facility Name", "Owner Address", "Facility Address", "BP2K Owner ID"
                                strTempKeyWord = "'%" + Replace(srchKeyword.ToString(), "'", "''") + "%'"
                                QueryOperator = "LIKE"
                            Case "Company Name", "Licensee Name", "Manager Name"
                                strTempKeyWord = srchKeyword.ToString
                                QueryOperator = "LIKE"
                            Case Else
                                RaiseEvent SearchErr("Please Enter valid Filter.", "Filter", Me.ToString)
                                bolDoSearch = False
                        End Select
                    End If

                    If bolDoSearch Then
                        ds = oSearchDB.DBGetDS(strTempKeyWord, oSearchInfo.Filter, QueryOperator)
                        If ds.Tables.Count = 0 Then
                            RaiseEvent SearchErr("No Records Found", "RESULTS", Me.ToString)
                            ds = Nothing
                        Else
                            If ds.Tables(0).Rows.Count = 0 Then
                                RaiseEvent SearchErr("No Records Found", "RESULTS", Me.ToString)
                                ds = Nothing
                            Else
                                RaiseEvent SearchResults(ds.Tables(0).Rows.Count, Me.ToString)
                            End If
                        End If
                    Else
                        ds = Nothing
                    End If
                End If
                Return ds
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function
#End Region
#Region "LookUp Operations"
        
        Public Function PopulateQuickSearchFilter(ByVal strParentId As String) As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("v_QUICKSEARCH_FILTER", strParentId)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function
        Private Function GetDataTable(ByVal DBViewName As String, Optional ByVal strModuleName As String = "") As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            Try
                If strModuleName = String.Empty Then
                    strSQL = "SELECT * FROM " & DBViewName
                Else
                    strSQL = "SELECT * FROM " & DBViewName & " WHERE PROPERTY_ID_PARENT = '" & strModuleName & "'"
                End If

                dsReturn = oSearchDB.DBGetsearchFilter(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                Else
                    dtReturn = Nothing
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#End Region
    End Class
End Namespace
