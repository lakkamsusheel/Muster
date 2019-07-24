'-------------------------------------------------------------------------------
' MUSTER.DataAccess.FlagDB
'   Provides the means for marshalling LustEvent state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        JVC     03/02/05    Original class definition.
'  1.1        AB      03/21/05    Added DBGetByFacilityID
'  1.2        MNR     04/22/05    Added condition to check if value is 0, 
'                                   returns new info/collection instead of accessing db
'                                   as there are no records with primary id 0
'
' Function                  Description
' GetAllInfo()      Returns an LustEventCollection containing all LustEvent objects in the repository
' DBGetByID(ID)     Returns an LustEvent object indicated by int arg ID
' DBGetDS(SQL)      Returns a resultant Dataset by running query specified by the string arg SQL
' Put(LustEvent)       Saves the LustEvent passed as an argument, to the DB
'-------------------------------------------------------------------------------
'
'

Imports System.Data.SqlClient
Imports Utils.DBUtils
Namespace MUSTER.DataAccess
    Public Class LustEventDB
        Private _strConn As String
        Private MusterException As Exceptions.MusterExceptions = New MUSTER.Exceptions.MusterExceptions

        Public Sub New(Optional ByVal strDBConn As String = "", Optional ByRef MusterXCEP As MUSTER.Exceptions.MusterExceptions = Nothing)
            Try
                If MusterXCEP Is Nothing Then
                    MusterException = New MUSTER.Exceptions.MusterExceptions
                Else
                    MusterException = MusterXCEP
                End If
                If strDBConn = String.Empty Then
                    Dim oCnn As New ConnectionSettings
                    _strConn = oCnn.cnString
                    oCnn = Nothing
                Else
                    _strConn = strDBConn
                End If
            Catch ex As Exception
                Throw ex
            End Try
        End Sub


        Public Function DBGetByID(ByVal nVal As Int64) As MUSTER.Info.LustEventInfo
            ' #Region "XDEOperation" ' Begin Template Expansion{7BAED731-2749-4C9B-9830-48689EE8C96C}
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection

            Try
                If nVal = 0 Then
                    Return New MUSTER.Info.LustEventInfo
                End If
                strSQL = "spGetTecEvent"
                strVal = nVal

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                '
                ' Repeat the following line as many times as necessary
                '
                Params("@EVENT_ID").Value = nVal
                Params("@FACILITY_ID").Value = DBNull.Value

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.LustEventInfo(AltIsDBNull(drSet.Item("Event_ID"), 0), _
                            AltIsDBNull(drSet.Item("LAST_GWS"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_LDR"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_PTT"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_OF_REPORT"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_SITE_ASSESSED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EVENT_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EVENT_STATUS"), 0), _
                            AltIsDBNull(drSet.Item("FACILITY_ID"), 0), _
                            AltIsDBNull(drSet.Item("FLAG_ID"), 0), _
                            AltIsDBNull(drSet.Item("MGPTF_STATUS"), 0), _
                            AltIsDBNull(drSet.Item("LEAKAGE_PRIORITY"), 0), _
                            AltIsDBNull(drSet.Item("EVENT_PROJECT_MANAGER_ID"), 0), _
                            AltIsDBNull(drSet.Item("RELEASE_STATUS"), 0), _
                            0, _
                            AltIsDBNull(drSet.Item("RelatedSites"), ""), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("SUSPECTED_SOURCE"), 0), _
                            AltIsDBNull(drSet.Item("CONFIRMED_ON"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("IDENTIFIED_BY"), 0), _
                            AltIsDBNull(drSet.Item("LOCATION"), 0), _
                            AltIsDBNull(drSet.Item("EXTENT"), 0), _
                            AltIsDBNull(drSet.Item("EVENT_PROJECT_MANAGER_ID"), 0), _
                            AltIsDBNull(drSet.Item("EVENT_STARTED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EVENT_ENDED"), "1/1/0001"), _
                            (AltIsDBNull(drSet.Item("TOC_SOIL"), False)), _
                            (AltIsDBNull(drSet.Item("SOIL_BTEX"), False)), _
                            (AltIsDBNull(drSet.Item("SOIL_PAH"), False)), _
                            (AltIsDBNull(drSet.Item("SOIL_TPH"), False)), _
                            (AltIsDBNull(drSet.Item("TOC_GROUNDWATER"), False)), _
                            (AltIsDBNull(drSet.Item("GW_BTEX"), False)), _
                            (AltIsDBNull(drSet.Item("GW_PAH"), False)), _
                            (AltIsDBNull(drSet.Item("GW_TPH"), False)), _
                            (AltIsDBNull(drSet.Item("FREE_PRODUCT"), False)), _
                            (AltIsDBNull(drSet.Item("FP_GASOLINE"), False)), _
                            (AltIsDBNull(drSet.Item("FP_DIESEL"), False)), _
                            (AltIsDBNull(drSet.Item("FP_KEROSENE"), False)), _
                            (AltIsDBNull(drSet.Item("FP_WASTE_OIL"), False)), _
                            (AltIsDBNull(drSet.Item("FP_UNKNOWN"), False)), _
                            (AltIsDBNull(drSet.Item("TOC_VAPOR"), False)), _
                            (AltIsDBNull(drSet.Item("VAPOR_BTEX"), False)), _
                            (AltIsDBNull(drSet.Item("VAPOR_PAH"), False)), _
                            AltIsDBNull(drSet.Item("EVENT_SEQUENCE"), 0), _
                            drSet.Item("TFChecklist"), _
                            drSet.Item("TankPipeList"), _
                            (AltIsDBNull(drSet.Item("PM_HEAD_ASSESS"), 0)), _
                            (AltIsDBNull(drSet.Item("PM_HEAD_DATE"), "1/1/0001")), _
                            (AltIsDBNull(drSet.Item("PM_HEAD_BY"), "")), _
                            (AltIsDBNull(drSet.Item("UST_CHIEF_ASSESS"), 0)), _
                            (AltIsDBNull(drSet.Item("UST_CHIEF_DATE"), "1/1/0001")), _
                            (AltIsDBNull(drSet.Item("UST_CHIEF_BY"), "")), _
                            (AltIsDBNull(drSet.Item("OPC_HEAD_ASSESS"), 0)), _
                            (AltIsDBNull(drSet.Item("OPC_HEAD_DATE"), "1/1/0001")), _
                            (AltIsDBNull(drSet.Item("OPC_HEAD_BY"), "")), _
                            (AltIsDBNull(drSet.Item("FOR_OPC_HEAD"), False)), _
                            (AltIsDBNull(drSet.Item("COMMISSION_ASSESS"), 0)), _
                            (AltIsDBNull(drSet.Item("COMMISSION_DATE"), "1/1/0001")), _
                            (AltIsDBNull(drSet.Item("COMMISSION_BY"), "")), _
                            (AltIsDBNull(drSet.Item("FOR_COMMISSION"), "0")), _
                            (AltIsDBNull(drSet.Item("ELIGIBITY_COMMENTS"), "")), drSet.Item("PMDesc"), drSet.Item("MGPTFStatusDesc"), drSet.Item("TechnicalStatusDesc"), _
                            (AltIsDBNull(drSet.Item("IRAC"), 0)), _
                            (AltIsDBNull(drSet.Item("ERAC"), 0)), _
                            drSet.Item("HOW_DISC_FAC_LEAK_DETECTION"), _
                            drSet.Item("HOW_DISC_SURFACE_SHEEN"), _
                            drSet.Item("HOW_DISC_GW_WELL"), _
                            drSet.Item("HOW_DISC_GW_CONTAMINATION"), _
                            drSet.Item("HOW_DISC_VAPORS"), _
                            drSet.Item("HOW_DISC_FREE_PRODUCT"), _
                            drSet.Item("HOW_DISC_SOIL_CONTAMINATION"), _
                            drSet.Item("HOW_DISC_FAILED_PTT"), _
                            drSet.Item("HOW_DISC_INVENTORY_SHORTAGE"), _
                            drSet.Item("HOW_DISC_TANK_CLOSURE"), _
                            drSet.Item("HOW_DISC_INSPECTION"), _
                            AltIsDBNull(drSet.Item("CAUSE"), 0) _
                            )
                    'AltIsDBNull(drSet.Item("HOW_DISCOVERED_ID"), 0), _

                Else

                    Return New MUSTER.Info.LustEventInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
            ' #End Region ' XDEOperation End Template Expansion{7BAED731-2749-4C9B-9830-48689EE8C96C}
        End Function

        Public Function DBGetByFacilityAndSequence(ByVal nFacility As Int64, ByVal nSeq As Int64) As MUSTER.Info.LustEventInfo
            Dim drSet As SqlDataReader
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection

            Try

                strSQL = "spGetTecEvent_ByFacilityAndSequence"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                '
                ' Repeat the following line as many times as necessary
                '
                Params("@SEQUENCE").Value = nSeq
                Params("@FACILITY_ID").Value = nFacility

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                If drSet.HasRows Then
                    drSet.Read()
                    Return New MUSTER.Info.LustEventInfo(AltIsDBNull(drSet.Item("Event_ID"), 0), _
                            AltIsDBNull(drSet.Item("LAST_GWS"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_LDR"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_PTT"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_OF_REPORT"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_SITE_ASSESSED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EVENT_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EVENT_STATUS"), 0), _
                            AltIsDBNull(drSet.Item("FACILITY_ID"), 0), _
                            AltIsDBNull(drSet.Item("FLAG_ID"), 0), _
                            AltIsDBNull(drSet.Item("MGPTF_STATUS"), 0), _
                            AltIsDBNull(drSet.Item("LEAKAGE_PRIORITY"), 0), _
                            AltIsDBNull(drSet.Item("EVENT_PROJECT_MANAGER_ID"), 0), _
                            AltIsDBNull(drSet.Item("RELEASE_STATUS"), 0), _
                            0, _
                            AltIsDBNull(drSet.Item("RelatedSites"), ""), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("SUSPECTED_SOURCE"), 0), _
                            AltIsDBNull(drSet.Item("CONFIRMED_ON"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("IDENTIFIED_BY"), 0), _
                            AltIsDBNull(drSet.Item("LOCATION"), 0), _
                            AltIsDBNull(drSet.Item("EXTENT"), 0), _
                            AltIsDBNull(drSet.Item("EVENT_PROJECT_MANAGER_ID"), 0), _
                            AltIsDBNull(drSet.Item("EVENT_STARTED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EVENT_ENDED"), "1/1/0001"), _
                            (AltIsDBNull(drSet.Item("TOC_SOIL"), False)), _
                            (AltIsDBNull(drSet.Item("SOIL_BTEX"), False)), _
                            (AltIsDBNull(drSet.Item("SOIL_PAH"), False)), _
                            (AltIsDBNull(drSet.Item("SOIL_TPH"), False)), _
                            (AltIsDBNull(drSet.Item("TOC_GROUNDWATER"), False)), _
                            (AltIsDBNull(drSet.Item("GW_BTEX"), False)), _
                            (AltIsDBNull(drSet.Item("GW_PAH"), False)), _
                            (AltIsDBNull(drSet.Item("GW_TPH"), False)), _
                            (AltIsDBNull(drSet.Item("FREE_PRODUCT"), False)), _
                            (AltIsDBNull(drSet.Item("FP_GASOLINE"), False)), _
                            (AltIsDBNull(drSet.Item("FP_DIESEL"), False)), _
                            (AltIsDBNull(drSet.Item("FP_KEROSENE"), False)), _
                            (AltIsDBNull(drSet.Item("FP_WASTE_OIL"), False)), _
                            (AltIsDBNull(drSet.Item("FP_UNKNOWN"), False)), _
                            (AltIsDBNull(drSet.Item("TOC_VAPOR"), False)), _
                            (AltIsDBNull(drSet.Item("VAPOR_BTEX"), False)), _
                            (AltIsDBNull(drSet.Item("VAPOR_PAH"), False)), _
                            AltIsDBNull(drSet.Item("EVENT_SEQUENCE"), 0), _
                            drSet.Item("TFChecklist"), _
                            drSet.Item("TankPipeList"), _
                            (AltIsDBNull(drSet.Item("PM_HEAD_ASSESS"), 0)), _
                            (AltIsDBNull(drSet.Item("PM_HEAD_DATE"), "1/1/0001")), _
                            (AltIsDBNull(drSet.Item("PM_HEAD_BY"), "")), _
                            (AltIsDBNull(drSet.Item("UST_CHIEF_ASSESS"), 0)), _
                            (AltIsDBNull(drSet.Item("UST_CHIEF_DATE"), "1/1/0001")), _
                            (AltIsDBNull(drSet.Item("UST_CHIEF_BY"), "")), _
                            (AltIsDBNull(drSet.Item("OPC_HEAD_ASSESS"), 0)), _
                            (AltIsDBNull(drSet.Item("OPC_HEAD_DATE"), "1/1/0001")), _
                            (AltIsDBNull(drSet.Item("OPC_HEAD_BY"), "")), _
                            (AltIsDBNull(drSet.Item("FOR_OPC_HEAD"), False)), _
                            (AltIsDBNull(drSet.Item("COMMISSION_ASSESS"), 0)), _
                            (AltIsDBNull(drSet.Item("COMMISSION_DATE"), "1/1/0001")), _
                            (AltIsDBNull(drSet.Item("COMMISSION_BY"), "")), _
                            (AltIsDBNull(drSet.Item("FOR_COMMISSION"), "0")), _
                            (AltIsDBNull(drSet.Item("ELIGIBITY_COMMENTS"), "")), drSet.Item("PMDesc"), drSet.Item("MGPTFStatusDesc"), drSet.Item("TechnicalStatusDesc"), _
                            (AltIsDBNull(drSet.Item("IRAC"), 0)), _
                            (AltIsDBNull(drSet.Item("ERAC"), 0)), _
                            drSet.Item("HOW_DISC_FAC_LEAK_DETECTION"), _
                            drSet.Item("HOW_DISC_SURFACE_SHEEN"), _
                            drSet.Item("HOW_DISC_GW_WELL"), _
                            drSet.Item("HOW_DISC_GW_CONTAMINATION"), _
                            drSet.Item("HOW_DISC_VAPORS"), _
                            drSet.Item("HOW_DISC_FREE_PRODUCT"), _
                            drSet.Item("HOW_DISC_SOIL_CONTAMINATION"), _
                            drSet.Item("HOW_DISC_FAILED_PTT"), _
                            drSet.Item("HOW_DISC_INVENTORY_SHORTAGE"), _
                            drSet.Item("HOW_DISC_TANK_CLOSURE"), _
                            drSet.Item("HOW_DISC_INSPECTION"), _
                            AltIsDBNull(drSet.Item("CAUSE"), 0) _
                            )

                Else

                    Return New MUSTER.Info.LustEventInfo
                End If
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function

        Public Function DBGetDS(ByVal strSQL As String) As DataSet
            ' #Region "XDEOperation" ' Begin Template Expansion{074F9FF7-4DD4-4F41-A9C1-F134E46BC49B}
            Dim dsData As DataSet
            Try
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            ' #End Region ' XDEOperation End Template Expansion{074F9FF7-4DD4-4F41-A9C1-F134E46BC49B}
        End Function

        Public Function DBExeNonQuery(ByVal strSQL As String)
            Dim dsData As DataSet
            Try
                SqlHelper.ExecuteNonQuery(_strConn, CommandType.Text, strSQL)
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function DBGetByFacilityID(ByVal nVal As Int64) As MUSTER.Info.LustEventCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Dim colEvents As New MUSTER.Info.LustEventCollection
            Try
                If nVal = 0 Then
                    Return colEvents
                End If
                strSQL = "spGetTecEvent"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@FACILITY_ID").Value = nVal
                Params("@EVENT_ID").Value = DBNull.Value

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                While drSet.Read
                    Dim oLustEventInfo As New MUSTER.Info.LustEventInfo(AltIsDBNull(drSet.Item("Event_ID"), 0), _
                            AltIsDBNull(drSet.Item("LAST_GWS"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_LDR"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_PTT"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_OF_REPORT"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_SITE_ASSESSED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EVENT_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EVENT_STATUS"), 0), _
                            AltIsDBNull(drSet.Item("FACILITY_ID"), 0), _
                            AltIsDBNull(drSet.Item("FLAG_ID"), 0), _
                            AltIsDBNull(drSet.Item("MGPTF_STATUS"), 0), _
                            AltIsDBNull(drSet.Item("LEAKAGE_PRIORITY"), 0), _
                            AltIsDBNull(drSet.Item("EVENT_PROJECT_MANAGER_ID"), 0), _
                            AltIsDBNull(drSet.Item("RELEASE_STATUS"), 0), _
                            0, _
                            AltIsDBNull(drSet.Item("RelatedSites"), ""), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("SUSPECTED_SOURCE"), 0), _
                            AltIsDBNull(drSet.Item("CONFIRMED_ON"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("IDENTIFIED_BY"), 0), _
                            AltIsDBNull(drSet.Item("LOCATION"), 0), _
                            AltIsDBNull(drSet.Item("EXTENT"), 0), _
                            AltIsDBNull(drSet.Item("EVENT_PROJECT_MANAGER_ID"), 0), _
                            AltIsDBNull(drSet.Item("EVENT_STARTED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EVENT_ENDED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("TOC_SOIL"), False), _
                            AltIsDBNull(drSet.Item("SOIL_BTEX"), False), _
                            AltIsDBNull(drSet.Item("SOIL_PAH"), False), _
                            AltIsDBNull(drSet.Item("SOIL_TPH"), False), _
                            AltIsDBNull(drSet.Item("TOC_GROUNDWATER"), False), _
                            AltIsDBNull(drSet.Item("GW_BTEX"), False), _
                            AltIsDBNull(drSet.Item("GW_PAH"), False), _
                            AltIsDBNull(drSet.Item("GW_TPH"), False), _
                            AltIsDBNull(drSet.Item("FREE_PRODUCT"), False), _
                            AltIsDBNull(drSet.Item("FP_GASOLINE"), False), _
                            AltIsDBNull(drSet.Item("FP_DIESEL"), False), _
                            AltIsDBNull(drSet.Item("FP_KEROSENE"), False), _
                            AltIsDBNull(drSet.Item("FP_WASTE_OIL"), False), _
                            AltIsDBNull(drSet.Item("FP_UNKNOWN"), False), _
                            AltIsDBNull(drSet.Item("TOC_VAPOR"), False), _
                            AltIsDBNull(drSet.Item("VAPOR_BTEX"), False), _
                            AltIsDBNull(drSet.Item("VAPOR_PAH"), False), _
                            AltIsDBNull(drSet.Item("EVENT_SEQUENCE"), 0), _
                            drSet.Item("TFChecklist"), _
                            drSet.Item("TankPipeList"), _
                            (AltIsDBNull(drSet.Item("PM_HEAD_ASSESS"), 0)), _
                            (AltIsDBNull(drSet.Item("PM_HEAD_DATE"), "1/1/0001")), _
                            (AltIsDBNull(drSet.Item("PM_HEAD_BY"), "")), _
                            (AltIsDBNull(drSet.Item("UST_CHIEF_ASSESS"), 0)), _
                            (AltIsDBNull(drSet.Item("UST_CHIEF_DATE"), "1/1/0001")), _
                            (AltIsDBNull(drSet.Item("UST_CHIEF_BY"), "")), _
                            (AltIsDBNull(drSet.Item("OPC_HEAD_ASSESS"), 0)), _
                            (AltIsDBNull(drSet.Item("OPC_HEAD_DATE"), "1/1/0001")), _
                            (AltIsDBNull(drSet.Item("OPC_HEAD_BY"), "")), _
                            (AltIsDBNull(drSet.Item("FOR_OPC_HEAD"), False)), _
                            (AltIsDBNull(drSet.Item("COMMISSION_ASSESS"), 0)), _
                            (AltIsDBNull(drSet.Item("COMMISSION_DATE"), "1/1/0001")), _
                            (AltIsDBNull(drSet.Item("COMMISSION_BY"), "")), _
                            (AltIsDBNull(drSet.Item("FOR_COMMISSION"), 0)), _
                            (AltIsDBNull(drSet.Item("ELIGIBITY_COMMENTS"), "")), drSet.Item("PMDesc"), drSet.Item("MGPTFStatusDesc"), drSet.Item("TechnicalStatusDesc"), _
                            (AltIsDBNull(drSet.Item("IRAC"), 0)), _
                            (AltIsDBNull(drSet.Item("ERAC"), 0)), _
                            drSet.Item("HOW_DISC_FAC_LEAK_DETECTION"), _
                            drSet.Item("HOW_DISC_SURFACE_SHEEN"), _
                            drSet.Item("HOW_DISC_GW_WELL"), _
                            drSet.Item("HOW_DISC_GW_CONTAMINATION"), _
                            drSet.Item("HOW_DISC_VAPORS"), _
                            drSet.Item("HOW_DISC_FREE_PRODUCT"), _
                            drSet.Item("HOW_DISC_SOIL_CONTAMINATION"), _
                            drSet.Item("HOW_DISC_FAILED_PTT"), _
                            drSet.Item("HOW_DISC_INVENTORY_SHORTAGE"), _
                            drSet.Item("HOW_DISC_TANK_CLOSURE"), _
                            drSet.Item("HOW_DISC_INSPECTION"), _
                            AltIsDBNull(drSet.Item("CAUSE"), 0) _
                    )
                    colEvents.Add(oLustEventInfo)
                End While

                Return colEvents
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function

        Public Function GetAllInfo() As MUSTER.Info.LustEventCollection
            ' #Region "XDEOperation" ' Begin Template Expansion{995B7456-17FF-498E-A293-B0175978420B}
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Try
                strSQL = "spGetTecEvent"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                '
                ' Repeat the following line as many times as necessary
                '
                Params("@EVENT_ID").Value = String.Empty
                Params("@FACILITY_ID").Value = DBNull.Value

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Dim colEntities As New MUSTER.Info.LustEventCollection
                While drSet.Read
                    '
                    ' Modify following statement as necessary to generate new info object
                    '
                    Dim oLustEventInfo As New MUSTER.Info.LustEventInfo(AltIsDBNull(drSet.Item("Event_ID"), 0), _
                            AltIsDBNull(drSet.Item("LAST_GWS"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_LDR"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_PTT"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_OF_REPORT"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_SITE_ASSESSED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EVENT_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EVENT_STATUS"), 0), _
                            AltIsDBNull(drSet.Item("FACILITY_ID"), 0), _
                            AltIsDBNull(drSet.Item("FLAG_ID"), 0), _
                            AltIsDBNull(drSet.Item("MGPTF_STATUS"), 0), _
                            AltIsDBNull(drSet.Item("LEAKAGE_PRIORITY"), 0), _
                            AltIsDBNull(drSet.Item("EVENT_PROJECT_MANAGER_ID"), 0), _
                            AltIsDBNull(drSet.Item("RELEASE_STATUS"), 0), _
                            0, _
                            AltIsDBNull(drSet.Item("RelatedSites"), ""), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), ""), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("SUSPECTED_SOURCE"), 0), _
                            AltIsDBNull(drSet.Item("CONFIRMED_ON"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("IDENTIFIED_BY"), 0), _
                            AltIsDBNull(drSet.Item("LOCATION"), 0), _
                            AltIsDBNull(drSet.Item("EXTENT"), 0), _
                            AltIsDBNull(drSet.Item("EVENT_PROJECT_MANAGER_ID"), 0), _
                            AltIsDBNull(drSet.Item("EVENT_STARTED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("EVENT_ENDED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("TOC_SOIL"), False), _
                            AltIsDBNull(drSet.Item("SOIL_BTEX"), False), _
                            AltIsDBNull(drSet.Item("SOIL_PAH"), False), _
                            AltIsDBNull(drSet.Item("SOIL_TPH"), False), _
                            AltIsDBNull(drSet.Item("TOC_GROUNDWATER"), False), _
                            AltIsDBNull(drSet.Item("GW_BTEX"), False), _
                            AltIsDBNull(drSet.Item("GW_PAH"), False), _
                            AltIsDBNull(drSet.Item("GW_TPH"), False), _
                            AltIsDBNull(drSet.Item("FREE_PRODUCT"), False), _
                            AltIsDBNull(drSet.Item("FP_GASOLINE"), False), _
                            AltIsDBNull(drSet.Item("FP_DIESEL"), False), _
                            AltIsDBNull(drSet.Item("FP_KEROSENE"), False), _
                            AltIsDBNull(drSet.Item("FP_WASTE_OIL"), False), _
                            AltIsDBNull(drSet.Item("FP_UNKNOWN"), False), _
                            AltIsDBNull(drSet.Item("TOC_VAPOR"), False), _
                            AltIsDBNull(drSet.Item("VAPOR_BTEX"), False), _
                            AltIsDBNull(drSet.Item("VAPOR_PAH"), False), _
                            AltIsDBNull(drSet.Item("EVENT_SEQUENCE"), 0), _
                            drSet.Item("TFChecklist"), _
                            drSet.Item("TankPipeList"), _
                            (AltIsDBNull(drSet.Item("PM_HEAD_ASSESS"), 0)), _
                            (AltIsDBNull(drSet.Item("PM_HEAD_DATE"), "1/1/0001")), _
                            (AltIsDBNull(drSet.Item("PM_HEAD_BY"), "")), _
                            (AltIsDBNull(drSet.Item("UST_CHIEF_ASSESS"), 0)), _
                            (AltIsDBNull(drSet.Item("UST_CHIEF_DATE"), "1/1/0001")), _
                            (AltIsDBNull(drSet.Item("UST_CHIEF_BY"), "")), _
                            (AltIsDBNull(drSet.Item("OPC_HEAD_ASSESS"), 0)), _
                            (AltIsDBNull(drSet.Item("OPC_HEAD_DATE"), "1/1/0001")), _
                            (AltIsDBNull(drSet.Item("OPC_HEAD_BY"), "")), _
                            (AltIsDBNull(drSet.Item("FOR_OPC_HEAD"), False)), _
                            (AltIsDBNull(drSet.Item("COMMISSION_ASSESS"), 0)), _
                            (AltIsDBNull(drSet.Item("COMMISSION_DATE"), "1/1/0001")), _
                            (AltIsDBNull(drSet.Item("COMMISSION_BY"), "")), _
                            (AltIsDBNull(drSet.Item("FOR_COMMISSION"), "0")), _
                            (AltIsDBNull(drSet.Item("ELIGIBITY_COMMENTS"), "")), drSet.Item("PMDesc"), drSet.Item("MGPTFStatusDesc"), drSet.Item("TechnicalStatusDesc"), _
                            (AltIsDBNull(drSet.Item("IRAC"), 0)), _
                            (AltIsDBNull(drSet.Item("ERAC"), 0)), _
                            drSet.Item("HOW_DISC_FAC_LEAK_DETECTION"), _
                            drSet.Item("HOW_DISC_SURFACE_SHEEN"), _
                            drSet.Item("HOW_DISC_GW_WELL"), _
                            drSet.Item("HOW_DISC_GW_CONTAMINATION"), _
                            drSet.Item("HOW_DISC_VAPORS"), _
                            drSet.Item("HOW_DISC_FREE_PRODUCT"), _
                            drSet.Item("HOW_DISC_SOIL_CONTAMINATION"), _
                            drSet.Item("HOW_DISC_FAILED_PTT"), _
                            drSet.Item("HOW_DISC_INVENTORY_SHORTAGE"), _
                            drSet.Item("HOW_DISC_TANK_CLOSURE"), _
                            drSet.Item("HOW_DISC_INSPECTION"), _
                            AltIsDBNull(drSet.Item("CAUSE"), 0) _
                   )
                    colEntities.Add(oLustEventInfo)
                End While

                Return colEntities
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet.IsClosed Then drSet.Close()
            End Try
            ' #End Region ' XDEOperation End Template Expansion{995B7456-17FF-498E-A293-B0175978420B}
        End Function

        Public Sub Put(ByRef oLustEvent As MUSTER.Info.LustEventInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            ' #Region "XDEOperation" ' Begin Template Expansion{CE44BE18-660C-4873-BF52-C01555F731C8}
            Dim Params() As SqlParameter
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.LUST_Event, Integer))) Then
                    returnVal = "You do not have rights to save Lust Event."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                ' #2188
                If oLustEvent.EventStatus = 625 And oLustEvent.EventStatusOriginal = 624 Then
                    oLustEvent.Priority = 0
                End If

                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutTecEvent")
                If oLustEvent.ID = 0 Then
                    Params(0).Value = System.DBNull.Value
                Else
                    Params(0).Value = oLustEvent.ID
                End If
                If oLustEvent.FacilityID = 0 Then
                    Params(1).Value = System.DBNull.Value
                Else
                    Params(1).Value = oLustEvent.FacilityID
                End If
                If oLustEvent.EventStatus = 0 Then
                    Params(2).Value = System.DBNull.Value
                Else
                    Params(2).Value = oLustEvent.EventStatus
                End If
                If oLustEvent.Priority = 0 Then
                    Params(3).Value = System.DBNull.Value
                Else
                    Params(3).Value = oLustEvent.Priority
                End If
                If oLustEvent.MGPTFStatus = 0 Then
                    Params(4).Value = System.DBNull.Value
                Else
                    Params(4).Value = oLustEvent.MGPTFStatus
                End If
                Params(5).Value = IIFIsDateNull(oLustEvent.LastLDR, DBNull.Value)
                Params(6).Value = IIFIsDateNull(oLustEvent.LastPTT, DBNull.Value)
                Params(7).Value = IIFIsDateNull(oLustEvent.LastGWS, DBNull.Value)
                'If oLustEvent.MGPTFStatus = 0 Then
                Params(8).Value = System.DBNull.Value
                'Else
                '    Params(8).Value = oLustEvent.MGPTFStatus
                'End If
                Params(9).Value = oLustEvent.Started
                Params(10).Value = IIFIsDateNull(oLustEvent.ReportDate, DBNull.Value)
                Params(11).Value = oLustEvent.SuspectedSource
                Params(12).Value = oLustEvent.PM
                Params(13).Value = oLustEvent.ReleaseStatus
                'Params(14).Value = oLustEvent.HowDiscoveredID
                Params(14).Value = IIFIsDateNull(oLustEvent.Confirmed, DBNull.Value)
                Params(15).Value = oLustEvent.IDENTIFIEDBY
                Params(16).Value = oLustEvent.Location
                Params(17).Value = oLustEvent.Extent
                Params(18).Value = IIFIsDateNull(oLustEvent.EventStarted, oLustEvent.Started)
                Params(19).Value = IIFIsDateNull(oLustEvent.EventEnded, DBNull.Value)
                Params(20).Value = oLustEvent.TOCSOIL
                Params(21).Value = oLustEvent.SOILBTEX
                Params(22).Value = oLustEvent.SOILPAH
                Params(23).Value = oLustEvent.SOILTPH
                Params(24).Value = oLustEvent.TOCGROUNDWATER
                Params(25).Value = oLustEvent.GWBTEX
                Params(26).Value = oLustEvent.GWPAH
                Params(27).Value = oLustEvent.GWTPH
                Params(28).Value = oLustEvent.FREEPRODUCT
                Params(29).Value = oLustEvent.FPGASOLINE
                Params(30).Value = oLustEvent.FPDIESEL
                Params(31).Value = oLustEvent.FPKEROSENE
                Params(32).Value = oLustEvent.FPWASTEOIL
                Params(33).Value = oLustEvent.FPUNKNOWN
                Params(34).Value = oLustEvent.TOCVAPOR
                Params(35).Value = oLustEvent.VAPORBTEX
                Params(36).Value = oLustEvent.VAPORPAH
                Params(37).Value = oLustEvent.EVENTSEQUENCE
                Params(38).Value = oLustEvent.Deleted
                If IsNothing(oLustEvent.TankandPipe) Then
                    Params(39).Value = String.Empty
                Else
                    Params(39).Value = oLustEvent.TankandPipe
                End If

                Params(40).Value = oLustEvent.TFCheckList
                Params(41).Value = oLustEvent.PM_HEAD_ASSESS
                Params(42).Value = IIFIsDateNull(oLustEvent.PM_HEAD_DATE, DBNull.Value)
                Params(43).Value = IIf(oLustEvent.PM_HEAD_BY Is Nothing, String.Empty, oLustEvent.PM_HEAD_BY)
                Params(44).Value = oLustEvent.UST_CHIEF_ASSESS
                Params(45).Value = IIFIsDateNull(oLustEvent.UST_CHIEF_DATE, DBNull.Value)
                Params(46).Value = IIf(oLustEvent.UST_CHIEF_BY Is Nothing, String.Empty, oLustEvent.UST_CHIEF_BY)
                Params(47).Value = oLustEvent.OPC_HEAD_ASSESS
                Params(48).Value = IIFIsDateNull(oLustEvent.OPC_HEAD_DATE, DBNull.Value)
                Params(49).Value = IIf(oLustEvent.OPC_HEAD_BY Is Nothing, String.Empty, oLustEvent.OPC_HEAD_BY)
                Params(50).Value = oLustEvent.COMMISSION_ASSESS
                Params(51).Value = IIFIsDateNull(oLustEvent.COMMISSION_DATE, DBNull.Value)
                Params(52).Value = IIf(oLustEvent.COMMISSION_BY Is Nothing, String.Empty, oLustEvent.COMMISSION_BY)
                Params(53).Value = oLustEvent.FOR_OPC_HEAD
                Params(54).Value = oLustEvent.FOR_COMMISSION
                Params(55).Value = IIf(oLustEvent.ELIGIBITY_COMMENTS Is Nothing, String.Empty, oLustEvent.ELIGIBITY_COMMENTS)
                Params(56).Value = IIf(oLustEvent.RelatedSites Is Nothing, String.Empty, oLustEvent.RelatedSites)

                Params(57).Value = oLustEvent.IRAC
                Params(58).Value = oLustEvent.ERAC

                If oLustEvent.ID <= 0 Then
                    Params(59).Value = oLustEvent.CreatedBy
                Else
                    Params(59).Value = oLustEvent.ModifiedBy
                End If
                Params(60).Value = System.DBNull.Value
                Params(61).Value = System.DBNull.Value
                Params(62).Value = System.DBNull.Value
                Params(63).Value = System.DBNull.Value

                Params(64).Value = oLustEvent.HowDiscFacLD
                Params(65).Value = oLustEvent.HowDiscSurfaceSheen
                Params(66).Value = oLustEvent.HowDiscGWWell
                Params(67).Value = oLustEvent.HowDiscGWContamination
                Params(68).Value = oLustEvent.HowDiscVapors
                Params(69).Value = oLustEvent.HowDiscFreeProduct
                Params(70).Value = oLustEvent.HowDiscSoilContamination
                Params(71).Value = oLustEvent.HowDiscFailedPTT
                Params(72).Value = oLustEvent.HowDiscInventoryShortage
                Params(73).Value = oLustEvent.HowDiscTankClosure
                Params(74).Value = oLustEvent.HowDiscInspection
                Params(75).Value = oLustEvent.Cause
                Params(76).Value = IIFIsDateNull(oLustEvent.CompAssDate, DBNull.Value)

                SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutTecEvent", Params)
                '
                ' Perform check for New ID and assign, if necessary
                If Params(0).Value <> oLustEvent.ID Then
                    oLustEvent.ID = Params(0).Value
                    oLustEvent.EVENTSEQUENCE = Params(37).Value
                End If
                oLustEvent.ModifiedBy = AltIsDBNull(Params(60).Value, String.Empty)
                oLustEvent.ModifiedOn = AltIsDBNull(Params(61).Value, CDate("01/01/0001"))
                oLustEvent.CreatedBy = AltIsDBNull(Params(62).Value, String.Empty)
                oLustEvent.CreatedOn = AltIsDBNull(Params(63).Value, CDate("01/01/0001"))
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            ' #End Region ' XDEOperation End Template Expansion{CE44BE18-660C-4873-BF52-C01555F731C8}
        End Sub

        Public Function DBGetTecOpenFinPO(ByVal nEventID As Integer) As DataSet
            ' #Region "XDEOperation" ' Begin Template Expansion{074F9FF7-4DD4-4F41-A9C1-F134E46BC49B}
            Dim dsData As DataSet
            Dim strVal As String
            Dim strSQL As String
            Dim Params As Collection
            Try
                strSQL = "spGetTecOpenFinPOs"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                '
                ' Repeat the following line as many times as necessary
                '
                Params("@EVENT_ID").Value = nEventID
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.StoredProcedure, strSQL, Params)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            ' #End Region ' XDEOperation End Template Expansion{074F9FF7-4DD4-4F41-A9C1-F134E46BC49B}
        End Function
    End Class
End Namespace



