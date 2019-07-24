'-------------------------------------------------------------------------------
' MUSTER.DataAccess.LustRemediationDB
'   Provides the means for marshalling Lust Remediation state to/from the repository
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        JVC     05/03/05    Original class definition.
'  1.1        AB      05/03/05    extended Original Class Definition
'
' Function                  Description
' DBGetByID(ID)         Returns an LustRemediation Object indicated by Lust Remediation ID
' DBGetByEventID(ID)    Returns an LustRemediation Collection indicated by Lust Event ID
' DBGetDS(SQL)          Returns a resultant Dataset by running query specified by the string arg SQL
' Put(LustEvent)        Saves the LustRemediation passed as an argument, to the DB
'-------------------------------------------------------------------------------
'
'
Imports System.Data.SqlClient

Imports Utils.DBUtils

Namespace MUSTER.DataAccess
    Public Class LustRemediationDB
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

        ' Retrieve a remediation system by it's RemSysID
        Public Function DBGetByID(ByVal nVal As Int64) As MUSTER.Info.LustRemediationInfo
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Dim narrayLength As Int16
            Dim nOption1 As Int64 = 0
            Dim nOption2 As Int64 = 0
            Dim nOption3 As Int64 = 0
            Dim aryOptEquip As Array
            Dim oLustRemediationInfo As New MUSTER.Info.LustRemediationInfo
            Try
                If nVal = 0 Then
                    Return oLustRemediationInfo
                End If
                strSQL = "spGetTecRemediationEvent_BySystemID"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                Params("@REM_SYSTEM_ID").Value = nVal
                'Params("@ACTIVITY_ID").Value = 0
                'Params("@EVENT_ID").Value = 0

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                While drSet.Read
                    If AltIsDBNull(drSet.Item("OptEquipment"), "") <> "" Then
                        aryOptEquip = Split(AltIsDBNull(drSet.Item("OptEquipment"), ""), "|")
                        If aryOptEquip.Length >= 1 Then
                            nOption1 = aryOptEquip(0)
                        End If
                        If aryOptEquip.Length >= 2 Then
                            nOption2 = aryOptEquip(1)
                        End If
                        If aryOptEquip.Length >= 3 Then
                            nOption3 = aryOptEquip(2)
                        End If
                    End If


                    oLustRemediationInfo = New MUSTER.Info.LustRemediationInfo(drSet.Item("REM_SYSTEM_ID"), _
                            drSet.Item("SYSTEM_SEQ"), _
                            drSet.Item("SYSTEM_DEC"), _
                            AltIsDBNull(drSet.Item("START_DATE"), "01/01/0001"), _
                            drSet.Item("REM_SYSTEM_TYPE"), _
                            drSet.Item("DESCRIPTION"), _
                            drSet.Item("MANUFACTURER"), _
                            AltIsDBNull(drSet.Item("OWS_SIZE"), ""), _
                            AltIsDBNull(drSet.Item("OWS_MAN_NAME"), ""), _
                            AltIsDBNull(drSet.Item("OWS_SERIAL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("OWS_MODEL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("OWS_USED_NEW"), 0), _
                            AltIsDBNull(drSet.Item("OWS_AGE_OF_COMPONENT"), ""), _
                            AltIsDBNull(drSet.Item("MOTOR_SIZE"), ""), _
                            AltIsDBNull(drSet.Item("MOTOR_MAN_NAME"), ""), _
                            AltIsDBNull(drSet.Item("MOTOR_SERIAL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("MOTOR_MODEL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("MOTOR_USED_NEW"), 0), _
                            AltIsDBNull(drSet.Item("MOTOR_AGE_OF_COMPONENT"), ""), _
                            AltIsDBNull(drSet.Item("STEPPER_SIZE"), ""), _
                            AltIsDBNull(drSet.Item("STEPPER_MAN_NAME"), ""), _
                            AltIsDBNull(drSet.Item("STEPPER_SERIAL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("STEPPER_MODEL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("STEPPER_USED_NEW"), 0), _
                            AltIsDBNull(drSet.Item("STEPPER_AGE_OF_COMPONENT"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP1_SIZE"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP1_MAN_NAME"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP1_SERIAL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP1_MODEL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP1_USED_NEW"), 0), _
                            AltIsDBNull(drSet.Item("VACPUMP1_AGE_OF_COMPONENT"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP1_SEAL"), 0), _
                            AltIsDBNull(drSet.Item("VACPUMP2_SIZE"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP2_MAN_NAME"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP2_SERIAL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP2_MODEL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP2_USED_NEW"), 0), _
                            AltIsDBNull(drSet.Item("VACPUMP2_AGE_OF_COMPONENT"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP2_SEAL"), 0), _
                            AltIsDBNull(drSet.Item("OWNED_LEASED"), 0), _
                            AltIsDBNull(drSet.Item("SYSTEM_OWNER"), ""), _
                            AltIsDBNull(drSet.Item("BUILDING_SIZE"), ""), _
                            AltIsDBNull(drSet.Item("MOUNT_TYPE_ID"), 0), _
                            AltIsDBNull(drSet.Item("SYSTEM_REFURBISHED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("OTHER"), ""), _
                            AltIsDBNull(drSet.Item("PURCHASE_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DELETED"), False), _
                            nOption1, _
                            nOption2, _
                            nOption3 _
                    )
                End While

                Return oLustRemediationInfo
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            Finally
                If Not drSet Is Nothing Then
                    If Not drSet.IsClosed Then drSet.Close()
                End If
            End Try
        End Function

        Public Function DBGetByActivityID(ByVal nVal As Int64) As MUSTER.Info.LustRemediationCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Dim nOption1 As Int64 = 0
            Dim nOption2 As Int64 = 0
            Dim nOption3 As Int64 = 0
            Dim aryOptEquip As Array
            Dim colEvents As New MUSTER.Info.LustRemediationCollection
            Try
                If nVal = 0 Then
                    Return colEvents
                End If
                strSQL = "spGetTecRemediationEvent"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                'Params("@REM_SYSTEM_ID").Value = 0
                Params("@ACTIVITY_ID").Value = nVal
                Params("@EVENT_ID").Value = 0

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                While drSet.Read

                    If AltIsDBNull(drSet.Item("OptEquipment"), "") <> "" Then
                        aryOptEquip = Split(AltIsDBNull(drSet.Item("OptEquipment"), ""), "|")
                        If aryOptEquip.Length >= 1 Then
                            nOption1 = aryOptEquip(0)
                        End If
                        If aryOptEquip.Length >= 2 Then
                            nOption2 = aryOptEquip(1)
                        End If
                        If aryOptEquip.Length >= 3 Then
                            nOption3 = aryOptEquip(2)
                        End If
                    End If

                    Dim oLustRemediationInfo As New MUSTER.Info.LustRemediationInfo(drSet.Item("REM_SYSTEM_ID"), _
                            drSet.Item("SYSTEM_SEQ"), _
                            drSet.Item("SYSTEM_DEC"), _
                            AltIsDBNull(drSet.Item("START_DATE"), "01/01/0001"), _
                            drSet.Item("REM_SYSTEM_TYPE"), _
                            drSet.Item("DESCRIPTION"), _
                            drSet.Item("Manufacturer"), _
                            AltIsDBNull(drSet.Item("OWS_SIZE"), ""), _
                            AltIsDBNull(drSet.Item("OWS_MAN_NAME"), ""), _
                            AltIsDBNull(drSet.Item("OWS_SERIAL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("OWS_MODEL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("OWS_USED_NEW"), 0), _
                            AltIsDBNull(drSet.Item("OWS_AGE_OF_COMPONENT"), ""), _
                            AltIsDBNull(drSet.Item("MOTOR_SIZE"), ""), _
                            AltIsDBNull(drSet.Item("MOTOR_MAN_NAME"), ""), _
                            AltIsDBNull(drSet.Item("MOTOR_SERIAL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("MOTOR_MODEL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("MOTOR_USED_NEW"), 0), _
                            AltIsDBNull(drSet.Item("MOTOR_AGE_OF_COMPONENT"), ""), _
                            AltIsDBNull(drSet.Item("STEPPER_SIZE"), ""), _
                            AltIsDBNull(drSet.Item("STEPPER_MAN_NAME"), ""), _
                            AltIsDBNull(drSet.Item("STEPPER_SERIAL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("STEPPER_MODEL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("STEPPER_USED_NEW"), 0), _
                            AltIsDBNull(drSet.Item("STEPPER_AGE_OF_COMPONENT"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP1_SIZE"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP1_MAN_NAME"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP1_SERIAL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP1_MODEL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP1_USED_NEW"), 0), _
                            AltIsDBNull(drSet.Item("VACPUMP1_AGE_OF_COMPONENT"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP1_SEAL"), 0), _
                            AltIsDBNull(drSet.Item("VACPUMP2_SIZE"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP2_MAN_NAME"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP2_SERIAL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP2_MODEL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP2_USED_NEW"), 0), _
                            AltIsDBNull(drSet.Item("VACPUMP2_AGE_OF_COMPONENT"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP2_SEAL"), 0), _
                            AltIsDBNull(drSet.Item("OWNED_LEASED"), 0), _
                            AltIsDBNull(drSet.Item("SYSTEM_OWNER"), ""), _
                            AltIsDBNull(drSet.Item("BUILDING_SIZE"), ""), _
                            AltIsDBNull(drSet.Item("MOUNT_TYPE_ID"), 0), _
                            AltIsDBNull(drSet.Item("SYSTEM_REFURBISHED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("OTHER"), ""), _
                            AltIsDBNull(drSet.Item("PURCHASE_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DELETED"), False), _
                            nOption1, _
                            nOption2, _
                            nOption3 _
                    )
                    colEvents.Add(oLustRemediationInfo)
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

        Public Function DBGetByEventID(ByVal nVal As Int64) As MUSTER.Info.LustRemediationCollection
            Dim strSQL As String
            Dim Params As Collection
            Dim drSet As SqlDataReader
            Dim nOption1 As Int64 = 0
            Dim nOption2 As Int64 = 0
            Dim nOption3 As Int64 = 0
            Dim aryOptEquip As Array
            Dim colEvents As New MUSTER.Info.LustRemediationCollection
            Try
                If nVal = 0 Then
                    Return colEvents
                End If
                strSQL = "spGetTecRemediationEvent"

                Params = SqlHelperParameterCache.GetSpParameterCol(_strConn, strSQL)
                'Params("@REM_SYSTEM_ID").Value = 0
                Params("@EVENT_ID").Value = nVal
                Params("@ACTIVITY_ID").Value = 0

                drSet = SqlHelper.ExecuteReader(_strConn, CommandType.StoredProcedure, strSQL, Params)
                While drSet.Read

                    If AltIsDBNull(drSet.Item("OptEquipment"), "") <> "" Then
                        aryOptEquip = Split(AltIsDBNull(drSet.Item("OptEquipment"), ""), "|")
                        If aryOptEquip.Length >= 1 Then
                            nOption1 = aryOptEquip(0)
                        End If
                        If aryOptEquip.Length >= 2 Then
                            nOption2 = aryOptEquip(1)
                        End If
                        If aryOptEquip.Length >= 3 Then
                            nOption3 = aryOptEquip(2)
                        End If
                    End If

                    Dim oLustRemediationInfo As New MUSTER.Info.LustRemediationInfo(drSet.Item("REM_SYSTEM_ID"), _
                            drSet.Item("SYSTEM_SEQ"), _
                            drSet.Item("SYSTEM_DEC"), _
                            AltIsDBNull(drSet.Item("START_DATE"), "01/01/0001"), _
                            drSet.Item("REM_SYSTEM_TYPE"), _
                            drSet.Item("DESCRIPTION"), _
                            drSet.Item("Manufacturer"), _
                            AltIsDBNull(drSet.Item("OWS_SIZE"), ""), _
                            AltIsDBNull(drSet.Item("OWS_MAN_NAME"), ""), _
                            AltIsDBNull(drSet.Item("OWS_SERIAL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("OWS_MODEL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("OWS_USED_NEW"), 0), _
                            AltIsDBNull(drSet.Item("OWS_AGE_OF_COMPONENT"), ""), _
                            AltIsDBNull(drSet.Item("MOTOR_SIZE"), ""), _
                            AltIsDBNull(drSet.Item("MOTOR_MAN_NAME"), ""), _
                            AltIsDBNull(drSet.Item("MOTOR_SERIAL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("MOTOR_MODEL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("MOTOR_USED_NEW"), 0), _
                            AltIsDBNull(drSet.Item("MOTOR_AGE_OF_COMPONENT"), ""), _
                            AltIsDBNull(drSet.Item("STEPPER_SIZE"), ""), _
                            AltIsDBNull(drSet.Item("STEPPER_MAN_NAME"), ""), _
                            AltIsDBNull(drSet.Item("STEPPER_SERIAL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("STEPPER_MODEL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("STEPPER_USED_NEW"), 0), _
                            AltIsDBNull(drSet.Item("STEPPER_AGE_OF_COMPONENT"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP1_SIZE"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP1_MAN_NAME"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP1_SERIAL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP1_MODEL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP1_USED_NEW"), 0), _
                            AltIsDBNull(drSet.Item("VACPUMP1_AGE_OF_COMPONENT"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP1_SEAL"), 0), _
                            AltIsDBNull(drSet.Item("VACPUMP2_SIZE"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP2_MAN_NAME"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP2_SERIAL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP2_MODEL_NUMBER"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP2_USED_NEW"), 0), _
                            AltIsDBNull(drSet.Item("VACPUMP2_AGE_OF_COMPONENT"), ""), _
                            AltIsDBNull(drSet.Item("VACPUMP2_SEAL"), 0), _
                            AltIsDBNull(drSet.Item("OWNED_LEASED"), 0), _
                            AltIsDBNull(drSet.Item("SYSTEM_OWNER"), ""), _
                            AltIsDBNull(drSet.Item("BUILDING_SIZE"), ""), _
                            AltIsDBNull(drSet.Item("MOUNT_TYPE_ID"), 0), _
                            AltIsDBNull(drSet.Item("SYSTEM_REFURBISHED_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("OTHER"), ""), _
                            AltIsDBNull(drSet.Item("PURCHASE_DATE"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("CREATED_BY"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_CREATED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("LAST_EDITED_BY"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DATE_LAST_EDITED"), "1/1/0001"), _
                            AltIsDBNull(drSet.Item("DELETED"), False), _
                            nOption1, _
                            nOption2, _
                            nOption3 _
                    )
                    colEvents.Add(oLustRemediationInfo)
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


        Public Sub Put(ByRef oLustRemediation As MUSTER.Info.LustRemediationInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim Params() As SqlParameter
            Try

                If Not (SqlHelper.HasWriteAccess(moduleID, staffID, CType(SqlHelper.EntityTypes.LustRemediation, Integer))) Then
                    returnVal = "You do not have rights to save Lust Remediation."
                    Exit Sub
                Else
                    returnVal = String.Empty
                End If

                Params = SqlHelperParameterCache.GetSpParameterSet(_strConn, "spPutLustRemediation")

                If oLustRemediation.ID <= 0 Then
                    Params(0).Value = 0
                Else
                    Params(0).Value = oLustRemediation.ID
                End If
                Params(1).Value = oLustRemediation.SystemSequence
                Params(2).Value = oLustRemediation.SystemDeclaration
                Params(3).Value = IIFIsDateNull(oLustRemediation.DatePlacedInUse, System.DBNull.Value)
                Params(4).Value = oLustRemediation.RemSysType
                Params(5).Value = IIf(IsNothing(oLustRemediation.Description), "", oLustRemediation.Description)
                Params(6).Value = IIf(IsNothing(oLustRemediation.Manufacturer), "", oLustRemediation.Manufacturer)

                Params(7).Value = IIf(IsNothing(oLustRemediation.OWSSize), "", oLustRemediation.OWSSize)
                Params(8).Value = IIf(IsNothing(oLustRemediation.OWSManName), "", oLustRemediation.OWSManName)
                Params(9).Value = IIf(IsNothing(oLustRemediation.OWSSerialNumber), "", oLustRemediation.OWSSerialNumber)
                Params(10).Value = IIf(IsNothing(oLustRemediation.OWSModelNumber), "", oLustRemediation.OWSModelNumber)
                Params(11).Value = oLustRemediation.OWSNewUsed
                Params(12).Value = IIf(IsNothing(oLustRemediation.OWSAgeofComp), "", oLustRemediation.OWSAgeofComp)

                Params(13).Value = IIf(IsNothing(oLustRemediation.MotorSize), "", oLustRemediation.MotorSize)
                Params(14).Value = IIf(IsNothing(oLustRemediation.MotorManName), "", oLustRemediation.MotorManName)
                Params(15).Value = IIf(IsNothing(oLustRemediation.MotorSerialNumber), "", oLustRemediation.MotorSerialNumber)
                Params(16).Value = IIf(IsNothing(oLustRemediation.MotorModelNumber), "", oLustRemediation.MotorModelNumber)
                Params(17).Value = oLustRemediation.MotorNewUsed
                Params(18).Value = IIf(IsNothing(oLustRemediation.MotorAgeofComp), "", oLustRemediation.MotorAgeofComp)

                Params(19).Value = IIf(IsNothing(oLustRemediation.StripperSize), "", oLustRemediation.StripperSize)
                Params(20).Value = IIf(IsNothing(oLustRemediation.StripperManName), "", oLustRemediation.StripperManName)
                Params(21).Value = IIf(IsNothing(oLustRemediation.StripperSerialNumber), "", oLustRemediation.StripperSerialNumber)
                Params(22).Value = IIf(IsNothing(oLustRemediation.StripperModelNumber), "", oLustRemediation.StripperModelNumber)
                Params(23).Value = oLustRemediation.StripperNewUsed
                Params(24).Value = IIf(IsNothing(oLustRemediation.StripperAgeofComp), "", oLustRemediation.StripperAgeofComp)

                Params(25).Value = IIf(IsNothing(oLustRemediation.VacPump1Size), "", oLustRemediation.VacPump1Size)
                Params(26).Value = IIf(IsNothing(oLustRemediation.VacPump1ManName), "", oLustRemediation.VacPump1ManName)
                Params(27).Value = IIf(IsNothing(oLustRemediation.VacPump1SerialNumber), "", oLustRemediation.VacPump1SerialNumber)
                Params(28).Value = IIf(IsNothing(oLustRemediation.VacPump1ModelNumber), "", oLustRemediation.VacPump1ModelNumber)
                Params(29).Value = oLustRemediation.VacPump1NewUsed
                Params(30).Value = IIf(IsNothing(oLustRemediation.VacPump1AgeofComp), "", oLustRemediation.VacPump1AgeofComp)
                Params(31).Value = oLustRemediation.VacPump1Seal


                Params(32).Value = IIf(IsNothing(oLustRemediation.VacPump2Size), "", oLustRemediation.VacPump2Size)
                Params(33).Value = IIf(IsNothing(oLustRemediation.VacPump2ManName), "", oLustRemediation.VacPump2ManName)
                Params(34).Value = IIf(IsNothing(oLustRemediation.VacPump2SerialNumber), "", oLustRemediation.VacPump2SerialNumber)
                Params(35).Value = IIf(IsNothing(oLustRemediation.VacPump2ModelNumber), "", oLustRemediation.VacPump2ModelNumber)
                Params(36).Value = oLustRemediation.VacPump2NewUsed
                Params(37).Value = IIf(IsNothing(oLustRemediation.VacPump2AgeofComp), "", oLustRemediation.VacPump2AgeofComp)
                Params(38).Value = oLustRemediation.VacPump2Seal

                Params(39).Value = oLustRemediation.Owned
                Params(40).Value = IIf(IsNothing(oLustRemediation.Owner), "", oLustRemediation.Owner)
                Params(41).Value = IIf(IsNothing(oLustRemediation.BuildingSize), "", oLustRemediation.BuildingSize)
                Params(42).Value = oLustRemediation.MountType
                Params(43).Value = IIFIsDateNull(oLustRemediation.RefurbDate, System.DBNull.Value)
                Params(44).Value = IIf(IsNothing(oLustRemediation.Notes), "", oLustRemediation.Notes)
                Params(45).Value = IIFIsDateNull(oLustRemediation.PurchaseDate, System.DBNull.Value)
                Params(46).Value = oLustRemediation.Deleted
                Params(47).Value = oLustRemediation.Option1
                Params(48).Value = oLustRemediation.Option2
                Params(49).Value = oLustRemediation.Option3
                If oLustRemediation.ID <= 0 Then
                    If oLustRemediation.CreatedBy Is Nothing Then
                        oLustRemediation.CreatedBy = oLustRemediation.ModifiedBy
                    End If
                    Params(50).Value = oLustRemediation.CreatedBy
                Else
                    Params(50).Value = oLustRemediation.ModifiedBy
                End If


                Dim nResult As Integer = SqlHelper.ExecuteNonQuery(_strConn, CommandType.StoredProcedure, "spPutLustRemediation", Params)

                If oLustRemediation.ID <= 0 Then
                    oLustRemediation.ID = Params(0).Value
                End If
                If oLustRemediation.SystemSequence <= 0 Then
                    oLustRemediation.SystemSequence = Params(1).Value
                End If
                If oLustRemediation.SystemDeclaration <= 0 Then
                    oLustRemediation.SystemDeclaration = Params(2).Value
                End If
                'oLustRemediation.ModifiedBy = AltIsDBNull(Params(36).Value, String.Empty)
                'oLustRemediation.ModifiedOn = AltIsDBNull(Params(37).Value, CDate("01/01/0001"))
                'oLustRemediation.CreatedBy = AltIsDBNull(Params(38).Value, String.Empty)
                'oLustRemediation.CreatedOn = AltIsDBNull(Params(39).Value, CDate("01/01/0001"))

            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

        Public Function DBGetDS(ByVal strSQL As String) As DataSet
            Dim dsData As DataSet
            Try
                dsData = SqlHelper.ExecuteDataset(_strConn, CommandType.Text, strSQL)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

    End Class
End Namespace