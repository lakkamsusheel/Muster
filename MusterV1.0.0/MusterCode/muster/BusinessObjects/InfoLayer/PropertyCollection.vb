'-------------------------------------------------------------------------------
' MUSTER.Info.MusterPropertyCollection
'   Provides a stongly-typed collection for storing Entity objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        JVC2      11/19/04    Original class definition.
'  1.1        MR        03/18/05    Removed Empty string assigned to Created On and Modified On.
'
' Function          Description

'-------------------------------------------------------------------------------
Namespace MUSTER.Info
    Public Class PropertyCollection
        Inherits DictionaryBase
#Region "Public Events"
        Public Event InfoChanged()
#End Region
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.PropertyInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.PropertyInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.PropertyInfo)
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(Value.Key, Value)
                    RaiseEvent InfoChanged()
                Else
                    MyBase.Dictionary.Item(index) = Value
                    If Value.IsDirty Then
                        RaiseEvent InfoChanged()
                    End If
                End If
            End Set
        End Property
        Public ReadOnly Property Values() As ICollection
            Get
                Return MyBase.Dictionary.Values
            End Get
        End Property
        Public ReadOnly Property GetKeys() As String()
            Get
                Dim KeyCol(MyBase.Dictionary.Keys.Count - 1) As String
                MyBase.Dictionary.Keys.CopyTo(KeyCol, 0)
                Array.Sort(KeyCol)
                Return KeyCol
            End Get
        End Property

        Public Sub Add(ByVal value As MUSTER.Info.PropertyInfo)
            Me.Item(value.Key) = value
        End Sub
        Public Function Contains(ByVal value As MUSTER.Info.PropertyInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.Key)
        End Function
        Public Function Contains(ByVal Name As String) As Boolean
            Return MyBase.Dictionary.Contains(Name)
        End Function
        Public Sub Remove(ByVal value As MUSTER.Info.PropertyInfo)
            MyBase.Dictionary.Remove(value.Key)
        End Sub




        Public Function PropertiesTable(Optional ByVal PropType As Int64 = 0) As DataTable
            Dim dr As DataRow

            Dim tblPropTypes As New DataTable
            Dim oPropertyInfo As MUSTER.Info.PropertyInfo
            Dim colSubProps As Collection
            Dim nIndex As String = IIf(PropType = 0, "PRIMARY", PropType.ToString & "S")

            Try

                tblPropTypes.Columns.Add("Property ID")
                tblPropTypes.Columns.Add("Property Name")
                tblPropTypes.Columns.Add("Parent Property")
                tblPropTypes.Columns.Add("Property Position", GetType(System.Int32))
                tblPropTypes.Columns.Add("Property Description")
                tblPropTypes.Columns.Add("Property Active", GetType(System.Boolean))
                tblPropTypes.Columns.Add("Created By")
                tblPropTypes.Columns.Add("Created On", Type.GetType("System.DateTime"))
                tblPropTypes.Columns.Add("Modified By")
                tblPropTypes.Columns.Add("Modified On", Type.GetType("System.DateTime"))
                tblPropTypes.Columns("Property ID").DefaultValue = 0
                tblPropTypes.Columns("Parent Property").DefaultValue = 0
                tblPropTypes.Columns("Property Active").DefaultValue = True
                tblPropTypes.Columns("Created By").DefaultValue = ""
                tblPropTypes.Columns("Modified By").DefaultValue = ""


                'tblPropTypes.Constraints.Add("MustNotOverlap", tblPropTypes.Columns("Property Position"), False)
                'tblPropTypes.Columns("Property Name").AllowDBNull = False
                'tblPropTypes.Columns("Property Position").AllowDBNull = False
                tblPropTypes.Columns("Property Position").Unique = True
                tblPropTypes.Columns("Property Name").Unique = True




                'If Not colProperties Is Nothing Then
                If Me.Count > 0 Then
                    Try
                        'colSubProps = colProperties(nIndex)
                        For Each oPropertyInfo In Me.Values
                            dr = tblPropTypes.NewRow()
                            dr("Property ID") = oPropertyInfo.ID
                            dr("Property Name") = oPropertyInfo.Name
                            dr("Property Description") = oPropertyInfo.PropDesc
                            dr("Parent Property") = oPropertyInfo.Parent_ID
                            dr("Property Position") = oPropertyInfo.PropPos
                            dr("Property Active") = oPropertyInfo.PropIsActive
                            dr("Created By") = oPropertyInfo.CreatedBy
                            dr("Created On") = oPropertyInfo.CreatedOn
                            dr("Modified By") = oPropertyInfo.ModifiedBy
                            dr("Modified On") = oPropertyInfo.ModifiedOn
                            tblPropTypes.Rows.Add(dr)
                        Next
                    Catch ex As Exception
                        '
                        ' There's nothing to do - if an exception was generated, then
                        '  simply return an empty datatable
                        '
                        Throw ex
                    End Try
                End If
                'End If
                Return tblPropTypes
            Catch ex As Exception
                Throw ex
            End Try

        End Function
#End Region
#Region "Overloaded Operators"
        Protected Overloads Sub OnInsert(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.PropertyInfo)) Then
                Throw New ArgumentException("Only Property may be validated in an Property collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.PropertyInfo)) Then
                Throw New ArgumentException("Only Property may be validated in an Property collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.PropertyInfo)) Then
                Throw New ArgumentException("Only Property may be validated in an Property collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.PropertyInfo)) Then
                Throw New ArgumentException("Only Property may be validated in an Property collection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace



