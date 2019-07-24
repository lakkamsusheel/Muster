'-------------------------------------------------------------------------------
' MUSTER.Info.TankCollection
'   Provides a stongly-typed collection for storing Tank objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        KJ      12/15/04      Original class definition.
'  1.1        KJ      12/29/04      Added event for data update notification.
'                                   Added firing of event in ITEM()
'  1.2        EN      01/21/05      Added srcSource in Event. 
'  1.3        MNR     01/25/05      Added Array.Sort(KeyCol) in GetKeys() Property
'  1.4        MNR     01/27/05      Added ToString in Subs calling Value.ID
'  1.5        MNR     01/28/05      Added ChangeKey operation for altering keys
'
' Function              Description
' Item(TankId)          Gets/Sets the TankInfo requested by the arg TankID
' Values()              Returns the collection of Tanks in the TankCollection
' GetKeys()             Returns an array of string containing the keys in the TankCollection
' Add(TankInfo)         Adds the Tank supplied to the internal TankCollection
' Contains(TankInfo)    Returns True/False to indicate if the supplied TankInfo is contained
'                           in the internal TankCollection
' Contains(Key)         Returns True/False to indicate if the supplied TankInfo key is contained
'                           in the internal TankCollection
' Remove(TankInfo)      Removes the TankInfo supplied from the internal TankCollection
'-------------------------------------------------------------------------------
'
' TODO - Add to app 1/3/05 - JVC2
' TODO - check properties and operations against list.
'

Namespace MUSTER.Info
    Public Class TankCollection
        Inherits DictionaryBase
        '#Region "Public Events"
        '        Public Event TankColChanged()
        '#End Region
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.TankInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.TankInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.TankInfo)
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(CType(Value.TankId, String), Value)
                    'RaiseEvent TankColChanged()
                Else
                    MyBase.Dictionary.Item(index) = Value
                    'If Value.IsDirty Then
                    'RaiseEvent TankColChanged()
                    'End If
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
        Public Sub Add(ByVal value As MUSTER.Info.TankInfo)
            Me.Item(CType(value.TankId, String)) = value
        End Sub
        Public Function Contains(ByVal value As MUSTER.Info.TankInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.TankId.ToString)
        End Function
        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function
        Public Sub Remove(ByVal value As MUSTER.Info.TankInfo)
            MyBase.Dictionary.Remove(value.TankId.ToString)
        End Sub
        Public Sub Remove(ByVal value As String)
            MyBase.Dictionary.Remove(value)
        End Sub
        Public Sub ChangeKey(ByVal OldKey As String, ByVal NewKey As String)
            If Me.Contains(OldKey) Then
                Dim MyInfo As New Object
                MyInfo = MyBase.Dictionary.Item(OldKey)
                Me.Remove(OldKey)
                Me.Add(MyInfo)
            End If
        End Sub
#End Region
#Region "Overloaded Operators"
        Protected Overloads Sub OnInsert(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.TankInfo)) Then
                Throw New ArgumentException("Only Tanks may be inserted into a Tank Collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.TankInfo)) Then
                Throw New ArgumentException("Only Tanks may be removed from a Tank collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.TankInfo)) Then
                Throw New ArgumentException("Only Tanks updated in a Tank collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.TankInfo)) Then
                Throw New ArgumentException("Only Tanks may be validated in a Tank collection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace

