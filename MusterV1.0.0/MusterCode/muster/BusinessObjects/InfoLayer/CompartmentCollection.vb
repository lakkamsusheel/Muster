'-------------------------------------------------------------------------------
' MUSTER.Info.CompartmentCollection
'   Provides a stongly-typed collection for storing Entity objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        EN        12/16/04    Original class definition.
'  1.1        MNR       01/25/05    Added Array.Sort(KeyCol) in GetKeys() Property
'  1.2        MNR       01/26/05    Modified Array size in GetKeys() Property
'  1.3        MNR       01/28/05    Added ChangeKey operation for altering keys
'
' Function          Description
' Item(NAME)        Gets/Sets the CompartmentInfo requested by the string arg NAME
' Values()          Returns the collection of Facilities in the CompartmentCollection
' GetKeys()         Returns an array of string containing the keys in the CompartmentCollection
' Add(CompartmentInfo)   Adds the Entity supplied to the internal CompartmentCollection
' Contains(CompartmentInfo) Returns True/False to indicate if the supplied CompartmentInfo is contained
'                           in the internal CompartmentCollection
' Contains(Key)     Returns True/False to indicate if the supplied CompartmentInfo key is contained
'                           in the internal CompartmentCollection
' Remove(CompartmentInfo) Removes the CompartmentInfo supplied from the internal CompartmentCollection
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    Public Class CompartmentCollection
        '
        Inherits DictionaryBase
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.CompartmentInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.CompartmentInfo)
            End Get

            Set(ByVal Value As MUSTER.Info.CompartmentInfo)
                If MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Item(index) = Value
                Else
                    MyBase.Dictionary.Add(CType(Value.ID, String), Value)
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
        Public Sub Add(ByVal value As MUSTER.Info.CompartmentInfo)
            Me.Item(CType(value.ID, String)) = value
        End Sub
        Public Function Contains(ByVal value As MUSTER.Info.CompartmentInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.ID)
        End Function
        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function
        Public Sub Remove(ByVal value As MUSTER.Info.CompartmentInfo)
            MyBase.Dictionary.Remove(value.ID)
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
            If Not value.GetType() Is (GetType(MUSTER.Info.CompartmentInfo)) Then
                Throw New ArgumentException("Only Compartment may be inserted into a Compartment collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.CompartmentInfo)) Then

                Throw New ArgumentException("Only Compartment may be removed from a user Compartment collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.CompartmentInfo)) Then
                Throw New ArgumentException("Only Compartment updated in a Compartment collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.CompartmentInfo)) Then
                Throw New ArgumentException("Only Compartment may be validated in a Compartment collection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace



