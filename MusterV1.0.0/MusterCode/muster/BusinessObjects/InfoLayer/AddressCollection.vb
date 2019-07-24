'-------------------------------------------------------------------------------
' MUSTER.Info.AddressCollection
'   Provides a stongly-typed collection for storing Address objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        KJ      12/06/04      Original class definition.
'  1.1        MNR     01/13/05      Added Events
'  1.2        MNR     01/25/05      Added Array.Sort(KeyCol) in GetKeys() Property
'  1.3        MNR     01/27/05      Added ToString in Subs calling Value.ID
'  1.4        MNR     01/28/05      Added ChangeKey operation for altering keys
'
' Function              Description
' Item(AddressId)       Gets/Sets the AddressInfo requested by the arg AddressID
' Values()              Returns the collection of Addresses in the AddressCollection
' GetKeys()             Returns an array of string containing the keys in the AddressCollection
' Add(AddressInfo)      Adds the Address supplied to the internal AddressCollection
' Contains(AddressInfo) Returns True/False to indicate if the supplied AddressInfo is contained
'                           in the internal AddressCollection
' Contains(Key)         Returns True/False to indicate if the supplied AddressInfo key is contained
'                           in the internal AddressCollection
' Remove(AddressInfo)   Removes the AddressInfo supplied from the internal AddressCollection
'-------------------------------------------------------------------------------
Namespace MUSTER.Info
    Public Class AddressCollection
        Inherits DictionaryBase
#Region "Public Events"
        Public Event AddressColChanged(ByVal strSrc As String)
#End Region
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As Muster.Info.AddressInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), Muster.Info.AddressInfo)
            End Get
            Set(ByVal Value As Muster.Info.AddressInfo)
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(Value.AddressId.ToString, Value)
                    RaiseEvent AddressColChanged(Me.ToString)
                Else
                    MyBase.Dictionary.Item(index) = Value
                    RaiseEvent AddressColChanged(Me.ToString)
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
        Public Sub Add(ByVal value As Muster.Info.AddressInfo)
            Me.Item(value.AddressId.ToString) = value
        End Sub
        Public Function Contains(ByVal value As Muster.Info.AddressInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.AddressId.ToString)
        End Function
        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function
        Public Sub Remove(ByVal value As Muster.Info.AddressInfo)
            MyBase.Dictionary.Remove(value.AddressId.ToString)
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
            If Not value.GetType() Is (GetType(Muster.Info.AddressInfo)) Then
                Throw New ArgumentException("Only Reports may be inserted into a report collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(Muster.Info.AddressInfo)) Then
                Throw New ArgumentException("Only Reports may be removed from a report collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(Muster.Info.AddressInfo)) Then
                Throw New ArgumentException("Only Reports updated in a report collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(Muster.Info.AddressInfo)) Then
                Throw New ArgumentException("Only Reports may be validated in a report collection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace

