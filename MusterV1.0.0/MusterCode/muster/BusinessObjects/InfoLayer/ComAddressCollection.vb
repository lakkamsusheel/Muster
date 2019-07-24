'-------------------------------------------------------------------------------
' MUSTER.Info.ComAddressCollection
'   Provides a stongly-typed collection for storing Address objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MR      5/24/05      Original class definition.
'
' Function              Description
' Item(AddressId)       Gets/Sets the ComAddressInfo requested by the arg AddressID
' Values()              Returns the collection of Addresses in the ComAddressCollection
' GetKeys()             Returns an array of string containing the keys in the ComAddressCollection
' Add(ComAddressInfo)      Adds the Address supplied to the internal ComAddressCollection
' Contains(ComAddressInfo) Returns True/False to indicate if the supplied ComAddressInfo is contained
'                           in the internal ComAddressCollection
' Contains(Key)         Returns True/False to indicate if the supplied ComAddressInfo key is contained
'                           in the internal AddressCollection
' Remove(ComAddressInfo)   Removes the ComAddressInfo supplied from the internal ComAddressCollection
'-------------------------------------------------------------------------------
Namespace MUSTER.Info
    Public Class ComAddressCollection
        Inherits DictionaryBase
#Region "Public Events"
        Public Event AddressColChanged(ByVal strSrc As String)
#End Region
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.ComAddressInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.ComAddressInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.ComAddressInfo)
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
        Public Sub Add(ByVal value As MUSTER.Info.ComAddressInfo)
            Me.Item(value.AddressId.ToString) = value
        End Sub
        Public Function Contains(ByVal value As MUSTER.Info.ComAddressInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.AddressId.ToString)
        End Function
        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function
        Public Sub Remove(ByVal value As MUSTER.Info.ComAddressInfo)
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
            If Not value.GetType() Is (GetType(MUSTER.Info.ComAddressInfo)) Then
                Throw New ArgumentException("Only Addresses may be inserted into a Address collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.ComAddressInfo)) Then
                Throw New ArgumentException("Only Addresses may be removed from a Address collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.ComAddressInfo)) Then
                Throw New ArgumentException("Only Addresses updated in a Address collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.ComAddressInfo)) Then
                Throw New ArgumentException("Only Addresses may be validated in a Address collection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace

