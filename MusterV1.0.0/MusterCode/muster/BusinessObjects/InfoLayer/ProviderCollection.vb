'-------------------------------------------------------------------------------
' MUSTER.Info.ProviderCollection
'   Provides a stongly-typed collection for storing Provider objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MR      5/21/05    Original class definition.
'
' Function          Description
' Item(NAME)        Gets/Sets the ProviderInfo requested by the string arg NAME
' Values()          Returns the collection of Providers in the ProviderCollection
' GetKeys()         Returns an array of string containing the keys in the ProviderCollection
' Add(ProviderInfo)   Adds the Provider supplied to the internal ProviderCollection
' Contains(ProviderInfo) Returns True/False to indicate if the supplied ProviderInfo is contained
'                           in the internal ProviderCollection
' Contains(Key)     Returns True/False to indicate if the supplied ProviderInfo key is contained
'                           in the internal ProviderCollection
' Remove(ProviderInfo) Removes the ProviderInfo supplied from the internal ProviderCollection
'------------------------------------------------------------------------------------
Namespace MUSTER.Info
    Public Class ProviderCollection
        Inherits DictionaryBase
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.ProviderInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.ProviderInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.ProviderInfo)
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(Value.ID.ToString, Value)
                Else
                    MyBase.Dictionary.Item(index) = Value
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
        Public Sub Add(ByVal value As MUSTER.Info.ProviderInfo)
            Me.Item(value.ID.ToString) = value
        End Sub
        Public Function Contains(ByVal value As MUSTER.Info.ProviderInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.ID.ToString)
        End Function
        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function
        Public Sub Remove(ByVal value As MUSTER.Info.ProviderInfo)
            MyBase.Dictionary.Remove(value.ID.ToString)
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
            If Not value.GetType() Is (GetType(MUSTER.Info.ProviderInfo)) Then
                Throw New ArgumentException("Only Providers may be inserted into a Provider collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.ProviderInfo)) Then
                Throw New ArgumentException("Only Providers may be removed from a Provider collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.ProviderInfo)) Then
                Throw New ArgumentException("Only Providers updated in a Provider collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.ProviderInfo)) Then
                Throw New ArgumentException("Only Providers may be validated in a Provider collection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace