
Namespace MUSTER.Info
    ' -------------------------------------------------------------------------------
    '       MUSTER.Info.RegistrationCollection
    '                   Provides the collection for storing RegistrationInfo objects.
    ' 
    '       Copyright (C) 2004 CIBER, Inc.
    '       All rights reserved.
    ' 
    '       Release   Initials    Date        Description
    '             1.0        JVC2      02/08/2005  Original framework from Rational XDE.
    ' 
    ' -------------------------------------------------------------------------------
    '
    Public Class RegistrationActivityCollection
        Inherits DictionaryBase
#Region "Public Events"
        Public Delegate Sub RegistrationColChangedEventHandler(ByVal ColIsDirty As Boolean)
        ' Event which is raised when a member is added to or removed from the collection
        ' 
        Public Event RegistrationColChanged As RegistrationColChangedEventHandler
#End Region
#Region "Exposed Operations"
        ' Used to insert a new RegistrationInfo into the collection
        Public Sub Add(ByVal value As MUSTER.Info.RegistrationActivityInfo)
            Me.Item(value.RegActionIndex.ToString) = value
        End Sub
        ' Used to determine if a member of the collection matches the supplied RegistrationInfo object
        Public Function Contains(ByVal value As MUSTER.Info.RegistrationActivityInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.RegActionIndex.ToString)
        End Function
        ' Used to determine if a key of the collection matches the supplied string
        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function
        ' Used to supply an array of string containing the keys of the collection.
        Public ReadOnly Property GetKeys() As String()
            Get
                Dim KeyCol(MyBase.Dictionary.Keys.Count - 1) As String
                MyBase.Dictionary.Keys.CopyTo(KeyCol, 0)
                Array.Sort(KeyCol)
                Return KeyCol
            End Get
        End Property
        ' Used to insert a RegistrationInfo object into the collection or return a RegistrationInfo object from the collection
        Default Public Property Item(ByVal index As String) As MUSTER.Info.RegistrationActivityInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.RegistrationActivityInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.RegistrationActivityInfo)
                If Not MyBase.Dictionary.Contains(Integer.Parse(index)) Then
                    MyBase.Dictionary.Add(Value.RegActionIndex.ToString, Value)
                    RaiseEvent RegistrationColChanged(True)
                Else
                    MyBase.Dictionary.Item(Integer.Parse(index)) = Value
                    RaiseEvent RegistrationColChanged(True)
                End If
            End Set
        End Property
        ' Used to remove a member from the collection which matches the supplied RegistrationInfo object
        Public Sub Remove(ByVal value As MUSTER.Info.RegistrationActivityInfo)
            MyBase.Dictionary.Remove(value.RegActionIndex.ToString)
        End Sub
        ' Used to remove a member from the collection which has it's key matching the supplied key
        Public Sub Remove(ByVal value As String)
            MyBase.Dictionary.Remove(value)
        End Sub
        ' Used to supply an ICollection corresponding to the members of the collection for iteration
        Public ReadOnly Property Values() As ICollection
            Get
                Return MyBase.Dictionary.Values
            End Get
        End Property
        Public Sub ChangeKey(ByVal OldKey As String, ByVal NewKey As String)
            If Me.Contains(OldKey) Then
                Dim MyInfo As New Object
                MyInfo = MyBase.Dictionary.Item(OldKey)
                Me.Remove(OldKey)
                Me.Add(MyInfo)
            End If
        End Sub
#End Region
#Region "Protected Operations"
        ' Checks that the inserted member is a RegistrationInfo object
        Protected Overloads Sub OnInsert(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.RegistrationActivityInfo)) Then
                Throw New ArgumentException("Only Registrations may be inserted into a Registration collection!", "value")
            End If
        End Sub
        ' Checks that the removed member is a RegistrationInfo object
        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.RegistrationActivityInfo)) Then
                Throw New ArgumentException("Only Registrations may be removed from a Registration collection!", "value")
            End If
        End Sub
        ' Checks that the member being modified is a RegistrationInfo object
        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.RegistrationActivityInfo)) Then
                Throw New ArgumentException("Only Registrations updated in a Registration collection!", "value")
            End If
        End Sub
        ' Checks that the member being validated is a RegistrationInfo object
        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.RegistrationActivityInfo)) Then
                Throw New ArgumentException("Only Registrations may be validated in a Registration collection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace
