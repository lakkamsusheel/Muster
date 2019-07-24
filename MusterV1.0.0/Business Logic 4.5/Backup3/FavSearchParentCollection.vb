'-------------------------------------------------------------------------------
' MUSTER.Info.ReportsCollection
'   Provides a stongly-typed collection for storing FavSearchParentCollection objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MR         12/5/04    Original class definition.
'  1.1        MR         1/7/05    Added Events for data update notification.
'                                  Added firing of event in ITEM()
'  1.2        JVC2      01/21/05   Added ChangeKey method.
'                                    Added cast of index to string in ITEM method
'  1.3        MNR       01/25/05    Added Array.Sort(KeyCol) in GetKeys() Property
'
' Operations
' Function          Description
' Item(NAME)            Gets/Sets the FavSearchParentCollection requested by the string arg NAME
' Values()              Returns the collection of FavSearchParentInfo objects in the FavSearchParentCollection
' GetKeys()             Returns an array of strings containing the keys in the FavSearchParentCollection
' Add(FavSearchParentInfo)  Adds the FavSearchParentInfo supplied to the internal FavSearchParentCollection.
'                           Note - this is the same as Item(Name) with the exception that it
'                              provides the key for the object to Item(Name).
' Contains(FavSearchParentInfo)
'                       Returns True/False to indicate if the supplied FavSearchParentInfo object is contained
'                           in the internal FavSearchParentCollection.
' Contains(Key)         Returns True/False to indicate if the supplied FavSearchParentInfo key is contained
'                           in the internal FavSearchParentCollection.
' Remove(FavSearchParentInfo) 
'                       Removes the FavSearchParentInfo supplied from the internal FavSearchParentCollection.
' Remove(Key)           Removes the FavSearchParentInfo with the supplied Key string from the 
'                           internal FavSearchParentCollection.
'-------------------------------------------------------------------------------
Namespace MUSTER.Info
    <Serializable()> _
Public Class FavSearchParentCollection
        Inherits DictionaryBase
#Region "Public Events"
        Public Event FavSearchColChanged()
#End Region
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.FavSearchParentInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.FavSearchParentInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.FavSearchParentInfo)
                Dim strIndex As String = Value.ID.ToString
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(CType(Value.ID, String), Value)
                    RaiseEvent FavSearchColChanged()
                Else
                    MyBase.Dictionary.Item(index) = Value
                    RaiseEvent FavSearchColChanged()
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
        Public Sub Add(ByVal value As MUSTER.Info.FavSearchParentInfo)
            Me.Item(CType(value.ID, String)) = value
        End Sub
        Public Function Contains(ByVal value As MUSTER.Info.FavSearchParentInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.ID)
        End Function
        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function
        Public Sub Remove(ByVal value As MUSTER.Info.FavSearchParentInfo)
            MyBase.Dictionary.Remove(CStr(value.ID))
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
            If Not value.GetType() Is (GetType(MUSTER.Info.FavSearchParentInfo)) Then
                Throw New ArgumentException("Only FavSearch Parent may be inserted into a FavSearchParent collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.FavSearchParentInfo)) Then
                Throw New ArgumentException("Only FavSearch Parent may be removed from a FavSearchParent collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.FavSearchParentInfo)) Then
                Throw New ArgumentException("Only FavSearch Parent updated in a FavSearchParent collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.FavSearchParentInfo)) Then
                Throw New ArgumentException("Only FavSearch Parent may be validated in a FavSearchParent collection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace
