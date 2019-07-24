'-------------------------------------------------------------------------------
' MUSTER.Info.ReportsCollection
'   Provides a stongly-typed collection for storing FavSearchChildInfo objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      Mohan       12/4/04     Original class definition.
'  1.1      Mohan        1/7/05     Added Events for data update notification.
'                                   Added firing of event in ITEM()
'  1.2      JVC2        01/21/05    Added ChangeKey operation for altering keys.
'  1.3      MNR         01/25/05    Added Array.Sort(KeyCol) in GetKeys() Property
'
' Operations
' Function          Description
' Item(NAME)        Gets/Sets the FavSearchChildInfo requested by the string arg NAME
' Values()          Returns the collection of FavSearchChildInfo objects in the FavSearchChildCollection
' GetKeys()         Returns an array of strings containing the keys in the FavSearchChildCollection
' Add(FavSearchChildInfo)  Adds the FavSearchChildInfo supplied to the internal FavSearchChildCollection.
'                       Note - this is the same as Item(Name) with the exception that it
'                              provides the key for the object to Item(Name).
' Contains(FavSearchChildInfo)
'                   Returns True/False to indicate if the supplied FavSearchChildInfo object is contained
'                           in the internal FavSearchChildCollection.
' Contains(Key)     Returns True/False to indicate if the supplied FavSearchChildInfo key is contained
'                           in the internal FavSearchChildCollection.
' Remove(FavSearchChildInfo) 
'                   Removes the FavSearchChildInfo supplied from the internal FavSearchChildCollection.
' Remove(Key)       Removes the FavSearchChildInfo with the supplied Key string from the 
'                           internal FavSearchChildCollection.
'-------------------------------------------------------------------------------
Namespace MUSTER.Info
    <Serializable()> _
Public Class FavSearchChildCollection
        Inherits DictionaryBase
#Region "Public Events"
        Public Event CriteriaColChanged()
#End Region
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.FavSearchChildInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.FavSearchChildInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.FavSearchChildInfo)
                Dim strIndex As String = Value.ID.ToString
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(CType(Value.ID, String), Value)
                    RaiseEvent CriteriaColChanged()
                Else
                    MyBase.Dictionary.Item(index) = Value
                    RaiseEvent CriteriaColChanged()
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
        Public Sub Add(ByVal value As MUSTER.Info.FavSearchChildInfo)
            Me.Item(CType(value.ID, String)) = value
        End Sub
        Public Function Contains(ByVal value As MUSTER.Info.FavSearchChildInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.ID)
        End Function
        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function
        Public Sub Remove(ByVal value As MUSTER.Info.FavSearchChildInfo)
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
            If Not value.GetType() Is (GetType(MUSTER.Info.FavSearchChildInfo)) Then
                Throw New ArgumentException("Only FavSearchChildInfo may be inserted into a FavSearchChild collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.FavSearchChildInfo)) Then
                Throw New ArgumentException("Only FavSearchChildInfo may be removed from a FavSearchChild collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.FavSearchChildInfo)) Then
                Throw New ArgumentException("Only FavSearchChildInfo updated in a FavSearchChild collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.FavSearchChildInfo)) Then
                Throw New ArgumentException("Only FavSearchChildInfo may be validated in a FavSearchChild collection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace
