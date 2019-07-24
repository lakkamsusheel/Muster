'-------------------------------------------------------------------------------
' MUSTER.Info.ReportsCollection
'   Provides a stongly-typed collection for storing ProfileInfo objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        JVC2      11/19/04    Original class definition.
'  1.1        JC        12/28/04    Added event for data update notification.
'                                   Added firing of event in ITEM()
'  1.2        MNR       01/25/05    Added Array.Sort(KeyCol) in GetKeys() Property
'
' Operations
' Function          Description
' Item(NAME)        Gets/Sets the ProfileInfo requested by the string arg NAME
' Values()          Returns the collection of ProfileInfo objects in the ProfileCollection
' GetKeys()         Returns an array of strings containing the keys in the ProfileCollection
' Add(ProfileInfo)  Adds the ProfileInfo supplied to the internal ProfileCollection.
'                       Note - this is the same as Item(Name) with the exception that it
'                              provides the key for the object to Item(Name).
' Contains(ProfileInfo)
'                   Returns True/False to indicate if the supplied ProfileInfo object is contained
'                           in the internal ProfileCollection.
' Contains(Key)     Returns True/False to indicate if the supplied ProfileInfo key is contained
'                           in the internal ProfileCollection.
' Remove(ProfileInfo) 
'                   Removes the ProfileInfo supplied from the internal ProfileCollection.
' Remove(Key)       Removes the ProfileInfo with the supplied Key string from the 
'                           internal ProfileCollection.
'
' Events
' Name              Description
' InfoChanged       Alerts the client that and insert/update/delete has taken place
'
'-------------------------------------------------------------------------------
'
'TODO - Update in Application 12/28 JVC2
'
Namespace MUSTER.Info
    Public Class ProfileCollection
        '
        Inherits DictionaryBase
#Region "Public Events"
        Public Event InfoChanged()
#End Region
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.ProfileInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.ProfileInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.ProfileInfo)
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(Value.ID, Value)
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

        Public Sub Add(ByVal value As MUSTER.Info.ProfileInfo)
            Me.Item(value.ID) = value
        End Sub

        Public Function Contains(ByVal value As MUSTER.Info.ProfileInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.ID)
        End Function

        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function

        Public Sub Remove(ByVal value As MUSTER.Info.ProfileInfo)
            MyBase.Dictionary.Remove(value.ID)
        End Sub


        Public Sub Remove(ByVal value As String)
            MyBase.Dictionary.Remove(value)
        End Sub

#End Region
#Region "Overloaded Operators"
        Protected Overloads Sub OnInsert(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.ProfileInfo)) Then
                Throw New ArgumentException("Only Profiles may be inserted into a profile collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.ProfileInfo)) Then
                Throw New ArgumentException("Only Profiles may be removed from a profile collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.ProfileInfo)) Then
                Throw New ArgumentException("Only Profiles updated in a profile collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.ProfileInfo)) Then
                Throw New ArgumentException("Only Profiles may be validated in a profile collection!", "value")
            End If
        End Sub


#End Region

    End Class
End Namespace

