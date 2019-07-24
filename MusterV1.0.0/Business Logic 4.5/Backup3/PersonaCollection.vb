'-------------------------------------------------------------------------------
' MUSTER.Info.PersonaCollection
'   Provides a stongly-typed collection for storing Entity objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0         EN     12/13/04      Original class definition.
'  1.1        MNR     01/13/05      Added Events
'  1.2        MNR     01/25/05      Added Array.Sort(KeyCol) in GetKeys() Property
'  1.3        MNR     01/26/05      Modified Array size in GetKeys() Property
'  1.4        MNR     01/28/05      Added ChangeKey operation for altering keys
'
' Function          Description
' Item(NAME)        Gets/Sets the PersonaInfo  requested by the string arg NAME
' Values()          Returns the collection of Person info in the PersonaCollection
' GetKeys()         Returns an array of string containing the keys in the PersonaCollection
' Add(FacilityInfo)   Adds the Entity supplied to the internal PersonaCollection
' Contains(Facilityinfo) Returns True/False to indicate if the supplied PersonaInfo is contained
'                           in the internal PersonaCollection
' Contains(Key)     Returns True/False to indicate if the supplied PersonaInfo key is contained
'                           in the internal PersonaCollection
' Remove(Facilityinfo) Removes the PersonaInfo supplied from the internal PersonaCollection
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    Public Class PersonaCollection
        '
        Inherits DictionaryBase
#Region "Public Events"
        Public Event PersonaColChanged(ByVal strSrc As String)
#End Region
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.PersonaInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.PersonaInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.PersonaInfo)
                If MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Item(index) = Value
                    RaiseEvent PersonaColChanged(Me.ToString)
                Else
                    MyBase.Dictionary.Add(Value.ID.ToString, Value)
                    RaiseEvent PersonaColChanged(Me.ToString)
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
        Public Sub Add(ByVal value As MUSTER.Info.PersonaInfo)
            Me.Item(value.ID.ToString) = value
        End Sub
        Public Function Contains(ByVal value As MUSTER.Info.PersonaInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.ID)
        End Function
        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function
        Public Sub Remove(ByVal value As MUSTER.Info.PersonaInfo)
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
            If Not value.GetType() Is (GetType(MUSTER.Info.PersonaInfo)) Then
                Throw New ArgumentException("Only Persona info may be inserted into a Persona collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.PersonaInfo)) Then

                Throw New ArgumentException("Only User Groups may be removed from a user group collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.PersonaInfo)) Then
                Throw New ArgumentException("Only User Groups updated in a user group collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.PersonaInfo)) Then
                Throw New ArgumentException("Only User Groups may be validated in a user group collection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace



