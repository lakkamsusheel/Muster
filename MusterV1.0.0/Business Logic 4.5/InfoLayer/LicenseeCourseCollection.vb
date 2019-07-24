'-------------------------------------------------------------------------------
' MUSTER.Info.LicenseeCourseCollection
'   Provides a stongly-typed collection for storing Entity objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0                              Original class definition
'  1.1        MR        06/04/2005  Added ChangeKey Function
'
' Function          Description
' Item(ID)          Gets/Sets the EntityInfo requested by the string arg ID
' Values()          Returns the collection of Entities in the EntityCollection
' GetKeys()         Returns an array of string containing the keys in the EntityCollection
' Add(Entity)       Adds the Entity supplied to the Entity Collection
' Contains(Entity)  Returns True/False to indicate if the supplied Entity is contained
'                   in the Entity Collection
' Contains(Key)     Returns True/False to indicate if the supplied Entity key is contained
'                   in the Entity Collection
' Remove(Entity)    Removes the Entity supplied from the Entity Collection
'
' NOTE: This file to be used as LicenseeCourse to build other objects.
'       Replace keyword "LicenseeCourse" with respective Object name.
'       Also - don't forget to update the revision history above!!!
'-------------------------------------------------------------------------------
Namespace MUSTER.Info
    Public Class LicenseeCourseCollection
        Inherits DictionaryBase
#Region "Public Events"
        Public Event LicenseeCourseColChanged(ByVal bolValue As Boolean)
#End Region
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.LicenseeCourseInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.LicenseeCourseInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.LicenseeCourseInfo)
                If Not MyBase.Dictionary.Contains(Integer.Parse(index)) Then
                    MyBase.Dictionary.Add(Value.ID.ToString, Value)
                    RaiseEvent LicenseeCourseColChanged(True)
                Else
                    MyBase.Dictionary.Item(Integer.Parse(index)) = Value
                    RaiseEvent LicenseeCourseColChanged(True)
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
        Public Sub Add(ByVal value As MUSTER.Info.LicenseeCourseInfo)
            Me.Item(value.ID.ToString) = value
        End Sub
        Public Function Contains(ByVal value As MUSTER.Info.LicenseeCourseInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.ID.ToString)
        End Function
        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function
        Public Sub Remove(ByVal value As MUSTER.Info.LicenseeCourseInfo)
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
            If Not value.GetType() Is (GetType(MUSTER.Info.LicenseeCourseInfo)) Then
                Throw New ArgumentException("Only LicenseeCourse may be inserted into a LicenseeCourse collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.LicenseeCourseInfo)) Then
                Throw New ArgumentException("Only LicenseeCourse may be removed from a LicenseeCourse collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.LicenseeCourseInfo)) Then
                Throw New ArgumentException("Only LicenseeCourse updated in a LicenseeCourse collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.LicenseeCourseInfo)) Then
                Throw New ArgumentException("Only LicenseeCourse may be validated in a LicenseeCourse collection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace