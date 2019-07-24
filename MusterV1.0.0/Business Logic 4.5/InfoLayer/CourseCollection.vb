'-------------------------------------------------------------------------------
' MUSTER.Info.CourseCollection
'   Provides a stongly-typed collection for storing Course objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MR      5/21/05    Original class definition.
'
' Function          Description
' Item(NAME)        Gets/Sets the CourseInfo requested by the string arg NAME
' Values()          Returns the collection of Courses in the CourseCollection
' GetKeys()         Returns an array of string containing the keys in the CourseCollection
' Add(CourseInfo)   Adds the Course supplied to the internal CourseCollection
' Contains(CourseInfo) Returns True/False to indicate if the supplied EntityInfo is contained
'                           in the internal CourseCollection
' Contains(Key)     Returns True/False to indicate if the supplied CourseInfo key is contained
'                           in the internal CourseCollection
' Remove(CourseInfo) Removes the CourseInfo supplied from the internal CourseCollection
'------------------------------------------------------------------------------------
Namespace MUSTER.Info
    Public Class CourseCollection
        Inherits DictionaryBase
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.CourseInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.CourseInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.CourseInfo)
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
        Public Sub Add(ByVal value As MUSTER.Info.CourseInfo)
            Me.Item(value.ID.ToString) = value
        End Sub
        Public Function Contains(ByVal value As MUSTER.Info.CourseInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.ID.ToString)
        End Function
        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function
        Public Sub Remove(ByVal value As MUSTER.Info.CourseInfo)
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
            If Not value.GetType() Is (GetType(MUSTER.Info.CourseInfo)) Then
                Throw New ArgumentException("Only Courses may be inserted into a Course collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.CourseInfo)) Then
                Throw New ArgumentException("Only Courses may be removed from a Course collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.CourseInfo)) Then
                Throw New ArgumentException("Only Courses updated in a Course collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.CourseInfo)) Then
                Throw New ArgumentException("Only Courses may be validated in a Course collection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace