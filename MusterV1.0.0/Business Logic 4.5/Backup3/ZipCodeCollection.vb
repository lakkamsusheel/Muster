'-------------------------------------------------------------------------------
' MUSTER.Info.ZipCodeCollection
'   Provides a stongly-typed collection for storing Entity objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        EN      12/16/04      Original class definition.
'  1.1        MNR     01/25/05      Added Array.Sort(KeyCol) in GetKeys() Property
'  1.2        MNR     01/26/05      Modified Array size in GetKeys() Property
'  1.3        MNR     01/27/05      Added ToString in Subs calling Value.ID
'
' Function          Description
' Item(NAME)        Gets/Sets the ZipCodeInfo requested by the string arg NAME
' Values()          Returns the collection of Letter in the ZipCodeCollection
' GetKeys()         Returns an array of string containing the keys in the ZipCodeCollection
' Add(ZipCodeInfo)   Adds the Letter supplied to the internal ZipCodeCollection
' Contains(ZipCodeInfo) Returns True/False to indicate if the supplied ZipCodeInfo is contained
'                           in the internal ZipCodeCollection
' Contains(Key)     Returns True/False to indicate if the supplied ZipCodeInfo key is contained
'                           in the internal ZipCodeCollection
' Remove(ZipCodeInfo) Removes the ZipCodeInfo supplied from the internal ZipCodeCollectionf
'-------------------------------------------------------------------------------
Namespace MUSTER.Info
    Public Class ZipCodeCollection
        Inherits DictionaryBase
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.ZipCodeInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.ZipCodeInfo)
            End Get

            Set(ByVal Value As MUSTER.Info.ZipCodeInfo)
                If MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Item(index) = Value
                Else
                    MyBase.Dictionary.Add(CType(Value.ID, String), Value)
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

        Public Sub Add(ByVal value As MUSTER.Info.ZipCodeInfo)
            Me.Item(CType(value.ID, String)) = value
        End Sub

        Public Function Contains(ByVal value As MUSTER.Info.ZipCodeInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.ID.ToString)
        End Function

        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function

        Public Sub Remove(ByVal value As MUSTER.Info.ZipCodeInfo)
            MyBase.Dictionary.Remove(CType(value.ID, String))
        End Sub

        Public Sub Remove(ByVal value As String)
            MyBase.Dictionary.Remove(value)
        End Sub

#End Region
#Region "Overloaded Operators"
        Protected Overloads Sub OnInsert(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.ZipCodeInfo)) Then
                Throw New ArgumentException("Only Zip Information  may be inserted into a ZipCodeCollection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.ZipCodeInfo)) Then

                Throw New ArgumentException("Only Zip information  may be removed from a ZipCodeCollection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.ZipCodeInfo)) Then
                Throw New ArgumentException("Only Zip Information updated in a ZipCodeCollection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.ZipCodeInfo)) Then
                Throw New ArgumentException("Only Zip may be validated in a user ZipCodeCollection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace



