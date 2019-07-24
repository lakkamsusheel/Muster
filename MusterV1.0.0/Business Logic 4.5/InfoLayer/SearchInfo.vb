'-------------------------------------------------------------------------------
' MUSTER.Info.SearchInfo
'   Provides the container to persist MUSTER Search state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MNR       12/08/04    Original class definition.
'
' Function          Description
' New()             Instantiates an empty SearchInfo object
' New(Keyword, Module, Filter)
'                   Instantiates a populated SearchInfo object
' Init()            Initializes a SearchInfo object's member variables to String.Empty
'
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class SearchInfo
#Region "Private member variables"
        Private strKeyword As String
        Private strModule As String
        Private strFilter As String
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Me.Init()
        End Sub
        Sub New(ByVal Keyword As String, _
                ByVal [Module] As String, _
                ByVal Filter As String)
            strKeyword = Keyword
            strModule = [Module]
            strFilter = Filter
        End Sub
#End Region
#Region "Private Operations"
        Private Sub Init()
            strKeyword = String.Empty
            strModule = String.Empty
            strFilter = String.Empty
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property Keyword() As String
            Get
                Return strKeyword
            End Get
            Set(ByVal Value As String)
                strKeyword = Value
            End Set
        End Property
        Public Property [Module]() As String
            Get
                Return strModule
            End Get
            Set(ByVal Value As String)
                strModule = Value
            End Set
        End Property
        Public Property Filter() As String
            Get
                Return strFilter
            End Get

            Set(ByVal Value As String)
                strFilter = Value
            End Set
        End Property
#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
    End Class
End Namespace
