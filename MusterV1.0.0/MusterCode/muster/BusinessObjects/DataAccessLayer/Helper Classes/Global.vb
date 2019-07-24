
'----------------------------------------------------------------

' Copyright (C) 2004 CIBER, Inc.

' All rights reserved.

'

' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY 

' OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT 

' LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR 

' FITNESS FOR A PARTICULAR PURPOSE.

'----------------------------------------------------------------

Imports System.Configuration

Class ConnectionSettings
    Private strConnStr As String
    Private LocalUserSettings As Microsoft.Win32.Registry

    Private _cnString As String = String.Empty

    Public Sub New()
        _cnString = FindConnStr()
    End Sub

    Public Property cnString()
        Get
            Return _cnString
        End Get
        Set(ByVal Value)


        End Set
    End Property

    Public Property ConnStr()
        Get
            Return _cnString
        End Get
        Set(ByVal Value)
            _cnString = Value
        End Set
    End Property

    Private Function FindConnStr() As String
        Return LocalUserSettings.CurrentUser.GetValue("MusterSQLConnection")
    End Function

End Class
