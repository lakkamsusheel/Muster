Namespace Utils
    Public Class DBUtils
        Public Shared Function IsNull(ByVal ObjectToCheck As Object, ByVal AlternateObject As Object) As Object
            If ObjectToCheck Is Nothing Then
                IsNull = AlternateObject
            Else
                IsNull = ObjectToCheck
            End If
        End Function
        Public Shared Function AltIsDBNull(ByVal ObjectToCheck As Object, ByVal AlternateObject As Object) As Object

            If Not ObjectToCheck Is Nothing AndAlso ObjectToCheck Is System.DBNull.Value Then
                AltIsDBNull = AlternateObject
            Else
                AltIsDBNull = ObjectToCheck
            End If
        End Function
        Public Shared Function IIFIsIntegerNull(ByVal ObjectToCheck As Object, ByVal AlternateObject As Object) As Object
            If Not ObjectToCheck Is Nothing AndAlso ObjectToCheck = 0 Then
                IIFIsIntegerNull = AlternateObject
            Else
                IIFIsIntegerNull = ObjectToCheck
            End If
        End Function


        'P1 02/01/05
        Public Shared Function IIFIsDateNull(ByVal ObjectToCheck As Object, ByVal AlternateObject As Object) As Object
            Dim dtTempDate As Date
            If Not ObjectToCheck Is Nothing AndAlso Date.Compare(ObjectToCheck, dtTempDate) = 0 Then
                IIFIsDateNull = AlternateObject
            Else
                IIFIsDateNull = ObjectToCheck
            End If
        End Function
        'Public Shared Function getAddress(ByVal AddressId As Integer) As Address
        '    Return getAddress
        'End Function
    End Class
End Namespace

