
' The interface for returning Created and Modified info on an object.
Public Interface iAccessors
    ReadOnly Property CreatedBy() As String
    ReadOnly Property CreatedOn() As Date
    ReadOnly Property ModifiedBy() As String
    ReadOnly Property ModifiedOn() As Date
End Interface

