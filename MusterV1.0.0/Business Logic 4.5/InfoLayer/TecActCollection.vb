
Namespace MUSTER.Info
    Public Class TecActCollection
        Inherits DictionaryBase
        Public Delegate Sub TecActColChangedEventHandler()
        ' Raised whenever a member is inserted into or removed from the TecAct collection
        Public Event TecActColChanged As TecActColChangedEventHandler
        Public Sub Add(ByVal value As MUSTER.Info.TecActInfo)
            Me.Item(value.ID.ToString) = value
        End Sub
        Public Sub ChangeKey(ByVal OldKey As String, ByVal NewKey As String)
            If Me.Contains(OldKey) Then
                Dim MyInfo As New Object
                MyInfo = MyBase.Dictionary.Item(OldKey)
                Me.Remove(OldKey)
                Me.Add(MyInfo)
            End If
        End Sub
        Public Function Contains(ByVal value As MUSTER.Info.TecActInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.ID.ToString)
        End Function
        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function
        Public ReadOnly Property GetKeys() As String()
            Get
                Dim KeyCol(MyBase.Dictionary.Keys.Count - 1) As String
                MyBase.Dictionary.Keys.CopyTo(KeyCol, 0)
                Array.Sort(KeyCol)
                Return KeyCol
            End Get
        End Property
        Default Public Property Item(ByVal index As String) As MUSTER.Info.TecActInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.TecActInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.TecActInfo)
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(Value.ID.ToString, Value)
                Else
                    MyBase.Dictionary.Item(index) = Value
                    RaiseEvent TecActColChanged()
                End If
            End Set
        End Property
        Public Sub Remove(ByVal value As MUSTER.Info.TecActInfo)
            MyBase.Dictionary.Remove(value.ID.ToString)
            RaiseEvent TecActColChanged()
        End Sub
        Public Sub Remove(ByVal value As String)
            MyBase.Dictionary.Remove(value)
            RaiseEvent TecActColChanged()
        End Sub
        Public ReadOnly Property Values() As ICollection
            Get
                Return MyBase.Dictionary.Values
            End Get
        End Property
        Protected Overloads Sub OnInsert(ByVal Key As Object, ByVal value As Object)
        End Sub
        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
        End Sub
        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
        End Sub
        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
        End Sub
    End Class
End Namespace
