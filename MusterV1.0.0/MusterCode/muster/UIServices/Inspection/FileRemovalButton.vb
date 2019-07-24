Public Class FileRemovalButton
    Inherits Button

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal fileName As String)

        Call Me.New()
        Me.filePath = fileName

    End Sub


    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Text = "Remove"

    End Sub

    'UserControl overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

#Region "Public members"
    Public Event Fileremoved(ByVal sender As Object, ByVal e As EventArgs)
    Public filePath As String = String.Empty
#End Region


#Region "Methods"
    Private Sub FileRemovalButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Click

        If filePath.Length > 0 AndAlso MsgBox("Are you sure you to remove this file?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Try
                If IO.File.Exists(filePath) Then
                    IO.File.Delete(filePath)
                    MsgBox("File successfully removed")
                    RaiseEvent Fileremoved(Me, New EventArgs)
                Else
                    MsgBox(String.Format("{0} does not exist", filePath))
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If

    End Sub
#End Region

End Class
