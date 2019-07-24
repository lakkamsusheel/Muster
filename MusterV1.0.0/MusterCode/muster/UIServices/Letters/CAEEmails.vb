Public Class CAEEmails
    Inherits System.Windows.Forms.Form
    Dim emailsDs As DataSet
    Dim docPath As String
    Dim letterName As String
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub
    Public Sub New(ByRef dsEmails As DataSet, ByVal strUnPrintedPath As String, ByVal letter As String)
        MyBase.New()
        InitializeComponent()
        emailsDs = dsEmails
        docPath = strUnPrintedPath
        letterName = letter
    End Sub

    'Form overrides dispose to clean up the component list.
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
        '
        'CAEEmails
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 273)
        Me.Name = "CAEEmails"
        Me.Text = "C&E Emails"

    End Sub

#End Region
    Private Sub CAEEmails_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim emails As String
        For Each row As DataRow In emailsDs.Tables(0).Rows
            If Not System.IO.File.Exists(docPath + letterName) Then
                letterName = letterName.ToUpper.Replace(".DOC", "_TEMPLATE.DOC")
            End If
            emails = emails + row("contact_name") + ": " + row("cEmail") + ", " + row("oEmail") + ", " + _
                     docPath + letterName

        Next
    End Sub
End Class
