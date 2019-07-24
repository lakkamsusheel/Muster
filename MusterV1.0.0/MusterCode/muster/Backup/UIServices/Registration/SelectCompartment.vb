Imports System
Imports Microsoft.VisualBasic
Imports System.IO

Public Class SelectCompartment
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

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
    Friend WithEvents Lblheader As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Lblheader = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Lblheader
        '
        Me.Lblheader.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lblheader.Location = New System.Drawing.Point(8, 8)
        Me.Lblheader.Name = "Lblheader"
        Me.Lblheader.Size = New System.Drawing.Size(496, 23)
        Me.Lblheader.TabIndex = 0
        Me.Lblheader.Text = "Please Select which compartment you would like"
        '
        'SelectCompartment
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(528, 214)
        Me.ControlBox = False
        Me.Controls.Add(Me.Lblheader)
        Me.Name = "SelectCompartment"
        Me.Text = "Select Compartment"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "private members"

    Private _componentTable As Info.CompartmentCollection
    Private _ComponentIDSelected As Integer = 1
    Private _x, _y As Integer

#End Region

#Region "Shared Methods"

    Shared Function selectCompartment(ByVal text As String, ByVal compartments As Info.CompartmentCollection) As Integer

        Dim frm As New SelectCompartment
        Dim id As Integer = -1

        frm.displayText = text
        frm.ComponentTable = compartments
        frm.ShowDialog()

        If frm.DialogResult = frm.DialogResult.OK Then
            id = frm.ComponentIdSelected
        End If

        frm.Dispose(True)
        frm = Nothing

        Return id

    End Function

#End Region

#Region "Public Properties"

    Public Property displayText() As String
        Get
            Return Me.Lblheader.Text
        End Get

        Set(ByVal Value As String)
            Me.Lblheader.Text = Value
        End Set
    End Property

    Public Property ComponentTable() As Info.CompartmentCollection

        Get
            Return _componentTable
        End Get

        Set(ByVal Value As Info.CompartmentCollection)
            _componentTable = Value
        End Set

    End Property

    Public ReadOnly Property ComponentIdSelected() As Integer
        Get
            Return _ComponentIDSelected
        End Get
    End Property

#End Region

#Region "private methods"

    Sub Loaddata()

        Dim p As New BusinessLogic.pProperty

        _x = 30
        _y = 50


        For Each key As String In _componentTable.GetKeys

            Try

                Dim newBtn As New Button



                With newBtn

                    .Name = String.Format("BtnComponent_{0}", _componentTable.Item(key).COMPARTMENTNumber)
                    .Text = String.Format(" Compartment {0}   capacity:{1}    Substance:{2}", _componentTable.Item(key).COMPARTMENTNumber, _
                                                 _componentTable.Item(key).Capacity, p.GetPropertyNameByID(_componentTable.Item(key).Substance))
                    .Location = New Point(_x, _y)
                    .Size = New Size(250, 30)
                    .Tag = _componentTable.Item(key).COMPARTMENTNumber

                End With


                _y += 65
                If (_y + 35 > Me.Height) Then
                    _x += 270
                    _y = 50

                    If _x + 30 > Me.Width Then
                        Me.Width = _x + 30
                    End If


                End If

                AddHandler newBtn.Click, AddressOf ButtonClicked

                Controls.Add(newBtn)

            Catch ex As Exception
                Throw New Exception(ex.ToString)
            End Try

        Next

        p = Nothing


    End Sub

#End Region

#Region "Form Events"

    Private Sub ButtonClicked(ByVal sender As Object, ByVal e As EventArgs)

        _ComponentIDSelected = DirectCast(sender, Button).Tag
        DialogResult = DialogResult.OK
        Close()

    End Sub

    Private Sub SelectCompartment_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If Not _componentTable Is Nothing AndAlso _componentTable.Count > 0 Then

            Loaddata()
        Else
            MsgBox(" This Form has detected that no Component data has been loaded for choosing a compartment")
            Me.Close()

        End If

    End Sub

#End Region

End Class
