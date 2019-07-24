'Option Strict On
'Option Explicit On 

Public Class CostFormat
    Inherits System.Windows.Forms.UserControl


    Public CostFormatType As String
    Private editingControl As TextBox
    Public GrandTotal As Double
    Private bolhandlingKeys As Boolean = False
    Private useNewCode As Boolean = True

    Private RecalcInProcess As Boolean
    Private taggingFormatSpecs As Boolean = False
    Private bolloading As Boolean
    ' Private bolFormatting As Boolean
    Private bolFixingDuringTextChange As Boolean = False

    Private fakeLabel As Label

    Private nDecimalPlaces As Int16
    Private dControlDict As Collections.Specialized.ListDictionary
    Private lControlList As Collections.ArrayList

    Private oLocalFinancialCommitment As MUSTER.BusinessLogic.pFinancialCommitment




#Region "CostformatSpec"
    Public Class CostFormatSpec

        Public DecimalPlace As Integer = 0
        Public PartOfTotal As Boolean = False
        Public Shadows visible As Boolean = False
        Public isReadOnly As Boolean = False
        Public subTotal As Boolean = False
        Public Total As Boolean = False
        Private parentcon As Control
        Private dataSource As Object

        Public cntrlBindingControl As Control = Nothing
        Public fieldText As String = String.Empty
        Public staticText As String = String.Empty


        Sub New(ByRef con As Control, ByVal decPoint As Integer, ByVal ofTotal As Boolean, ByVal isVisible As Boolean, ByVal ReadOnlyField As Boolean, ByVal text As String, ByVal isSubTotal As Boolean, ByVal isTotal As Boolean, ByVal textField As String, Optional ByVal datasourceFrom As Object = Nothing, Optional ByVal BindingControl As TextBox = Nothing)

            DecimalPlace = decPoint
            PartOfTotal = ofTotal
            visible = isVisible
            isReadOnly = ReadOnlyField
            staticText = text.Trim
            fieldText = textField.Trim
            Total = isTotal
            subTotal = isSubTotal
            dataSource = datasourceFrom


            parentcon = con


            parentcon.DataBindings.Clear()

            Try

                If Not BindingControl Is Nothing AndAlso isReadOnly AndAlso text = String.Empty And TypeOf parentcon Is TextBox Then
                    cntrlBindingControl = BindingControl
                    AddHandler cntrlBindingControl.TextChanged, AddressOf AssignChange
                End If

                If Not dataSource Is Nothing AndAlso fieldText <> String.Empty Then
                    parentcon.DataBindings.Add("text", dataSource, fieldText)
                End If
            Catch ex As Exception
                MsgBox(String.Format("Binding error: {0}.{1}This control will bot bind to the data object.", ex.Message, vbCrLf), MsgBoxStyle.OKOnly, "Bindig issues")
            End Try


        End Sub

        Sub dispose()
            Try
                RemoveHandler cntrlBindingControl.TextChanged, AddressOf AssignChange
            Catch
            End Try
            parentcon = Nothing
            cntrlBindingControl = Nothing

        End Sub

        Public Function SetupControl(ByRef con As Control) As Control

            Try

                If TypeOf con Is TextBox Then
                    With DirectCast(con, TextBox)



                        If IsNumeric(staticText) Then
                            .Text = FormatNumber(staticText, DecimalPlace, TriState.True, TriState.False, TriState.True)
                        Else

                            .Text = IIf(staticText <> String.Empty, staticText, IIf(IsNumeric(.Text), FormatNumber(.Text, DecimalPlace, TriState.True, TriState.False, TriState.True), "0"))

                        End If
                        .ReadOnly = isReadOnly
                        .Visible = visible


                        If isReadOnly Then
                            .TabStop = False
                            .BorderStyle = BorderStyle.FixedSingle

                        End If


                        If subTotal Then
                            con.Font = New Font("Ariel", 8, FontStyle.Bold, GraphicsUnit.Point)

                        ElseIf Total Then
                            con.Font = New Font("Ariel", 9, FontStyle.Bold, GraphicsUnit.Point)
                        Else
                            con.Font = New Font("Ariel", 8, FontStyle.Regular, GraphicsUnit.Point)
                        End If

                    End With

                ElseIf TypeOf con Is Label Then
                    With DirectCast(con, Label)

                        If subTotal Then
                            con.Font = New Font("Ariel", 8, FontStyle.Bold, GraphicsUnit.Point)
                        ElseIf Total Then
                            con.Font = New Font("Ariel", 9, FontStyle.Bold, GraphicsUnit.Point)
                        Else
                            con.Font = New Font("Ariel", 8, FontStyle.Regular, GraphicsUnit.Point)
                        End If

                        .Text = IIf(staticText <> String.Empty, staticText, .Text)
                        .TabStop = False
                        .Visible = visible


                    End With

                End If

                Return con
            Catch ex As Exception
                Throw ex

            End Try


        End Function

        Sub AssignChange(ByVal sender As Object, ByVal e As EventArgs)

            If Not dataSource Is Nothing AndAlso fieldText <> String.Empty Then
                parentcon.DataBindings.Clear()

                If IsNumeric(DirectCast(sender, Control).Text) Then
                    parentcon.Text = FormatNumber(DirectCast(sender, Control).Text, DecimalPlace, TriState.True, TriState.False, TriState.True)
                Else
                    parentcon.Text = DirectCast(sender, Control).Text

                End If


                Dim inf As Reflection.PropertyInfo = dataSource.GetType.GetProperty(fieldText, Reflection.BindingFlags.IgnoreCase Or Reflection.BindingFlags.GetProperty Or Reflection.BindingFlags.Public Or Reflection.BindingFlags.Instance)

                If TypeOf inf.GetValue(dataSource, Nothing) Is Integer Then
                    inf.SetValue(dataSource, Convert.ToInt32(parentcon.Text), Nothing)
                ElseIf TypeOf inf.GetValue(dataSource, Nothing) Is Double Then
                    inf.SetValue(dataSource, Convert.ToDouble(parentcon.Text), Nothing)
                Else
                    inf.SetValue(dataSource, parentcon.Text, Nothing)
                End If


                inf = Nothing

                parentcon.DataBindings.Add("text", dataSource, fieldText)


            End If




        End Sub

    End Class
#End Region




    Private ReadOnly Property TextControl(ByVal row As Int32, ByVal col As Int32) As Control
        Get
            If controlDict.Contains(String.Format("{0},{1}", row, col)) Then
                Return controlDict.Item(String.Format("{0},{1}", row, col))
            Else
                Return fakeLabel
            End If
        End Get
    End Property

    Public Property controlDict() As Collections.Specialized.ListDictionary
        Get
            If dControlDict Is Nothing Then
                SetControlDictionary()
            End If
            Return dControlDict

        End Get
        Set(ByVal Value As Collections.Specialized.ListDictionary)
            dControlDict = Value
        End Set
    End Property

    Public Property controllist() As Collections.ArrayList
        Get
            If lControlList Is Nothing Then
                SetControlDictionary()
            End If
            Return lControlList

        End Get
        Set(ByVal Value As Collections.ArrayList)
            lControlList = Value
        End Set
    End Property


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

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
    Friend WithEvents lblCol1Row1 As System.Windows.Forms.Label
    Friend WithEvents lblCol1Row2 As System.Windows.Forms.Label
    Friend WithEvents lblCol1Row3 As System.Windows.Forms.Label
    Friend WithEvents lblCol1Row4 As System.Windows.Forms.Label
    Friend WithEvents lblCol1Row5 As System.Windows.Forms.Label
    Friend WithEvents lblCol1Row6 As System.Windows.Forms.Label
    Friend WithEvents lblCol1Row7 As System.Windows.Forms.Label
    Friend WithEvents lblCol1Row8 As System.Windows.Forms.Label
    Friend WithEvents txtCol2Row1 As System.Windows.Forms.TextBox
    Friend WithEvents txtCol2Row2 As System.Windows.Forms.TextBox
    Friend WithEvents txtCol2Row3 As System.Windows.Forms.TextBox
    Friend WithEvents txtCol2Row4 As System.Windows.Forms.TextBox
    Friend WithEvents txtCol2Row5 As System.Windows.Forms.TextBox
    Friend WithEvents txtCol2Row6 As System.Windows.Forms.TextBox
    Friend WithEvents txtCol2Row7 As System.Windows.Forms.TextBox
    Friend WithEvents txtCol2Row8 As System.Windows.Forms.TextBox
    Friend WithEvents lblCol3Row1 As System.Windows.Forms.Label
    Friend WithEvents lblCol3Row2 As System.Windows.Forms.Label
    Friend WithEvents lblCol3Row3 As System.Windows.Forms.Label
    Friend WithEvents lblCol3Row4 As System.Windows.Forms.Label
    Friend WithEvents lblCol3Row5 As System.Windows.Forms.Label
    Friend WithEvents lblCol3Row6 As System.Windows.Forms.Label
    Friend WithEvents lblCol3Row7 As System.Windows.Forms.Label
    Friend WithEvents lblCol3Row8 As System.Windows.Forms.Label
    Friend WithEvents txtCol4Row1 As System.Windows.Forms.TextBox
    Friend WithEvents txtCol4Row2 As System.Windows.Forms.TextBox
    Friend WithEvents txtCol4Row3 As System.Windows.Forms.TextBox
    Friend WithEvents txtCol4Row4 As System.Windows.Forms.TextBox
    Friend WithEvents txtCol4Row5 As System.Windows.Forms.TextBox
    Friend WithEvents txtCol4Row6 As System.Windows.Forms.TextBox
    Friend WithEvents txtCol4Row7 As System.Windows.Forms.TextBox
    Friend WithEvents txtCol4Row8 As System.Windows.Forms.TextBox
    Friend WithEvents lblCol5Row1 As System.Windows.Forms.Label
    Friend WithEvents txtCol6Row1 As System.Windows.Forms.TextBox
    Friend WithEvents lblCol5Row2 As System.Windows.Forms.Label
    Friend WithEvents txtCol6Row2 As System.Windows.Forms.TextBox
    Friend WithEvents txtCol6Row3 As System.Windows.Forms.TextBox
    Friend WithEvents lblCol5Row3 As System.Windows.Forms.Label
    Friend WithEvents txtCol6Row4 As System.Windows.Forms.TextBox
    Friend WithEvents lblCol5Row4 As System.Windows.Forms.Label
    Friend WithEvents txtCol6Row5 As System.Windows.Forms.TextBox
    Friend WithEvents lblCol5Row5 As System.Windows.Forms.Label
    Friend WithEvents txtCol6Row6 As System.Windows.Forms.TextBox
    Friend WithEvents lblCol5Row6 As System.Windows.Forms.Label
    Friend WithEvents txtCol6Row7 As System.Windows.Forms.TextBox
    Friend WithEvents lblCol5Row7 As System.Windows.Forms.Label
    Friend WithEvents txtCol6Row8 As System.Windows.Forms.TextBox
    Friend WithEvents lblCol5Row8 As System.Windows.Forms.Label
    Friend WithEvents lblCol1Row9 As System.Windows.Forms.Label
    Friend WithEvents txtCol2Row9 As System.Windows.Forms.TextBox
    Friend WithEvents lblCol3Row9 As System.Windows.Forms.Label
    Friend WithEvents txtCol4Row9 As System.Windows.Forms.TextBox
    Friend WithEvents lblCol5Row9 As System.Windows.Forms.Label
    Friend WithEvents txtCol6Row9 As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblCol1Row1 = New System.Windows.Forms.Label
        Me.lblCol1Row2 = New System.Windows.Forms.Label
        Me.lblCol1Row3 = New System.Windows.Forms.Label
        Me.lblCol1Row4 = New System.Windows.Forms.Label
        Me.lblCol1Row5 = New System.Windows.Forms.Label
        Me.lblCol1Row6 = New System.Windows.Forms.Label
        Me.lblCol1Row7 = New System.Windows.Forms.Label
        Me.lblCol1Row8 = New System.Windows.Forms.Label
        Me.txtCol2Row1 = New System.Windows.Forms.TextBox
        Me.txtCol2Row2 = New System.Windows.Forms.TextBox
        Me.txtCol2Row3 = New System.Windows.Forms.TextBox
        Me.txtCol2Row4 = New System.Windows.Forms.TextBox
        Me.txtCol2Row5 = New System.Windows.Forms.TextBox
        Me.txtCol2Row6 = New System.Windows.Forms.TextBox
        Me.txtCol2Row7 = New System.Windows.Forms.TextBox
        Me.txtCol2Row8 = New System.Windows.Forms.TextBox
        Me.lblCol3Row1 = New System.Windows.Forms.Label
        Me.lblCol3Row2 = New System.Windows.Forms.Label
        Me.lblCol3Row3 = New System.Windows.Forms.Label
        Me.lblCol3Row4 = New System.Windows.Forms.Label
        Me.lblCol3Row5 = New System.Windows.Forms.Label
        Me.lblCol3Row6 = New System.Windows.Forms.Label
        Me.lblCol3Row7 = New System.Windows.Forms.Label
        Me.lblCol3Row8 = New System.Windows.Forms.Label
        Me.txtCol4Row1 = New System.Windows.Forms.TextBox
        Me.txtCol4Row2 = New System.Windows.Forms.TextBox
        Me.txtCol4Row3 = New System.Windows.Forms.TextBox
        Me.txtCol4Row4 = New System.Windows.Forms.TextBox
        Me.txtCol4Row5 = New System.Windows.Forms.TextBox
        Me.txtCol4Row6 = New System.Windows.Forms.TextBox
        Me.txtCol4Row7 = New System.Windows.Forms.TextBox
        Me.txtCol4Row8 = New System.Windows.Forms.TextBox
        Me.lblCol5Row1 = New System.Windows.Forms.Label
        Me.txtCol6Row1 = New System.Windows.Forms.TextBox
        Me.lblCol5Row2 = New System.Windows.Forms.Label
        Me.txtCol6Row2 = New System.Windows.Forms.TextBox
        Me.txtCol6Row3 = New System.Windows.Forms.TextBox
        Me.lblCol5Row3 = New System.Windows.Forms.Label
        Me.txtCol6Row4 = New System.Windows.Forms.TextBox
        Me.lblCol5Row4 = New System.Windows.Forms.Label
        Me.txtCol6Row5 = New System.Windows.Forms.TextBox
        Me.lblCol5Row5 = New System.Windows.Forms.Label
        Me.txtCol6Row6 = New System.Windows.Forms.TextBox
        Me.lblCol5Row6 = New System.Windows.Forms.Label
        Me.txtCol6Row7 = New System.Windows.Forms.TextBox
        Me.lblCol5Row7 = New System.Windows.Forms.Label
        Me.txtCol6Row8 = New System.Windows.Forms.TextBox
        Me.lblCol5Row8 = New System.Windows.Forms.Label
        Me.lblCol1Row9 = New System.Windows.Forms.Label
        Me.txtCol2Row9 = New System.Windows.Forms.TextBox
        Me.lblCol3Row9 = New System.Windows.Forms.Label
        Me.txtCol4Row9 = New System.Windows.Forms.TextBox
        Me.lblCol5Row9 = New System.Windows.Forms.Label
        Me.txtCol6Row9 = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'lblCol1Row1
        '
        Me.lblCol1Row1.Location = New System.Drawing.Point(16, 8)
        Me.lblCol1Row1.Name = "lblCol1Row1"
        Me.lblCol1Row1.Size = New System.Drawing.Size(240, 23)
        Me.lblCol1Row1.TabIndex = 0
        '
        'lblCol1Row2
        '
        Me.lblCol1Row2.Location = New System.Drawing.Point(16, 32)
        Me.lblCol1Row2.Name = "lblCol1Row2"
        Me.lblCol1Row2.Size = New System.Drawing.Size(240, 23)
        Me.lblCol1Row2.TabIndex = 1
        '
        'lblCol1Row3
        '
        Me.lblCol1Row3.Location = New System.Drawing.Point(16, 56)
        Me.lblCol1Row3.Name = "lblCol1Row3"
        Me.lblCol1Row3.Size = New System.Drawing.Size(240, 23)
        Me.lblCol1Row3.TabIndex = 2
        '
        'lblCol1Row4
        '
        Me.lblCol1Row4.Location = New System.Drawing.Point(16, 80)
        Me.lblCol1Row4.Name = "lblCol1Row4"
        Me.lblCol1Row4.Size = New System.Drawing.Size(240, 23)
        Me.lblCol1Row4.TabIndex = 3
        '
        'lblCol1Row5
        '
        Me.lblCol1Row5.Location = New System.Drawing.Point(16, 104)
        Me.lblCol1Row5.Name = "lblCol1Row5"
        Me.lblCol1Row5.Size = New System.Drawing.Size(240, 23)
        Me.lblCol1Row5.TabIndex = 4
        '
        'lblCol1Row6
        '
        Me.lblCol1Row6.Location = New System.Drawing.Point(16, 128)
        Me.lblCol1Row6.Name = "lblCol1Row6"
        Me.lblCol1Row6.Size = New System.Drawing.Size(240, 23)
        Me.lblCol1Row6.TabIndex = 5
        '
        'lblCol1Row7
        '
        Me.lblCol1Row7.Location = New System.Drawing.Point(16, 152)
        Me.lblCol1Row7.Name = "lblCol1Row7"
        Me.lblCol1Row7.Size = New System.Drawing.Size(240, 23)
        Me.lblCol1Row7.TabIndex = 6
        '
        'lblCol1Row8
        '
        Me.lblCol1Row8.Location = New System.Drawing.Point(16, 176)
        Me.lblCol1Row8.Name = "lblCol1Row8"
        Me.lblCol1Row8.Size = New System.Drawing.Size(240, 23)
        Me.lblCol1Row8.TabIndex = 7
        '
        'txtCol2Row1
        '
        Me.txtCol2Row1.Location = New System.Drawing.Point(264, 8)
        Me.txtCol2Row1.Name = "txtCol2Row1"
        Me.txtCol2Row1.Size = New System.Drawing.Size(80, 20)
        Me.txtCol2Row1.TabIndex = 1
        Me.txtCol2Row1.Text = "0.00"
        Me.txtCol2Row1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol2Row1.Visible = False
        '
        'txtCol2Row2
        '
        Me.txtCol2Row2.Location = New System.Drawing.Point(264, 32)
        Me.txtCol2Row2.Name = "txtCol2Row2"
        Me.txtCol2Row2.Size = New System.Drawing.Size(80, 20)
        Me.txtCol2Row2.TabIndex = 4
        Me.txtCol2Row2.Text = "0.00"
        Me.txtCol2Row2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol2Row2.Visible = False
        '
        'txtCol2Row3
        '
        Me.txtCol2Row3.Location = New System.Drawing.Point(264, 56)
        Me.txtCol2Row3.Name = "txtCol2Row3"
        Me.txtCol2Row3.Size = New System.Drawing.Size(80, 20)
        Me.txtCol2Row3.TabIndex = 7
        Me.txtCol2Row3.Text = "0.00"
        Me.txtCol2Row3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol2Row3.Visible = False
        '
        'txtCol2Row4
        '
        Me.txtCol2Row4.Location = New System.Drawing.Point(264, 80)
        Me.txtCol2Row4.Name = "txtCol2Row4"
        Me.txtCol2Row4.Size = New System.Drawing.Size(80, 20)
        Me.txtCol2Row4.TabIndex = 10
        Me.txtCol2Row4.Text = "0.00"
        Me.txtCol2Row4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol2Row4.Visible = False
        '
        'txtCol2Row5
        '
        Me.txtCol2Row5.Location = New System.Drawing.Point(264, 104)
        Me.txtCol2Row5.Name = "txtCol2Row5"
        Me.txtCol2Row5.Size = New System.Drawing.Size(80, 20)
        Me.txtCol2Row5.TabIndex = 13
        Me.txtCol2Row5.Text = "0.00"
        Me.txtCol2Row5.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol2Row5.Visible = False
        '
        'txtCol2Row6
        '
        Me.txtCol2Row6.Location = New System.Drawing.Point(264, 128)
        Me.txtCol2Row6.Name = "txtCol2Row6"
        Me.txtCol2Row6.Size = New System.Drawing.Size(80, 20)
        Me.txtCol2Row6.TabIndex = 16
        Me.txtCol2Row6.Text = "0.00"
        Me.txtCol2Row6.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol2Row6.Visible = False
        '
        'txtCol2Row7
        '
        Me.txtCol2Row7.Location = New System.Drawing.Point(264, 152)
        Me.txtCol2Row7.Name = "txtCol2Row7"
        Me.txtCol2Row7.Size = New System.Drawing.Size(80, 20)
        Me.txtCol2Row7.TabIndex = 19
        Me.txtCol2Row7.Text = "0.00"
        Me.txtCol2Row7.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol2Row7.Visible = False
        '
        'txtCol2Row8
        '
        Me.txtCol2Row8.Location = New System.Drawing.Point(264, 176)
        Me.txtCol2Row8.Name = "txtCol2Row8"
        Me.txtCol2Row8.Size = New System.Drawing.Size(80, 20)
        Me.txtCol2Row8.TabIndex = 22
        Me.txtCol2Row8.Text = "0.00"
        Me.txtCol2Row8.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol2Row8.Visible = False
        '
        'lblCol3Row1
        '
        Me.lblCol3Row1.Location = New System.Drawing.Point(360, 8)
        Me.lblCol3Row1.Name = "lblCol3Row1"
        Me.lblCol3Row1.Size = New System.Drawing.Size(40, 23)
        Me.lblCol3Row1.TabIndex = 16
        Me.lblCol3Row1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblCol3Row2
        '
        Me.lblCol3Row2.Location = New System.Drawing.Point(360, 32)
        Me.lblCol3Row2.Name = "lblCol3Row2"
        Me.lblCol3Row2.Size = New System.Drawing.Size(40, 23)
        Me.lblCol3Row2.TabIndex = 17
        Me.lblCol3Row2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblCol3Row3
        '
        Me.lblCol3Row3.Location = New System.Drawing.Point(360, 56)
        Me.lblCol3Row3.Name = "lblCol3Row3"
        Me.lblCol3Row3.Size = New System.Drawing.Size(40, 23)
        Me.lblCol3Row3.TabIndex = 18
        Me.lblCol3Row3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblCol3Row4
        '
        Me.lblCol3Row4.Location = New System.Drawing.Point(360, 80)
        Me.lblCol3Row4.Name = "lblCol3Row4"
        Me.lblCol3Row4.Size = New System.Drawing.Size(40, 23)
        Me.lblCol3Row4.TabIndex = 19
        Me.lblCol3Row4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblCol3Row5
        '
        Me.lblCol3Row5.Location = New System.Drawing.Point(360, 104)
        Me.lblCol3Row5.Name = "lblCol3Row5"
        Me.lblCol3Row5.Size = New System.Drawing.Size(40, 23)
        Me.lblCol3Row5.TabIndex = 20
        Me.lblCol3Row5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblCol3Row6
        '
        Me.lblCol3Row6.Location = New System.Drawing.Point(360, 128)
        Me.lblCol3Row6.Name = "lblCol3Row6"
        Me.lblCol3Row6.Size = New System.Drawing.Size(40, 23)
        Me.lblCol3Row6.TabIndex = 21
        Me.lblCol3Row6.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblCol3Row7
        '
        Me.lblCol3Row7.Location = New System.Drawing.Point(360, 152)
        Me.lblCol3Row7.Name = "lblCol3Row7"
        Me.lblCol3Row7.Size = New System.Drawing.Size(40, 23)
        Me.lblCol3Row7.TabIndex = 22
        Me.lblCol3Row7.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblCol3Row8
        '
        Me.lblCol3Row8.Location = New System.Drawing.Point(360, 176)
        Me.lblCol3Row8.Name = "lblCol3Row8"
        Me.lblCol3Row8.Size = New System.Drawing.Size(40, 23)
        Me.lblCol3Row8.TabIndex = 23
        Me.lblCol3Row8.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txtCol4Row1
        '
        Me.txtCol4Row1.Location = New System.Drawing.Point(408, 8)
        Me.txtCol4Row1.Name = "txtCol4Row1"
        Me.txtCol4Row1.Size = New System.Drawing.Size(32, 20)
        Me.txtCol4Row1.TabIndex = 2
        Me.txtCol4Row1.Text = "0"
        Me.txtCol4Row1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol4Row1.Visible = False
        '
        'txtCol4Row2
        '
        Me.txtCol4Row2.Location = New System.Drawing.Point(408, 32)
        Me.txtCol4Row2.Name = "txtCol4Row2"
        Me.txtCol4Row2.Size = New System.Drawing.Size(32, 20)
        Me.txtCol4Row2.TabIndex = 5
        Me.txtCol4Row2.Text = "0"
        Me.txtCol4Row2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol4Row2.Visible = False
        '
        'txtCol4Row3
        '
        Me.txtCol4Row3.Location = New System.Drawing.Point(408, 56)
        Me.txtCol4Row3.Name = "txtCol4Row3"
        Me.txtCol4Row3.Size = New System.Drawing.Size(32, 20)
        Me.txtCol4Row3.TabIndex = 8
        Me.txtCol4Row3.Text = "0"
        Me.txtCol4Row3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol4Row3.Visible = False
        '
        'txtCol4Row4
        '
        Me.txtCol4Row4.Location = New System.Drawing.Point(408, 80)
        Me.txtCol4Row4.Name = "txtCol4Row4"
        Me.txtCol4Row4.Size = New System.Drawing.Size(32, 20)
        Me.txtCol4Row4.TabIndex = 11
        Me.txtCol4Row4.Text = "0"
        Me.txtCol4Row4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol4Row4.Visible = False
        '
        'txtCol4Row5
        '
        Me.txtCol4Row5.Location = New System.Drawing.Point(408, 104)
        Me.txtCol4Row5.Name = "txtCol4Row5"
        Me.txtCol4Row5.Size = New System.Drawing.Size(32, 20)
        Me.txtCol4Row5.TabIndex = 14
        Me.txtCol4Row5.Text = "0"
        Me.txtCol4Row5.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol4Row5.Visible = False
        '
        'txtCol4Row6
        '
        Me.txtCol4Row6.Location = New System.Drawing.Point(408, 128)
        Me.txtCol4Row6.Name = "txtCol4Row6"
        Me.txtCol4Row6.Size = New System.Drawing.Size(32, 20)
        Me.txtCol4Row6.TabIndex = 17
        Me.txtCol4Row6.Text = "0"
        Me.txtCol4Row6.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol4Row6.Visible = False
        '
        'txtCol4Row7
        '
        Me.txtCol4Row7.Location = New System.Drawing.Point(408, 152)
        Me.txtCol4Row7.Name = "txtCol4Row7"
        Me.txtCol4Row7.Size = New System.Drawing.Size(32, 20)
        Me.txtCol4Row7.TabIndex = 20
        Me.txtCol4Row7.Text = "0"
        Me.txtCol4Row7.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol4Row7.Visible = False
        '
        'txtCol4Row8
        '
        Me.txtCol4Row8.Location = New System.Drawing.Point(408, 176)
        Me.txtCol4Row8.Name = "txtCol4Row8"
        Me.txtCol4Row8.Size = New System.Drawing.Size(32, 20)
        Me.txtCol4Row8.TabIndex = 23
        Me.txtCol4Row8.Text = "0"
        Me.txtCol4Row8.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol4Row8.Visible = False
        '
        'lblCol5Row1
        '
        Me.lblCol5Row1.Location = New System.Drawing.Point(456, 8)
        Me.lblCol5Row1.Name = "lblCol5Row1"
        Me.lblCol5Row1.Size = New System.Drawing.Size(40, 23)
        Me.lblCol5Row1.TabIndex = 32
        Me.lblCol5Row1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txtCol6Row1
        '
        Me.txtCol6Row1.Location = New System.Drawing.Point(504, 8)
        Me.txtCol6Row1.Name = "txtCol6Row1"
        Me.txtCol6Row1.Size = New System.Drawing.Size(96, 20)
        Me.txtCol6Row1.TabIndex = 3
        Me.txtCol6Row1.Text = "0.00"
        Me.txtCol6Row1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol6Row1.Visible = False
        '
        'lblCol5Row2
        '
        Me.lblCol5Row2.Location = New System.Drawing.Point(456, 32)
        Me.lblCol5Row2.Name = "lblCol5Row2"
        Me.lblCol5Row2.Size = New System.Drawing.Size(40, 23)
        Me.lblCol5Row2.TabIndex = 34
        Me.lblCol5Row2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txtCol6Row2
        '
        Me.txtCol6Row2.Location = New System.Drawing.Point(504, 32)
        Me.txtCol6Row2.Name = "txtCol6Row2"
        Me.txtCol6Row2.Size = New System.Drawing.Size(96, 20)
        Me.txtCol6Row2.TabIndex = 6
        Me.txtCol6Row2.Text = "0.00"
        Me.txtCol6Row2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol6Row2.Visible = False
        '
        'txtCol6Row3
        '
        Me.txtCol6Row3.Location = New System.Drawing.Point(504, 56)
        Me.txtCol6Row3.Name = "txtCol6Row3"
        Me.txtCol6Row3.Size = New System.Drawing.Size(96, 20)
        Me.txtCol6Row3.TabIndex = 9
        Me.txtCol6Row3.Text = "0.00"
        Me.txtCol6Row3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol6Row3.Visible = False
        '
        'lblCol5Row3
        '
        Me.lblCol5Row3.Location = New System.Drawing.Point(456, 56)
        Me.lblCol5Row3.Name = "lblCol5Row3"
        Me.lblCol5Row3.Size = New System.Drawing.Size(40, 23)
        Me.lblCol5Row3.TabIndex = 36
        Me.lblCol5Row3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txtCol6Row4
        '
        Me.txtCol6Row4.Location = New System.Drawing.Point(504, 80)
        Me.txtCol6Row4.Name = "txtCol6Row4"
        Me.txtCol6Row4.Size = New System.Drawing.Size(96, 20)
        Me.txtCol6Row4.TabIndex = 12
        Me.txtCol6Row4.Text = "0.00"
        Me.txtCol6Row4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol6Row4.Visible = False
        '
        'lblCol5Row4
        '
        Me.lblCol5Row4.Location = New System.Drawing.Point(456, 80)
        Me.lblCol5Row4.Name = "lblCol5Row4"
        Me.lblCol5Row4.Size = New System.Drawing.Size(40, 23)
        Me.lblCol5Row4.TabIndex = 38
        Me.lblCol5Row4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txtCol6Row5
        '
        Me.txtCol6Row5.Location = New System.Drawing.Point(504, 104)
        Me.txtCol6Row5.Name = "txtCol6Row5"
        Me.txtCol6Row5.Size = New System.Drawing.Size(96, 20)
        Me.txtCol6Row5.TabIndex = 15
        Me.txtCol6Row5.Text = "0.00"
        Me.txtCol6Row5.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol6Row5.Visible = False
        '
        'lblCol5Row5
        '
        Me.lblCol5Row5.Location = New System.Drawing.Point(456, 104)
        Me.lblCol5Row5.Name = "lblCol5Row5"
        Me.lblCol5Row5.Size = New System.Drawing.Size(40, 23)
        Me.lblCol5Row5.TabIndex = 40
        Me.lblCol5Row5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txtCol6Row6
        '
        Me.txtCol6Row6.Location = New System.Drawing.Point(504, 128)
        Me.txtCol6Row6.Name = "txtCol6Row6"
        Me.txtCol6Row6.Size = New System.Drawing.Size(96, 20)
        Me.txtCol6Row6.TabIndex = 18
        Me.txtCol6Row6.Text = "0.00"
        Me.txtCol6Row6.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol6Row6.Visible = False
        '
        'lblCol5Row6
        '
        Me.lblCol5Row6.Location = New System.Drawing.Point(456, 128)
        Me.lblCol5Row6.Name = "lblCol5Row6"
        Me.lblCol5Row6.Size = New System.Drawing.Size(40, 23)
        Me.lblCol5Row6.TabIndex = 42
        Me.lblCol5Row6.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txtCol6Row7
        '
        Me.txtCol6Row7.Location = New System.Drawing.Point(504, 152)
        Me.txtCol6Row7.Name = "txtCol6Row7"
        Me.txtCol6Row7.Size = New System.Drawing.Size(96, 20)
        Me.txtCol6Row7.TabIndex = 21
        Me.txtCol6Row7.Text = "0.00"
        Me.txtCol6Row7.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol6Row7.Visible = False
        '
        'lblCol5Row7
        '
        Me.lblCol5Row7.Location = New System.Drawing.Point(456, 152)
        Me.lblCol5Row7.Name = "lblCol5Row7"
        Me.lblCol5Row7.Size = New System.Drawing.Size(40, 23)
        Me.lblCol5Row7.TabIndex = 44
        Me.lblCol5Row7.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txtCol6Row8
        '
        Me.txtCol6Row8.Location = New System.Drawing.Point(504, 176)
        Me.txtCol6Row8.Name = "txtCol6Row8"
        Me.txtCol6Row8.Size = New System.Drawing.Size(96, 20)
        Me.txtCol6Row8.TabIndex = 24
        Me.txtCol6Row8.Text = "0.00"
        Me.txtCol6Row8.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol6Row8.Visible = False
        '
        'lblCol5Row8
        '
        Me.lblCol5Row8.Location = New System.Drawing.Point(456, 176)
        Me.lblCol5Row8.Name = "lblCol5Row8"
        Me.lblCol5Row8.Size = New System.Drawing.Size(40, 23)
        Me.lblCol5Row8.TabIndex = 46
        Me.lblCol5Row8.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblCol1Row9
        '
        Me.lblCol1Row9.Location = New System.Drawing.Point(16, 200)
        Me.lblCol1Row9.Name = "lblCol1Row9"
        Me.lblCol1Row9.Size = New System.Drawing.Size(240, 23)
        Me.lblCol1Row9.TabIndex = 8
        '
        'txtCol2Row9
        '
        Me.txtCol2Row9.Location = New System.Drawing.Point(264, 200)
        Me.txtCol2Row9.Name = "txtCol2Row9"
        Me.txtCol2Row9.Size = New System.Drawing.Size(80, 20)
        Me.txtCol2Row9.TabIndex = 25
        Me.txtCol2Row9.Text = "0.00"
        Me.txtCol2Row9.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol2Row9.Visible = False
        '
        'lblCol3Row9
        '
        Me.lblCol3Row9.Location = New System.Drawing.Point(360, 200)
        Me.lblCol3Row9.Name = "lblCol3Row9"
        Me.lblCol3Row9.Size = New System.Drawing.Size(40, 23)
        Me.lblCol3Row9.TabIndex = 23
        Me.lblCol3Row9.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txtCol4Row9
        '
        Me.txtCol4Row9.Location = New System.Drawing.Point(408, 200)
        Me.txtCol4Row9.Name = "txtCol4Row9"
        Me.txtCol4Row9.Size = New System.Drawing.Size(32, 20)
        Me.txtCol4Row9.TabIndex = 26
        Me.txtCol4Row9.Text = "0"
        Me.txtCol4Row9.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol4Row9.Visible = False
        '
        'lblCol5Row9
        '
        Me.lblCol5Row9.Location = New System.Drawing.Point(456, 200)
        Me.lblCol5Row9.Name = "lblCol5Row9"
        Me.lblCol5Row9.Size = New System.Drawing.Size(40, 23)
        Me.lblCol5Row9.TabIndex = 46
        Me.lblCol5Row9.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txtCol6Row9
        '
        Me.txtCol6Row9.Location = New System.Drawing.Point(504, 200)
        Me.txtCol6Row9.Name = "txtCol6Row9"
        Me.txtCol6Row9.Size = New System.Drawing.Size(96, 20)
        Me.txtCol6Row9.TabIndex = 27
        Me.txtCol6Row9.Text = "0.00"
        Me.txtCol6Row9.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtCol6Row9.Visible = False
        '
        'CostFormat
        '
        Me.Controls.Add(Me.txtCol6Row8)
        Me.Controls.Add(Me.lblCol5Row8)
        Me.Controls.Add(Me.txtCol6Row7)
        Me.Controls.Add(Me.lblCol5Row7)
        Me.Controls.Add(Me.txtCol6Row6)
        Me.Controls.Add(Me.lblCol5Row6)
        Me.Controls.Add(Me.txtCol6Row5)
        Me.Controls.Add(Me.lblCol5Row5)
        Me.Controls.Add(Me.txtCol6Row4)
        Me.Controls.Add(Me.lblCol5Row4)
        Me.Controls.Add(Me.txtCol6Row3)
        Me.Controls.Add(Me.lblCol5Row3)
        Me.Controls.Add(Me.txtCol6Row2)
        Me.Controls.Add(Me.lblCol5Row2)
        Me.Controls.Add(Me.txtCol6Row1)
        Me.Controls.Add(Me.lblCol5Row1)
        Me.Controls.Add(Me.txtCol4Row8)
        Me.Controls.Add(Me.txtCol4Row7)
        Me.Controls.Add(Me.txtCol4Row6)
        Me.Controls.Add(Me.txtCol4Row5)
        Me.Controls.Add(Me.txtCol4Row4)
        Me.Controls.Add(Me.txtCol4Row3)
        Me.Controls.Add(Me.txtCol4Row2)
        Me.Controls.Add(Me.txtCol4Row1)
        Me.Controls.Add(Me.lblCol3Row8)
        Me.Controls.Add(Me.lblCol3Row7)
        Me.Controls.Add(Me.lblCol3Row6)
        Me.Controls.Add(Me.lblCol3Row5)
        Me.Controls.Add(Me.lblCol3Row4)
        Me.Controls.Add(Me.lblCol3Row3)
        Me.Controls.Add(Me.lblCol3Row2)
        Me.Controls.Add(Me.lblCol3Row1)
        Me.Controls.Add(Me.txtCol2Row8)
        Me.Controls.Add(Me.txtCol2Row7)
        Me.Controls.Add(Me.txtCol2Row6)
        Me.Controls.Add(Me.txtCol2Row5)
        Me.Controls.Add(Me.txtCol2Row4)
        Me.Controls.Add(Me.txtCol2Row3)
        Me.Controls.Add(Me.txtCol2Row2)
        Me.Controls.Add(Me.txtCol2Row1)
        Me.Controls.Add(Me.lblCol1Row8)
        Me.Controls.Add(Me.lblCol1Row7)
        Me.Controls.Add(Me.lblCol1Row6)
        Me.Controls.Add(Me.lblCol1Row5)
        Me.Controls.Add(Me.lblCol1Row4)
        Me.Controls.Add(Me.lblCol1Row3)
        Me.Controls.Add(Me.lblCol1Row2)
        Me.Controls.Add(Me.lblCol1Row1)
        Me.Controls.Add(Me.lblCol1Row9)
        Me.Controls.Add(Me.txtCol2Row9)
        Me.Controls.Add(Me.lblCol3Row9)
        Me.Controls.Add(Me.txtCol4Row9)
        Me.Controls.Add(Me.lblCol5Row9)
        Me.Controls.Add(Me.txtCol6Row9)
        Me.Name = "CostFormat"
        Me.Size = New System.Drawing.Size(600, 232)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Page Events "

    Overloads Sub dispose()
        dControlDict = Nothing
        lControlList = Nothing
    End Sub

    Private Sub CostFormat_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        bolloading = False

        SetControlDictionary()
    End Sub


    Private Sub DataFields_LostFocus(ByVal sender As Object, ByVal e As EventArgs)

        DirectCast(sender, TextBox).TextAlign = HorizontalAlignment.Right

        editingControl = Nothing

        DataFields_TextChanged(sender, e)
        'all 2 except
        'K,O,W = 0 on txtCol2Row2
        'H,L = 0 on txtcol2Row4
        'G = 0 on txtcol2Row6

    End Sub

    Private Sub DataFields_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)

        editingControl = sender

        editingControl.TextAlign = HorizontalAlignment.Left

        If Not IsNumeric(editingControl.Text) OrElse (IsNumeric(editingControl.Text) AndAlso Convert.ToDouble(editingControl.Text.Replace("$", String.Empty)) = 0) Then
            editingControl.Text = String.Empty
        End If

    End Sub

    Private Sub DataFields_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try

            If Not DirectCast(sender, TextBox).ReadOnly AndAlso Not bolFixingDuringTextChange AndAlso TypeOf sender Is TextBox AndAlso ((Not DirectCast(sender, Control) Is editingControl) Or bolloading) Then

                bolFixingDuringTextChange = True

                With DirectCast(sender, TextBox)

                    Dim spec As CostFormatSpec = .Tag
                    Dim textStr As String


                    If Not spec Is Nothing Then
                        textStr = "$0.00"
                        If IsNumeric(.Text.Replace("$", String.Empty)) Then textStr = .Text.Replace("$", String.Empty)
                        .Text = String.Format("{1}{0}", FormatNumber(textStr, spec.DecimalPlace, TriState.True, TriState.False, TriState.True), IIf(spec.DecimalPlace >= 2, "$", String.Empty))
                    Else
                        textStr = "0"
                        If IsNumeric(.Text) Then textStr = .Text
                        .Text = FormatNumber(textStr, 0, TriState.True, TriState.False, TriState.True)
                    End If
                End With

                bolFixingDuringTextChange = False

            End If

            If bolloading OrElse RecalcInProcess OrElse bolFixingDuringTextChange Then
                Exit Sub
            End If





            'If (spec Is Nothing OrElse spec.DecimalPlace <= 0) AndAlso useNewCode Then
            'If Textstr.IndexOf(".") > -1 Then
            'MsgBox("Only Whole Numbers Allowed In This Field", MsgBoxStyle.Exclamation, "Integer Only")
            '.Text = FormatNumber(Textstr, spec.DecimalPlace, TriState.True, TriState.False, TriState.True)
            'End If
            'ElseIf Not IsNumeric(.Text) AndAlso .Text.Trim <> String.Empty Then
            'MsgBox("Only Numerical Allowed In This Field", MsgBoxStyle.Exclamation, "Numerics Only")
            'If Not spec Is Nothing Then
            '.Text = FormatNumber(textstr, spec.DecimalPlace, TriState.True, TriState.False, TriState.True)
            'Else
            '   .Text = "0.00"
            'End If
            ' ElseIf Not spec Is Nothing AndAlso textstr.IndexOf(".") = -1 AndAlso spec.DecimalPlace > 0 AndAlso useNewCode Then
            'bolFixingDuringTextChange = True
            '.Text = FormatNumber(textstr, spec.DecimalPlace, TriState.True, TriState.False, TriState.True)
            'bolFixingDuringTextChange = False
            'End If

            If Not Me.bolFixingDuringTextChange Then
                ReCalcAll()
                PushToObject()
            End If
        Catch ex As Exception
            Throw ex
        End Try


    End Sub


#End Region

#Region " Processes "

    Private Sub SetControlDictionary()

        Dim indx As String

        controlDict = New Collections.Specialized.ListDictionary
        controllist = New Collections.ArrayList

        For Each Con As Control In Controls

            With Con.Name.ToUpper
                If (.StartsWith("TXT") Or .StartsWith("LBL")) AndAlso .IndexOf("COL") = 3 AndAlso .IndexOf("ROW") > 6 AndAlso Not .EndsWith("ROW") Then

                    indx = String.Format("{0},{1}", .Substring(.LastIndexOf("ROW") + 3), .Substring(6, .IndexOf("ROW") - 6))
                    controlDict.Add(indx, Con)
                    controllist.Add(Con)

                    If .StartsWith("TXT") Then
                        RemoveHandler DirectCast(Con, TextBox).Validated, AddressOf DataFields_LostFocus
                        AddHandler DirectCast(Con, TextBox).Validated, AddressOf DataFields_LostFocus

                        RemoveHandler DirectCast(Con, TextBox).TextChanged, AddressOf DataFields_TextChanged
                        AddHandler DirectCast(Con, TextBox).TextChanged, AddressOf DataFields_TextChanged

                        RemoveHandler DirectCast(Con, TextBox).Enter, AddressOf DataFields_Enter
                        AddHandler DirectCast(Con, TextBox).Enter, AddressOf DataFields_Enter


                    End If



                End If
            End With

        Next

    End Sub



    Public Sub AssignCommitmentObject(ByRef oPassedFinancialCommitment As MUSTER.BusinessLogic.pFinancialCommitment)
        oLocalFinancialCommitment = oPassedFinancialCommitment
    End Sub


    Public Sub SetReadonly(ByVal bolReadonly As Boolean)

        For Each con As Control In controllist

            If TypeOf con Is TextBox Then
                DirectCast(con, TextBox).ReadOnly = bolReadonly
            End If

        Next

    End Sub


    Public Sub SetDisplay(Optional ByVal ClearObject As Boolean = True)
        bolloading = True


        Try

            If useNewCode Then
                Me.ResetControl()

                If Not ParentForm Is Nothing Then
                    ParentForm.Text = "Committment"
                End If



                Dim act As New BusinessLogic.pFinancialActivity
                Dim data As DataTable = act.GetCostFormatDesign(CostFormatType)


                If Not data Is Nothing AndAlso data.Rows.Count > 0 Then

                    Dim cnt As Integer = 1
                    Dim col4mem As TextBox = Nothing

                    For Each row As DataRow In data.Rows

                        Dim primaryText As String = row("PrimaryText").ToString
                        Dim Col3text As String = row("Col3Text").ToString
                        Dim Col5text As String = row("Col5Text").ToString
                        Dim Col2text As String = row("Col2Text").ToString
                        Dim Col4text As String = row("Col4Text").ToString
                        Dim Col6text As String = row("Col6Text").ToString

                        Dim Col2Field As String = row("Col2Field").ToString
                        Dim Col4Field As String = row("Col4field").ToString
                        Dim Col6Field As String = row("Col6Field").ToString

                        Dim Col2DecPoint As Integer = DirectCast(row("Col2DecPoint"), Int32)
                        Dim Col4DecPoint As Integer = DirectCast(row("Col4DecPoint"), Int32)
                        Dim Col6DecPoint As Integer = DirectCast(row("Col6DecPoint"), Int32)

                        Dim Col2ReadOnly As Boolean = DirectCast(row("Col2ReadOnly"), Boolean)
                        Dim Col4ReadOnly As Boolean = DirectCast(row("Col4ReadOnly"), Boolean)

                        Dim Col2Visible As Boolean = DirectCast(row("Col2Visible"), Boolean)
                        Dim Col4Visible As Boolean = DirectCast(row("Col4Visible"), Boolean)
                        Dim Col6Visible As Boolean = DirectCast(row("Col6Visible"), Boolean)

                        Dim Col2GrandTotal As Boolean = DirectCast(row("Col2GrandTotal"), Boolean)
                        Dim Col4GrandTotal As Boolean = DirectCast(row("Col4GrandTotal"), Boolean)
                        Dim Col6GrandTotal As Boolean = DirectCast(row("Col6GrandTotal"), Boolean)

                        Dim isTotal As Boolean = DirectCast(row("IsTotal"), Boolean)
                        Dim isSubTotal As Boolean = DirectCast(row("isSubTotal"), Boolean)

                        If Col4text.Length > 0 Then col4mem = TextControl(cnt, 4)


                        If isSubTotal Or isTotal Then
                            col4mem = Nothing
                        End If

                        TextControl(cnt, 1).Tag = New CostFormatSpec(TextControl(cnt, 1), 0, False, True, True, primaryText, isSubTotal, isTotal, String.Empty, oLocalFinancialCommitment)
                        TextControl(cnt, 3).Tag = New CostFormatSpec(TextControl(cnt, 3), 0, False, Not Col3text = String.Empty, True, Col3text, False, False, String.Empty, oLocalFinancialCommitment)
                        TextControl(cnt, 5).Tag = New CostFormatSpec(TextControl(cnt, 5), 0, False, Not Col5text = String.Empty, True, Col5text, False, False, String.Empty, oLocalFinancialCommitment)

                        TextControl(cnt, 2).Tag = New CostFormatSpec(TextControl(cnt, 2), Col2DecPoint, Col2GrandTotal, Col2Visible, Col2ReadOnly, Col2text, isSubTotal, isTotal, Col2Field, oLocalFinancialCommitment)
                        TextControl(cnt, 4).Tag = New CostFormatSpec(TextControl(cnt, 4), Col4DecPoint, Col4GrandTotal, Col4Visible, Col4ReadOnly, Col4text, isSubTotal, isTotal, Col4Field, oLocalFinancialCommitment, col4mem)
                        TextControl(cnt, 6).Tag = New CostFormatSpec(TextControl(cnt, 6), Col6DecPoint, Col6GrandTotal, Col6Visible, True, Col6text, isSubTotal, isTotal, Col6Field, oLocalFinancialCommitment)



                        cnt += 1


                    Next

                End If


                For Each con As Control In controllist
                    If Not con.Tag Is Nothing AndAlso TypeOf con.Tag Is CostFormatSpec Then
                        con = DirectCast(con.Tag, CostFormatSpec).SetupControl(con)
                    End If

                Next

            Else

                ResetControl()

                ResetData(ClearObject)

                setDisplay_old()

            End If

            bolloading = False

            Me.Height = Me.lblCol1Row9.Height + Me.lblCol1Row9.Top + 5
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try



    End Sub

    Private Sub ReCalcAll()
        RecalcInProcess = True
        Try

            If Not useNewCode Then
                RecalcAll_old()
            Else

                Dim subTotal() As Double = {0.0, 0.0}
                Dim GrandTotals() As Double = {0.0, 0.0}
                Dim aggregative As String = String.Empty
                Dim row As Integer = 1
                Dim col As Integer = 1


                For row = 1 To 9
                    For col = 1 To 6


                        Dim indx = String.Format("{0},{1}", row, col)
                        If controlDict.Contains(indx) Then

                            Dim con As Control = DirectCast(controlDict.Item(indx), Control)
                            If TypeOf con.Tag Is CostFormatSpec Then

                                With DirectCast(con.Tag, CostFormatSpec)

                                    If (Not .subTotal AndAlso Not .Total) Then

                                        If col = 3 Then
                                            aggregative = con.Text.Trim
                                        End If

                                        If col = 6 AndAlso con.Visible Then
                                            Dim a As String = TextControl(row, 2).Text.Replace("$", String.Empty)
                                            Dim b As String = TextControl(row, 4).Text.Replace("$", String.Empty)
                                            Dim aa As Double
                                            Dim bb As Double


                                            aa = Convert.ToDouble(IIf(IsNumeric(a), a, "0.0"))
                                            bb = Convert.ToDouble(IIf(IsNumeric(b), b, "0.0"))

                                            Select Case aggregative.ToUpper
                                                Case "+", "AND"
                                                    con.Text = FormatNumber(aa + bb, .DecimalPlace, TriState.False, TriState.False, TriState.True)
                                                Case "-", "MINUS"
                                                    con.Text = FormatNumber(aa - bb, .DecimalPlace, TriState.False, TriState.False, TriState.True)
                                                Case "*", "PER", String.Empty, "X"
                                                    con.Text = FormatNumber(aa * bb, .DecimalPlace, TriState.False, TriState.False, TriState.True)
                                                Case "/", "FOR"
                                                    con.Text = FormatNumber(aa / bb, .DecimalPlace, TriState.False, TriState.False, TriState.True)
                                                Case Else
                                                    con.Text = FormatNumber(Math.Max(aa, bb), .DecimalPlace, TriState.False, TriState.False, TriState.True)
                                            End Select

                                            con.Text = String.Format("${0}", con.Text)

                                        End If

                                        If (col = 2 Or col = 6) AndAlso IsNumeric(con.Text) AndAlso con.Visible AndAlso Not .subTotal Then
                                            subTotal(IIf(col = 2, 0, 1)) += Convert.ToDouble(con.Text.Replace("$", String.Empty))
                                        End If


                                    End If

                                    If .subTotal Then

                                        If (col = 2 AndAlso Not TextControl(row, 6).Visible) Or (col = 6 AndAlso con.Visible) Then
                                            If Not .Total Then
                                                GrandTotals(IIf(col = 2, 0, 1)) += subTotal((IIf(col = 2, 0, 1)))
                                                con.Text = String.Format("${0}", FormatNumber(subTotal(IIf(col = 2, 0, 1)), .DecimalPlace, TriState.False, TriState.False, TriState.True))

                                            ElseIf IsNumeric(con.Text) Then
                                                GrandTotals(IIf(col = 2, 0, 1)) += IIf(GrandTotals(IIf(col = 2, 0, 1)) <= 0, subTotal((IIf(col = 2, 0, 1))), 0)

                                                GrandTotals(IIf(col = 2, 0, 1)) *= Convert.ToDouble(con.Text.Replace("$", String.Empty))
                                            Else

                                                GrandTotals(IIf(col = 2, 0, 1)) *= 0.0


                                            End If

                                            subTotal(IIf(col = 2, 0, 1)) = 0.0

                                        End If

                                        If col = 6 AndAlso con.Visible Then
                                            TextControl(row, 2).Visible = False
                                        End If

                                        If (col = 3 Or col = 4 Or col = 5) Then
                                            con.Visible = False
                                        End If

                                    End If

                                    If .Total And Not .subTotal Then

                                        Dim arrayPlace As Integer = Convert.ToInt32(IIf(col = 2, 0, 1))

                                        If (col = 2 AndAlso Not TextControl(row, 6).Visible) Or (col = 6 AndAlso con.Visible) Then
                                            con.Text = String.Format("${0}", FormatNumber(GrandTotals(arrayPlace) + subTotal(arrayPlace), .DecimalPlace, TriState.False, TriState.False, TriState.True))
                                            subTotal(arrayPlace) = 0.0
                                            ' GrandTotals(arrayPlace) = 0.0

                                        End If

                                        If col = 6 AndAlso con.Visible Then
                                            TextControl(row, 2).Visible = False
                                        End If

                                        If (col = 3 Or col = 4 Or col = 5) Then
                                            con.Visible = False
                                        End If

                                    End If

                                    If .PartOfTotal AndAlso IsNumeric(con.Text) Then
                                        GrandTotal = Convert.ToDouble(con.Text.Replace("$", String.Empty))
                                    End If

                                End With
                            End If
                        End If

                    Next
                Next


            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Calculate Commitment " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            RecalcInProcess = False
        End Try


    End Sub

    Private Sub ResetControl()



        Dim spec As CostFormatSpec

        For Each con As Control In controllist

            If TypeOf con Is TextBox Then

                With DirectCast(con, TextBox)

                    spec = .Tag

                    .Visible = False
                    .TabStop = True
                    .ReadOnly = False

                    With .Name.ToUpper
                        If .IndexOf("COL2") > -1 OrElse .IndexOf("COL6") > -1 OrElse (.Substring(.LastIndexOf("ROW"), 1) >= "2" AndAlso .Substring(.LastIndexOf("ROW"), 1) <= "4") Then
                            DirectCast(con, TextBox).BorderStyle = BorderStyle.Fixed3D

                        Else
                            DirectCast(con, TextBox).BorderStyle = BorderStyle.None
                        End If

                    End With

                    If (spec Is Nothing OrElse spec.DecimalPlace > 0) OrElse .Name.ToUpper.IndexOf("COL4") <= -1 Then
                        '   .Text = "0.00"
                    Else
                        .Text = "0"
                    End If

                    .Tag = Nothing

                End With

            End If

            If TypeOf con Is Label Then

                With DirectCast(con, Label)

                    spec = .Tag
                    .Text = String.Empty
                    .Visible = False
                    .Tag = Nothing


                End With

            End If
        Next

    End Sub

    Public Sub LoadCommitment()

        Try
            bolloading = True

            If Not useNewCode Then

                LoadCommitment_old()

            Else

                '   For Each con As Control In controllist

                '  con.DataBindings.Clear()


                ' Next


            End If

            bolloading = False
            ReCalcAll()
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot LoadCommitment " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub PushToObject()

        Try


            If Not useNewCode Then
                PushToObject_old()
            Else
                'taken care of by data binding
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot PushToObject " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub


#End Region

#Region "Old Code"

    Private Sub setDisplay_old()

        If Not ParentForm Is Nothing Then
            ParentForm.Text = "Committment (Old Code)"
        End If


        lblCol1Row1.Visible = True
        lblCol1Row2.Visible = True
        lblCol1Row3.Visible = True
        lblCol1Row4.Visible = True
        lblCol1Row5.Visible = True
        lblCol1Row6.Visible = True
        lblCol1Row7.Visible = True
        lblCol1Row8.Visible = True
        lblCol1Row9.Visible = True

        Select Case CostFormatType

            Case "A"
                lblCol1Row1.Text = "ERAC Services"
                txtCol2Row1.Visible = True

                lblCol1Row2.Text = "Drilling and Laboratory Services"
                txtCol2Row2.Visible = True

                lblCol1Row3.Text = "Estimated Markup"
                txtCol2Row3.Visible = True

                lblCol1Row4.Text = "Total"
                txtCol2Row4.Visible = True
                txtCol2Row4.ReadOnly = True
                txtCol2Row4.TabStop = False
                txtCol2Row4.BorderStyle = BorderStyle.FixedSingle

            Case "B"
                lblCol1Row1.Text = "ERAC Services"
                txtCol2Row1.Visible = True

                lblCol1Row2.Text = "Drilling and Laboratory Services"
                txtCol2Row2.Visible = True

                lblCol1Row3.Text = "Subtotal"
                txtCol2Row3.Visible = True
                txtCol2Row3.ReadOnly = True
                txtCol2Row3.BorderStyle = BorderStyle.FixedSingle
                txtCol2Row3.TabStop = False

                lblCol1Row4.Text = "Fixed-Fee"
                txtCol2Row4.Visible = True

                lblCol1Row5.Text = "Markup"
                txtCol2Row5.Visible = True

                lblCol1Row6.Text = "Total"
                txtCol2Row6.Visible = True
                txtCol2Row6.ReadOnly = True
                txtCol2Row6.TabStop = False
                txtCol2Row6.BorderStyle = BorderStyle.FixedSingle

            Case "C"
                lblCol1Row1.Text = "ERAC Services"
                txtCol2Row1.Visible = True

            Case "CR"
                lblCol1Row1.Text = "Cost Recovery"
                txtCol2Row1.Visible = True

            Case "D"
                lblCol1Row1.Text = "ERAC Services"
                txtCol2Row1.Visible = True

                lblCol1Row2.Text = "Fixed-Fee"
                txtCol2Row2.Visible = True

                lblCol1Row3.Text = "Estimated Markup"
                txtCol2Row3.Visible = True

                lblCol1Row4.Text = "Total"
                txtCol2Row4.Visible = True
                txtCol2Row4.ReadOnly = True
                txtCol2Row4.TabStop = False
                txtCol2Row4.BorderStyle = BorderStyle.FixedSingle

            Case "E"
                lblCol1Row1.Text = "ERAC Services"
                txtCol2Row1.Visible = True

                lblCol1Row2.Text = "Laboratory Services"
                txtCol2Row2.Visible = True

                lblCol1Row3.Text = "Estimated Markup"
                txtCol2Row3.Visible = True

                lblCol1Row4.Text = "Total"
                txtCol2Row4.Visible = True
                txtCol2Row4.ReadOnly = True
                txtCol2Row4.TabStop = False
                txtCol2Row4.BorderStyle = BorderStyle.FixedSingle

            Case "F"
                lblCol1Row1.Text = "ERAC Services"
                txtCol2Row1.Visible = True

                lblCol1Row2.Text = "Laboratory Services"
                txtCol2Row2.Visible = True

                lblCol1Row3.Text = "Subtotal"
                txtCol2Row3.Visible = True
                txtCol2Row3.ReadOnly = True
                txtCol2Row3.TabStop = False
                txtCol2Row3.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row4.Text = "Fixed-Fee"
                txtCol2Row4.Visible = True

                lblCol1Row5.Text = "Markup"
                txtCol2Row5.Visible = True

                lblCol1Row6.Text = "Total"
                txtCol2Row6.Visible = True
                txtCol2Row6.ReadOnly = True
                txtCol2Row6.TabStop = False
                txtCol2Row6.BorderStyle = BorderStyle.FixedSingle


            Case "G"
                lblCol1Row1.Text = "ERAC Services"
                txtCol2Row1.Visible = True

                lblCol1Row2.Text = "Laboratory Services"
                txtCol2Row2.Visible = True

                lblCol1Row3.Text = "Subtotal"
                txtCol2Row3.Visible = True
                txtCol2Row3.ReadOnly = True
                txtCol2Row3.TabStop = False
                txtCol2Row3.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row4.Text = "Fixed-Fee"
                txtCol2Row4.Visible = True

                lblCol1Row5.Text = "Total Per Event"
                txtCol2Row5.Visible = True
                txtCol2Row5.ReadOnly = True
                txtCol2Row5.TabStop = False
                txtCol2Row5.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row6.Text = "Number of Events"
                txtCol2Row6.Visible = True

                lblCol1Row7.Text = "Total"
                txtCol2Row7.Visible = True
                txtCol2Row7.ReadOnly = True
                txtCol2Row7.TabStop = False
                txtCol2Row7.BorderStyle = BorderStyle.FixedSingle

            Case "H"
                lblCol1Row1.Text = "ERAC Services"
                txtCol2Row1.Visible = True

                lblCol1Row2.Text = "Laboratory Services"
                txtCol2Row2.Visible = True

                lblCol1Row3.Text = "Subtotal"
                txtCol2Row3.Visible = True
                txtCol2Row3.ReadOnly = True
                txtCol2Row3.TabStop = False
                txtCol2Row3.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row4.Text = "Number of Events"
                txtCol2Row4.Visible = True

                lblCol1Row5.Text = "Total"
                txtCol2Row5.Visible = True
                txtCol2Row5.ReadOnly = True
                txtCol2Row5.TabStop = False
                txtCol2Row5.BorderStyle = BorderStyle.FixedSingle

            Case "I"
                lblCol1Row1.Text = "Well Abandonment Services"
                txtCol2Row1.Visible = True

            Case "IR"
                lblCol1Row1.Text = "IRAC Services"
                txtCol2Row1.Visible = True

            Case "J"
                lblCol1Row1.Text = "ERAC Services"
                txtCol2Row1.Visible = True

                lblCol1Row2.Text = "Well Abandonment Subcontractor"
                txtCol2Row2.Visible = True

                lblCol1Row3.Text = "Subtotal"
                txtCol2Row3.Visible = True
                txtCol2Row3.ReadOnly = True
                txtCol2Row3.TabStop = False
                txtCol2Row3.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row4.Text = "Fixed-Fee"
                txtCol2Row4.Visible = True

                lblCol1Row5.Text = "Markup"
                txtCol2Row5.Visible = True

                lblCol1Row6.Text = "Total"
                txtCol2Row6.Visible = True
                txtCol2Row6.ReadOnly = True
                txtCol2Row6.TabStop = False
                txtCol2Row6.BorderStyle = BorderStyle.FixedSingle

            Case "K"
                lblCol1Row1.Text = "Free Product Recovery Services"
                txtCol2Row1.Visible = True

                lblCol1Row2.Text = "Number of Events"
                txtCol2Row2.Visible = True

                lblCol1Row3.Text = "Total"
                txtCol2Row3.Visible = True
                txtCol2Row3.ReadOnly = True
                txtCol2Row3.TabStop = False
                txtCol2Row3.BorderStyle = BorderStyle.FixedSingle

            Case "L"
                lblCol1Row1.Text = "ERAC Services"
                txtCol2Row1.Visible = True

                lblCol1Row2.Text = "Vacuuming Contractor"
                txtCol2Row2.Visible = True

                lblCol1Row3.Text = "Subtotal"
                txtCol2Row3.Visible = True
                txtCol2Row3.ReadOnly = True
                txtCol2Row3.TabStop = False
                txtCol2Row3.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row4.Text = "Number of Events"
                txtCol2Row4.Visible = True

                lblCol1Row5.Text = "Total"
                txtCol2Row5.Visible = True
                txtCol2Row5.ReadOnly = True
                txtCol2Row5.TabStop = False
                txtCol2Row5.BorderStyle = BorderStyle.FixedSingle

            Case "M"
                lblCol1Row1.Text = "Precision Tank Tightness Testing Services"
                txtCol2Row1.Visible = True

            Case "N"
                lblCol1Row1.Text = "ERAC Vacuum Services"
                txtCol2Row1.Visible = True
                lblCol3Row1.Text = "per"
                txtCol4Row1.Visible = True
                lblCol5Row1.Text = "events"

                lblCol1Row2.Text = "Vacuum Contractor Services"
                txtCol2Row2.Visible = True
                lblCol3Row2.Text = "per"
                txtCol4Row2.Visible = True
                txtCol4Row2.ReadOnly = True
                txtCol4Row2.TabStop = False
                txtCol4Row2.BorderStyle = BorderStyle.FixedSingle
                lblCol5Row2.Text = "events"

                lblCol1Row3.Text = "Subtotal"
                txtCol2Row3.Visible = True
                txtCol2Row3.ReadOnly = True
                txtCol2Row3.TabStop = False
                txtCol2Row3.BorderStyle = BorderStyle.FixedSingle
                lblCol3Row3.Text = "per"
                txtCol4Row3.Visible = True
                txtCol4Row3.ReadOnly = True
                txtCol4Row3.TabStop = False
                lblCol5Row3.Text = "events"
                txtCol4Row3.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row4.Text = "Total"
                txtCol2Row4.Visible = True
                txtCol2Row4.ReadOnly = True
                txtCol2Row4.TabStop = False
                txtCol2Row4.BorderStyle = BorderStyle.FixedSingle
                lblCol3Row4.Text = "per"
                txtCol4Row4.Visible = True
                txtCol4Row4.ReadOnly = True
                txtCol4Row4.TabStop = False
                lblCol5Row4.Text = "vacuum events"
                lblCol5Row4.Width = 80
                txtCol4Row4.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row5.Text = "ERAC Sampling Services"
                txtCol2Row5.Visible = True

                lblCol1Row6.Text = "Laboratory Services"
                txtCol2Row6.Visible = True

                lblCol1Row7.Text = "Total for Sampling Services"
                txtCol2Row7.Visible = True
                txtCol2Row7.ReadOnly = True
                txtCol2Row7.TabStop = False
                txtCol2Row7.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row8.Text = "Grand Total"
                txtCol2Row8.Visible = True
                txtCol2Row8.ReadOnly = True
                txtCol2Row8.TabStop = False
                txtCol2Row8.BorderStyle = BorderStyle.FixedSingle

            Case "O"
                lblCol1Row1.Text = "Free Product Recovery Services"
                txtCol2Row1.Visible = True

                lblCol1Row2.Text = "Number of Events"
                txtCol2Row2.Visible = True

                lblCol1Row3.Text = "Subtotal"
                txtCol2Row3.Visible = True
                txtCol2Row3.ReadOnly = True
                txtCol2Row3.TabStop = False
                txtCol2Row3.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row4.Text = "ERAC Sampling Services"
                txtCol2Row4.Visible = True

                lblCol1Row5.Text = "Laboratory Services"
                txtCol2Row5.Visible = True

                lblCol1Row6.Text = "Subtotal"
                txtCol2Row6.Visible = True
                txtCol2Row6.ReadOnly = True
                txtCol2Row6.TabStop = False
                txtCol2Row6.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row7.Text = "Total"
                txtCol2Row7.Visible = True
                txtCol2Row7.ReadOnly = True
                txtCol2Row7.TabStop = False
                txtCol2Row7.BorderStyle = BorderStyle.FixedSingle

            Case "P"
                lblCol1Row1.Text = "ERAC Services"
                txtCol2Row1.Visible = True

                lblCol1Row2.Text = "Laboratory Services"
                txtCol2Row2.Visible = True

                lblCol1Row3.Text = "Subtotal"
                txtCol2Row3.Visible = True
                txtCol2Row3.ReadOnly = True
                txtCol2Row3.TabStop = False
                txtCol2Row3.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row4.Text = "IRAC Services Estimated Charge"
                txtCol2Row4.Visible = True

                lblCol1Row5.Text = "Markup"
                txtCol2Row5.Visible = True

                lblCol1Row6.Text = "Total"
                txtCol2Row6.Visible = True
                txtCol2Row6.ReadOnly = True
                txtCol2Row6.TabStop = False
                txtCol2Row6.BorderStyle = BorderStyle.FixedSingle

            Case "Q"
                lblCol1Row1.Text = "ERAC Services"
                txtCol2Row1.Visible = True

                lblCol1Row2.Text = "Laboratory Services"
                txtCol2Row2.Visible = True

                lblCol1Row3.Text = "Subtotal"
                txtCol2Row3.Visible = True
                txtCol2Row3.ReadOnly = True
                txtCol2Row3.TabStop = False
                txtCol2Row3.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row4.Text = "Fixed-Fee"
                txtCol2Row4.Visible = True

                lblCol1Row5.Text = "Total"
                txtCol2Row5.Visible = True
                txtCol2Row5.ReadOnly = True
                txtCol2Row5.TabStop = False
                txtCol2Row5.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row6.Text = "IRAC Services Estimated Charge"
                txtCol2Row6.Visible = True

                lblCol1Row7.Text = "Markup"
                txtCol2Row7.Visible = True

                lblCol1Row8.Text = "Grand Total"
                txtCol2Row8.Visible = True
                txtCol2Row8.ReadOnly = True
                txtCol2Row8.TabStop = False
                txtCol2Row8.BorderStyle = BorderStyle.FixedSingle

            Case "R"
                lblCol1Row1.Text = "ERAC Services"
                txtCol2Row1.Visible = True

                lblCol1Row2.Text = "Subcontractor Services"
                txtCol2Row2.Visible = True

                lblCol1Row3.Text = "Estimated Markup"
                txtCol2Row3.Visible = True

                lblCol1Row4.Text = "Total"
                txtCol2Row4.Visible = True
                txtCol2Row4.ReadOnly = True
                txtCol2Row4.TabStop = False
                txtCol2Row4.BorderStyle = BorderStyle.FixedSingle

            Case "S"
                lblCol1Row1.Text = "ERAC Services"
                txtCol2Row1.Visible = True

                lblCol1Row2.Text = "ORC Injection Contractor Services"
                txtCol2Row2.Visible = True

                lblCol1Row3.Text = "Subtotal"
                txtCol2Row3.Visible = True
                txtCol2Row3.ReadOnly = True
                txtCol2Row3.TabStop = False
                txtCol2Row3.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row4.Text = "Fixed-Fee"
                txtCol2Row4.Visible = True

                lblCol1Row5.Text = "Markup"
                txtCol2Row5.Visible = True

                lblCol1Row6.Text = "Total"
                txtCol2Row6.Visible = True
                txtCol2Row6.ReadOnly = True
                txtCol2Row6.TabStop = False
                txtCol2Row6.BorderStyle = BorderStyle.FixedSingle

            Case "T"
                lblCol1Row1.Text = "ERAC Services"
                txtCol2Row1.Visible = True

                lblCol1Row2.Text = "Subcontractor Services"
                txtCol2Row2.Visible = True

                lblCol1Row3.Text = "Estimated Markup"
                txtCol2Row3.Visible = True

                lblCol1Row4.Text = "Total"
                txtCol2Row4.Visible = True
                txtCol2Row4.ReadOnly = True
                txtCol2Row4.TabStop = False
                txtCol2Row4.BorderStyle = BorderStyle.FixedSingle

            Case "U"
                lblCol1Row1.Text = "ERAC Services"
                txtCol2Row1.Visible = True

                lblCol1Row2.Text = "Subcontractor Services"
                txtCol2Row2.Visible = True

                lblCol1Row3.Text = "Subtotal"
                txtCol2Row3.Visible = True
                txtCol2Row3.ReadOnly = True
                txtCol2Row3.TabStop = False
                txtCol2Row3.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row4.Text = "Fixed-Fee"
                txtCol2Row4.Visible = True

                lblCol1Row5.Text = "Markup"
                txtCol2Row5.Visible = True

                lblCol1Row6.Text = "Total"
                txtCol2Row6.Visible = True
                txtCol2Row6.ReadOnly = True
                txtCol2Row6.TabStop = False
                txtCol2Row6.BorderStyle = BorderStyle.FixedSingle

            Case "V"
                lblCol1Row1.Text = "Remediation Contractor Services"
                txtCol2Row1.Visible = True

            Case "W"
                lblCol1Row1.Text = "ERAC Services"
                txtCol2Row1.Visible = True

                lblCol1Row2.Text = "Number of Events"
                txtCol2Row2.Visible = True

                lblCol1Row3.Text = "Total"
                txtCol2Row3.Visible = True
                txtCol2Row3.ReadOnly = True
                txtCol2Row3.TabStop = False
                txtCol2Row3.BorderStyle = BorderStyle.FixedSingle

            Case "X"
                lblCol1Row1.Text = "ERAC System Installation Services"
                txtCol6Row1.Visible = True

                lblCol1Row2.Text = "Subcontractor System Installation Services"
                txtCol6Row2.Visible = True

                lblCol1Row3.Text = "Monthly System Use Rate"
                txtCol2Row3.Visible = True
                lblCol3Row3.Text = "X"
                txtCol4Row3.Visible = True
                lblCol5Row3.Text = "mo. ="
                txtCol6Row3.Visible = True
                txtCol6Row3.ReadOnly = True
                txtCol6Row3.TabStop = False
                txtCol6Row3.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row4.Text = "Monthly O&&M, Sampling"
                txtCol2Row4.Visible = True
                lblCol3Row4.Text = "X"
                txtCol4Row4.Visible = True
                lblCol5Row4.Text = "mo. ="
                txtCol6Row4.Visible = True
                txtCol6Row4.ReadOnly = True
                txtCol6Row4.TabStop = False
                txtCol6Row4.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row5.Text = "Triannual O&M, Sampling, Reporting"
                txtCol2Row5.Visible = True
                lblCol3Row5.Text = "X"
                txtCol4Row5.Visible = True
                lblCol5Row5.Text = "mo. ="
                txtCol6Row5.Visible = True
                txtCol6Row5.ReadOnly = True
                txtCol6Row5.TabStop = False
                txtCol6Row5.BorderStyle = BorderStyle.FixedSingle


                lblCol1Row6.Text = "Estimated Triannual Laboratory Analysis"
                txtCol2Row6.Visible = True
                lblCol3Row6.Text = "X"
                txtCol4Row6.Visible = True
                lblCol5Row6.Text = "mo. ="
                txtCol6Row6.Visible = True
                txtCol6Row6.ReadOnly = True
                txtCol6Row6.TabStop = False
                txtCol6Row6.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row7.Text = "Estimated Electrical/Sewer/Water/Tax Charges"
                txtCol2Row7.Visible = True
                lblCol3Row7.Text = "X"
                txtCol4Row7.Visible = True
                lblCol5Row7.Text = "mo. ="
                txtCol6Row7.Visible = True
                txtCol6Row7.ReadOnly = True
                txtCol6Row7.TabStop = False
                txtCol6Row7.BorderStyle = BorderStyle.FixedSingle


                lblCol1Row8.Text = "Total Eligible Reimbursement"
                txtCol6Row8.Visible = True
                txtCol6Row8.ReadOnly = True
                txtCol6Row8.TabStop = False
                txtCol6Row8.BorderStyle = BorderStyle.FixedSingle


            Case "X3"
                lblCol1Row1.Text = "ERAC System Installation Services"
                txtCol6Row1.Visible = True

                lblCol1Row2.Text = "Subcontractor System Installation Services"
                txtCol6Row2.Visible = True

                lblCol1Row3.Text = "Monthly System Use Rate"
                txtCol2Row3.Visible = True
                lblCol3Row3.Text = "X"
                txtCol4Row3.Visible = True
                txtCol4Row3.Text = "6"
                txtCol4Row3.ReadOnly = True
                lblCol5Row3.Text = "mo. ="
                txtCol6Row3.Visible = True
                txtCol6Row3.ReadOnly = True
                txtCol6Row3.TabStop = False
                txtCol6Row3.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row4.Text = "Monthly O&&M, Sampling"
                txtCol2Row4.Visible = True
                lblCol3Row4.Text = "X"
                txtCol4Row4.Visible = True
                txtCol4Row4.Text = "6"
                txtCol4Row4.ReadOnly = True
                lblCol5Row4.Text = "mo. ="
                txtCol6Row4.Visible = True
                txtCol6Row4.ReadOnly = True
                txtCol6Row4.TabStop = False
                txtCol6Row4.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row5.Text = "Finalization Report"
                txtCol2Row5.Visible = True
                lblCol3Row5.Text = "X"
                txtCol4Row5.Visible = True
                txtCol4Row5.Text = "6"
                txtCol4Row5.ReadOnly = True

                lblCol5Row5.Text = "mo. ="
                txtCol6Row5.Visible = True
                txtCol6Row5.ReadOnly = True
                txtCol6Row5.TabStop = False
                txtCol6Row5.BorderStyle = BorderStyle.FixedSingle


                lblCol1Row6.Text = "Estimated Triannual Laboratory Analysis"
                txtCol2Row6.Visible = True
                lblCol3Row6.Text = "X"
                txtCol4Row6.Visible = True
                lblCol5Row6.Text = "mo. ="
                txtCol6Row6.Visible = True
                txtCol6Row6.ReadOnly = True
                txtCol6Row6.TabStop = False
                txtCol6Row6.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row6.Text = "Estimated Electrical/Sewer/Water/Tax Charges"
                txtCol2Row6.Visible = True
                lblCol3Row6.Text = "X"
                txtCol4Row6.Visible = True
                txtCol4Row6.Text = "6"
                txtCol4Row6.ReadOnly = True
                lblCol5Row6.Text = "mo. ="
                txtCol6Row6.Visible = True
                txtCol6Row6.ReadOnly = True
                txtCol6Row6.TabStop = False
                txtCol6Row6.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row7.Text = "Total Eligible Reimbursement"
                txtCol6Row7.Visible = True
                txtCol6Row7.ReadOnly = True
                txtCol6Row7.TabStop = False
                txtCol6Row7.BorderStyle = BorderStyle.FixedSingle

            Case "X2"
                lblCol1Row1.Text = "Monthly System Use Rate"
                txtCol2Row1.Visible = True
                lblCol3Row1.Text = "X"
                txtCol4Row1.Visible = True
                lblCol5Row1.Text = "mo. ="
                txtCol6Row1.Visible = True
                txtCol6Row1.ReadOnly = True
                txtCol6Row1.TabStop = False
                txtCol6Row1.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row2.Text = "Monthly O&M, Sampling"
                txtCol2Row2.Visible = True
                lblCol3Row2.Text = "X"
                txtCol4Row2.Visible = True
                lblCol5Row2.Text = "mo. ="
                txtCol6Row2.Visible = True
                txtCol6Row2.ReadOnly = True
                txtCol6Row2.TabStop = False
                txtCol6Row2.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row3.Text = "Triannual O&M, Sampling, Reporting"
                txtCol2Row3.Visible = True
                lblCol3Row3.Text = "X"
                txtCol4Row3.Visible = True
                lblCol5Row3.Text = "mo. ="
                txtCol6Row3.Visible = True
                txtCol6Row3.ReadOnly = True
                txtCol6Row3.TabStop = False
                txtCol6Row3.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row4.Text = "Total"
                txtCol6Row4.Visible = True
                txtCol6Row4.ReadOnly = True
                txtCol6Row4.TabStop = False
                txtCol6Row4.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row5.Text = "Estimated Triannual Laboratory Analysis"
                txtCol2Row5.Visible = True
                lblCol3Row5.Text = "X"
                txtCol4Row5.Visible = True
                lblCol5Row5.Text = "mo. ="
                txtCol6Row5.Visible = True
                txtCol6Row5.ReadOnly = True
                txtCol6Row5.TabStop = False
                txtCol6Row5.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row6.Text = "Estimated Electrical/Sewer/Water/Tax Charges"
                txtCol2Row6.Visible = True
                lblCol3Row6.Text = "X"
                txtCol4Row6.Visible = True
                lblCol5Row6.Text = "mo. ="
                txtCol6Row6.Visible = True
                txtCol6Row6.ReadOnly = True
                txtCol6Row6.TabStop = False
                txtCol6Row6.BorderStyle = BorderStyle.FixedSingle

                lblCol1Row7.Text = "Markup"
                txtCol6Row7.Visible = True

                lblCol1Row8.Text = "Total Eligible Reimbursement"
                txtCol6Row8.Visible = True
                txtCol6Row8.ReadOnly = True
                txtCol6Row8.TabStop = False
                txtCol6Row8.BorderStyle = BorderStyle.FixedSingle

            Case "Y"
                lblCol1Row1.Text = "ERAC Vacuum Services"
                txtCol2Row1.Visible = True
                lblCol3Row1.Text = "per"
                txtCol4Row1.Visible = True
                lblCol5Row1.Text = "events"

                lblCol1Row2.Text = "Vacuum Contractor Services"
                txtCol2Row2.Visible = True
                lblCol3Row2.Text = "per"
                txtCol4Row2.Visible = True
                txtCol4Row2.ReadOnly = True
                txtCol4Row2.TabStop = False
                txtCol4Row2.BorderStyle = BorderStyle.FixedSingle
                lblCol5Row2.Text = "events"

                lblCol1Row3.Text = "Subtotal"
                txtCol2Row3.Visible = True
                txtCol2Row3.ReadOnly = True
                txtCol2Row3.TabStop = False
                txtCol2Row3.BorderStyle = BorderStyle.FixedSingle
                lblCol3Row3.Text = "per"
                txtCol4Row3.Visible = True
                txtCol4Row3.ReadOnly = True
                txtCol4Row3.TabStop = False
                txtCol4Row3.BorderStyle = BorderStyle.FixedSingle
                lblCol5Row3.Text = "events"

                lblCol1Row4.Text = "Total"
                txtCol2Row4.Visible = True
                txtCol2Row4.ReadOnly = True
                txtCol2Row4.TabStop = False
                txtCol2Row4.BorderStyle = BorderStyle.FixedSingle
                lblCol3Row4.Text = "per"
                txtCol4Row4.Visible = True
                txtCol4Row4.ReadOnly = True
                txtCol4Row4.TabStop = False
                txtCol4Row4.BorderStyle = BorderStyle.FixedSingle
                lblCol5Row4.Text = "vacuum events"
                lblCol5Row4.Width = 80
            Case "Z"
                lblCol1Row1.Text = "Third Party Settlement"
                txtCol2Row1.Visible = True
            Case Else
                If Not bolloading Then
                    lblCol1Row1.Text = "Unsupported Cost Format"
                End If
        End Select

    End Sub

    Private Sub RecalcAll_old()

        Select Case CostFormatType

            Case "A"
                txtCol2Row4.Text = FormatNumber(CDbl(txtCol2Row1.Text) + CDbl(txtCol2Row2.Text) + CDbl(txtCol2Row3.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol2Row4.Text)


            Case "B"
                txtCol2Row3.Text = FormatNumber(CDbl(txtCol2Row1.Text) + CDbl(txtCol2Row2.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row6.Text = FormatNumber(CDbl(txtCol2Row3.Text) + CDbl(txtCol2Row4.Text) + CDbl(txtCol2Row5.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol2Row6.Text)

            Case "C"
                GrandTotal = CDbl(txtCol2Row1.Text)

            Case "CR"
                GrandTotal = CDbl(txtCol2Row1.Text)

            Case "D"
                txtCol2Row4.Text = FormatNumber(CDbl(txtCol2Row1.Text) + CDbl(txtCol2Row2.Text) + CDbl(txtCol2Row3.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol2Row4.Text)

            Case "E"
                txtCol2Row4.Text = FormatNumber(CDbl(txtCol2Row1.Text) + CDbl(txtCol2Row2.Text) + CDbl(txtCol2Row3.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol2Row4.Text)

            Case "F"
                txtCol2Row3.Text = FormatNumber(CDbl(txtCol2Row1.Text) + CDbl(txtCol2Row2.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row6.Text = FormatNumber(CDbl(txtCol2Row3.Text) + CDbl(txtCol2Row4.Text) + CDbl(txtCol2Row5.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol2Row6.Text)

            Case "G"
                txtCol2Row3.Text = FormatNumber(CDbl(txtCol2Row1.Text) + CDbl(txtCol2Row2.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row5.Text = FormatNumber(CDbl(txtCol2Row3.Text) + CDbl(txtCol2Row4.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row7.Text = FormatNumber(CDbl(txtCol2Row5.Text) * CDbl(txtCol2Row6.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol2Row7.Text)

            Case "H"
                txtCol2Row3.Text = FormatNumber(CDbl(txtCol2Row1.Text) + CDbl(txtCol2Row2.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row5.Text = FormatNumber(CDbl(txtCol2Row3.Text) * CDbl(txtCol2Row4.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol2Row5.Text)

            Case "I"
                GrandTotal = CDbl(txtCol2Row1.Text)

            Case "IR"
                GrandTotal = CDbl(txtCol2Row1.Text)

            Case "J"
                txtCol2Row3.Text = FormatNumber(CDbl(txtCol2Row1.Text) + CDbl(txtCol2Row2.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row6.Text = FormatNumber(CDbl(txtCol2Row3.Text) + CDbl(txtCol2Row4.Text) + CDbl(txtCol2Row5.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol2Row6.Text)

            Case "K"
                txtCol2Row3.Text = FormatNumber(CDbl(txtCol2Row1.Text) * CDbl(txtCol2Row2.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol2Row3.Text)

            Case "L"
                txtCol2Row3.Text = FormatNumber(CDbl(txtCol2Row1.Text) + CDbl(txtCol2Row2.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row5.Text = FormatNumber(CDbl(txtCol2Row3.Text) * CDbl(txtCol2Row4.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol2Row5.Text)

            Case "M"
                GrandTotal = CDbl(txtCol2Row1.Text)

            Case "N"
                txtCol4Row2.Text = txtCol4Row1.Text
                txtCol4Row3.Text = txtCol4Row1.Text
                txtCol4Row4.Text = txtCol4Row1.Text

                txtCol2Row3.Text = FormatNumber(CDbl(txtCol2Row1.Text) + CDbl(txtCol2Row2.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row4.Text = FormatNumber(CDbl(txtCol2Row3.Text) * CDbl(txtCol4Row3.Text), 2, TriState.False, TriState.False, TriState.True)

                txtCol2Row7.Text = FormatNumber(CDbl(txtCol2Row5.Text) + CDbl(txtCol2Row6.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row8.Text = FormatNumber(CDbl(txtCol2Row4.Text) + CDbl(txtCol2Row7.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol2Row8.Text)

            Case "O"
                'txtCol2Row3.Text = FormatNumber(CDbl(txtCol2Row1.Text) * CDbl(txtCol2Row2.Text), 2, TriState.False, TriState.False, TriState.True)
                'txtCol2Row5.Text = FormatNumber(CDbl(txtCol2Row3.Text) + CDbl(txtCol2Row4.Text), 2, TriState.False, TriState.False, TriState.True)
                'txtCol2Row7.Text = FormatNumber(CDbl(txtCol2Row5.Text) + CDbl(txtCol2Row6.Text), 2, TriState.False, TriState.False, TriState.True)
                'GrandTotal = CDbl(txtCol2Row7.Text)
                txtCol2Row3.Text = FormatNumber(CDbl(txtCol2Row1.Text) * CDbl(txtCol2Row2.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row6.Text = FormatNumber(CDbl(txtCol2Row4.Text) + CDbl(txtCol2Row5.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row7.Text = FormatNumber(CDbl(txtCol2Row3.Text) + CDbl(txtCol2Row6.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol2Row7.Text)

            Case "P"
                txtCol2Row3.Text = FormatNumber(CDbl(txtCol2Row1.Text) + CDbl(txtCol2Row2.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row6.Text = FormatNumber(CDbl(txtCol2Row3.Text) + CDbl(txtCol2Row4.Text) + CDbl(txtCol2Row5.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol2Row6.Text)

            Case "Q"
                txtCol2Row3.Text = FormatNumber(CDbl(txtCol2Row1.Text) + CDbl(txtCol2Row2.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row5.Text = FormatNumber(CDbl(txtCol2Row3.Text) + CDbl(txtCol2Row4.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row8.Text = FormatNumber(CDbl(txtCol2Row5.Text) + CDbl(txtCol2Row6.Text) + CDbl(txtCol2Row7.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol2Row8.Text)

            Case "R"
                txtCol2Row4.Text = FormatNumber(CDbl(txtCol2Row1.Text) + CDbl(txtCol2Row2.Text) + CDbl(txtCol2Row3.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol2Row4.Text)

            Case "S"
                txtCol2Row3.Text = FormatNumber(CDbl(txtCol2Row1.Text) + CDbl(txtCol2Row2.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row6.Text = FormatNumber(CDbl(txtCol2Row3.Text) + CDbl(txtCol2Row4.Text) + CDbl(txtCol2Row5.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol2Row6.Text)

            Case "T"
                txtCol2Row4.Text = FormatNumber(CDbl(txtCol2Row1.Text) + CDbl(txtCol2Row2.Text) + CDbl(txtCol2Row3.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol2Row4.Text)

            Case "U"
                txtCol2Row3.Text = FormatNumber(CDbl(txtCol2Row1.Text) + CDbl(txtCol2Row2.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row6.Text = FormatNumber(CDbl(txtCol2Row3.Text) + CDbl(txtCol2Row4.Text) + CDbl(txtCol2Row5.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol2Row6.Text)

            Case "V"
                GrandTotal = CDbl(txtCol2Row1.Text)

            Case "W"
                txtCol2Row3.Text = FormatNumber(CDbl(txtCol2Row1.Text) * CDbl(txtCol2Row2.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol2Row3.Text)

            Case "X"
                txtCol6Row3.Text = FormatNumber(CDbl(txtCol2Row3.Text) * CDbl(txtCol4Row3.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol6Row4.Text = FormatNumber(CDbl(txtCol2Row4.Text) * CDbl(txtCol4Row4.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol6Row5.Text = FormatNumber(CDbl(txtCol2Row5.Text) * CDbl(txtCol4Row5.Text), 2, TriState.False, TriState.False, TriState.True)

                txtCol6Row6.Text = FormatNumber(CDbl(txtCol2Row6.Text) * CDbl(txtCol4Row6.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol6Row7.Text = FormatNumber(CDbl(txtCol2Row7.Text) * CDbl(txtCol4Row7.Text), 2, TriState.False, TriState.False, TriState.True)

                txtCol6Row8.Text = FormatNumber(CDbl(txtCol6Row1.Text) + CDbl(txtCol6Row2.Text) + CDbl(txtCol6Row3.Text) + CDbl(txtCol6Row4.Text) + CDbl(txtCol6Row5.Text) + CDbl(txtCol6Row6.Text) + CDbl(txtCol6Row7.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol6Row8.Text)

            Case "X3"
                txtCol6Row3.Text = FormatNumber(CDbl(txtCol2Row3.Text) * CDbl(txtCol4Row3.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol6Row4.Text = FormatNumber(CDbl(txtCol2Row4.Text) * CDbl(txtCol4Row4.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol6Row5.Text = FormatNumber(CDbl(txtCol2Row5.Text) * CDbl(txtCol4Row5.Text), 2, TriState.False, TriState.False, TriState.True)

                txtCol6Row6.Text = FormatNumber(CDbl(txtCol2Row6.Text) * CDbl(txtCol4Row6.Text), 2, TriState.False, TriState.False, TriState.True)

                txtCol6Row7.Text = FormatNumber(CDbl(txtCol6Row1.Text) + CDbl(txtCol6Row2.Text) + CDbl(txtCol6Row3.Text) + CDbl(txtCol6Row4.Text) + CDbl(txtCol6Row5.Text) + CDbl(txtCol6Row6.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol6Row7.Text)


            Case "X2"
                txtCol6Row1.Text = FormatNumber(CDbl(txtCol2Row1.Text) * CDbl(txtCol4Row1.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol6Row2.Text = FormatNumber(CDbl(txtCol2Row2.Text) * CDbl(txtCol4Row2.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol6Row3.Text = FormatNumber(CDbl(txtCol2Row3.Text) * CDbl(txtCol4Row3.Text), 2, TriState.False, TriState.False, TriState.True)

                txtCol6Row4.Text = FormatNumber(CDbl(txtCol6Row1.Text) + CDbl(txtCol6Row2.Text) + CDbl(txtCol6Row3.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol6Row5.Text = FormatNumber(CDbl(txtCol2Row5.Text) * CDbl(txtCol4Row5.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol6Row6.Text = FormatNumber(CDbl(txtCol2Row6.Text) * CDbl(txtCol4Row6.Text), 2, TriState.False, TriState.False, TriState.True)

                txtCol6Row8.Text = FormatNumber(CDbl(txtCol6Row4.Text) + CDbl(txtCol6Row5.Text) + CDbl(txtCol6Row6.Text) + CDbl(txtCol6Row7.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol6Row8.Text)

            Case "Y"
                txtCol4Row2.Text = txtCol4Row1.Text
                txtCol4Row3.Text = txtCol4Row1.Text
                txtCol4Row4.Text = txtCol4Row1.Text

                txtCol2Row3.Text = FormatNumber(CDbl(txtCol2Row1.Text) + CDbl(txtCol2Row2.Text), 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row4.Text = FormatNumber(CDbl(txtCol2Row3.Text) * CDbl(txtCol4Row3.Text), 2, TriState.False, TriState.False, TriState.True)
                GrandTotal = CDbl(txtCol2Row4.Text)

            Case "Z"
                GrandTotal = CDbl(txtCol2Row1.Text)

            Case Else
                GrandTotal = 0
        End Select

    End Sub

    Private Sub LoadCommitment_old()


        Select Case oLocalFinancialCommitment.Case_Letter

            Case "A"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.ERACServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row2.Text = FormatNumber(oLocalFinancialCommitment.LaboratoryServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row3.Text = FormatNumber(oLocalFinancialCommitment.Markup, 2, TriState.False, TriState.False, TriState.True)

            Case "B"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.ERACServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row2.Text = FormatNumber(oLocalFinancialCommitment.LaboratoryServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row4.Text = FormatNumber(oLocalFinancialCommitment.FixedFee, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row5.Text = FormatNumber(oLocalFinancialCommitment.Markup, 2, TriState.False, TriState.False, TriState.True)

            Case "C"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.ERACServices, 2, TriState.False, TriState.False, TriState.True)

            Case "CR"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.CostRecovery, 2, TriState.False, TriState.False, TriState.True)

            Case "D"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.ERACServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row2.Text = FormatNumber(oLocalFinancialCommitment.FixedFee, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row3.Text = FormatNumber(oLocalFinancialCommitment.Markup, 2, TriState.False, TriState.False, TriState.True)

            Case "E"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.ERACServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row2.Text = FormatNumber(oLocalFinancialCommitment.LaboratoryServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row3.Text = FormatNumber(oLocalFinancialCommitment.Markup, 2, TriState.False, TriState.False, TriState.True)

            Case "F"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.ERACServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row2.Text = FormatNumber(oLocalFinancialCommitment.LaboratoryServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row4.Text = FormatNumber(oLocalFinancialCommitment.FixedFee, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row5.Text = FormatNumber(oLocalFinancialCommitment.Markup, 2, TriState.False, TriState.False, TriState.True)

            Case "G"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.ERACServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row2.Text = FormatNumber(oLocalFinancialCommitment.LaboratoryServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row4.Text = FormatNumber(oLocalFinancialCommitment.FixedFee, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row6.Text = FormatNumber(oLocalFinancialCommitment.NumberofEvents, -1, TriState.False, TriState.False, TriState.True)

            Case "H"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.ERACServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row2.Text = FormatNumber(oLocalFinancialCommitment.LaboratoryServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row4.Text = FormatNumber(oLocalFinancialCommitment.NumberofEvents, -1, TriState.False, TriState.False, TriState.True)

            Case "I"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.WellAbandonment, 2, TriState.False, TriState.False, TriState.True)

            Case "IR"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.IRACServicesEstimate, 2, TriState.False, TriState.False, TriState.True)

            Case "J"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.ERACServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row2.Text = FormatNumber(oLocalFinancialCommitment.WellAbandonment, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row4.Text = FormatNumber(oLocalFinancialCommitment.FixedFee, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row5.Text = FormatNumber(oLocalFinancialCommitment.Markup, 2, TriState.False, TriState.False, TriState.True)

            Case "K"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.FreeProductRecovery, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row2.Text = FormatNumber(oLocalFinancialCommitment.NumberofEvents, -1, TriState.False, TriState.False, TriState.True)

            Case "L"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.ERACServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row2.Text = FormatNumber(oLocalFinancialCommitment.VacuumContServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row4.Text = FormatNumber(oLocalFinancialCommitment.NumberofEvents, -1, TriState.False, TriState.False, TriState.True)

            Case "M"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.PTTTesting, 2, TriState.False, TriState.False, TriState.True)

            Case "N"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.ERACVacuum, 2, TriState.False, TriState.False, TriState.True)
                txtCol4Row1.Text = oLocalFinancialCommitment.ERACVacuumCnt.ToString
                txtCol2Row2.Text = FormatNumber(oLocalFinancialCommitment.VacuumContServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol4Row2.Text = oLocalFinancialCommitment.VacuumContServicesCnt.ToString
                txtCol2Row5.Text = FormatNumber(oLocalFinancialCommitment.ERACSampling, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row6.Text = FormatNumber(oLocalFinancialCommitment.LaboratoryServices, 2, TriState.False, TriState.False, TriState.True)

                txtCol4Row2.Text = txtCol4Row1.Text
                txtCol4Row3.Text = txtCol4Row1.Text
                txtCol4Row4.Text = txtCol4Row1.Text

            Case "O"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.FreeProductRecovery, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row2.Text = oLocalFinancialCommitment.NumberofEvents.ToString
                txtCol2Row4.Text = FormatNumber(oLocalFinancialCommitment.ERACSampling, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row5.Text = FormatNumber(oLocalFinancialCommitment.LaboratoryServices, 2, TriState.False, TriState.False, TriState.True)

            Case "P"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.ERACServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row2.Text = FormatNumber(oLocalFinancialCommitment.LaboratoryServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row4.Text = FormatNumber(oLocalFinancialCommitment.IRACServicesEstimate, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row5.Text = FormatNumber(oLocalFinancialCommitment.Markup, 2, TriState.False, TriState.False, TriState.True)

            Case "Q"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.ERACServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row2.Text = FormatNumber(oLocalFinancialCommitment.LaboratoryServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row4.Text = FormatNumber(oLocalFinancialCommitment.FixedFee, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row6.Text = FormatNumber(oLocalFinancialCommitment.IRACServicesEstimate, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row7.Text = FormatNumber(oLocalFinancialCommitment.Markup, 2, TriState.False, TriState.False, TriState.True)

            Case "R"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.ERACServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row2.Text = FormatNumber(oLocalFinancialCommitment.SubContractorSvcs, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row3.Text = FormatNumber(oLocalFinancialCommitment.Markup, 2, TriState.False, TriState.False, TriState.True)

            Case "S"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.ERACServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row2.Text = FormatNumber(oLocalFinancialCommitment.ORCContractorSvcs, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row4.Text = FormatNumber(oLocalFinancialCommitment.FixedFee, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row5.Text = FormatNumber(oLocalFinancialCommitment.Markup, 2, TriState.False, TriState.False, TriState.True)

            Case "T"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.ERACServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row2.Text = FormatNumber(oLocalFinancialCommitment.SubContractorSvcs, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row3.Text = FormatNumber(oLocalFinancialCommitment.Markup, 2, TriState.False, TriState.False, TriState.True)

            Case "U"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.ERACServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row2.Text = FormatNumber(oLocalFinancialCommitment.SubContractorSvcs, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row4.Text = FormatNumber(oLocalFinancialCommitment.FixedFee, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row5.Text = FormatNumber(oLocalFinancialCommitment.Markup, 2, TriState.False, TriState.False, TriState.True)

            Case "V"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.REMContractorSvcs, 2, TriState.False, TriState.False, TriState.True)

            Case "W"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.ERACServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol2Row2.Text = oLocalFinancialCommitment.NumberofEvents.ToString

            Case "X"
                txtCol6Row1.Text = FormatNumber(oLocalFinancialCommitment.PreInstallSetup, 2, TriState.False, TriState.False, TriState.True)
                txtCol6Row2.Text = FormatNumber(oLocalFinancialCommitment.InstallSetup, 2, TriState.False, TriState.False, TriState.True)

                txtCol2Row3.Text = FormatNumber(oLocalFinancialCommitment.MonthlySystemUse, 2, TriState.False, TriState.False, TriState.True)
                txtCol4Row3.Text = oLocalFinancialCommitment.MonthlySystemUseCnt.ToString
                txtCol2Row4.Text = FormatNumber(oLocalFinancialCommitment.MonthlyOMSampling, 2, TriState.False, TriState.False, TriState.True)
                txtCol4Row4.Text = oLocalFinancialCommitment.MonthlyOMSamplingCnt.ToString
                txtCol2Row5.Text = FormatNumber(oLocalFinancialCommitment.TriAnnualOMSampling, 2, TriState.False, TriState.False, TriState.True)
                txtCol4Row5.Text = oLocalFinancialCommitment.TriAnnualOMSamplingCnt.ToString
                txtCol2Row6.Text = FormatNumber(oLocalFinancialCommitment.EstimateTriAnnualLab, 2, TriState.False, TriState.False, TriState.True)
                txtCol4Row6.Text = oLocalFinancialCommitment.EstimateTriAnnualLabCnt.ToString
                txtCol2Row7.Text = FormatNumber(oLocalFinancialCommitment.EstimateUtilities, 2, TriState.False, TriState.False, TriState.True)
                txtCol4Row7.Text = oLocalFinancialCommitment.EstimateUtilitiesCnt.ToString

            Case "X3"
                txtCol6Row1.Text = FormatNumber(oLocalFinancialCommitment.PreInstallSetup, 2, TriState.False, TriState.False, TriState.True)
                txtCol6Row2.Text = FormatNumber(oLocalFinancialCommitment.InstallSetup, 2, TriState.False, TriState.False, TriState.True)

                txtCol2Row3.Text = FormatNumber(oLocalFinancialCommitment.MonthlySystemUse, 2, TriState.False, TriState.False, TriState.True)
                txtCol4Row3.Text = oLocalFinancialCommitment.MonthlySystemUseCnt.ToString
                txtCol2Row4.Text = FormatNumber(oLocalFinancialCommitment.MonthlyOMSampling, 2, TriState.False, TriState.False, TriState.True)
                txtCol4Row4.Text = oLocalFinancialCommitment.MonthlyOMSamplingCnt.ToString
                txtCol2Row5.Text = FormatNumber(oLocalFinancialCommitment.TriAnnualOMSampling, 2, TriState.False, TriState.False, TriState.True)
                txtCol4Row5.Text = oLocalFinancialCommitment.TriAnnualOMSamplingCnt.ToString
                txtCol2Row6.Text = FormatNumber(oLocalFinancialCommitment.EstimateUtilities, 2, TriState.False, TriState.False, TriState.True)
                txtCol4Row6.Text = oLocalFinancialCommitment.EstimateUtilitiesCnt.ToString

            Case "X2"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.MonthlySystemUse, 2, TriState.False, TriState.False, TriState.True)
                txtCol4Row1.Text = oLocalFinancialCommitment.MonthlySystemUseCnt.ToString
                txtCol2Row2.Text = FormatNumber(oLocalFinancialCommitment.MonthlyOMSampling, 2, TriState.False, TriState.False, TriState.True)
                txtCol4Row2.Text = oLocalFinancialCommitment.MonthlyOMSamplingCnt.ToString
                txtCol2Row3.Text = FormatNumber(oLocalFinancialCommitment.TriAnnualOMSampling, 2, TriState.False, TriState.False, TriState.True)
                txtCol4Row3.Text = oLocalFinancialCommitment.TriAnnualOMSamplingCnt.ToString
                txtCol2Row5.Text = FormatNumber(oLocalFinancialCommitment.EstimateTriAnnualLab, 2, TriState.False, TriState.False, TriState.True)
                txtCol4Row5.Text = oLocalFinancialCommitment.EstimateTriAnnualLabCnt.ToString
                txtCol2Row6.Text = FormatNumber(oLocalFinancialCommitment.EstimateUtilities, 2, TriState.False, TriState.False, TriState.True)
                txtCol4Row6.Text = oLocalFinancialCommitment.EstimateUtilitiesCnt.ToString
                txtCol6Row7.Text = FormatNumber(oLocalFinancialCommitment.Markup, 2, TriState.False, TriState.False, TriState.True)

            Case "Y"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.ERACVacuum, 2, TriState.False, TriState.False, TriState.True)
                txtCol4Row1.Text = oLocalFinancialCommitment.ERACVacuumCnt.ToString
                txtCol2Row2.Text = FormatNumber(oLocalFinancialCommitment.VacuumContServices, 2, TriState.False, TriState.False, TriState.True)
                txtCol4Row2.Text = oLocalFinancialCommitment.VacuumContServicesCnt

                txtCol4Row2.Text = txtCol4Row1.Text
                txtCol4Row3.Text = txtCol4Row1.Text
                txtCol4Row4.Text = txtCol4Row1.Text

            Case "Z"
                txtCol2Row1.Text = FormatNumber(oLocalFinancialCommitment.ThirdPartySettlement, 2, TriState.False, TriState.False, TriState.True)

        End Select

    End Sub

    Private Sub PushToObject_old()
        Select Case CostFormatType

            Case "A"
                oLocalFinancialCommitment.ERACServices = CDbl(txtCol2Row1.Text)
                oLocalFinancialCommitment.LaboratoryServices = CDbl(txtCol2Row2.Text)
                oLocalFinancialCommitment.Markup = CDbl(txtCol2Row3.Text)

            Case "B"
                oLocalFinancialCommitment.ERACServices = CDbl(txtCol2Row1.Text)
                oLocalFinancialCommitment.LaboratoryServices = CDbl(txtCol2Row2.Text)
                oLocalFinancialCommitment.FixedFee = CDbl(txtCol2Row4.Text)
                oLocalFinancialCommitment.Markup = CDbl(txtCol2Row5.Text)

            Case "C"
                oLocalFinancialCommitment.ERACServices = CDbl(txtCol2Row1.Text)

            Case "CR"
                oLocalFinancialCommitment.CostRecovery = txtCol2Row1.Text

            Case "D"
                oLocalFinancialCommitment.ERACServices = txtCol2Row1.Text
                oLocalFinancialCommitment.FixedFee = txtCol2Row2.Text
                oLocalFinancialCommitment.Markup = txtCol2Row3.Text

            Case "E"
                oLocalFinancialCommitment.ERACServices = txtCol2Row1.Text
                oLocalFinancialCommitment.LaboratoryServices = txtCol2Row2.Text
                oLocalFinancialCommitment.Markup = txtCol2Row3.Text

            Case "F"
                oLocalFinancialCommitment.ERACServices = txtCol2Row1.Text
                oLocalFinancialCommitment.LaboratoryServices = txtCol2Row2.Text
                oLocalFinancialCommitment.FixedFee = txtCol2Row4.Text
                oLocalFinancialCommitment.Markup = txtCol2Row5.Text

            Case "G"
                oLocalFinancialCommitment.ERACServices = txtCol2Row1.Text
                oLocalFinancialCommitment.LaboratoryServices = txtCol2Row2.Text
                oLocalFinancialCommitment.FixedFee = txtCol2Row4.Text
                oLocalFinancialCommitment.NumberofEvents = txtCol2Row6.Text

            Case "H"
                oLocalFinancialCommitment.ERACServices = txtCol2Row1.Text
                oLocalFinancialCommitment.LaboratoryServices = txtCol2Row2.Text
                oLocalFinancialCommitment.NumberofEvents = txtCol2Row4.Text

            Case "I"
                oLocalFinancialCommitment.WellAbandonment = txtCol2Row1.Text

            Case "IR"
                oLocalFinancialCommitment.IRACServicesEstimate = txtCol2Row1.Text

            Case "J"
                oLocalFinancialCommitment.ERACServices = txtCol2Row1.Text
                oLocalFinancialCommitment.WellAbandonment = txtCol2Row2.Text
                oLocalFinancialCommitment.FixedFee = txtCol2Row4.Text
                oLocalFinancialCommitment.Markup = txtCol2Row5.Text

            Case "K"
                oLocalFinancialCommitment.FreeProductRecovery = txtCol2Row1.Text
                oLocalFinancialCommitment.NumberofEvents = txtCol2Row2.Text

            Case "L"
                oLocalFinancialCommitment.ERACServices = txtCol2Row1.Text
                oLocalFinancialCommitment.VacuumContServices = txtCol2Row2.Text
                oLocalFinancialCommitment.NumberofEvents = txtCol2Row4.Text

            Case "M"
                oLocalFinancialCommitment.PTTTesting = txtCol2Row1.Text

            Case "N"
                oLocalFinancialCommitment.ERACVacuum = txtCol2Row1.Text
                oLocalFinancialCommitment.ERACVacuumCnt = txtCol4Row1.Text
                oLocalFinancialCommitment.VacuumContServices = txtCol2Row2.Text
                oLocalFinancialCommitment.VacuumContServicesCnt = txtCol4Row2.Text
                oLocalFinancialCommitment.ERACSampling = txtCol2Row5.Text
                oLocalFinancialCommitment.LaboratoryServices = txtCol2Row6.Text

            Case "O"
                oLocalFinancialCommitment.FreeProductRecovery = txtCol2Row1.Text
                oLocalFinancialCommitment.NumberofEvents = txtCol2Row2.Text
                oLocalFinancialCommitment.ERACSampling = txtCol2Row4.Text
                oLocalFinancialCommitment.LaboratoryServices = txtCol2Row5.Text

            Case "P"
                oLocalFinancialCommitment.ERACServices = txtCol2Row1.Text
                oLocalFinancialCommitment.LaboratoryServices = txtCol2Row2.Text
                oLocalFinancialCommitment.IRACServicesEstimate = txtCol2Row4.Text
                oLocalFinancialCommitment.Markup = txtCol2Row5.Text

            Case "Q"
                oLocalFinancialCommitment.ERACServices = txtCol2Row1.Text
                oLocalFinancialCommitment.LaboratoryServices = txtCol2Row2.Text
                oLocalFinancialCommitment.FixedFee = txtCol2Row4.Text
                oLocalFinancialCommitment.IRACServicesEstimate = txtCol2Row6.Text
                oLocalFinancialCommitment.Markup = txtCol2Row7.Text

            Case "R"
                oLocalFinancialCommitment.ERACServices = txtCol2Row1.Text
                oLocalFinancialCommitment.SubContractorSvcs = txtCol2Row2.Text
                oLocalFinancialCommitment.Markup = txtCol2Row3.Text

            Case "S"
                oLocalFinancialCommitment.ERACServices = txtCol2Row1.Text
                oLocalFinancialCommitment.ORCContractorSvcs = txtCol2Row2.Text
                oLocalFinancialCommitment.FixedFee = txtCol2Row4.Text
                oLocalFinancialCommitment.Markup = txtCol2Row5.Text

            Case "T"
                oLocalFinancialCommitment.ERACServices = txtCol2Row1.Text
                oLocalFinancialCommitment.SubContractorSvcs = txtCol2Row2.Text
                oLocalFinancialCommitment.Markup = txtCol2Row3.Text

            Case "U"
                oLocalFinancialCommitment.ERACServices = txtCol2Row1.Text
                oLocalFinancialCommitment.SubContractorSvcs = txtCol2Row2.Text
                oLocalFinancialCommitment.FixedFee = txtCol2Row4.Text
                oLocalFinancialCommitment.Markup = txtCol2Row5.Text

            Case "V"
                oLocalFinancialCommitment.REMContractorSvcs = txtCol2Row1.Text

            Case "W"
                oLocalFinancialCommitment.ERACServices = txtCol2Row1.Text
                oLocalFinancialCommitment.NumberofEvents = txtCol2Row2.Text

            Case "X"
                oLocalFinancialCommitment.PreInstallSetup = CDbl(txtCol6Row1.Text)
                oLocalFinancialCommitment.InstallSetup = CDbl(txtCol6Row2.Text)
                oLocalFinancialCommitment.MonthlySystemUse = txtCol2Row3.Text
                oLocalFinancialCommitment.MonthlySystemUseCnt = txtCol4Row3.Text
                oLocalFinancialCommitment.MonthlyOMSampling = txtCol2Row4.Text
                oLocalFinancialCommitment.MonthlyOMSamplingCnt = txtCol4Row4.Text
                oLocalFinancialCommitment.TriAnnualOMSampling = txtCol2Row5.Text
                oLocalFinancialCommitment.TriAnnualOMSamplingCnt = txtCol4Row5.Text
                oLocalFinancialCommitment.EstimateTriAnnualLab = txtCol2Row6.Text
                oLocalFinancialCommitment.EstimateTriAnnualLabCnt = txtCol4Row6.Text
                oLocalFinancialCommitment.EstimateUtilities = txtCol2Row7.Text
                oLocalFinancialCommitment.EstimateUtilitiesCnt = txtCol4Row7.Text


            Case "X3"
                oLocalFinancialCommitment.PreInstallSetup = CDbl(txtCol6Row1.Text)
                oLocalFinancialCommitment.InstallSetup = CDbl(txtCol6Row2.Text)
                oLocalFinancialCommitment.MonthlySystemUse = txtCol2Row3.Text
                oLocalFinancialCommitment.MonthlySystemUseCnt = txtCol4Row3.Text
                oLocalFinancialCommitment.MonthlyOMSampling = txtCol2Row4.Text
                oLocalFinancialCommitment.MonthlyOMSamplingCnt = txtCol4Row4.Text
                oLocalFinancialCommitment.TriAnnualOMSampling = txtCol2Row5.Text
                oLocalFinancialCommitment.TriAnnualOMSamplingCnt = txtCol4Row5.Text
                oLocalFinancialCommitment.EstimateUtilities = txtCol2Row6.Text
                oLocalFinancialCommitment.EstimateUtilitiesCnt = txtCol4Row6.Text

            Case "X2"
                oLocalFinancialCommitment.MonthlySystemUse = txtCol2Row1.Text
                oLocalFinancialCommitment.MonthlySystemUseCnt = txtCol4Row1.Text
                oLocalFinancialCommitment.MonthlyOMSampling = txtCol2Row2.Text
                oLocalFinancialCommitment.MonthlyOMSamplingCnt = txtCol4Row2.Text
                oLocalFinancialCommitment.TriAnnualOMSampling = txtCol2Row3.Text
                oLocalFinancialCommitment.TriAnnualOMSamplingCnt = txtCol4Row3.Text
                oLocalFinancialCommitment.EstimateTriAnnualLab = txtCol2Row5.Text
                oLocalFinancialCommitment.EstimateTriAnnualLabCnt = txtCol4Row5.Text
                oLocalFinancialCommitment.EstimateUtilities = txtCol2Row6.Text
                oLocalFinancialCommitment.EstimateUtilitiesCnt = txtCol4Row6.Text
                oLocalFinancialCommitment.Markup = txtCol6Row7.Text

            Case "Y"
                oLocalFinancialCommitment.ERACVacuum = txtCol2Row1.Text
                oLocalFinancialCommitment.ERACVacuumCnt = txtCol4Row1.Text
                oLocalFinancialCommitment.VacuumContServices = txtCol2Row2.Text
                oLocalFinancialCommitment.VacuumContServicesCnt = txtCol4Row2.Text

            Case "Z"
                oLocalFinancialCommitment.ThirdPartySettlement = txtCol2Row1.Text

        End Select

    End Sub


    Private Sub ResetData(Optional ByVal ClearObject As Boolean = True)

        lblCol5Row4.Width = 40

        If ClearObject Then

            With oLocalFinancialCommitment

                .ERACSampling = 0
                .ERACServices = 0
                .ERACVacuum = 0
                .ERACVacuumCnt = 0
                .FixedFee = 0
                .FreeProductRecovery = 0
                .IRACServicesEstimate = 0
                .LaboratoryServices = 0
                .NumberofEvents = 0
                .ORCContractorSvcs = 0
                .InstallSetup = 0
                .PTTTesting = 0
                .REMContractorSvcs = 0
                .SubContractorSvcs = 0
                .VacuumContServices = 0
                .VacuumContServicesCnt = 0
                .WellAbandonment = 0
                .PreInstallSetup = 0
                .MonthlySystemUse = 0
                .MonthlySystemUseCnt = 0
                .MonthlyOMSampling = 0
                .MonthlyOMSamplingCnt = 0
                .TriAnnualOMSampling = 0
                .TriAnnualOMSamplingCnt = 0
                .EstimateTriAnnualLab = 0
                .EstimateTriAnnualLabCnt = 0
                .EstimateUtilities = 0
                .EstimateUtilitiesCnt = 0
                .ThirdPartySettlement = 0
                .CostRecovery = 0
                .Markup = 0

            End With
        End If

        GrandTotal = 0

    End Sub


#End Region



    Private Sub CostFormat_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.DoubleClick

        Exit Sub

        useNewCode = Not useNewCode

        If Not useNewCode AndAlso Not controllist Is Nothing Then
            For Each con As Control In controllist
                con.DataBindings.Clear()

                If TypeOf con.Tag Is CostFormatSpec Then
                    DirectCast(con.Tag, CostFormatSpec).dispose()
                End If
                con.Tag = Nothing

            Next
        End If

        Me.SetDisplay(False)
        Me.LoadCommitment()


    End Sub
End Class
