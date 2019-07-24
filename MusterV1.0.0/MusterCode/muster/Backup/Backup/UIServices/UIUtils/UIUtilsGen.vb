Imports Microsoft.Win32.Registry
Imports System.Runtime.InteropServices
Imports System.Drawing.Printing
Imports System.IO
Imports System
Imports System.Text
Imports System.Drawing.Imaging.BitmapData


Public Class UIUtilsGen
    '-------------------------------------------------------------------------------
    ' MUSTER.MUSTER.UIUtilsGen
    '   Provides various utility functions to the MUSTER UI
    '
    ' Copyright (C) 2004 CIBER, Inc.
    ' All rights reserved.
    '
    ' Release   Initials    Date        Description
    '  1.0        ??      8/??/04    Original class definition.
    '  1.1       JVC2     2/8/2005   Moved functions from UIUtils.vb
    '  1.2       AB       03/21/05   Altered PopulateOwnerFacilities to handle Technical
    '  1.3       AB       03/22/05   Added GetLUSTEventsForFacility
    '-------------------------------------------------------------------------------
    '
    ' TODO - Intergate with application JVC2 2/9/05
    '
    'This function sets empty values for date pickers
    Private Shared strName As String = String.Empty
    Private Shared currentNotes As String = String.Empty

#Region "Imported DLL functions"

    Public Enum IEIFLAG As Integer
        ASYNC = &H1
        CACHE = &H2
        ASPECT = &H4
        OFFLINE = &H8
        GLEAM = &H10
        SCREEN = &H20
        ORIGSIZE = &H40
        NOSTAMP = &H80
        NOBORDER = &H100
        QUALITY = &H200
    End Enum




    <ComImportAttribute(), _
    GuidAttribute("000214eb-0000-0000-c000-000000000046"), _
    InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)> _
Friend Interface IExtractIcon
        <PreserveSig()> _
        Function Extract(ByVal pszFile As IntPtr, _
                 ByVal nIconIndex As Integer, _
                 ByVal phiconLarge As IntPtr, _
                 ByVal phiconSmall As IntPtr, _
                 ByVal nIconSize As Integer) As Integer

        Function GetIconLocation(ByVal uFlags As Integer, _
                    <MarshalAs(UnmanagedType.LPStr)> _
                     ByRef szIconFile As StringBuilder, _
                     ByVal cchMax As Integer, _
                     ByRef piIndex As Integer, _
                     ByRef pwFlags As Integer) As Integer
    End Interface



    ''' <summary>
    ''' Exposes methods that request a thumbnail image from a Shell folder.
    ''' </summary>
    <ComImportAttribute(), _
    GuidAttribute("BB2E617C-0920-11d1-9A0B-00C04FC2D6C1"), _
    InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)> _
    Friend Interface IExtractImage

        ''' <summary>
        ''' Gets a path to the image that is to be extracted.
        ''' </summary>
        ''' <returns>This method may return a COM-defined error code or one of the following: S_OK if successful, or E_PENDING.</returns>
        Function GetLocation( _
            <Out(), MarshalAs(UnmanagedType.LPWStr)> ByVal pszPathBuffer As System.Text.StringBuilder, _
            ByVal cch As Integer, _
            ByRef pdwPriority As Integer, _
            ByRef prgSize As SIZE, _
            ByVal dwRecClrDepth As Integer, _
            ByRef pdwFlags As Integer) As Integer

        ''' <summary>
        ''' Requests an image from an object, such as an item in a Shell folder.
        ''' </summary>
        ''' <returns>Returns S_OK if successful, or a COM-defined error code otherwise.</returns>
        Function Extract(<Out()> ByRef phBmpImage As IntPtr) As Integer
    End Interface




    <StructLayout(LayoutKind.Sequential)> _
    Public Structure STRRET_CSTR
        Public uType As Integer
        <FieldOffset(4), MarshalAs(UnmanagedType.LPWStr)> _
        Public pOleStr As String
        <FieldOffset(4)> _
        Public uOffset As Integer
        <FieldOffset(4), MarshalAs(UnmanagedType.ByValArray, SizeConst:=520)> _
        Public strName As Byte()
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure SIZE
        Public cx As Integer
        Public cy As Integer
    End Structure



    <ComImportAttribute(), _
    GuidAttribute("000214E6-0000-0000-C000-000000000046"), _
    InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)> _
    Public Interface IShellFolder

        Sub ParseDisplayName( _
          ByVal hWnd As IntPtr, _
          ByVal pbc As IntPtr, _
          ByVal pszDisplayName As String, _
          ByRef pchEaten As Integer, _
          ByRef ppidl As System.IntPtr, _
          ByRef pdwAttributes As Integer)

        Sub EnumObjects( _
          ByVal hwndOwner As IntPtr, _
          <MarshalAs(UnmanagedType.U4)> ByVal grfFlags As Integer, _
          <Out()> ByRef ppenumIDList As IntPtr)

        Sub BindToObject( _
          ByVal pidl As IntPtr, _
          ByVal pbcReserved As IntPtr, _
          ByRef riid As Guid, _
          ByRef ppvOut As IShellFolder)

        Sub BindToStorage( _
          ByVal pidl As IntPtr, _
          ByVal pbcReserved As IntPtr, _
          ByRef riid As Guid, _
          <Out()> ByVal ppvObj As IntPtr)

        <PreserveSig()> _
        Function CompareIDs( _
          ByVal lParam As IntPtr, _
          ByVal pidl1 As IntPtr, _
          ByVal pidl2 As IntPtr) As Integer

        Sub CreateViewObject( _
          ByVal hwndOwner As IntPtr, _
          ByRef riid As Guid, _
          ByVal ppvOut As Object)

        Sub GetAttributesOf( _
          ByVal cidl As Integer, _
          ByVal apidl As IntPtr, _
          <MarshalAs(UnmanagedType.U4)> ByRef rgfInOut As Integer)

        Sub GetUIObjectOf( _
          ByVal hwndOwner As IntPtr, _
          ByVal cidl As Integer, _
          ByRef apidl As IntPtr, _
          ByRef riid As Guid, _
          <Out()> ByVal prgfInOut As Integer, _
          <Out(), MarshalAs(UnmanagedType.IUnknown)> ByRef ppvOut As Object)

        Sub GetDisplayNameOf( _
          ByVal pidl As IntPtr, _
          <MarshalAs(UnmanagedType.U4)> ByVal uFlags As Integer, _
          ByRef lpName As STRRET_CSTR)

        Sub SetNameOf( _
          ByVal hwndOwner As IntPtr, _
          ByVal pidl As IntPtr, _
          <MarshalAs(UnmanagedType.LPWStr)> ByVal lpszName As String, _
          <MarshalAs(UnmanagedType.U4)> ByVal uFlags As Integer, _
          ByRef ppidlOut As IntPtr)

    End Interface

    Public Class ShellInterop
        <DllImport("shell32.dll", CharSet:=CharSet.Auto)> _
        Public Shared Function SHGetDesktopFolder( _
          <Out()> ByRef ppshf As IShellFolder) As Integer
        End Function
    End Class

    <DllImport("gdi32.DLL", EntryPoint:="BitBlt", _
 SetLastError:=True, CharSet:=CharSet.Unicode, _
 ExactSpelling:=True, _
 CallingConvention:=CallingConvention.StdCall)> _
 Private Shared Function BitBlt(ByVal hdcDest As IntPtr, ByVal nXDest As Integer, ByVal nYDest As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hdcSrc As IntPtr, ByVal nXSrc As Integer, ByVal nYSrc As Integer, ByVal dwRop As System.Int32) As Boolean
    End Function


#End Region



#Region "String Methods"
    Friend Shared Function TitleCaseString(ByVal str As String) As String
        Try
            Return StrConv(str, VbStrConv.ProperCase)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Friend Shared Function TodayString() As String

        Return Now.ToString("MMddyy_HHmm")

    End Function
#End Region
#Region "Date Methods"
    Friend Shared Sub CreateEmptyFormatDatePicker(ByRef dtPicker As DateTimePicker)
        Try
            dtPicker.CustomFormat = "__/__/____"
            dtPicker.Format = DateTimePickerFormat.Custom
            dtPicker.Checked = False
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    'This function toggles/switches the date format based on the state of the 
    'Date time Picker
    Friend Shared Sub ToggleDateFormat(ByRef dtPicker As DateTimePicker)
        Try

            If dtPicker.Checked Then
                Dim dtTemp As Date = dtPicker.Value
                If dtPicker.Format <> DateTimePickerFormat.Short Then
                    dtPicker.Format = DateTimePickerFormat.Short
                    dtPicker.Value = dtTemp
                End If
            Else
                dtPicker.Tag = Nothing
                CreateEmptyFormatDatePicker(dtPicker)
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Shared Function GetDatePickerValue(ByVal dtPick As DateTimePicker) As Date
        Try
            Dim dtPickValue As Date
            If dtPick.Checked And dtPick.Enabled Then
                dtPickValue = dtPick.Value
            End If
            Return dtPickValue
        Catch ex As Exception
            MsgBox("Cannot get the Date Picker Value")
        End Try
    End Function
    Friend Shared Sub SetDatePickerValue(ByVal dtPick As DateTimePicker, ByVal dtValue As DateTime)
        Try
            If Date.Compare(dtValue, CDate("01/01/0001")) = 0 Then
                CreateEmptyFormatDatePicker(dtPick)
            Else
                dtPick.Format = DateTimePickerFormat.Short
                dtPick.Checked = True
                dtPick.Value = dtValue

                If dtPick.Value < "1/1/1910" Then
                    dtPick.Visible = False
                Else
                    dtPick.Visible = True
                End If


            End If
        Catch ex As Exception
            MsgBox("Cannot Set Value for " + dtPick.Name + vbCrLf + ex.Message)
        End Try
    End Sub
    Public Sub RejectFutureDate(ByVal dtpicks As Collection)
        Dim dtPick As Object
        Try
            For Each dtPick In dtpicks
                AddHandler CType(dtPick, DateTimePicker).CloseUp, AddressOf RejectFutureDate_Handler
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub RejectFutureDate_Handler(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dtPick As DateTimePicker = CType(sender, DateTimePicker)
        Dim dtValue As Date = dtPick.Value
        Try
            If DateDiff(DateInterval.Day, Today(), dtValue) > 0 Then
                If IsNothing(dtPick.Tag) Then
                    dtPick.Checked = False
                Else
                    If Date.Compare(dtPick.Tag, CDate("01/01/0001")) = 0 Then
                        dtPick.Checked = False
                    Else
                        dtPick.Value = dtPick.Tag
                    End If

                End If

                MsgBox("The date selected cannot be greater than today")
                dtPick.Refresh()
            Else
                If dtPick.Format <> DateTimePickerFormat.Short Then
                    dtPick.Format = DateTimePickerFormat.Short
                    dtPick.Value = dtValue
                End If
                dtPick.Tag = dtValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub RetainCurrentDateValue(ByVal objControl As Control)
        Try
            Dim currentControl As Control
            Dim tmpdtpick As System.Windows.Forms.DateTimePicker
            Dim myEnumerator As System.Collections.IEnumerator = _
            objControl.Controls.GetEnumerator()
            Try
                While myEnumerator.MoveNext()
                    currentControl = myEnumerator.Current
                    If currentControl.GetType.ToString.ToLower = "system.Windows.Forms.DateTimePicker".ToLower Then
                        tmpdtpick = CType(currentControl, System.Windows.Forms.DateTimePicker)
                        AddHandler tmpdtpick.DropDown, AddressOf RetainCurrentDateValue_Handler
                    Else
                        If currentControl.Controls.Count > 0 Then
                            RetainCurrentDateValue(currentControl)
                        End If
                    End If
                End While
            Catch ex As Exception
                Throw ex
            End Try

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub RetainCurrentDateValue_Handler(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dtPick As DateTimePicker = CType(sender, DateTimePicker)
        Try
            If IsDate(dtPick.Text) Then
                dtPick.Tag = CType(dtPick.Text, Date)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
#End Region
#Region "ComboBox Methods"
    Friend Shared Function GetComboBoxValue(ByVal cmbField As ComboBox) As Object
        Try

            Dim cmbSelectedValue As Object
            If cmbField.Enabled Then
                cmbSelectedValue = cmbField.SelectedValue
            End If
            Return cmbSelectedValue
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Friend Shared Function GetComboBoxValueInt(ByVal cmbField As ComboBox) As Integer
        Try
            Dim cmbSelectedValue As Object = GetComboBoxValue(cmbField)
            If IsNothing(cmbSelectedValue) Then
                cmbSelectedValue = 0
            End If
            Return CType(cmbSelectedValue, Integer)
        Catch ex As Exception
            MsgBox("Cannot get the Combobox Value for " + cmbField.Name + vbCrLf + ex.Message)
        End Try
    End Function
    Friend Shared Function GetComboBoxValueString(ByVal cmbField As ComboBox) As String
        Try

            Dim cmbSelectedValue As Object = GetComboBoxValue(cmbField)
            If IsNothing(cmbSelectedValue) Then
                cmbSelectedValue = ""
            End If
            Return CType(cmbSelectedValue, String)
        Catch ex As Exception
            MsgBox("Cannot get the Combobox Value for " + cmbField.Name + vbCrLf + ex.Message)
        End Try
    End Function
    Friend Shared Function GetComboBoxText(ByVal cmbField As ComboBox) As String
        Try
            If cmbField.Enabled Then
                Return cmbField.Text
            Else
                Return ""
            End If

        Catch ex As Exception
            MsgBox("Cannot get the Combobox Text for " + cmbField.Name + vbCrLf + ex.Message)
        End Try
    End Function
    Friend Shared Sub SetComboboxItemByValue(ByVal cmb As ComboBox, ByVal oSelectedValue As Object, Optional ByVal overRideEnabled As Boolean = False)
        Try
            If cmb.Enabled Or overRideEnabled Then
                If cmb.Items.Count > 0 Then
                    If oSelectedValue > 0 Then
                        cmb.SelectedValue = oSelectedValue
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox("Cannot Set Combobox Value for " + cmb.Name + vbCrLf + ex.Message)
        End Try
    End Sub
    Friend Shared Sub SetComboboxItemByText(ByVal cmb As ComboBox, ByVal strSelectedText As String)
        Try
            If cmb.Enabled Then
                cmb.Text = strSelectedText
            End If
        Catch ex As Exception
            MsgBox("Cannot Set Combobox Value for " + cmb.Name + vbCrLf + ex.Message)
        End Try
    End Sub
    'Friend Shared Sub FillComboBox(ByVal cmb As ComboBox, ByVal Lookups As ArrayList)
    '    Dim i As Integer
    '    Dim dtableSrc As DataTable

    '    Try
    '        dtableSrc = ArrayListToDataTable(Lookups)
    '        cmb.DataSource = dtableSrc
    '        cmb.DisplayMember = "Type"
    '        cmb.ValueMember = "Id"

    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub
    Friend Shared Sub FillComboBox(ByVal cmb As ComboBox, ByVal dtableSrc As DataTable)

        Try
            cmb.DataSource = dtableSrc
            cmb.DisplayMember = "Type"
            cmb.ValueMember = "Id"

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Sub ClearComboBox(ByRef objControl As Control)
        Dim currentControl As Control
        Dim tmpCmb As System.Windows.Forms.ComboBox
        Dim myEnumerator As System.Collections.IEnumerator = _
        objControl.Controls.GetEnumerator()
        Try
            While myEnumerator.MoveNext()
                currentControl = myEnumerator.Current
                If currentControl.GetType.ToString.ToLower = "system.Windows.Forms.ComboBox".ToLower Then
                    tmpCmb = CType(currentControl, System.Windows.Forms.ComboBox)
                    AddHandler tmpCmb.KeyPress, AddressOf ComboBoxDelegates
                Else
                    If currentControl.Controls.Count > 0 Then
                        ClearComboBox(currentControl)
                    End If
                End If

            End While
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub ComboBoxDelegates(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim cmbBox As ComboBox = CType(sender, ComboBox)
        If e.KeyChar = Microsoft.VisualBasic.ChrW(8) Then
            cmbBox.SelectedIndex = 0
            cmbBox.SelectedIndex = -1
        End If
    End Sub
    Friend Shared Sub ValidateComboBoxItemByValue(ByVal cmb As ComboBox, ByVal oSelectedValue As Object)
        Try
            If cmb.Items.Count > 0 Then
                If oSelectedValue > 0 Then
                    cmb.SelectedValue = oSelectedValue
                Else
                    cmb.SelectedIndex = -1
                    cmb.SelectedIndex = -1
                End If
            End If
        Catch ex As Exception
            MsgBox("Cannot Set Combobox Value for " + cmb.Name + vbCrLf + ex.Message)
        End Try
    End Sub
    Friend Shared Function ComboBoxContainsValueSourceIsDataTable(ByVal cmb As ComboBox, ByVal val As String, ByVal valueMember As String, ByVal operand As String) As Boolean
        Dim dt As DataTable
        Dim returnVal As Boolean = False
        Try
            dt = cmb.DataSource
            If Not dt Is Nothing Then
                If dt.Columns.Contains(valueMember) Then
                    If dt.Select(valueMember + " " + operand + " " + val).Length > 0 Then returnVal = True
                End If
            End If
        Catch ex As Exception
            MsgBox(cmb.Name + vbCrLf + ex.Message)
        End Try
        Return returnVal
    End Function
    Friend Shared Function ComboBoxContainsValueSourceIsDataSet(ByVal cmb As ComboBox, ByVal val As String, ByVal valueMember As String, ByVal operand As String) As Boolean
        Dim ds As DataSet
        Dim dt As DataTable
        Dim returnVal As Boolean = False
        Try
            ds = cmb.DataSource
            If Not ds Is Nothing Then
                If ds.Tables.Count > 0 Then
                    dt = ds.Tables(0)
                    If dt.Columns.Contains(valueMember) Then
                        If dt.Select(valueMember + " " + operand + " " + val).Length > 0 Then returnVal = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(cmb.Name + vbCrLf + ex.Message)
        End Try
        Return returnVal
    End Function
#End Region
#Region "Validation Methods"
    Friend Shared Function IsPathValid(ByVal inputPath As String) As String
        Dim str As String = String.Empty
        Try
            inputPath = IIf(IsNothing(inputPath), "", inputPath)
            Dim strRegex As String = "([A-Za-z]:[^/:*?<>|]+|\\)[^/:*?<>|]+\\[^/:*?<>|]+"
            Dim rx As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex(strRegex)
            If rx.IsMatch(inputPath) Then

            Else

                If inputPath.IndexOfAny(System.IO.Path.InvalidPathChars) <> -1 Or inputPath.IndexOfAny("*?") >= 0 Then
                    str += "/ : * ? < > | are not allowed in " + inputPath + vbCrLf
                Else
                    'If System.IO.Path.IsPathRooted(inputPath) = False Then
                    '    str += inputPath + " contains no root information." + vbCrLf
                    'End If
                End If
                If str.Length > 0 Then
                    Return str
                End If
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Shared Function IsEmailValid(ByVal inputEmail As String) As Boolean
        Try
            inputEmail = IIf(IsNothing(inputEmail), "", inputEmail)
            Dim strRegex As String = "^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
            Dim rx As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex(strRegex)
            If rx.IsMatch(inputEmail) Then
                Return (True)
            Else
                Return (False)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Friend Shared Function IsPhoneValid(ByVal inputPhone As String) As Boolean
        Try
            Dim strRegex As String = "(\(\d\d\d\))?\s*(\d\d\d)\s*[\-]?\s*(\d\d\d\d)"
            '"^\(?\d{3}\)?\s|-\d{3}-\d{4}$" -  matches (555) 555-5555, or 555-555-5555
            Dim rx As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex(strRegex)
            If rx.IsMatch(inputPhone) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region
#Region "Lookup Methods"

    ''' <summary>
    ''' This method starts at the specified directory, and traverses all subdirectories.
    ''' It returns a List of those directories.
    ''' </summary>
    Public Shared Function GetFilesRecursive(ByVal initial As String, ByVal file As String) As Collections.ArrayList
        ' This list stores the results.
        Dim result As New Collections.ArrayList

        ' This stack stores the directories to process.
        Dim stack As New Collections.Stack

        ' Add the initial directory
        stack.Push(initial)

        ' Continue processing for each stacked directory
        Do While (stack.Count > 0)
            ' Get top directory string
            Dim dir As String = stack.Pop
            Try
                ' Add all immediate file paths
                result.AddRange(Directory.GetFiles(dir, file))

                ' Loop through all subdirectories and add them to the stack.
                Dim directoryName As String
                For Each directoryName In Directory.GetDirectories(dir)
                    stack.Push(directoryName)
                Next

            Catch ex As Exception
            End Try
        Loop

        ' Return the list
        Return result
    End Function

    Friend Shared Sub PopulateOwnerType(ByVal cmb As ComboBox, ByRef pOwn As Object)
        Try
            cmb.DisplayMember = "PROPERTY_NAME"
            cmb.ValueMember = "PROPERTY_ID"
            cmb.DataSource = pOwn.PopulateOwnerType()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Friend Shared Sub PopulateOwnerEntityList(ByVal cmb As ComboBox, ByRef pOwn As BusinessLogic.pOwner, ByVal showAll As Boolean)
        Try
            cmb.DisplayMember = "PROPERTY_NAME"
            cmb.ValueMember = "PROPERTY_ID"
            cmb.DataSource = pOwn.PopulateOpenEntitiesForOwnership(showAll)
            If pOwn.BPersona.PersonId <> 0 Then
                cmb.SelectedValue = String.Format("P|{0}", pOwn.BPersona.PersonId)
            ElseIf pOwn.BPersona.OrgID <> 0 Then
                cmb.SelectedValue = String.Format("O|{0}", pOwn.BPersona.OrgID)
            Else
                cmb.SelectedValue = "-1"
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Friend Shared Sub PopulateOrgEntityType(ByVal cmb As ComboBox, ByRef pOwn As Object)
        Try
            cmb.DisplayMember = "PROPERTY_NAME"
            cmb.ValueMember = "PROPERTY_ID"
            cmb.DataSource = pOwn.BPersona.PopulateEntityCode()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Shared Sub PopulateFacilityType(ByVal cmb As ComboBox, ByRef pFacilities As Object)
        Try
            cmb.DisplayMember = "PROPERTY_NAME"
            cmb.ValueMember = "PROPERTY_ID"
            cmb.DataSource = pFacilities.PopulateFacilityType()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Shared Sub PopulateFacilityDatum(ByVal cmb As ComboBox, ByRef pFacilities As Object)
        Try
            cmb.DisplayMember = "PROPERTY_NAME"
            cmb.ValueMember = "PROPERTY_ID"
            cmb.DataSource = pFacilities.PopulateFacilityDatum
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Shared Sub PopulateFacilityMethod(ByVal cmb As ComboBox, ByRef pFacilities As Object)
        Try
            cmb.DisplayMember = "PROPERTY_NAME"
            cmb.ValueMember = "PROPERTY_ID"
            cmb.DataSource = pFacilities.PopulateFacilityMethod()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Friend Shared Sub PopulateFacilityLocationType(ByVal cmb As ComboBox, ByRef pFacilities As Object)
        Try
            cmb.DisplayMember = "PROPERTY_NAME"
            cmb.ValueMember = "PROPERTY_ID"
            cmb.DataSource = pFacilities.PopulateFacilityLocationType()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Screen captures"

    Public Shared Img As Bitmap

    Public Shared Sub CaptureScreen(ByVal frm As Form, ByVal notes As String)
        Dim g1 As Graphics = frm.CreateGraphics()
        Dim dc1 As IntPtr

        g1 = frm.CreateGraphics

        dc1 = g1.GetHdc()

        CaptureScreen(dc1, frm.ClientRectangle.Width, (frm.ClientRectangle.Height), notes)

        g1.ReleaseHdc(dc1)


    End Sub

    Public Shared Sub CaptureScreen(ByVal pointer As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal notes As String)

        Dim dc2 As IntPtr
        Dim MyImage As Bitmap

        Dim g2 As Graphics

        MyImage = New Bitmap(x, y)
        Img = MyImage

        g2 = Graphics.FromImage(MyImage)
        dc2 = g2.GetHdc()

        Dim pd As PrintDocument

        Try

            BitBlt(dc2, 0, 0, x, y, pointer, 0, 0, 13369376)

            g2.ReleaseHdc(dc2)

            pd = New PrintDocument


            currentNotes = notes
            AddHandler pd.PrintPage, AddressOf pd_PrintPage

            pd.Print()

        Catch ex As Exception

        Finally



            If Not Img Is Nothing Then
                Img.Dispose()
            End If


            If Not pd Is Nothing Then

                RemoveHandler pd.PrintPage, AddressOf pd_PrintPage

                pd.Dispose()
            End If

            If Not MyImage Is Nothing Then
                MyImage.Dispose()
            End If


        End Try




    End Sub


    Public Shared Sub FormatNotes(ByRef str As String, ByVal charPerLine As Integer)

        Dim newStr As New StringBuilder(str)
        Dim gg As Integer = 0
        Dim g As Integer = 0
        Dim endString As Integer

        'newStr.Replace(vbCrLf, String.Empty)

        endString = newStr.Length

        While g < endString

            If gg = charPerLine AndAlso g > 0 Then
                While newStr.Chars(g - 1) <> " " AndAlso g > 0
                    g -= 1
                End While

                If g = 0 Then Exit While

                newStr.Insert(g, vbCrLf)

                endString = newStr.Length

                g += vbCrLf.Length
                gg = -1
            ElseIf (g) <= (endString - 1) AndAlso g > 4 AndAlso newStr.ToString.Substring(g - 1, 2) = vbCrLf Then
                g += vbCrLf.Length
                gg = -1
            Else
                g += 1
            End If

            gg += 1

        End While

        str = newStr.ToString

        newStr.Length = 0

    End Sub

    'this method will be called each time when pd.printpage event occurs
    Public Shared Sub pd_PrintPage(ByVal sender As Object, ByVal e As PrintPageEventArgs)

        e.PageSettings.Landscape = False
        e.Graphics.DrawImage(Img, 0, 0, 960, 700)
        e.Graphics.DrawLine(Drawing.Pens.Black, 0, 705, 965, 705)

        UIUtilsGen.FormatNotes(currentNotes, 180)

        e.Graphics.DrawString(currentNotes, New Font(Drawing.FontFamily.GenericSansSerif, 7, FontStyle.Bold, GraphicsUnit.Pixel), Drawing.Brushes.Black, New Drawing.RectangleF(5, 715, 960, 1080))
        currentNotes = String.Empty
        e.HasMorePages = False

    End Sub


#End Region
#Region "Common Methods"

#Region "Tickler Entity Opening handlers"

    Public Shared Sub ActivateEntity(ByVal frm As Object)

        Dim grid As Infragistics.Win.UltraWinGrid.RowsCollection
        Dim prop As Reflection.PropertyInfo = Nothing

        With BootStrap._container

            prop = frm.GetType.GetProperty(.GoToGrid, Reflection.BindingFlags.Public Or Reflection.BindingFlags.GetProperty Or Reflection.BindingFlags.IgnoreCase Or Reflection.BindingFlags.Instance)

            If .GoToEntityCode > -1 AndAlso Not prop Is Nothing Then


                grid = DirectCast(prop.GetValue(frm, Nothing), Infragistics.Win.UltraWinGrid.UltraGrid).Rows

                If Not grid Is Nothing Then

                    For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In grid

                        If row.Cells(0).Value = .GoToEntityCode Then

                            .GoToEntityCode = -1

                            DirectCast(prop.GetValue(frm, Nothing), Infragistics.Win.UltraWinGrid.UltraGrid).ActiveRow = row

                            If Not frm.GetType.GetMember("PerformClick").GetUpperBound(0) = -1 Then

                                frm.GetType.InvokeMember("PerformClick", Reflection.BindingFlags.InvokeMethod, Nothing, frm, Nothing)

                            End If

                        End If

                    Next

                End If

                .GoToEntityCode = -1

            End If

        End With

        prop = Nothing

    End Sub


#End Region

#Region "Contacts"

    Public Class ContactMessageException
        Inherits Exception

        Sub New(ByVal message As String)
            MyBase.New(message)
        End Sub
    End Class

    Public Shared Function ModifyContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal nModuleID As Integer, ByRef PConStruct As BusinessLogic.pContactStruct) As Boolean

        Dim contactFrm As Contacts

        Try

            If ugGrid.Enabled Then

                If ugGrid.Rows.Count <= 0 Then Exit Function

                If ugGrid.ActiveRow Is Nothing Then
                    Throw New Exception("Select row to Modify.")
                End If


                Dim dr As Infragistics.Win.UltraWinGrid.UltraGridRow = ugGrid.ActiveRow
                Dim moduleNameOnRow As String = dr.Cells("Module").Value
                Dim moduleName As String = UIUtilsGen.GetModuleNameByID(nModuleID)

                If moduleNameOnRow.ToUpper = "C & E" Then
                    moduleNameOnRow = "CAE"
                End If
                If moduleNameOnRow.ToUpper <> moduleName.ToUpper And (moduleNameOnRow.ToUpper <> "C & E" AndAlso moduleName.ToUpper <> "CAE") Then
                    Throw New ContactMessageException(String.Format("Cannot modify {0} Contacts in {1}", moduleNameOnRow, moduleName))
                ElseIf dr.CellAppearance.BackColor.ToString = Color.Gainsboro.ToString Then
                    Throw New ContactMessageException(String.Format("Cannot modify this {0} Registration contact from here. ", moduleName))
                End If

                contactFrm = New Contacts(Integer.Parse(dr.Cells("EntityID").Value), CInt(dr.Cells("EntityType").Value), dr.Cells("Module").Value, CInt(dr.Cells("ContactID").Value), dr, PConStruct, "MODIFY")

                contactFrm.ShowDialog()

                contactFrm.Dispose()

                Return True
            Else
                Return False

            End If


        Catch ex As ContactMessageException

            MsgBox(ex.Message)

            Return False

        Catch ex As Exception
            Throw ex

        Finally
            contactFrm = Nothing

        End Try
    End Function

    Public Shared Function AssociateContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal nEntityID As Integer, ByVal nEntityType As Integer, ByVal nModuleID As Integer, ByRef PConStruct As BusinessLogic.pContactStruct) As Boolean

        Dim contactFrm As Contacts

        Try

            If ugGrid.Enabled Then

                If ugGrid.Rows.Count <= 0 Then Exit Function

                If ugGrid.ActiveRow Is Nothing Then
                    Throw New ContactMessageException("Select row to Associate.")
                End If
                Dim dr As Infragistics.Win.UltraWinGrid.UltraGridRow = ugGrid.ActiveRow
                Dim moduleName As String = UIUtilsGen.GetModuleNameByID(nModuleID)

                If ((CInt(ugGrid.ActiveRow.Cells("EntityID").Value) = nEntityID) And (CInt(ugGrid.ActiveRow.Cells("ModuleID").Value) = nModuleID)) Then
                    Throw New ContactMessageException("Selected contact is already associated with the current entity")
                End If

                contactFrm = New Contacts(nEntityID, nEntityType, moduleName, CInt(dr.Cells("ContactID").Value), dr, PConStruct, "ASSOCIATE")

                contactFrm.ShowDialog()

                contactFrm.Dispose()

                Return True
            End If

            Return False

        Catch ex As ContactMessageException
            MsgBox(ex.Message)

            Return False
        Catch ex As Exception
            Throw ex
        Finally
            contactFrm = Nothing
        End Try
    End Function

    Public Shared Function DeleteContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal nEntityID As Integer, ByVal nModuleID As Integer, ByRef PConStruct As BusinessLogic.pContactStruct) As Boolean
        Try
            If ugGrid.Enabled Then
                Dim result As DialogResult
                If ugGrid.Rows.Count <= 0 Then Exit Function

                If ugGrid.ActiveRow Is Nothing Then
                    Throw New ContactMessageException("Select row to Delete.")
                End If

                Dim moduleName As String = UIUtilsGen.GetModuleNameByID(nModuleID)

                If ugGrid.ActiveRow.CellAppearance.BackColor.ToString = Color.Gainsboro.ToString Then
                    Throw New ContactMessageException(String.Format("Cannot unassociate this {0} contact as it is for a different entity.", moduleName))
                End If

                If (CInt(ugGrid.ActiveRow.Cells("EntityID").Value) <> nEntityID) Or (CInt(ugGrid.ActiveRow.Cells("ModuleID").Value) <> nModuleID) Then
                    Throw New ContactMessageException("Selected contact is not associated with the current entity and cannot be deleted")
                End If

                result = MessageBox.Show("Are you sure you wish to unassociate this contact?", "MUSTER", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If result = DialogResult.No Then Exit Function

                Dim returnval As String = String.Empty

                PConStruct.Remove(ugGrid.ActiveRow.Cells("EntityAssocID").Text, CType(nModuleID, Integer), MusterContainer.AppUser.UserKey, returnval, MusterContainer.AppUser.ID)

                Try
                    If UIUtilsGen.HasRights(returnval) Then
                        ugGrid.ActiveRow.Delete(False)
                    End If
                Catch ex As Exception
                    Throw New ContactMessageException(ex.ToString)
                End Try
                Return True
            End If


            Return False

        Catch ex As ContactMessageException

            MsgBox(ex.Message)

            Return False

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Sub LoadContacts(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal EntityID As Integer, ByVal EntityType As Integer, ByVal pConStruct As BusinessLogic.pContactStruct, ByVal moduleID As Integer, _
    Optional ByVal strEntities As String = "", Optional ByVal bolActive As Boolean = False, Optional ByVal strentityassocid As String = "", Optional ByVal nRelatedEntityType As Integer = 0, Optional ByVal defaultview As Data.DataView = Nothing)
        Dim dsContactsLocal As DataSet
        Try
            'dsContacts = pConStruct.GetAll()
            If defaultview Is Nothing Then
                dsContactsLocal = pConStruct.GetFilteredContacts(EntityID, moduleID, strEntities, bolActive, strentityassocid, EntityType, nRelatedEntityType)
            End If



            'dsContacts.Tables(0).DefaultView.RowFilter = "MODULEID = 614 And ENTITYID = " + EntityID.ToString

            If Not EntityID = 0 Then


                If defaultview Is Nothing Then
                    ugGrid.DataSource = dsContactsLocal.Tables(0).DefaultView      'dsContactsLocal.Tables(0).DefaultView 'dsContacts.Tables(0).DefaultView    
                Else
                    ugGrid.DataSource = defaultview
                End If

                '------ change column headings ---------------------------------------------------------
                ugGrid.DisplayLayout.Bands(0).Columns("CONTACT_Name").Header.Caption = "Contact Name"
                ugGrid.DisplayLayout.Bands(0).Columns("Address").Header.Caption = "Address"
                ugGrid.DisplayLayout.Bands(0).Columns("Assoc Company").Header.Caption = "Assoc Company"
                ugGrid.DisplayLayout.Bands(0).Columns("Phone Contact").Header.Caption = "Phone Contact"
                ugGrid.DisplayLayout.Bands(0).Columns("IsPersonType").Header.Caption = "Person/Company"
                ugGrid.DisplayLayout.Bands(0).Columns("Associated").Header.Caption = "Associated"
                ugGrid.DisplayLayout.Bands(0).Columns("E-mail").Header.Caption = "E-mail"
                ugGrid.DisplayLayout.Bands(0).Columns("ACTIVE").Header.Caption = "Active"
                ugGrid.DisplayLayout.Bands(0).Columns("MODULE").Header.Caption = "Module"
                ugGrid.DisplayLayout.Bands(0).Columns("RELATIONSHIP").Header.Caption = "Relationship"
                ugGrid.DisplayLayout.Bands(0).Columns("Type").Header.Caption = "Contact Type"

                '------ column widths ------------------------------------------------------------
                ugGrid.DisplayLayout.Bands(0).Columns("CONTACT_Name").Width = 185
                ugGrid.DisplayLayout.Bands(0).Columns("Module").Width = 80
                ugGrid.DisplayLayout.Bands(0).Columns("Type").Width = 150
                ugGrid.DisplayLayout.Bands(0).Columns("Relationship").Width = 110


                ugGrid.DisplayLayout.Bands(0).Columns("Address").Width = 300
                ugGrid.DisplayLayout.Bands(0).Columns("CC").Width = 40
                ugGrid.DisplayLayout.Bands(0).Columns("Assoc Company").Width = 160
                ugGrid.DisplayLayout.Bands(0).Columns("Phone Contact").Width = 300
                ugGrid.DisplayLayout.Bands(0).Columns("Associated").Width = 85
                ugGrid.DisplayLayout.Bands(0).Columns("Active").Width = 70
                ugGrid.DisplayLayout.Bands(0).Columns("IsPersonType").Width = 80


                '------ assign hidden columns ----------------------------------------------------
                ugGrid.DisplayLayout.Bands(0).Columns("ContactID").Hidden = True
                ugGrid.DisplayLayout.Bands(0).Columns("EntityAssocID").Hidden = True

                ugGrid.DisplayLayout.Bands(0).Columns("address_one").Hidden = True
                ugGrid.DisplayLayout.Bands(0).Columns("address_Two").Hidden = True
                ugGrid.DisplayLayout.Bands(0).Columns("City").Hidden = True
                ugGrid.DisplayLayout.Bands(0).Columns("State").Hidden = True
                ugGrid.DisplayLayout.Bands(0).Columns("ContactType").Hidden = True
                ugGrid.DisplayLayout.Bands(0).Columns("IsPerson").Hidden = True


                ugGrid.DisplayLayout.Bands(0).Columns("Zip").Hidden = True
                ugGrid.DisplayLayout.Bands(0).Columns("ModuleID").Hidden = True
                ugGrid.DisplayLayout.Bands(0).Columns("EntityType").Hidden = True
                ugGrid.DisplayLayout.Bands(0).Columns("Alias Used").Hidden = True
                ugGrid.DisplayLayout.Bands(0).Columns("Address Alias Used").Hidden = True
                ugGrid.DisplayLayout.Bands(0).Columns("LetterContactType").Hidden = True

                If EntityType <> 26 Then
                    ugGrid.DisplayLayout.Bands(0).Columns("EntityID").Hidden = True
                End If

                For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In ugGrid.Rows

                    If row.Cells("Associated").Value.ToString.ToUpper = "READ ONLY" Then
                        row.Appearance.ForeColor = Color.LightSlateGray
                        row.Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.False
                        row.Appearance.BackColor = Color.Gainsboro

                    Else
                        row.Appearance.ForeColor = Color.Black
                        row.Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.Default
                        row.Appearance.BackColor = Color.White


                        If row.Cells("Alias Used").Value = "Yes" Then
                            row.Cells("CONTACT_Name").Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
                            row.Cells("CONTACT_Name").Appearance.FontData.SizeInPoints = 9
                            row.Cells("CONTACT_Name").Appearance.ForeColor = Color.Red
                        Else
                            row.Cells("CONTACT_Name").Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.Default
                            row.Cells("CONTACT_Name").Appearance.FontData.SizeInPoints = 8
                            row.Cells("CONTACT_Name").Appearance.ForeColor = Color.Black

                        End If

                        If row.Cells("Address Alias Used").Value = "Yes" Then
                            row.Cells("Address").Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
                            row.Cells("Address").Appearance.FontData.SizeInPoints = 9
                            row.Cells("Address").Appearance.ForeColor = Color.Red
                        Else
                            row.Cells("Address").Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.Default
                            row.Cells("Address").Appearance.FontData.SizeInPoints = 8
                            row.Cells("Address").Appearance.ForeColor = Color.Black
                        End If
                    End If


                Next


                '--------------------------------------------------------------------------------
            End If

        Catch ex As Exception

            ugGrid.DataSource = Nothing
            ugGrid.Controls.Clear()
            Dim pnl As New TextBox
            pnl.Top = 0
            pnl.Left = 0
            pnl.Width = ugGrid.Width
            pnl.Height = ugGrid.Height
            pnl.Text = "Please wait until the SQL database is updated with the new contact code"
            ugGrid.Controls.Add(pnl)
            ugGrid.Enabled = False

        End Try
    End Sub
#End Region


#Region "Owner"

    Friend Shared Sub PopulateOwnerInfo(ByVal OwnerID As Integer, ByRef pown As Object, ByRef frm As Form)
        Dim MyGuid As System.Guid

        Try
            Dim obj As Object
            Dim mtype As BusinessLogic.pFacility.FacilityModule



            If TypeOf frm Is Technical Then
                obj = CType(frm, Technical)
                mtype = BusinessLogic.pFacility.FacilityModule.Technical
            ElseIf TypeOf frm Is Registration Then
                obj = CType(frm, Registration)
                mtype = BusinessLogic.pFacility.FacilityModule.Inspection
            ElseIf TypeOf frm Is Closure Then
                obj = CType(frm, Closure)
                mtype = BusinessLogic.pFacility.FacilityModule.Closure
            ElseIf TypeOf frm Is Financial Then
                obj = CType(frm, Financial)
                mtype = BusinessLogic.pFacility.FacilityModule.Financial
            ElseIf TypeOf frm Is Fees Then
                obj = CType(frm, Fees)
                mtype = BusinessLogic.pFacility.FacilityModule.Fees
            ElseIf TypeOf frm Is CandE Then
                obj = CType(frm, CandE)
                mtype = BusinessLogic.pFacility.FacilityModule.Closure
            Else
                mtype = BusinessLogic.pFacility.FacilityModule.ALL
            End If

            With obj

                .lblOwnerLastEditedBy.Text = String.Empty
                .lblOwnerLastEditedOn.Text = String.Empty

                pown.Retrieve(OwnerID, "SELF", False, True)

                MyGuid = obj.MyGuid

                .lblOwnerIDValue.Text = IIf(pown.id > 0, pown.ID, String.Empty)
                .Tag = pown.ID

                If pown.OrganizationID = 0 Then
                    strName = IIf(pown.BPersona.Title.Trim.Length > 0, pown.BPersona.Title.ToString() + " ", "") + pown.BPersona.FirstName.ToString() + " " + IIf(pown.BPersona.MiddleName.Trim.Length > 0, pown.BPersona.MiddleName.ToString() + " ", "") + pown.BPersona.LastName.ToString() + IIf(pown.BPersona.Suffix.Trim.Length > 0, " " + pown.BPersona.Suffix.ToString(), "")
                ElseIf pown.PersonID = 0 Then
                    strName = pown.Organization.Company
                End If

                .txtOwnerName.Text = strName
                .txtOwnerAddress.Tag = pown.AddressId
                .txtOwnerAddress.Text = IIf(pown.AddressId > 0, FormatAddress(pown.Addresses), String.Empty)
                .txtOwnerEmail.Text = pown.EmailAddress

                .mskTxtOwnerPhone.SelText = IIf(pown.PhoneNumberOne.Length = 0, "", Trim(pown.PhoneNumberOne))
                .mskTxtOwnerPhone2.SelText = IIf(pown.PhoneNumberTwo.Length = 0, "", Trim(pown.PhoneNumberTwo))
                .mskTxtOwnerFax.SelText = IIf(pown.Fax.Length = 0, "", Trim(pown.Fax))

                UIUtilsGen.ValidateComboBoxItemByValue(obj.cmbOwnerType, pown.OwnerType)

                If Not mtype = BusinessLogic.pFacility.FacilityModule.Fees Then

                    .chkOwnerAgencyInterest.Checked = pown.EnsiteAgencyInterestID

                    If pown.EnsiteOrganizationID <> 0 AndAlso pown.EnsitePersonID = 0 Then
                        .txtOwnerAIID.Text = pown.EnsiteOrganizationID.ToString
                    ElseIf pown.EnsiteOrganizationID = 0 AndAlso pown.EnsitePersonID <> 0 Then
                        .txtOwnerAIID.Text = pown.EnsitePersonID.ToString
                    Else
                        .txtOwnerAIID.Text = String.Empty
                    End If

                End If

                .lblOwnerLastEditedBy.Text = String.Format("Last Edited By : {0}", IIf(pown.ModifiedBy = String.Empty, pown.CreatedBy, pown.ModifiedBy))
                .lblOwnerLastEditedOn.Text = String.Format("Last Edited On : {0:d}", IIf(pown.ModifiedOn = CDate("01/01/0001"), pown.CreatedOn, pown.ModifiedOn))

                'To Generate New Owner Registration letter.
                If mtype = BusinessLogic.pFacility.FacilityModule.Registration Then

                    .txtOwnerAIID.Enabled = True
                    .chkCAPParticipant.Checked = pown.ComplianceStatus
                    .lblNewOwnerSnippetValue.Text = IIf(pown.OwnerL2CSnippet, 1, 0)
                    .lblCAPParticipationLevel.Text = pown.CAPParticipationLevel

                End If

                If mtype = BusinessLogic.pFacility.FacilityModule.Compliance OrElse mtype = BusinessLogic.pFacility.FacilityModule.Closure OrElse _
                    mtype = BusinessLogic.pFacility.FacilityModule.Financial OrElse mtype = BusinessLogic.pFacility.FacilityModule.Technical Then

                    .lblCAPParticipationLevel.Text = pown.CAPParticipationLevel
                End If

            End With

            With MusterContainer.AppSemaphores

                .Retrieve(MyGuid.ToString, "OwnerID", obj.lblOwnerIDValue.Text, frm.Name)
                .Retrieve(MyGuid.ToString, "OwnerName", obj.txtOwnerName.Text, frm.Name)
                .Retrieve(MyGuid.ToString, "OwnerAddress", obj.txtOwnerAddress.Text, frm.Name)

            End With


            If Not (TypeOf frm Is Fees Or TypeOf frm Is CandE) Then
                PopulateOwnerFacilities(pown, frm, pown.ID)
            End If

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Friend Shared Sub PopulateOwnerFacilities(ByRef pOwn As Object, ByRef frm As Form, Optional ByVal nOwnerId As Integer = 0)
        Dim dtOwnFacilities As DataTable
        Dim dtDrOwnFacilities As DataRow
        Dim flag As Boolean
        Dim strFacAddress() As String
        Dim MyGuid As System.Guid
        Dim strFacilityAddress As String
        Dim str As String = String.Empty
        Dim rowcount As Integer = 0
        Dim obj As Object
        Try

            If TypeOf frm Is Technical Then
                obj = CType(frm, Technical)
            ElseIf TypeOf frm Is Registration Then
                obj = CType(frm, Registration)
            ElseIf TypeOf frm Is Closure Then
                obj = CType(frm, Closure)
            ElseIf TypeOf frm Is Financial Then
                obj = CType(frm, Financial)
            ElseIf TypeOf frm Is Fees Then
                obj = CType(frm, Fees)
            ElseIf TypeOf frm Is CandE Then
                obj = CType(frm, CandE)
            End If

            If TypeOf frm Is Technical Then
                dtOwnFacilities = pOwn.FacilitiesLUSTSummaryTable
            ElseIf TypeOf frm Is Financial Then
                dtOwnFacilities = pOwn.FacilitiesFinancialSummaryTable
            ElseIf TypeOf frm Is CandE Then
                dtOwnFacilities = pOwn.GetFacilitiesCAESummary
            Else
                dtOwnFacilities = pOwn.FacilitiesTankStatusTable
            End If

            If nOwnerId <> 0 Then
            Else
                flag = True
            End If
            MyGuid = obj.MyGuid
            If Not IsNothing(dtOwnFacilities) Then
                obj.ugFacilityList.DataSource = Nothing
                obj.ugFacilityList.DataBind()
                obj.ugFacilityList.DataSource = dtOwnFacilities
                obj.ugFacilityList.DataBind()

                obj.ugFacilityList.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
                obj.ugFacilityList.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti

                obj.ugFacilityList.DisplayLayout.Bands(0).Columns(UCase("FacilityID")).Width = 50
                obj.ugFacilityList.DisplayLayout.Bands(0).Columns(UCase("FacilityID")).Header.Caption = "Facility ID"
                obj.ugFacilityList.DisplayLayout.Bands(0).Columns(UCase("Facility Name")).Width = 150
                obj.ugFacilityList.DisplayLayout.Bands(0).Columns("ADDRESS").Width = 200
                obj.ugFacilityList.DisplayLayout.Bands(0).Columns("ADDRESS").Header.Caption = "Address"
                obj.ugFacilityList.DisplayLayout.Bands(0).Columns("CITY").Width = 100
                obj.ugFacilityList.DisplayLayout.Bands(0).Columns("CITY").Header.Caption = "City"
                obj.ugFacilityList.DisplayLayout.Bands(0).Columns("COUNTY").Width = 100
                obj.ugFacilityList.DisplayLayout.Bands(0).Columns("COUNTY").Header.Caption = "County"

                If Not (TypeOf frm Is Technical Or TypeOf frm Is Financial) Then
                    obj.ugFacilityList.DisplayLayout.Bands(0).Columns("CIU").Width = 40
                    obj.ugFacilityList.DisplayLayout.Bands(0).Columns("TOSI").Width = 40
                    obj.ugFacilityList.DisplayLayout.Bands(0).Columns("TOS").Width = 40
                    'obj.ugFacilityList.DisplayLayout.Bands(0).Columns("CP").Width = 40
                    obj.ugFacilityList.DisplayLayout.Bands(0).Columns("CP").Hidden = True
                    obj.ugFacilityList.DisplayLayout.Bands(0).Columns("POU").Width = 40
                    obj.ugFacilityList.DisplayLayout.Bands(0).Columns(UCase("Total")).Width = 75
                End If

                If obj.ugFacilityList.Rows.Count > 0 Then
                    obj.ugFacilityList.ActiveRow = obj.ugFacilityList.Rows(0)
                End If
                obj.ugFacilityList.DisplayLayout.Bands(0).Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
                obj.strFacilityIdTags = String.Empty
                For Each dtDrOwnFacilities In dtOwnFacilities.Rows
                    If rowcount < dtOwnFacilities.Rows.Count - 1 Then
                        str = ","
                    Else
                        str = ""
                    End If
                    obj.strFacilityIdTags += dtDrOwnFacilities("FacilityID").ToString + str
                    rowcount += 1
                Next
                MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "OwnerFacilities", obj.strFacilityIdTags, "Registration")
                If flag = True Then
                    Dim i As Integer = 0

                    strFacAddress = obj.txtFacilityAddress.Text.Split(vbCrLf)
                    strFacilityAddress = String.Empty
                    For i = 0 To UBound(strFacAddress)
                        If i = 0 Or i = UBound(strFacAddress) Then
                            strFacilityAddress += Trim(strFacAddress(i).ToString)

                        End If
                    Next

                    MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "FacilityID", obj.lblFacilityIDValue.Text, frm.Name)
                    MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "FacilityName", obj.txtFacilityName.Text, frm.Name)
                    MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "FacilityAddress", Trim(strFacilityAddress), frm.Name)
                    flag = False

                End If


            End If
            obj.lblNoOfFacilitiesValue.Text = obj.ugFacilityList.Rows.Count
            If pOwn.Active = True Then
                obj.lblOwnerActiveOrNot.Text = "ACTIVE" 'IIf(ugFacilityList.Rows.Count = 0, "INACTIVE", "ACTIVE")
            Else
                obj.lblOwnerActiveOrNot.Text = "INACTIVE"
            End If


        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    'if any form is using the owner, leave alone. else remove from collection
    Friend Shared Sub RemoveOwner(ByRef pown As Object, ByVal frm As Form)
        Dim frmMuster As MusterContainer
        Dim frmTech As Technical
        Dim frmClo As Closure
        Dim frmReg As Registration
        Dim frmFinancial As Financial
        Dim frmFees As Fees
        Dim frmCompany As Company
        Try
            Dim obj As Object
            If TypeOf frm Is Technical Then
                obj = CType(frm, Technical)
            ElseIf TypeOf frm Is Registration Then
                obj = CType(frm, Registration)
            ElseIf TypeOf frm Is Closure Then
                obj = CType(frm, Closure)
            ElseIf TypeOf frm Is Financial Then
                obj = CType(frm, Financial)
            ElseIf TypeOf frm Is Fees Then
                obj = CType(frm, Fees)
            ElseIf TypeOf frm Is Company Then
                obj = CType(frm, Company)
            ElseIf TypeOf frm Is CandE Then
                obj = CType(frm, CandE)
            End If

            frmMuster = obj.MdiParent

            For Each frmChild As Form In frmMuster.MdiChildren
                If frmChild.Name <> obj.Name Then
                    Select Case frmChild.Name.ToUpper
                        Case "REGISTRATION"
                            frmReg = CType(frmChild, Registration)
                            If frmReg.lblOwnerIDValue.Text.Trim = pown.ID.ToString Then
                                Exit Sub
                            End If
                        Case "TECHNICAL"
                            frmTech = CType(frmChild, Technical)
                            If frmTech.lblOwnerIDValue.Text.Trim = pown.ID.ToString Then
                                Exit Sub
                            End If
                        Case "CLOSURE"
                            frmClo = CType(frmChild, Closure)
                            If frmClo.lblOwnerIDValue.Text.Trim = pown.ID.ToString Then
                                Exit Sub
                            End If
                        Case "FINANCIAL"
                            frmFinancial = CType(frmChild, Financial)
                            If frmFinancial.lblOwnerIDValue.Text.Trim = pown.ID.ToString Then
                                Exit Sub
                            End If
                        Case "FEES"
                            frmFees = CType(frmChild, Fees)
                            If frmFees.lblOwnerIDValue.Text.Trim = pown.ID.ToString Then
                                Exit Sub
                            End If
                        Case "COMPANY"
                            frmCompany = CType(frmChild, Company)
                            If frmCompany.txtCompanyID.Text.Trim = pown.ID.ToString Then
                                Exit Sub
                            End If
                    End Select
                End If
            Next
            pown.Remove(pown.ID)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Shared Sub UpdateOwnerCAPInfo(ByRef frm As Form, ByRef owner As MUSTER.BusinessLogic.pOwner)
        Dim obj As Object
        If TypeOf frm Is Technical Then
            obj = CType(frm, Technical)
        ElseIf TypeOf frm Is Registration Then
            obj = CType(frm, Registration)
        ElseIf TypeOf frm Is Closure Then
            obj = CType(frm, Closure)
        ElseIf TypeOf frm Is Financial Then
            obj = CType(frm, Financial)
        ElseIf TypeOf frm Is Fees Then
            obj = CType(frm, Fees)
        ElseIf TypeOf frm Is CandE Then
            obj = CType(frm, CandE)
        End If
        If Not owner Is Nothing Then
            Dim ds As DataSet = owner.RunSQLQuery("SELECT ISNULL(CAP_PARTICIPATION_LEVEL,'NONE - 0/0 (Compliant/Candidate)') FROM TBLREG_OWNER WHERE OWNER_ID = " + owner.ID.ToString)
            owner.CAPParticipationLevel = ds.Tables(0).Rows(0)(0)
            If Not obj Is Nothing Then
                obj.lblCAPParticipationLevel.Text = owner.CAPParticipationLevel
            End If
        End If
    End Sub
#Region "Owner Name and Address UI handlers"


    Friend Shared Sub SwapOrgPersonDisplay(ByRef frm As Form)
        Dim obj As Object
        Try
            If TypeOf frm Is Registration Then
                obj = CType(frm, Registration)
            ElseIf TypeOf frm Is Closure Then
                obj = CType(frm, Closure)
            End If
            If obj.rdOwnerPerson.Checked Then
                obj.pnlOwnerPerson.Location = New Point(obj.pnlPersonOrganization.Location.X, obj.pnlPersonOrganization.Location.Y + obj.pnlPersonOrganization.Height)
                obj.pnlOwnerName.Height = obj.pnlPersonOrganization.Height + obj.pnlOwnerPerson.Height + 5 + obj.pnlOwnerNameButton.Height
                obj.pnlOwnerNameButton.Location = New Point(obj.pnlOwnerPerson.Location.X, obj.pnlOwnerPerson.Location.Y + obj.pnlOwnerPerson.Height)
                obj.pnlOwnerOrg.Visible = False
                obj.pnlOwnerPerson.Visible = True
            Else
                obj.pnlOwnerOrg.Location = New Point(obj.pnlPersonOrganization.Location.X, obj.pnlPersonOrganization.Location.Y + obj.pnlPersonOrganization.Height)
                obj.pnlOwnerName.Height = obj.pnlPersonOrganization.Height + obj.pnlOwnerOrg.Height + obj.pnlOwnerNameButton.Height
                obj.pnlOwnerNameButton.Location = New Point(obj.pnlOwnerOrg.Location.X, obj.pnlOwnerOrg.Location.Y + obj.pnlOwnerOrg.Height)
                obj.pnlOwnerOrg.Visible = True
                obj.pnlOwnerPerson.Visible = False
                If obj.bolNewPersona = True Then
                    obj.FormLoading = True
                    ' obj.cmbOwnerOrgEntityCode.SelectedIndex = -1
                    obj.bolNewPersona = False
                    obj.FormLoading = False
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally

        End Try
    End Sub
    Friend Shared Sub rdOwnerPersonClick(ByRef frm As Form, ByRef pOwn As Object)
        Dim msgResult As MsgBoxResult
        Dim obj As Object
        Try
            If TypeOf frm Is Closure Then
                obj = CType(frm, Closure)
            ElseIf TypeOf frm Is Registration Then
                obj = CType(frm, Registration)
            End If
            If pOwn.BPersona.OrgID <> 0 Or pOwn.BPersona.IsDirty Then
                If obj.rdOwnerOrg.Tag = True And obj.rdOwnerPerson.Tag = False Then
                    msgResult = MsgBox(" Do you want to change from Organization to Person", MsgBoxStyle.YesNo, "Persona")
                    If msgResult = MsgBoxResult.Yes Then
                        ClearPersona(frm)
                        'ClearBPersonaOrganization()
                        pOwn.BPersona.Clear()
                        obj.bolNewPersona = True
                        SwapOrgPersonDisplay(frm)
                        obj.rdOwnerOrg.Tag = False
                        obj.rdOwnerPerson.Tag = True
                        obj.cmbOwnerNameTitle.Focus()
                        Dim nAIID As Integer = 0
                        If pOwn.EnsiteOrganizationID <> 0 Then nAIID = pOwn.EnsiteOrganizationID
                        pOwn.EnsiteOrganizationID = 0
                        pOwn.EnsitePersonID = nAIID
                    Else
                        obj.rdOwnerOrg.Checked = True
                        obj.txtOwnerOrgName.Focus()
                    End If
                End If
            Else
                SwapOrgPersonDisplay(frm)
                obj.cmbOwnerNameTitle.Focus()
                '    ClearPersona()
            End If

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Friend Shared Sub rdOwnerOrgClick(ByRef frm As Form, ByRef pOwn As Object)

        Dim msgResult As MsgBoxResult
        Dim obj As Object
        Try
            If TypeOf frm Is Closure Then
                obj = CType(frm, Closure)
            ElseIf TypeOf frm Is Registration Then
                obj = CType(frm, Registration)
            End If
            If pOwn.BPersona.PersonId <> 0 Or pOwn.BPersona.IsDirty Then
                If obj.rdOwnerOrg.Tag = False And obj.rdOwnerPerson.Tag = True Then
                    msgResult = MsgBox(" Do you want to change from Person to Organization", MsgBoxStyle.YesNo, "Persona")
                    If msgResult = MsgBoxResult.Yes Then
                        ClearPersona(frm)
                        pOwn.BPersona.Clear()
                        obj.bolNewPersona = True
                        SwapOrgPersonDisplay(frm)
                        obj.rdOwnerOrg.Tag = True
                        obj.rdOwnerPerson.Tag = False
                        obj.txtOwnerOrgName.Focus()
                        Dim nAIID As Integer = 0
                        If pOwn.EnsitePersonID <> 0 Then nAIID = pOwn.EnsitePersonID
                        pOwn.EnsiteOrganizationID = nAIID
                        pOwn.EnsitePersonID = 0
                    Else
                        obj.rdOwnerPerson.Checked = True
                        obj.cmbOwnerNameTitle.Focus()
                    End If
                End If
            Else
                SwapOrgPersonDisplay(frm)
                obj.txtOwnerOrgName.Focus()
                obj.FormLoading = True
                ' obj.cmbOwnerOrgEntityCode.SelectedIndex = -1
                obj.FormLoading = False
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Friend Shared Sub setupOwnername(ByRef frm As Form, ByRef pOwn As Object)
        Dim obj As Object
        Try
            If TypeOf frm Is Closure Then
                obj = CType(frm, Closure)
            ElseIf TypeOf frm Is Registration Then
                obj = CType(frm, Registration)
                obj.pnlPersonOrganization.BackColor = SystemColors.ControlLightLight
                obj.pnlOwnerName.BackColor = SystemColors.ControlLightLight
                obj.pnlOwnerPerson.BackColor = SystemColors.ControlLightLight
                obj.pnlOwnerOrg.BackColor = SystemColors.ControlLightLight
            End If
            If Not obj.rdOwnerOrg.Checked Then
                obj.cmbOwnerNameTitle.Focus()
                'obj.rdOwnerPerson.Focus()
            Else
                obj.txtOwnerOrgName.Focus()
            End If
            If obj.txtOwnerName.Text = String.Empty Then
                obj.bolNewPersona = True
                CheckUncheckPersonaOrg(frm, False, False)
                ClearPersona(frm)

            Else
                obj.bolNewPersona = False
                ResetOwnerName(frm, pOwn)
            End If
            obj.pnlOwnerName.Location = New Point(obj.txtOwnerName.Location.X, obj.txtOwnerName.Location.Y)
            SwapOrgPersonDisplay(frm)
            obj.pnlOwnerName.BringToFront()
            obj.pnlOwnerName.Visible = True
            If pOwn.BPersona.Org_Entity_Code = 0 Then
                obj.FormLoading = True
                '  obj.cmbOwnerOrgEntityCode.SelectedIndex = -1
                obj.FormLoading = False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Shared Sub FillPersona(ByRef cntrl As Control, ByRef pOwn As BusinessLogic.pOwner)
        Try

            Select Case cntrl.Name
                Case "cmbOwnerNameTitle"
                    If CType(cntrl, System.Windows.Forms.ComboBox).SelectedIndex <> -1 Then
                        pOwn.BPersona.Title = CType(cntrl, System.Windows.Forms.ComboBox).Text
                    Else
                        pOwn.BPersona.Title = String.Empty
                    End If
                Case "txtOwnerFirstName"
                    If cntrl.Text <> String.Empty Then
                        pOwn.BPersona.FirstName = cntrl.Text
                    Else
                        pOwn.BPersona.FirstName = String.Empty
                    End If
                Case "txtOwnerLastName"
                    If cntrl.Text <> String.Empty Then
                        pOwn.BPersona.LastName = cntrl.Text
                    Else
                        pOwn.BPersona.LastName = String.Empty
                    End If
                Case "txtOwnerMiddleName"
                    If cntrl.Text <> String.Empty Then
                        pOwn.BPersona.MiddleName = cntrl.Text
                    Else
                        pOwn.BPersona.MiddleName = String.Empty
                    End If
                Case "cmbOwnerNameSuffix"
                    If CType(cntrl, System.Windows.Forms.ComboBox).SelectedIndex <> -1 Then
                        pOwn.BPersona.Suffix = CType(cntrl, System.Windows.Forms.ComboBox).Text
                    Else
                        pOwn.BPersona.Suffix = String.Empty
                    End If
                Case "txtOwnerOrgName"
                    If cntrl.Text <> String.Empty Then
                        pOwn.BPersona.Company = cntrl.Text
                    Else
                        pOwn.BPersona.Company = String.Empty
                    End If
                    ' Case "cmbOwnerOrgEntityCode"
                    '    If CType(cntrl, System.Windows.Forms.ComboBox).SelectedValue > 0 Then
                    '   pOwn.BPersona.Org_Entity_Code = CType(cntrl, System.Windows.Forms.ComboBox).SelectedValue
                    '  Else
                    '     pOwn.BPersona.Org_Entity_Code = 0

                    ' End If
            End Select

            If pOwn.OrganizationID > 0 AndAlso Not pOwn.Organization Is Nothing AndAlso pOwn.Organization.Org_Entity_Code = 0 Then
                pOwn.Organization.Org_Entity_Code = 539
            End If



        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Shared Sub SetOwnerName(ByRef frm As Form)
        Dim strOwnerName As String = String.Empty
        Dim obj As Object
        Try

            If TypeOf frm Is Closure Then
                obj = CType(frm, Closure)
            ElseIf TypeOf frm Is Registration Then
                obj = CType(frm, Registration)
            End If
            If obj.rdOwnerOrg.Checked Then
                strOwnerName = obj.txtOwnerOrgName.Text
            Else
                strOwnerName = IIf(obj.cmbOwnerNameTitle.Text.Trim.Length > 0, obj.cmbOwnerNameTitle.Text + " ", "") + obj.txtOwnerFirstName.Text + " " + IIf(obj.txtOwnerMiddleName.Text.Trim.Length > 0, obj.txtOwnerMiddleName.Text + " ", "") + obj.txtOwnerLastName.Text + IIf(obj.cmbOwnerNameSuffix.Text.Trim.Length > 0, " " + obj.cmbOwnerNameSuffix.Text, "")
            End If
            obj.txtOwnerName.Text = strOwnerName
        Catch ex As Exception
            Throw ex
        End Try


    End Sub
    Friend Shared Sub CheckUncheckPersonaOrg(ByRef frm As Form, ByVal bolPerson As Boolean, ByVal bolOrg As Boolean)
        Dim obj As Object
        Try


            If TypeOf frm Is Closure Then
                obj = CType(frm, Closure)
            ElseIf TypeOf frm Is Registration Then
                obj = CType(frm, Registration)
            End If
            obj.rdOwnerOrg.Checked = bolOrg
            obj.rdOwnerPerson.Checked = bolPerson
            obj.rdOwnerOrg.Tag = bolOrg
            obj.rdOwnerPerson.Tag = bolPerson
            SwapOrgPersonDisplay(frm)
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Friend Shared Sub ClearPersona(ByRef frm As Form)
        Dim obj As Object
        Try
            If TypeOf frm Is Closure Then
                obj = CType(frm, Closure)
            ElseIf TypeOf frm Is Registration Then
                obj = CType(frm, Registration)
            End If
            If obj.rdOwnerPerson.Checked = False And obj.rdOwnerOrg.Checked = False Then
                obj.rdOwnerPerson.Checked = True
            End If
            obj.FormLoading = True
            obj.txtOwnerOrgName.Text = String.Empty
            obj.cmbOwnerNameTitle.SelectedIndex = -1
            obj.cmbOwnerNameSuffix.SelectedIndex = -1
            obj.txtOwnerFirstName.Text = String.Empty
            obj.txtOwnerLastName.Text = String.Empty
            obj.txtOwnerMiddleName.Text = String.Empty
            'obj.cmbOwnerOrgEntityCode.SelectedIndex = -1
            'obj.cmbOwnerOrgEntityCode.SelectedIndex = -1
            obj.FormLoading = False
        Catch ex As Exception
            Throw ex
        End Try


    End Sub

    Friend Shared Function ResetOwnerName(ByRef frm As Form, ByRef pOwn As Object) As String
        Dim oPersonaInfo As MUSTER.Info.PersonaInfo
        Dim strOwnerName As String = String.Empty
        Dim obj As Object
        Try
            If TypeOf frm Is Closure Then
                obj = CType(frm, Closure)
            ElseIf TypeOf frm Is Registration Then
                obj = CType(frm, Registration)
            End If
            If pOwn.PersonID = 0 Then
                CheckUncheckPersonaOrg(frm, False, True)
                oPersonaInfo = pOwn.Organization()
                obj.txtOwnerOrgName.Text = IIf(IsNothing(pOwn.BPersona.Company), String.Empty, CStr(pOwn.BPersona.Company))
                'UIUtilsGen.ValidateComboBoxItemByValue(obj.cmbOwnerOrgEntityCode, pOwn.BPersona.Org_Entity_Code)
                strOwnerName = obj.txtOwnerOrgName.Text
            Else
                CheckUncheckPersonaOrg(frm, True, False)
                oPersonaInfo = pOwn.Persona()
                obj.cmbOwnerNameTitle.Text = IIf(IsNothing(Trim(pOwn.BPersona.Title)), String.Empty, CStr(Trim(pOwn.BPersona.Title)))
                obj.txtOwnerFirstName.Text = IIf(IsNothing(pOwn.BPersona.FirstName), String.Empty, CStr(pOwn.BPersona.FirstName))
                obj.txtOwnerLastName.Text = IIf(IsNothing(pOwn.BPersona.LastName), String.Empty, CStr(pOwn.BPersona.LastName))
                obj.cmbOwnerNameSuffix.Text = IIf(pOwn.BPersona.Suffix = String.Empty, String.Empty, CStr(Trim(pOwn.BPersona.Suffix)))
                obj.txtOwnerMiddleName.Text = IIf(IsNothing(pOwn.BPersona.MiddleName), String.Empty, CStr(pOwn.BPersona.MiddleName))
                strOwnerName = IIf(obj.cmbOwnerNameTitle.Text.Trim.Length > 0, obj.cmbOwnerNameTitle.Text.ToString() + " ", "") + obj.txtOwnerFirstName.Text.ToString() + " " + IIf(obj.txtOwnerMiddleName.Text.Trim.Length > 0, obj.txtOwnerMiddleName.Text.ToString() + " ", "") + obj.txtOwnerLastName.Text.ToString() + IIf(obj.cmbOwnerNameSuffix.Text.Trim.Length > 0, " " + obj.cmbOwnerNameSuffix.Text.ToString(), "")
            End If
            Return strOwnerName
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Friend Shared Sub OwnerNameCancel(ByRef frm As Form, ByRef pOwn As Object)
        Try
            If Not pOwn.BPersona Is Nothing Then
                pOwn.BPersona.Reset()
                ClearPersona(frm)
                If pOwn.BPersona.PersonId <> 0 Or pOwn.BPersona.OrgID <> 0 Then
                    ResetOwnerName(frm, pOwn)
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#End Region

#Region "facility old code"

    'Friend Shared Sub PopulateFacilityInfo(ByRef frm As Form, ByRef pFacilities As Object, ByVal [Module] As String, ByRef frmPage As Windows.Forms.TabPage, ByRef [Guid] As System.Guid, Optional ByVal FacilityID As Integer = 0)
    'Friend Shared Sub PopulateFacilityInfo(ByRef frm As Form, ByRef ownerInfo As MUSTER.Info.OwnerInfo, ByRef pFacilities As BusinessLogic.pFacility, Optional ByVal FacilityID As Integer = 0)
    'Dim oAddressInfo As MUSTER.Info.AddressInfo

    ''case registration old code
    'obj.PipeVisible()
    'obj.TankVisible()
    ''obj.btnAddTank.Enabled = True
    'obj.GetTankandPipeForFacility()


    ''older code
    '    If Not pFacilities Is Nothing Then
    '        obj.lblFacilityIDValue.Text = pFacilities.ID
    '        If pFacilities.ID < 0 Then
    '            obj.lblFacilityIDValue.Visible = False
    '        Else
    '            obj.lblFacilityIDValue.Visible = True
    '        End If
    '        obj.lblDateTransfered.Text = IIf(InfoRepository.Utils.IsDateNull(pFacilities.DateTransferred), String.Empty, pFacilities.DateTransferred)
    '        obj.txtFacilityName.Text = pFacilities.Name
    '        'oAddressInfo = pFacilities.FacilityAddress
    '        obj.txtFacilityAddress.Tag = pFacilities.AddressId
    '        UIUtilsGen.SetComboboxItemByValue(obj.cmbFacilityType, pFacilities.FacilityType)
    '        strFacilityAddress = pFacilities.FacilityAddresses.AddressLine1 + IIf(pFacilities.FacilityAddresses.AddressLine2.Length = 0, String.Empty, vbCrLf + pFacilities.FacilityAddresses.AddressLine2.ToString) + IIf(pFacilities.FacilityAddresses.City.Trim.ToString <> String.Empty, vbCrLf + pFacilities.FacilityAddresses.City.Trim.ToString + " ,", String.Empty) + pFacilities.FacilityAddresses.State.Trim.ToString + IIf(pFacilities.FacilityAddresses.Zip.Trim.ToString <> String.Empty, " " + pFacilities.FacilityAddresses.Zip.Trim.ToString, String.Empty)
    '        obj.txtFacilityAddress.Text = strFacilityAddress
    '        UIUtilsGen.SetMaskEditText(obj.mskTxtFacilityPhone, pFacilities.Phone)
    '        obj.mskTxtFacilityPhone.Tag = obj.mskTxtFacilityPhone.SelText.ToString
    '        obj.mskTxtFacilityFax.ResetText()
    '        obj.mskTxtFacilityFax.Tag = String.Empty
    '        UIUtilsGen.SetMaskEditText(obj.mskTxtFacilityFax, pFacilities.Fax)
    '        obj.mskTxtFacilityFax.Tag = obj.mskTxtFacilityFax.SelText.ToString
    '        UIUtilsGen.ValidateComboBoxItemByValue(obj.cmbFacilityType, pFacilities.FacilityType)
    '        If obj.FormLoading = True Then obj.FormLoading = False
    '        obj.txtFacilityLatDegree.Text = IIf(pFacilities.LatitudeDegree < 0, String.Empty, pFacilities.LatitudeDegree)
    '        obj.txtFacilityLatMin.Text = IIf(pFacilities.LatitudeMinutes < 0, String.Empty, pFacilities.LatitudeMinutes)
    '        obj.txtFacilityLatSec.Text = IIf(pFacilities.LatitudeSeconds < 0, String.Empty, pFacilities.LatitudeSeconds)
    '        obj.txtFacilityLongDegree.Text = IIf(pFacilities.LongitudeDegree < 0, String.Empty, pFacilities.LongitudeDegree)
    '        obj.txtFacilityLongMin.Text = IIf(pFacilities.LongitudeMinutes < 0, String.Empty, pFacilities.LongitudeMinutes)
    '        obj.txtFacilityLongSec.Text = IIf(pFacilities.LongitudeSeconds < 0, String.Empty, pFacilities.LongitudeSeconds)
    '        'Check_If_Datum_Enable(sender, e)
    '        UIUtilsGen.ValidateComboBoxItemByValue(obj.cmbFacilityDatum, pFacilities.Datum)
    '        UIUtilsGen.ValidateComboBoxItemByValue(obj.cmbFacilityMethod, pFacilities.Method)
    '        UIUtilsGen.ValidateComboBoxItemByValue(obj.cmbFacilityLocationType, pFacilities.LocationType)
    '        UIUtilsGen.SetDatePickerValue(obj.dtPickFacilityRecvd, pFacilities.DateReceived)
    '        UIUtilsGen.SetDatePickerValue(obj.dtFacilityPowerOff, pFacilities.DatePowerOff)
    '        obj.txtFuelBrand.Text = IIf(pFacilities.FuelBrand.Length = 0, "", pFacilities.FuelBrand)
    '        obj.lblFacilityStatusValue.Text = pFacilities.FacilityStatusDescription
    '        obj.lblFacilityStatusValue.Tag = pFacilities.FacilityStatus
    '        If obj.lblFacilityStatusValue.Text.ToUpper = "Active".Trim.ToUpper Then
    '            obj.lblFacilityStatusValue.BackColor = Color.Green
    '        ElseIf obj.lblFacilityStatusValue.Text.ToUpper = "CLOSED".Trim.ToUpper Then
    '            obj.lblFacilityStatusValue.BackColor = Color.White
    '        ElseIf obj.lblFacilityStatusValue.Text.ToUpper = "Pre 88".Trim.ToUpper Then
    '            obj.lblFacilityStatusValue.BackColor = Color.Orange
    '        Else
    '            obj.lblFacilityStatusValue.Backcolor = Color.Transparent
    '        End If
    '        ' obj.lblDateTransfered.Text = IIf(InfoRepository.Utils.IsDateNull(pFacilities.DateTransferred), String.Empty, pFacilities.DateTransferred)
    '        If TypeOf frm Is Registration Then
    '            If pFacilities.SignatureOnNF = False Then
    '                obj.SignatureFlag = True
    '            End If
    '        End If
    '        obj.chkSignatureofNF.Checked = pFacilities.SignatureOnNF
    '        If obj.chkSignatureofNF.Checked Then
    '            obj.txtDueByNF.Text = String.Empty
    '        Else
    '            obj.txtDueByNF.Text = "Due"
    '        End If
    '        obj.chkCAPCandidate.Checked = pFacilities.CAPCandidate
    '        If pFacilities.GetCapStatus(pFacilities.ID, False) = 1 Then
    '            obj.lblCAPStatusValue.BackColor = Color.Green
    '            obj.lblCAPStatusValue.Text = "Compliant"
    '        Else
    '            obj.lblCAPStatusValue.BackColor = Color.Red
    '            obj.lblCAPStatusValue.Text = "Not Compliant"
    '        End If
    '        obj.chkUpcomingInstall.Checked = pFacilities.UpcomingInstallation
    '        obj.chkUpcomingInstall.Tag = pFacilities.UpcomingInstallation
    '        UIUtilsGen.SetDatePickerValue(obj.dtPickUpcomingInstallDateValue, pFacilities.UpcomingInstallationDate)
    '        obj.dtPickUpcomingInstallDateValue.Enabled = True
    '        If obj.lblOwnerIDValue.Text = String.Empty Then
    '            obj.PopulateOwnerInfo(pFacilities.OwnerID)
    '        End If
    '        obj.Tag = obj.lblFacilityIDValue.Text
    '        'obj.Text = IIf(TypeOf frm Is Registration, "Registration", "Technical") & " - Facility Detail - (" & obj.txtFacilityName.Text & ")"
    '        If TypeOf frm Is Registration Then
    '            obj.Text = "Registration" & " - Facility Detail - (" & obj.txtFacilityName.Text & ")"
    '        ElseIf TypeOf frm Is Technical Then
    '            obj.Text = "Technical" & " - Facility Detail - (" & obj.txtFacilityName.Text & ")"
    '        ElseIf TypeOf frm Is Closure Then
    '            obj.Text = "Closure" & " - Facility Detail - (" & obj.txtFacilityName.Text & ")"
    '        ElseIf TypeOf frm Is Financial Then
    '            obj.Text = "Financial" & " - Facility Detail - (" & obj.txtFacilityName.Text & ")"
    '            'P1 05/14/05 end
    '        End If
    '        If TypeOf frm Is Registration Then
    '            obj.lblOwnerLastEditedBy.Text = "Last Edited By : " & IIf(pFacilities.ModifiedBy = String.Empty, pFacilities.CreatedBy.ToString, pFacilities.ModifiedBy.ToString)
    '            obj.lblOwnerLastEditedOn.Text = "Last Edited On : " & IIf(pFacilities.ModifiedOn = CDate("01/01/0001"), pFacilities.CreatedOn.ToString, pFacilities.ModifiedOn.ToString)
    '            obj.PipeVisible()
    '            obj.TankVisible()
    '            'obj.btnAddTank.Enabled = True
    '            obj.GetTankandPipeForFacility()
    '        ElseIf TypeOf frm Is Technical Then
    '            obj.GetLUSTEventsForFacility()
    '        End If
    '        MusterContainer.AppSemaphores.Retrieve(fGUID.ToString, "FacilityDetails", obj.lblFacilityIDValue.Text, frm.Name)
    '        MusterContainer.AppSemaphores.Retrieve(fGUID.ToString, "FacilityID", obj.lblFacilityIDValue.Text, frm.Name)
    '        MusterContainer.AppSemaphores.Retrieve(fGUID.ToString, "FacilityName", obj.txtFacilityName.Text, frm.Name)
    '        MusterContainer.AppSemaphores.Retrieve(fGUID.ToString, "FacilityAddress", strFacilityAddress, frm.Name)
    '        'Me.EnableDiableFacilityControls()
    '    End If

    'Catch ex As Exception
    '    Throw ex
    'End Try


    'End Sub
#End Region
#Region "Facility"

    Public Shared Sub OperatorTextChnage(ByVal sender As Object, ByVal e As EventArgs)

        With DirectCast(sender, TextBox)
            If .Text.ToUpper.IndexOf(" (SOC") > -1 OrElse .Text.ToUpper.IndexOf(" (OOC)") > -1 Then
                .BackColor = Color.DarkRed
            Else
                .BackColor = Color.White
            End If

        End With
    End Sub


    Friend Shared Sub PopulateFacilityInfo(ByRef frm As Form, ByRef ownerInfo As MUSTER.Info.OwnerInfo, ByRef pFacilities As BusinessLogic.pFacility, Optional ByVal FacilityID As Integer = 0)

        Dim fGUID As System.Guid
        Dim sender As Object
        Dim e As System.ComponentModel.CancelEventArgs
        Dim obj As Object
        Dim mtype As BusinessLogic.pFacility.FacilityModule
        Dim CandEFlag As Short

        Try
            CandEFlag = 0
            If TypeOf frm Is Technical Then
                obj = CType(frm, Technical)
                mtype = BusinessLogic.pFacility.FacilityModule.Technical
            ElseIf TypeOf frm Is Registration Then
                obj = CType(frm, Registration)
                mtype = BusinessLogic.pFacility.FacilityModule.Registration
            ElseIf TypeOf frm Is Closure Then
                obj = CType(frm, Closure)
                mtype = BusinessLogic.pFacility.FacilityModule.Closure
            ElseIf TypeOf frm Is Financial Then
                obj = CType(frm, Financial)
                mtype = BusinessLogic.pFacility.FacilityModule.Financial
            ElseIf TypeOf frm Is Fees Then
                obj = CType(frm, Fees)
                mtype = BusinessLogic.pFacility.FacilityModule.Fees
            ElseIf TypeOf frm Is CandE Then
                obj = CType(frm, CandE)
                CandEFlag = 1
                mtype = BusinessLogic.pFacility.FacilityModule.Compliance
            Else
                mtype = BusinessLogic.pFacility.FacilityModule.Inspection
            End If

            fGUID = obj.MyGuid
            If CandEFlag = 0 Then
                If obj.lblFacilityIDValue.Text Is Nothing OrElse obj.lblFacilityIDValue.Text <> FacilityID.ToString() Then
                    obj.txtFacilityAddress.Text = String.Empty
                End If
            Else
                obj.txtFacilityAddress.Text = String.Empty
            End If

            If mtype <> BusinessLogic.pFacility.FacilityModule.Compliance AndAlso mtype <> BusinessLogic.pFacility.FacilityModule.Inspection Then
                pFacilities.Retrieve(ownerInfo, IIf(FacilityID <> 0, FacilityID, pFacilities.ID), , "FACILITY", False, True)
            Else
                pFacilities.Retrieve(ownerInfo, IIf(FacilityID <> 0, FacilityID, pFacilities.ID), , "FACILITY", False, True)
            End If


            If Not pFacilities Is Nothing Then

                With obj



                    If Not (mtype = BusinessLogic.pFacility.FacilityModule.Fees Or mtype = BusinessLogic.pFacility.FacilityModule.Compliance) Then
                        .lblDateTransfered.Text = IIf(Date.Compare(pFacilities.DateTransferred, CDate("01/01/0001")) = 0, String.Empty, pFacilities.DateTransferred)
                    End If

                    If mtype = BusinessLogic.pFacility.FacilityModule.Compliance OrElse _
                       mtype = BusinessLogic.pFacility.FacilityModule.Inspection Then

                        RemoveHandler DirectCast(.txtDesOp, TextBox).TextChanged, AddressOf OperatorTextChnage
                        AddHandler DirectCast(.txtDesOp, TextBox).TextChanged, AddressOf OperatorTextChnage

                        DirectCast(.txtDesOp, TextBox).Text = pFacilities.DesignatedOperator

                    ElseIf mtype = BusinessLogic.pFacility.FacilityModule.Registration Then

                        RemoveHandler DirectCast(.txtDesignatedOperator, TextBox).TextChanged, AddressOf OperatorTextChnage
                        AddHandler DirectCast(.txtDesignatedOperator, TextBox).TextChanged, AddressOf OperatorTextChnage

                    End If


                    .lblFacilityIDValue.Text = pFacilities.ID
                    .lblFacilityIDValue.Visible = (pFacilities.ID >= 0)

                    .txtFacilityName.Text = pFacilities.Name
                    If obj.lblFacilityIDValue.Text Is Nothing OrElse obj.lblFacilityIDValue.Text <> FacilityID.ToString() Then
                        pFacilities.AddressID = 0
                    End If
                    .txtFacilityAddress.Tag = pFacilities.AddressID


                    UIUtilsGen.SetComboboxItemByValue(obj.cmbFacilityType, pFacilities.FacilityType, True)

                    .txtFacilitySIC.Text = .cmbFacilityType.Text


                    If pFacilities.FacilityAddresses.AddressId = 0 Then
                        pFacilities.FacilityAddresses.Retrieve(pFacilities.AddressID)
                    End If


                    Dim newAddress As String = IIf(pFacilities.AddressID > 0, FormatAddress(pFacilities.FacilityAddresses, True), String.Empty)

                    If .txtFacilityAddress.Text = String.Empty Then
                        .txtFacilityAddress.Text = newAddress
                    End If



                    .chkLustSite.checked = (pFacilities.CurrentLUSTStatus > 0)
                    .chkLustSite.enabled = False


                    UIUtilsGen.ValidateComboBoxItemByValue(.cmbFacilityType, pFacilities.FacilityType)
                    UIUtilsGen.SetMaskEditText(.mskTxtFacilityPhone, pFacilities.Phone)

                    .mskTxtFacilityPhone.Tag = .mskTxtFacilityPhone.SelText.ToString
                    .mskTxtFacilityFax.ResetText()
                    .mskTxtFacilityFax.Tag = String.Empty

                    UIUtilsGen.SetMaskEditText(.mskTxtFacilityFax, pFacilities.Fax)
                    .mskTxtFacilityFax.Tag = .mskTxtFacilityFax.SelText.ToString

                    UIUtilsGen.SetDatePickerValue(.dtPickFacilityRecvd, pFacilities.DateReceived)

                    .lblFacilityStatusValue.Text = pFacilities.FacilityStatusDescription
                    .lblFacilityStatusValue.Tag = pFacilities.FacilityStatus

                    With .lblFacilityStatusValue
                        If .Text.ToUpper = "Active".Trim.ToUpper Then
                            .BackColor = Color.Green
                        ElseIf .Text.ToUpper = "CLOSED".Trim.ToUpper Then
                            .BackColor = Color.White
                        ElseIf .Text.ToUpper = "Pre 88".Trim.ToUpper Then
                            .BackColor = Color.Orange
                        Else
                            .Backcolor = Color.Transparent
                        End If
                    End With



                    .FormLoading = False

                    .txtFacilityLatDegree.Text = IIf(pFacilities.LatitudeDegree < 0, String.Empty, pFacilities.LatitudeDegree)
                    .txtFacilityLatMin.Text = IIf(pFacilities.LatitudeMinutes < 0, String.Empty, pFacilities.LatitudeMinutes)
                    .txtFacilityLatSec.Text = IIf(pFacilities.LatitudeSeconds < 0, String.Empty, FormatNumber(pFacilities.LatitudeSeconds, 2, TriState.True, TriState.False, TriState.True))
                    .txtFacilityLongDegree.Text = IIf(pFacilities.LongitudeDegree < 0, String.Empty, pFacilities.LongitudeDegree)
                    .txtFacilityLongMin.Text = IIf(pFacilities.LongitudeMinutes < 0, String.Empty, pFacilities.LongitudeMinutes)
                    .txtFacilityLongSec.Text = IIf(pFacilities.LongitudeSeconds < 0, String.Empty, FormatNumber(pFacilities.LongitudeSeconds, 2, TriState.True, TriState.False, TriState.True))

                    If mtype <> BusinessLogic.pFacility.FacilityModule.Fees Then

                        .FormLoading = True

                        UIUtilsGen.ValidateComboBoxItemByValue(.cmbFacilityDatum, pFacilities.Datum)
                        UIUtilsGen.ValidateComboBoxItemByValue(.cmbFacilityMethod, pFacilities.Method)
                        UIUtilsGen.ValidateComboBoxItemByValue(.cmbFacilityLocationType, pFacilities.LocationType)
                        UIUtilsGen.SetDatePickerValue(.dtFacilityPowerOff, pFacilities.DatePowerOff)

                        .txtFuelBrand.Text = IIf(pFacilities.FuelBrand.Length = 0, "", pFacilities.FuelBrand)

                        .chkSignatureofNF.Checked = pFacilities.SignatureOnNF

                        If mtype <> BusinessLogic.pFacility.FacilityModule.Compliance Then
                            .txtDueByNF.Text = IIf(.chkSignatureofNF.checked, String.Empty, "Due")
                            .dtPickUpcomingInstallDateValue.Enabled = False

                        Else
                            .dtPickUpcomingInstallDateValue.Enabled = True
                        End If

                        PopulateFacilityCapInfo(frm, pFacilities)

                        .chkUpcomingInstall.Checked = pFacilities.UpcomingInstallation
                        .chkUpcomingInstall.Tag = pFacilities.UpcomingInstallation

                        Dim bolLoadingLocal As Boolean = .FormLoading



                        If pFacilities.UpcomingInstallation OrElse pFacilities.UpcomingInstallationDate > CDate("1/1/1960") Then
                            .dtPickUpcomingInstallDateValue.Enabled = True
                            .dtPickUpcomingInstallDateValue.Checked = True
                        End If

                        UIUtilsGen.SetDatePickerValue(.dtPickUpcomingInstallDateValue, pFacilities.UpcomingInstallationDate)
                        .FormLoading = bolLoadingLocal

                    End If

                    .txtFacilityAIID.Text = pFacilities.AIID.ToString
                    .chkCAPCandidate.Checked = pFacilities.CAPCandidate

                    If .lblOwnerIDValue.Text = String.Empty Then
                        .PopulateOwnerInfo(pFacilities.OwnerID)
                    End If

                    .Tag = obj.lblFacilityIDValue.Text

                    Dim textStr As String = String.Empty

                    Select Case mtype

                        Case BusinessLogic.pFacility.FacilityModule.Registration
                            textStr = "Registration  - Facility Detail - "

                        Case BusinessLogic.pFacility.FacilityModule.Technical
                            textStr = "Technical - Facility Detail - "
                            obj.GetLUSTEventsForFacility()

                        Case BusinessLogic.pFacility.FacilityModule.Closure
                            textStr = "Closure - Facility Detail - "
                            .txtMGPTF.text = pFacilities.Current_MGPTF_Status

                        Case BusinessLogic.pFacility.FacilityModule.Financial
                            textStr = "Financial - Facility Detail - "
                            obj.GetFinancialEventsForFacility()

                        Case BusinessLogic.pFacility.FacilityModule.Fees
                            textStr = "Fees - Facility Detail - "
                            'P1 05/29/05 end

                        Case BusinessLogic.pFacility.FacilityModule.Compliance
                            textStr = "C & E - Facility Detail - "

                    End Select

                    .text = String.Format("{0} {1} ({2})", textStr, .lblFacilityIDValue.Text, .txtFacilityName.Text)


                    Dim byObj As Object = pFacilities.ModuleModifiedBy(mtype, pFacilities.ID)
                    Dim onObj As Object = pFacilities.ModuleModifiedOn(mtype, pFacilities.ID)

                    If onObj Is Nothing Then

                        .lblOwnerLastEditedBy.Text = String.Format("Facility Created By: {0}", pFacilities.CreatedBy)
                        .lblOwnerLastEditedOn.Text = String.Format("Created On : {0:d}", pFacilities.CreatedOn)

                    Else
                        .lblOwnerLastEditedBy.Text = String.Format("Last Edited By : {0}", byObj)
                        .lblOwnerLastEditedOn.Text = String.Format("Last Edited On : {0:d}", onObj)
                    End If

                    .FormLoading = False


                End With

                With MusterContainer.AppSemaphores
                    .Retrieve(fGUID.ToString, "FacilityDetails", obj.lblFacilityIDValue.Text, frm.Name)
                    .Retrieve(fGUID.ToString, "FacilityID", obj.lblFacilityIDValue.Text, frm.Name)
                    .Retrieve(fGUID.ToString, "FacilityName", obj.txtFacilityName.Text, frm.Name)
                    .Retrieve(fGUID.ToString, "FacilityAddress", FormatAddress(pFacilities.FacilityAddresses, True), frm.Name)
                End With

            End If



            UIUtilsGen.ActivateEntity(frm)

        Catch ex As Exception
            Throw ex
        End Try


    End Sub
    Friend Shared Sub PopulateFacilityCapInfo(ByRef frm As Form, ByVal pFacilities As MUSTER.BusinessLogic.pFacility)
        Dim obj As Object
        If TypeOf frm Is Technical Then
            obj = CType(frm, Technical)
        ElseIf TypeOf frm Is Registration Then
            obj = CType(frm, Registration)
        ElseIf TypeOf frm Is Closure Then
            obj = CType(frm, Closure)
        ElseIf TypeOf frm Is Financial Then
            obj = CType(frm, Financial)
        ElseIf TypeOf frm Is Fees Then
            obj = CType(frm, Fees)
        ElseIf TypeOf frm Is CandE Then
            obj = CType(frm, CandE)
        End If



        If pFacilities.CAPCandidate Then
            If pFacilities.CapStatus = 1 Then
                obj.lblCAPStatusValue.BackColor = Color.Green
                obj.lblCAPStatusValue.Text = "CAP Compliant"
            Else
                obj.lblCAPStatusValue.BackColor = Color.Red
                obj.lblCAPStatusValue.Text = "Not CAP Compliant"
            End If
        Else
            obj.lblCAPStatusValue.Text = String.Empty
            obj.lblCAPStatusValue.BackColor = System.Drawing.SystemColors.Control
        End If

    End Sub
    ' manju - 8/22/05 - modified byval to byref
    Friend Shared Sub Check_If_Datum_Enable(ByRef frm As Form)
        Dim obj As Object
        Try
            If TypeOf frm Is Technical Then
                obj = CType(frm, Technical)
            ElseIf TypeOf frm Is Registration Then
                obj = CType(frm, Registration)
            ElseIf TypeOf frm Is Closure Then
                obj = CType(frm, Closure)
            ElseIf TypeOf frm Is CandE Then
                obj = CType(frm, CandE)
            End If

            If IsNumeric(obj.txtFacilityLatDegree.Text) And _
                IsNumeric(obj.txtFacilityLatMin.Text) And _
                IsNumeric(obj.txtFacilityLatSec.Text) And _
                IsNumeric(obj.txtFacilityLongDegree.Text) And _
                IsNumeric(obj.txtFacilityLongMin.Text) And _
                IsNumeric(obj.txtFacilityLongSec.Text) AndAlso (TypeOf frm Is Registration Or TypeOf frm Is Technical) Then

                obj.cmbFacilityDatum.Enabled = True
                obj.cmbFacilityMethod.Enabled = True
                obj.cmbFacilityLocationType.Enabled = True
            Else
                obj.cmbFacilityDatum.Enabled = False
                obj.cmbFacilityMethod.Enabled = False
                obj.cmbFacilityLocationType.Enabled = False
                Dim bolLoadingLocal As Boolean = obj.FormLoading
                obj.FormLoading = True

                If Not TypeOf frm Is Closure Then
                    obj.cmbFacilityDatum.SelectedIndex = -1
                    If obj.cmbFacilityDatum.SelectedIndex <> -1 Then
                        obj.cmbFacilityDatum.SelectedIndex = -1
                    End If
                    obj.cmbFacilityMethod.SelectedIndex = -1
                    If obj.cmbFacilityMethod.SelectedIndex <> -1 Then
                        obj.cmbFacilityMethod.SelectedIndex = -1
                    End If
                    obj.cmbFacilityLocationType.SelectedIndex = -1
                    If obj.cmbFacilityLocationType.SelectedIndex <> -1 Then
                        obj.cmbFacilityLocationType.SelectedIndex = -1
                    End If
                End If

                obj.FormLoading = bolLoadingLocal
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Shared Sub ClearFacilityForm(ByRef frm As Form)
        Dim obj As Object
        Try
            If TypeOf frm Is Registration Then
                obj = CType(frm, Registration)
            ElseIf TypeOf frm Is Closure Then
                obj = CType(frm, Closure)
            ElseIf TypeOf frm Is CandE Then
                obj = CType(frm, CandE)
            End If

            obj.FormLoading = True
            UIUtilsGen.ClearFields(obj.pnl_FacilityDetail)
            obj.lblFacilityIDValue.Text = String.Empty
            obj.lblDateTransfered.Text = String.Empty
            'lblTotalNoOfTanksValue.Text = "0"
            'lblTotalNoOfTanksValue2.Text = "0"
            obj.txtFacilityAddress.Tag = 0
            obj.lblFacilityStatusValue.Text = String.Empty
            obj.lblFacilityStatusValue.BackColor = Nothing
            obj.cmbFacilityType.SelectedIndex = -1
            obj.cmbFacilityType.SelectedIndex = -1
            obj.cmbFacilityDatum.Enabled = False
            obj.cmbFacilityLocationType.Enabled = False
            obj.cmbFacilityMethod.Enabled = False
            obj.cmbFacilityDatum.SelectedIndex = -1
            obj.cmbFacilityMethod.SelectedIndex = -1
            obj.cmbFacilityLocationType.SelectedIndex = -1
            obj.lblCAPStatusValue.BackColor = Nothing
            obj.lblCAPStatusValue.Text = String.Empty
            obj.chkCAPCandidate.Checked = False
            obj.chkUpcomingInstall.Tag = False
            'dgPipesAndTanks.DataSource = Nothing
            'dgPipesAndTanks2.DataSource = Nothing
            'Me.lblOwnerLastEditedBy.Text = "Last Edited By : "
            'Me.lblOwnerLastEditedOn.Text = "Last Edited On : "
            obj.FormLoading = False
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "Letters"




    Friend Shared Function CreateDocument(ByVal strModule As String, ByVal Doc_Path As String, ByVal strDocName As String, ByVal strTemplatePath As String, Optional ByVal colParams As Specialized.NameValueCollection = Nothing) As Word.Application
        Try
            If Not strModule = "Registration" Then
                Dim oWord As Word.Application = MusterContainer.GetWordApp

                If Not oWord Is Nothing Then

                    Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
                    ltrGen.CreateLetter(strModule, strDocName, colParams, strTemplatePath, Doc_Path & strDocName, oWord)


                    Return oWord
                Else
                    Return Nothing
                End If

            End If
            Return Nothing
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Friend Shared Function CreateAndSaveDocument(ByVal strModule As String, ByVal nEntityID As Integer, ByVal nEntityType As Integer, ByVal Doc_Path As String, ByVal strDocName As String, ByVal strDocType As String, ByVal strTemplatePath As String, ByVal strDocPath As String, ByVal strDocDesc As String, ByVal ModuleID As Integer, Optional ByVal colParams As Specialized.NameValueCollection = Nothing, Optional ByVal eventID As Int64 = 0, Optional ByVal eventSequence As Integer = 0, Optional ByVal eventType As Integer = 0)
        Try
            Dim oWord As Word.Application = MusterContainer.GetWordApp

            If Not oWord Is Nothing Then

                If Not strModule = "Registration" Then
                    Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
                    ltrGen.CreateLetter(strModule, strDocName, colParams, strTemplatePath, Doc_Path & strDocName, oWord)
                    oWord.Visible = True
                End If

                Try
                    SaveDocument(nEntityID, nEntityType, strDocName, strDocType, Doc_Path, strDocDesc, ModuleID, eventID, eventSequence, eventType)
                Catch
                End Try
            End If
            oWord = Nothing

        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Friend Shared Sub OpenInPDFFile(ByVal strpath As String)

        Try
            Dim NeedAdobe As Boolean = False
            Try
                Dim LocalUserSettings As Microsoft.Win32.Registry

                Dim strSQLPath = LocalUserSettings.LocalMachine.OpenSubKey("SOFTWARE", False).OpenSubKey("Adobe", False).OpenSubKey("Acrobat Reader", False)
            Catch ex As Exception

                Try
                    Dim LocalUserSettings As Microsoft.Win32.Registry

                    Dim strSQLPath = LocalUserSettings.LocalMachine.OpenSubKey("SOFTWARE", False).OpenSubKey("Adobe", False).OpenSubKey("Adobe Acrobat", False)

                Catch ex2 As Exception

                    NeedAdobe = True

                End Try

            End Try

            If NeedAdobe Then

                If MsgBox("Your PC does not have adobe reader installed. Would you like to go to the Adobe website to download the application before viewing this document?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    Process.Start("http://www.adobe.com/downloads/")
                    MsgBox("Press OK to view document again", MsgBoxStyle.OKOnly)
                    Process.Start(strpath)
                End If
            Else
                Process.Start(strpath)
            End If

        Catch ex As Exception
            Throw ex
        End Try

    End Sub


    Public Shared Function DisplayLargeImage(ByRef PicForm As pictureViewer, ByVal img As Image, ByVal info As FileInfo)
        If img Is Nothing Then
            Exit Function
        Else

            Try

                PicForm = New pictureViewer(info, img)

                PicForm.Show()

            Catch ex As Exception
                Throw ex
            End Try


        End If

    End Function


    Public Shared Function GetThumbnail(ByVal path As String) As Image

        Dim desktopFolder As IShellFolder
        Dim someFolder As IShellFolder
        Dim extract As IExtractImage
        Dim pidl As IntPtr
        Dim filePidl As IntPtr
        Dim MAX_PATH As IntPtr
        Dim img As Image


        Try
            'Manually define the IIDs for IShellFolder and IExtractImage
            Dim IID_IShellFolder = New Guid("000214E6-0000-0000-C000-000000000046")
            Dim IID_IExtractImage = New Guid("BB2E617C-0920-11d1-9A0B-00C04FC2D6C1")

            'Divide the file name into a path and file name
            Dim folderName = path.Substring(0, path.LastIndexOf("\") + 1)
            Dim shortFileName = path.Substring(path.LastIndexOf("\") + 1)


            'Get the desktop IShellFolder
            ShellInterop.SHGetDesktopFolder(desktopFolder)

            'Get the parent folder IShellFolder
            desktopFolder.ParseDisplayName(IntPtr.Zero, IntPtr.Zero, folderName, 0, pidl, 0)
            desktopFolder.BindToObject(pidl, IntPtr.Zero, IID_IShellFolder, someFolder)

            'Get the file's IExtractImage
            someFolder.ParseDisplayName(IntPtr.Zero, IntPtr.Zero, shortFileName, 0, filePidl, 0)
            someFolder.GetUIObjectOf(IntPtr.Zero, 1, filePidl, IID_IExtractImage, 0, extract)

            'Set the size
            Dim size As Size
            size.cx = 500
            size.cy = 500

            Dim flags = IEIFLAG.ORIGSIZE Or IEIFLAG.QUALITY Or IEIFLAG.OFFLINE


            Dim bmp As IntPtr

            Dim thepath As New StringBuilder(540, 540)
            thepath.Append(path)

            'Interop will throw an exception if one of these calls fail.
            Try

                extract.GetLocation(thepath, thepath.MaxCapacity + 1, 0, size, 32, flags)

                extract.Extract(bmp)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try


            'Free the global memory we allocated for the path string
            thepath.Length = 0

            'Free the pidls. The Runtime Callable Wrappers 
            'should automatically release the COM objects
            Marshal.FreeCoTaskMem(pidl)
            Marshal.FreeCoTaskMem(filePidl)

            If Not bmp.Equals(IntPtr.Zero) Then
                img = Image.FromHbitmap(bmp)
                ' GetThumbnailImage = Image(bmp)

            Else
                img = Nothing

            End If
        Catch ex As Exception
            Throw ex
        End Try

        Return img
    End Function


    Public Shared Function ThumbnailCallback() As Boolean
        Return False
    End Function












    'Generic Function to Put Document Information for all the Letters.
    Friend Shared Function SaveDocument(ByVal nEntityID As Integer, ByVal nEntityType As Integer, ByVal strDocName As String, ByVal strDocType As String, ByVal strDocPath As String, ByVal strDocDescription As String, ByVal ModuleID As Integer, ByVal eventID As Int64, ByVal eventSequence As Integer, ByVal eventType As Integer) As Integer

        Dim ltrInfo As MUSTER.Info.LetterInfo
        Try
            ltrInfo = New MUSTER.Info.LetterInfo(0, _
                                                 strDocName.Trim, _
                                                 strDocType, _
                                                 strDocPath, _
                                                 nEntityType, _
                                                 nEntityID, _
                                                 strDocDescription, _
                                                1, _
                                                CDate("01/01/0001"), _
                                                False, _
                                                MusterContainer.AppUser.ID, _
                                                CDate("01/01/0001"), _
                                                String.Empty, _
                                                 CDate("01/01/0001"), _
                                                MusterContainer.AppUser.ID, _
                                                ModuleID, _
                                                eventID, _
                                                eventSequence, _
                                                eventType)
            MusterContainer.pLetter.Add(ltrInfo)
            MusterContainer.pLetter.Save()

            Return MusterContainer.pLetter.ID

        Catch ex As Exception
            Throw ex
        Finally
            ltrInfo = Nothing
        End Try
    End Function
    Friend Shared Sub DeleteDocument(ByVal strDocName As String, ByVal strOwningUser As String, Optional ByVal bolDeleted As Boolean = False)
        Try
            MusterContainer.pLetter.RetrieveByDocName(strDocName, strOwningUser, bolDeleted)
            If MusterContainer.pLetter.ID > 0 Then
                MusterContainer.pLetter.Deleted = True
                'MusterContainer.pLetter.ModifiedBy = MusterContainer.AppUser.ID
                MusterContainer.pLetter.Save()
            End If
            MusterContainer.pLetter.Remove(MusterContainer.pLetter.ID)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "Owner Summary"
    Friend Shared Sub PopulateOwnerSummary(ByRef pown As Object, ByRef frm As Form)
        Dim dsSet As DataSet
        Dim dtTotals As DataTable
        Try
            Dim obj As Object
            If TypeOf frm Is Technical Then
                obj = CType(frm, Technical)
            ElseIf TypeOf frm Is Registration Then
                obj = CType(frm, Registration)
            ElseIf TypeOf frm Is Closure Then
                obj = CType(frm, Closure)
            ElseIf TypeOf frm Is Financial Then
                obj = CType(frm, Financial)
            ElseIf TypeOf frm Is Fees Then
                obj = CType(frm, Fees)
            ElseIf TypeOf frm Is CandE Then
                obj = CType(frm, CandE)
            End If
            dsSet = pown.GetOwnerSummary
            obj.UCOwnerSummary.LustSites = dsSet.Tables(0)
            obj.UCOwnerSummary.FinancialSites = dsSet.Tables(1)
            obj.UCOwnerSummary.Fees = dsSet.Tables(2)
            obj.UCOwnerSummary.Penalities = dsSet.Tables(3)
            obj.UCOwnerSummary.PreviousFacilities = dsSet.Tables(4)
            dtTotals = pown.GetOwnerSummaryFeesTotals

            If IsNothing(dtTotals) Then
                obj.ucownersummary.PriorBalance = ""
                obj.ucownersummary.CurrentFees = ""
                obj.ucownersummary.LateFees = ""
                obj.ucownersummary.TotalDue = ""
                obj.ucownersummary.CurrentPayments = ""
                obj.ucownersummary.CurrentCredits = ""
                obj.ucownersummary.CurrentAdjustments = ""
                obj.ucownersummary.ToDateBalance = ""

            Else
                obj.ucownersummary.PriorBalance = dtTotals.Rows(0).Item("PriorBalanceTotal").ToString
                obj.ucownersummary.CurrentFees = dtTotals.Rows(0).Item("CurrentFeesTotal").ToString
                obj.ucownersummary.LateFees = dtTotals.Rows(0).Item("LateFeesTotal").ToString
                obj.ucownersummary.TotalDue = dtTotals.Rows(0).Item("TotalDueTotal").ToString
                obj.ucownersummary.CurrentPayments = dtTotals.Rows(0).Item("CurrentPaymentsTotal").ToString
                obj.ucownersummary.CurrentCredits = dtTotals.Rows(0).Item("CurrentCreditTotal").ToString
                obj.ucownersummary.CurrentAdjustments = dtTotals.Rows(0).Item("CurrentAdjustmentsTotal").ToString
                obj.ucownersummary.ToDateBalance = dtTotals.Rows(0).Item("ToDateBalanceTotal").ToString
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
    Friend Shared Function FormatAddress(ByVal addr As MUSTER.BusinessLogic.pAddress, Optional ByVal isFacilityAddress As Boolean = False) As String
        Try
            Dim strAddress As String = String.Empty
            strAddress += addr.AddressLine1.Trim
            If addr.AddressLine2.Trim <> String.Empty Then
                strAddress += vbCrLf + addr.AddressLine2.Trim
            End If
            strAddress += vbCrLf + _
                            addr.City.Trim + ", " + _
                            addr.State.Trim + " " + _
                            addr.Zip.Trim + IIf(addr.PhysicalTown.ToUpper.Trim <> addr.City.ToUpper.Trim AndAlso addr.PhysicalTown.Length > 0, "  (" + addr.PhysicalTown + ")", String.Empty)
            If isFacilityAddress Then
                strAddress += vbCrLf + vbCrLf + addr.County + " County"
            End If
            Return strAddress.Trim
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Friend Shared Sub ClearFields(ByVal objControl As Control)
        Try
            Dim currentControl As Control
            Dim currentSubControl As Control
            Dim str As String
            Dim tmpCmb As System.Windows.Forms.ComboBox
            Dim tmpTxt As System.Windows.Forms.TextBox
            Dim tmpRd As System.Windows.Forms.RadioButton
            Dim tmpDtPick As System.Windows.Forms.DateTimePicker
            Dim tmpMaksedEdit As AxMSMask.AxMaskEdBox
            Dim tmpChk As System.Windows.Forms.CheckBox
            Dim strtmpMask As String
            Dim myEnumerator As System.Collections.IEnumerator = _
                       objControl.Controls.GetEnumerator()

            While myEnumerator.MoveNext()
                currentControl = myEnumerator.Current
                If currentControl.GetType.ToString.ToLower = "system.Windows.Forms.ComboBox".ToLower Then
                    tmpCmb = CType(currentControl, System.Windows.Forms.ComboBox)
                    tmpCmb.SelectedIndex = -1
                    tmpCmb.Tag = ""
                ElseIf currentControl.GetType.ToString.ToLower = "system.Windows.Forms.TextBox".ToLower Then
                    tmpTxt = CType(currentControl, System.Windows.Forms.TextBox)
                    If tmpTxt.Visible Then
                        tmpTxt.Text = String.Empty
                        tmpTxt.Tag = ""
                    End If

                ElseIf currentControl.GetType.ToString.ToLower = "system.Windows.Forms.DateTimePicker".ToLower Then
                    tmpDtPick = CType(currentControl, System.Windows.Forms.DateTimePicker)
                    tmpDtPick.Text = ""
                    tmpDtPick.CustomFormat = "__/__/____"
                    tmpDtPick.Format = DateTimePickerFormat.Custom
                    tmpDtPick.Checked = True
                    tmpDtPick.Checked = False
                    tmpDtPick.Tag = Nothing

                    '----- control is a masked edit text box -----------------
                ElseIf currentControl.GetType.ToString.ToLower = "axmsmask.axmaskedbox" Then
                    tmpMaksedEdit = CType(currentControl, AxMSMask.AxMaskEdBox)
                    strtmpMask = tmpMaksedEdit.Mask
                    tmpMaksedEdit.Mask = ""
                    tmpMaksedEdit.CtlText = ""
                    tmpMaksedEdit.Mask = strtmpMask
                    tmpMaksedEdit.Tag = ""

                    '----- control is a radio button -------------------------
                ElseIf currentControl.GetType.ToString.ToLower = "system.Windows.Forms.RadioButton".ToLower Then
                    tmpRd = CType(currentControl, System.Windows.Forms.RadioButton)
                    tmpRd.Checked = False

                    '----- control is a text box -----------------------------
                ElseIf currentControl.GetType.ToString.ToLower = "system.Windows.Forms.CheckBox".ToLower Then
                    tmpChk = CType(currentControl, System.Windows.Forms.CheckBox)
                    tmpChk.Checked = False

                Else
                    If currentControl.Controls.Count > 0 Then
                        ClearFields(currentControl)
                    End If
                End If

            End While
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Shared Function CorrectZipFormat(ByVal strZip As String) As String
        Try
            If strZip.Substring(6, 4) = "xxxx" Or strZip.Substring(6, 4) = "0000" Then
                strZip = strZip.Substring(0, 5)
            Else
                strZip = strZip.Replace("x", "0")
            End If
            Return strZip
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    'Public Shared Function ArrayListToDataTable(ByVal ArrLst As ArrayList) As DataTable
    '    Dim i As Int16

    '    Dim dsTable As New DataTable
    '    Dim dsDR As DataRow
    '    Try
    '        dsTable.Columns.Add("Id")
    '        dsTable.Columns(0).DataType = System.Type.GetType("System.Int32")
    '        dsTable.Columns.Add("Type")
    '        dsTable.Columns(1).DataType = System.Type.GetType("System.String")
    '        dsDR = dsTable.NewRow
    '        dsDR.Item(0) = 0
    '        dsDR.Item(1) = ""
    '        dsTable.Rows.Add(dsDR)
    '        For i = 0 To ArrLst.Count - 1
    '            Dim LP As InfoRepository.LookupProperty
    '            LP = ArrLst(i)
    '            dsDR = dsTable.NewRow
    '            dsDR.Item(0) = LP.Id
    '            dsDR.Item(1) = LP.Type
    '            dsTable.Rows.Add(dsDR)
    '        Next

    '        Return dsTable

    '    Catch ex As Exception
    '        Throw ex
    '    End Try

    'End Function
    Friend Shared Function GetWordApp() As Word.Application
        'Dim WordApp As Word.Application
        'Try
        '    WordApp = GetObject(, "Word.Application")
        'Catch ex As Exception
        '    If ex.Message.ToUpper = "Cannot Create ActiveX Component.".ToUpper Then
        '        WordApp = New Word.Application
        '    Else
        '        Throw ex
        '    End If
        'End Try
        Return MusterContainer.GetWordApp

    End Function
    Friend Shared Function CreateWordObject() As Word.Application
        Dim TempApp As Word.Application
        Dim WordApp As Word.Application
        Try
            'Test if object is already created before calling CreateObject:
            If TypeName(WordApp) <> "Application" Then
                TempApp = New Word.Application
                WordApp = New Word.Application
                TempApp.Quit()
                TempApp = Nothing
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return WordApp
    End Function
    'Friend Shared Function GetDistinctArrayListItems(ByVal MasterArrLst As ArrayList) As ArrayList
    '    Dim FilteredItems As New ArrayList
    '    Dim LstItem As InfoRepository.LookupProperty
    '    Dim strRepeatedText As String = ""
    '    Dim MasterEnumerator As System.Collections.IEnumerator = MasterArrLst.GetEnumerator()

    '    GetDistinctArrayListItems = Nothing

    '    Try
    '        While MasterEnumerator.MoveNext()
    '            LstItem = MasterEnumerator.Current
    '            If Not LstItem.Type = strRepeatedText Then
    '                FilteredItems.Add(LstItem)
    '                strRepeatedText = LstItem.Type
    '            End If
    '        End While
    '        GetDistinctArrayListItems = FilteredItems
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        MasterEnumerator = Nothing
    '    End Try
    'End Function
    Public Shared Sub ShowHideControl(ByVal ObjControl As Control)
        If ObjControl.Visible Then
            ObjControl.Visible = False
        Else
            ObjControl.Visible = True
        End If
    End Sub
    Public Shared Sub EnableDisableControl(ByVal objControl As Control)
        If objControl.Enabled Then
            objControl.Enabled = False
        Else
            objControl.Enabled = True
        End If
    End Sub
    Friend Shared Function SetMaskEditText(ByVal msk As AxMSMask.AxMaskEdBox, ByVal value As String)
        msk.Mask = String.Empty
        msk.SelText = String.Empty
        msk.CtlText = String.Empty
        msk.Mask = "(###)###-####"
        If value.Length > 0 Then
            msk.SelText = value
        Else
            msk.SelText = String.Empty
        End If
    End Function
    Friend Shared Sub FillStringObjectValues(ByRef currentObj As Object, ByVal value As String)
        If value = "(___)___-____" Then
            currentObj = String.Empty
            Exit Sub
        End If
        If value.Length > 0 Then
            currentObj = value
        Else
            currentObj = String.Empty
        End If
    End Sub
    Friend Shared Sub FillDateobjectValues(ByRef currentObj As Object, ByVal value As String)

        If value.Length > 0 And value <> "__/__/____" Then
            currentObj = CType(value, Date)
        Else
            currentObj = "#12:00:00AM#"
        End If
    End Sub
    Friend Shared Sub FillSingleObjectValues(ByRef currentObj As Object, ByVal value As String)
        If value.Length > 0 Then
            currentObj = CSng(value)
        Else
            currentObj = -1.0
        End If
    End Sub
    Friend Shared Sub FillDoubleObjectValues(ByRef currentObj As Object, ByVal value As String)
        If value.Length > 0 Then
            currentObj = CDbl(value)
        Else
            currentObj = -1.0
        End If
    End Sub
    'Friend Shared Function AddEmptyRowToCmbDatatable(ByRef cmbox As ComboBox, ByRef dt As DataTable) As DataTable
    '    Dim dr As DataRow
    '    Try
    '        cmbox.DataSource = Nothing
    '        If Not dt Is Nothing Then
    '            dr = dt.NewRow
    '            dr.Item("PROPERTY_NAME") = String.Empty
    '            dr.Item("PROPERTY_ID") = 0
    '            dt.Rows.InsertAt(dr, 0)
    '            cmbox.DataSource = dt
    '            cmbox.DisplayMember = "PROPERTY_NAME"
    '            cmbox.ValueMember = "PROPERTY_ID"
    '        End If
    '        cmbox.SelectedIndex = 0
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Function
    'Friend Shared Function RemoveEmptyRowFromCmbDatatable(ByRef cmbox As ComboBox, ByRef dt As DataTable) As DataTable
    '    Dim dr As DataRow
    '    Try
    '        cmbox.DataSource = Nothing
    '        If Not dt Is Nothing Then
    '            If dt.Rows(0).Item("PROPERTY_NAME") = String.Empty Then
    '                dt.Rows.RemoveAt(0)
    '            End If
    '            cmbox.DataSource = dt
    '            cmbox.DisplayMember = "PROPERTY_NAME"
    '            cmbox.ValueMember = "PROPERTY_ID"
    '        End If
    '        cmbox.SelectedIndex = 0
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Function
    Friend Shared Sub Delay(Optional ByVal msecs As Double = 0.0, Optional ByVal sec As Double = 0.0, Optional ByVal mins As Double = 0.0)
        Dim MyTime As DateTime
        MyTime = Now.AddMilliseconds(msecs).AddSeconds(sec).AddMilliseconds(msecs).AddSeconds(sec).AddMinutes(mins)
        Do Until Now > MyTime
        Loop
    End Sub
    Friend Shared Function GetMonth(ByVal dtDate As Date) As String
        Select Case dtDate.Month
            Case "1"
                Return ("JANUARY")
            Case "2"
                Return ("FEBRUARY")
            Case "3"
                Return ("MARCH")
            Case "4"
                Return ("APRIL")
            Case "5"
                Return ("MAY")
            Case "6"
                Return ("JUNE")
            Case "7"
                Return ("JULY")
            Case "8"
                Return ("AUGUST")
            Case "9"
                Return ("SEPTEMBER")
            Case "10"
                Return ("OCTOBER")
            Case "11"
                Return ("NOVEMBER")
            Case "12"
                Return ("DECEMBER")
        End Select
    End Function
    Public Shared Function GetModuleIDByName(ByVal moduleName As String) As Integer
        Select Case moduleName
            Case "Registration"
                Return 612
            Case "CAE"
                Return 613
            Case "Technical"
                Return 614
            Case "Inspection"
                Return 615
            Case "Financial"
                Return 616
            Case "Closure"
                Return 891
            Case "Fees"
                Return 892
            Case "Company"
                Return 893
            Case "ContactManagement"
                Return 894
            Case "Admin"
                Return 1303
            Case "Global"
                Return 1311
            Case "FeeAdmin"
                Return 1312
            Case Else
                Return -1
        End Select

    End Function

    Public Shared Function GetModuleNameByID(ByVal moduleID As Integer) As String
        Select Case moduleID
            Case 612
                Return "Registration"
            Case 613
                Return "CAE"
            Case 614
                Return "Technical"
            Case 615
                Return "Inspection"
            Case 616
                Return "Financial"
            Case 891
                Return "Closure"
            Case 892
                Return "Fees"
            Case 893
                Return "Company"
            Case 894
                Return "ContactManagement"
            Case 1303
                Return "Admin"
            Case 1311
                Return "Global"
            Case 1312
                Return "FeeAdmin"
            Case Else
                Return ""
        End Select

    End Function

    Friend Shared Function DataTableContainsValue(ByVal dt As DataTable, ByVal val As String, ByVal valueMember As String, ByVal operand As String) As Boolean
        Dim returnVal As Boolean = False
        Try
            If Not dt Is Nothing Then
                If dt.Columns.Contains(valueMember) Then
                    If dt.Select(valueMember + " " + operand + " " + val).Length > 0 Then returnVal = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return returnVal
    End Function
    Friend Shared Function DataSetContainsValue(ByVal ds As DataSet, ByVal val As String, ByVal valueMember As String, ByVal operand As String) As Boolean
        Dim dt As DataTable
        Dim returnVal As Boolean = False
        Try
            If Not ds Is Nothing Then
                If ds.Tables.Count > 0 Then
                    dt = ds.Tables(0)
                    If dt.Columns.Contains(valueMember) Then
                        If dt.Select(valueMember + " " + operand + " " + val).Length > 0 Then returnVal = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return returnVal
    End Function
#End Region
#Region "Constants"
    Friend Const CommentsButton_HasCmts = "Orange"
    Friend Const CommentsButton_NoCmts = "Control"
    ' File Path Keys
    Friend Const FilePathKey_FacImages = "SYSTEM|COMMON_PATHS|Facilities|NONE"
    Friend Const FilePathKey_LicenseesImages = "SYSTEM|COMMON_PATHS|Licensees|NONE"
    Friend Const FilePathKey_Sketches = "SYSTEM|COMMON_PATHS|Sketches|NONE"
    Friend Const FilePathKey_Templates = "SYSTEM|COMMON_PATHS|Templates|NONE"
    Friend Const FilePathKey_SystemGenerated = "SYSTEM|COMMON_PATHS|System_Generated|NONE"
    Friend Const FilePathKey_ManuallyCreated = "SYSTEM|COMMON_PATHS|Manually_Created|NONE"
    Friend Const FilePathKey_SystemArchive = "SYSTEM|COMMON_PATHS|System_Archive|NONE"
    Friend Const FilePathKey_Reports = "SYSTEM|COMMON_PATHS|Reports|NONE"
    Friend Const FilePathKey_DBSync = "SYSTEM|COMMON_PATHS|DBSync|NONE"
#End Region
#Region "Enumeration"
    Public Enum LicenseeCourseType
        Install = 920
        Closure = 921
    End Enum
    Public Enum TankPipeStatus
        CIU = 424
        TOS = 425
        POU = 426
        TOSI = 429
        Unregulated = 430
    End Enum
    Public Enum EntityTypes
        MUSTER = 1
        Agency = 2
        Citation = 3
        Contact = 4
        Contractor = 5
        Facility = 6
        LUST_Event = 7
        NONE = 8
        Owner = 9
        Pipe = 10
        Report = 11
        Tank = 12
        Violation = 13
        Organization = 14
        User = 15
        Persona = 16
        Flag = 17
        Calendar = 18
        Letter = 19
        Profile = 20
        Address = 21
        ClosureEvent = 22
        LustActivity = 23
        LustDocument = 24
        Fees = 25
        Company = 26
        Licensee = 27
        Provider = 28
        Financial = 29
        Inspection = 30
        CAE = 31
        FinancialEvent = 32
        FinancialCommitment = 33
        FinancialInvoice = 34
        FinancialReimbursement = 35
        CAEOwnerComplianceEvent = 36
        CAELicenseeCompliantEvent = 37
        TechnicalActivity = 38
        TechnicalDocument = 39
        LustRemediation = 40
        CAEFacilityCompliantEvent = 41
        Comment = 42
    End Enum
    Public Enum ActivityTypes
        AddOwner = 1
        'AddFacility = 2
        AddTank = 3
        'AddPipe = 4
        'AddContact = 5
        TransferOwnership = 6
        TankStatusTOSI = 7
        'TankStatusCIU = 8
        'NeedUIReminderLetter = 9
        'TransferAcknowledgement = 10
        UpComingInstall = 11
        SignatureRequired = 12
        Fees = 13
        SecondLetterForSigRequired = 14
    End Enum
    Public Enum ModuleID
        Registration = 612
        CAE = 613
        Technical = 614
        Inspection = 615
        Financial = 616
        Closure = 891
        Fees = 892
        Company = 893
        ContactManagement = 894
        Admin = 1303
        Global = 1311
        FeeAdmin = 1312
        CAPProcess = 1314
        TechAdmin = 1637
        FinAdmin = 1638
        CompanyAdmin = 1652
    End Enum
    Public Enum OCELetterTemplateNum
        ' OCE Creation
        DiscrepanciesOnly = 1
        CAT3_NoPrior_NOV = 2
        CAT2_NoPrior_NOV_Workshop = 3
        CAT2_1_CAT3_NOV_Workshop = 3
        CAT1_CAT1_CAT2_1_CAT3_NOV_AgreedOrder = 4
        CAT1_NoPrior_NOV_Workshop_AgreedOrder = 5
        CAT1_1_CAT3_NOV_Workshop_AgreedOrder = 5
        CAT2_CAT1_CAT2_1_CAT3_NOV_AgreedOrder = 6
        CAT3_CAT1_CAT2_1_CAT3_NOV_AgreedOrder = 7
        CAT3_1_CAT3_NOV_Workshop = 8
        ' OCE Escalation
        NOV_A_AgreedOrder = 9
        NOV_A_2ndNotice = 10
        NOV_A_ShowCauseHearing = 11
        NOV_AgreedOrder_B_2ndNotice = 12
        NOV_AgreedOrder_B_AgreedOrder = 13
        NOV_AgreedOrder_B_ShowCauseHearing = 14
        NOV_Workshop_C_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsNoShowORChoseToPayAndEscalatedStatusIsAgreedOrder = 15
        NOV_Workshop_C_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsPassORWaiveAndEscalatedStatusIsAgreedOrder = 16
        NOV_Workshop_C_AgreedOrder_WhenCurrentStatusIs2ndNoticeAndEscalatedStatusIsAgreedOrder = 17
        NOV_Workshop_C_ShowCauseHearing = 18
        NOV_Workshop_C_2ndNotice = 19
        NOV_Workshop_AgreedOrder_D_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsNoShowORChoseToPayAndEscalatedStatusIsAgreedOrder = 20
        NOV_Workshop_AgreedOrder_D_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsPassORWaiveAndEscalatedStatusIsAgreedOrder = 21
        NOV_Workshop_AgreedOrder_D_2ndNotice = 22
        NOV_Workshop_AgreedOrder_D_AgreedOrder_WhenCurrentStatusIs2ndNoticeAndEscalatedStatusIsAgreedOrder = 23
        NOV_Workshop_AgreedOrder_D_ShowCauseHearing = 24
        Hearing_CommissionHearing_WhenCurrentStatusIsShowCauseHearing = 25
        Hearing_ShowCauseAgreedOrder = 26
        Hearing_CommissionHearing_WhenCurrentStatusIsShowCauseAgreedOrder = 27
        Hearing_CommissionHearingNFARescinded = 28
        Hearing_AdministrativeOrder = 29
        NFA_NFA = 30
        ' OCE Rescind
        NFARescind = 39
        ' OCE Creation
        ViolationWithin90Days_Discrepancy_WhenOwnerHasNonDiscrepCitation = 41
        ViolationWithin90Days_Discrepancy_WhenOwnerHasDiscrepCitationsOnly = 42
        ' OCE Escalation
        NOV_AgreedOrder_AgreedOrder = 43
        NOV_AgreedOrder_ShowCauseHearing = 44
        RedTag_Notice = 45
        RedTag_Warning = 46
        NOV_AgreedOrder_AfterRedTag = 47
        StandAloneAgreedOrder = 48

    End Enum
    Public Enum LCELetterTemplateNum
        ' LCE Creation
        NOV = 31
        ' LCE Escalation
        NOV_ShowCauseHearing = 32
        Hearings_CommissionHearing_WhenCurrentStatusIsShowCauseHearingAndEscalatedStatusIsCommissionHearing = 33
        Hearings_ShowCauseAgreedOrder = 34
        Hearings_CommissionHearing_WhenCurrentStatusIsShowCauseAgreedOrderAndEscalatedStatusIsCommissionHearing = 35
        Hearings_CommissionHearingNFARescinded = 36
        Hearings_AdministrativeOrder = 37
        NFA_NFA = 38
        ' LCE Rescind
        NFARescind = 40
    End Enum
    Friend Shared Function GetLetterTemplateNumPropertyID(ByVal templateNum As Integer, Optional ByVal useSecondLetter As Boolean = False) As Integer
        Dim retVal As Integer = 0
        Select Case templateNum
            Case OCELetterTemplateNum.DiscrepanciesOnly
                retVal = 1261
            Case OCELetterTemplateNum.CAT3_NoPrior_NOV
                retVal = 1262
            Case OCELetterTemplateNum.CAT2_NoPrior_NOV_Workshop, OCELetterTemplateNum.CAT2_1_CAT3_NOV_Workshop
                If Not useSecondLetter Then
                    retVal = 1263
                Else
                    retVal = 1264
                End If
            Case OCELetterTemplateNum.CAT1_CAT1_CAT2_1_CAT3_NOV_AgreedOrder
                retVal = 1265
            Case OCELetterTemplateNum.CAT1_NoPrior_NOV_Workshop_AgreedOrder, OCELetterTemplateNum.CAT1_1_CAT3_NOV_Workshop_AgreedOrder
                If Not useSecondLetter Then
                    retVal = 1266
                Else
                    retVal = 1267
                End If
            Case OCELetterTemplateNum.CAT2_CAT1_CAT2_1_CAT3_NOV_AgreedOrder
                retVal = 1268
            Case OCELetterTemplateNum.CAT3_CAT1_CAT2_1_CAT3_NOV_AgreedOrder
                retVal = 1269
            Case OCELetterTemplateNum.CAT3_1_CAT3_NOV_Workshop
                retVal = 1270
            Case OCELetterTemplateNum.NOV_A_AgreedOrder
                retVal = 1271
            Case OCELetterTemplateNum.NOV_A_2ndNotice
                retVal = 1272
            Case OCELetterTemplateNum.NOV_A_ShowCauseHearing
                retVal = 1273
            Case OCELetterTemplateNum.NOV_AgreedOrder_B_2ndNotice
                retVal = 1274
            Case OCELetterTemplateNum.NOV_AgreedOrder_B_AgreedOrder
                retVal = 1275
            Case OCELetterTemplateNum.NOV_AgreedOrder_B_ShowCauseHearing
                retVal = 1276
            Case OCELetterTemplateNum.NOV_Workshop_C_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsNoShowORChoseToPayAndEscalatedStatusIsAgreedOrder
                retVal = 1277
            Case OCELetterTemplateNum.NOV_Workshop_C_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsPassORWaiveAndEscalatedStatusIsAgreedOrder
                retVal = 1278
            Case OCELetterTemplateNum.NOV_Workshop_C_AgreedOrder_WhenCurrentStatusIs2ndNoticeAndEscalatedStatusIsAgreedOrder
                retVal = 1279
            Case OCELetterTemplateNum.NOV_Workshop_C_ShowCauseHearing
                retVal = 1280
            Case OCELetterTemplateNum.NOV_Workshop_C_2ndNotice
                retVal = 1281
            Case OCELetterTemplateNum.NOV_Workshop_AgreedOrder_D_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsNoShowORChoseToPayAndEscalatedStatusIsAgreedOrder
                retVal = 1282
            Case OCELetterTemplateNum.NOV_Workshop_AgreedOrder_D_AgreedOrder_WhenCurrentStatusIsNew_WorkshopResultIsPassORWaiveAndEscalatedStatusIsAgreedOrder
                retVal = 1283
            Case OCELetterTemplateNum.NOV_Workshop_AgreedOrder_D_2ndNotice
                retVal = 1284
            Case OCELetterTemplateNum.NOV_Workshop_AgreedOrder_D_AgreedOrder_WhenCurrentStatusIs2ndNoticeAndEscalatedStatusIsAgreedOrder
                retVal = 1285
            Case OCELetterTemplateNum.NOV_Workshop_AgreedOrder_D_ShowCauseHearing
                retVal = 1286
            Case OCELetterTemplateNum.Hearing_CommissionHearing_WhenCurrentStatusIsShowCauseHearing
                retVal = 1287
            Case OCELetterTemplateNum.Hearing_ShowCauseAgreedOrder
                retVal = 1288
            Case OCELetterTemplateNum.Hearing_CommissionHearing_WhenCurrentStatusIsShowCauseAgreedOrder
                retVal = 1289
            Case OCELetterTemplateNum.Hearing_CommissionHearingNFARescinded ' 28
                retVal = 1290
            Case OCELetterTemplateNum.Hearing_AdministrativeOrder ' 29
                retVal = 1291
            Case OCELetterTemplateNum.NFA_NFA ' 30
                retVal = 1292
            Case LCELetterTemplateNum.NOV ' 31
                retVal = 1293
            Case LCELetterTemplateNum.NOV_ShowCauseHearing ' 32
                retVal = 1294
            Case LCELetterTemplateNum.Hearings_CommissionHearing_WhenCurrentStatusIsShowCauseHearingAndEscalatedStatusIsCommissionHearing
                retVal = 1295
            Case LCELetterTemplateNum.Hearings_ShowCauseAgreedOrder
                retVal = 1296
            Case LCELetterTemplateNum.Hearings_CommissionHearing_WhenCurrentStatusIsShowCauseAgreedOrderAndEscalatedStatusIsCommissionHearing
                retVal = 1297
            Case LCELetterTemplateNum.Hearings_CommissionHearingNFARescinded
                retVal = 1298
            Case LCELetterTemplateNum.Hearings_AdministrativeOrder
                retVal = 1299
            Case LCELetterTemplateNum.NFA_NFA
                retVal = 1300
            Case OCELetterTemplateNum.NFARescind
                retVal = 1301
            Case LCELetterTemplateNum.NFARescind
                retVal = 1302
            Case OCELetterTemplateNum.ViolationWithin90Days_Discrepancy_WhenOwnerHasNonDiscrepCitation ' 41
                retVal = 1567
            Case OCELetterTemplateNum.ViolationWithin90Days_Discrepancy_WhenOwnerHasDiscrepCitationsOnly ' 42
                retVal = 1568
            Case OCELetterTemplateNum.NOV_AgreedOrder_AgreedOrder ' 43
                retVal = 1569
            Case OCELetterTemplateNum.NOV_AgreedOrder_ShowCauseHearing ' 44
                retVal = 1570
            Case OCELetterTemplateNum.RedTag_Notice   ' 45
                retVal = 1664
            Case OCELetterTemplateNum.RedTag_Warning   '46
                retVal = 1668
            Case OCELetterTemplateNum.NOV_AgreedOrder_AfterRedTag   ' 47
                retVal = 1673
            Case OCELetterTemplateNum.StandAloneAgreedOrder    ' 48
                retVal = 48
        End Select
        Return retVal
    End Function
#End Region
#Region "Security Fuctions"

    Friend Shared Function HasRights(ByVal returnVal As String, Optional ByVal pass As Boolean = False) As Boolean

        If Not returnVal = String.Empty Then
            If Not pass Then
                MessageBox.Show(returnVal.ToString(), "ACCESS DENIED", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                Throw New Exception(returnVal.ToString)
            End If

            Return False
        Else
            Return True
        End If

    End Function

#End Region
#Region "Envelopes And Letters"
    Public Shared Function CreateEnvelopes(ByVal strName As String, ByVal arrAddress() As String, ByVal strModule As String, ByVal nEntityID As Integer, Optional ByVal strContact As String = "")
        Dim colParams As New Specialized.NameValueCollection
        Dim strTempPath As String
        Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
        Dim strDOC_NAME As String = String.Empty
        Dim strToday As String = String.Empty
        Dim sAddress() As String
        Dim DOC_PATH As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_SystemGenerated).ProfileValue & "\"
        Dim TmpltPath As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_Templates).ProfileValue & "\"
        Dim input
        Dim nLocation As Integer
        Dim strTempAddress As String
        Try
            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If
            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm")) + CStr(Format(Now, "ss"))
            strDOC_NAME = strModule + "Envelope" + "_" + CStr(Trim(nEntityID.ToString)) + "_" + strToday + ".doc"
            'sAddress = strAddress.Split(",")
            sAddress = arrAddress
            colParams.Add("<CONTACT>", strContact)

            colParams.Add("<NAME>", strName)
            colParams.Add("<ADDRESS1>", sAddress(0))
            If sAddress(1) = String.Empty Then
                colParams.Add("<ADDRESS2>", sAddress(2) + " " + sAddress(3) + " " + sAddress(4))
                colParams.Add("<CITYSTATEZIP>", String.Empty)
            Else
                colParams.Add("<ADDRESS2>", sAddress(1))
                colParams.Add("<CITYSTATEZIP>", sAddress(2) + " " + sAddress(3) + " " + sAddress(4))
            End If
            strTempPath = TmpltPath & "Global\EnvelopeTemplate.doc"
            Dim oWord As Word.Application = MusterContainer.GetWordApp

            If Not oWord Is Nothing Then
                ltrGen.CreateEnvelope("Envelope", colParams, strTempPath, String.Empty, oWord)

                'ltrGen.CreateLetter("Global", "Envelope", colParams, strTempPath, DOC_PATH & strDOC_NAME, oWord)
                oWord.Visible = True
            End If
            oWord = Nothing
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Shared Function CreateLabels(ByVal strName As String, ByVal arrAddress() As String, ByVal strModule As String, ByVal nEntityID As Integer)
        Dim colParams As New Specialized.NameValueCollection
        Dim strTempPath As String
        Dim ltrGen As New MUSTER.BusinessLogic.pLetterGen
        Dim strDOC_NAME As String = String.Empty
        Dim strToday As String = String.Empty
        Dim DOC_PATH As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_SystemGenerated).ProfileValue & "\"
        Dim TmpltPath As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_Templates).ProfileValue & "\"
        Dim sAddress() As String
        Dim input
        Dim nRow As Integer = 1
        Dim nColumn As Integer = 1
        Dim strTempAddress As String
        Try
            If DOC_PATH = "\" Then
                Throw New Exception("Document Path Unspecified. Please give the path before generating the letter.")
            End If

            strToday = CStr(Format(Now, "MM")) + CStr(Format(Now, "dd")) + CStr(Format(Now, "yy")) + "_" + CStr(Format(Now, "HH")) + CStr(Format(Now, "mm")) + CStr(Format(Now, "ss"))
            strDOC_NAME = strModule + "Label" + "_" + CStr(Trim(nEntityID.ToString)) + "_" + strToday + ".doc"
            'sAddress = strAddress.Split(",")
            sAddress = arrAddress
            If MsgBox("Do you want to create one label and indicate it's location on label page", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                colParams.Add("<NAME>", strName)
                colParams.Add("<ADDRESS1>", sAddress(0))
                If sAddress(1) = String.Empty Then
                    colParams.Add("<ADDRESS2>", sAddress(2) + " " + sAddress(3) + " " + sAddress(4))
                    colParams.Add("<CITYSTATEZIP>", String.Empty)
                Else
                    colParams.Add("<ADDRESS2>", sAddress(1))
                    colParams.Add("<CITYSTATEZIP>", sAddress(2) + " " + sAddress(3) + " " + sAddress(4))
                End If
                strTempPath = TmpltPath & "Global\AddressedLabels.doc"
                Dim oWord As Word.Application = MusterContainer.GetWordApp

                If Not oWord Is Nothing Then
                    ltrGen.CreateLabels("Global", "Label", colParams, strTempPath, String.Empty, oWord)

                    'ltrGen.CreateLetter("Global", "Label", colParams, strTempPath, DOC_PATH & strDOC_NAME, oWord)
                    oword.Visible = True
                End If
                oWord = Nothing

            Else
                input = InputBox("Enter row  (1-10 numeric only):")
                If input <> String.Empty Then
                    nRow = Integer.Parse(input)
                    If nRow > 10 Or nRow <= 0 Then
                        MsgBox("Invalid Location. Should be from 1 to 10")
                        Exit Function
                    End If
                End If
                input = String.Empty
                input = InputBox("Enter Column (1-2 numeric only)")
                If input <> String.Empty Then
                    nColumn = Integer.Parse(input)
                    If nColumn = 2 Then
                        nColumn = 3
                    End If
                    If nColumn <> 3 And nColumn <> 1 Then
                        MsgBox("Invalid Column Location. Should be either 1 or 2")
                        Exit Function
                    End If
                End If
                strTempAddress = strName + vbCrLf + sAddress(0) + vbCrLf + IIf(sAddress(1) = String.Empty, "", sAddress(1) + vbCrLf) + sAddress(2) + ", " + sAddress(3) + " " + sAddress(4)
                strTempPath = TmpltPath & "Global\Labels.doc"
                Dim oWord As Word.Application = MusterContainer.GetWordApp

                If Not oWord Is Nothing Then

                    ltrGen.CreateLabels("Global", "Label", colParams, strTempPath, DOC_PATH & strDOC_NAME, oword, strTempAddress, nRow, nColumn)
                    oword.Visible = True
                End If
                oWord = Nothing




            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region
#Region "UltraGrid Initialize"
    Friend Shared Sub ug_InitializePrintPreview(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelablePrintPreviewEventArgs, _
                                                Optional ByVal strHeaderText As String = "", _
                                                Optional ByVal columnClipMode As Infragistics.Win.UltraWinGrid.ColumnClipMode = Infragistics.Win.UltraWinGrid.ColumnClipMode.RepeatClippedColumns, _
                                                Optional ByVal autoFitColumns As Boolean = True, _
                                                Optional ByVal printRange As System.Drawing.Printing.PrintRange = Printing.PrintRange.AllPages, _
                                                Optional ByVal landscape As Boolean = True, _
                                                Optional ByVal pageHeaderHeight As Integer = 20, _
                                                Optional ByVal pageHeaderFontBold As Infragistics.Win.DefaultableBoolean = Infragistics.Win.DefaultableBoolean.True, _
                                                Optional ByVal pageHeaderFontSize As Single = 10, _
                                                Optional ByVal fitWidthToPages As Integer = 0)
        e.DefaultLogicalPageLayoutInfo.ColumnClipMode = columnClipMode
        e.PrintLayout.AutoFitColumns = True
        e.PrintDocument.PrinterSettings.PrintRange = printRange
        e.PrintDocument.DefaultPageSettings.Landscape = landscape
        e.DefaultLogicalPageLayoutInfo.PageHeader = strHeaderText
        e.DefaultLogicalPageLayoutInfo.PageHeaderHeight = pageHeaderHeight
        e.DefaultLogicalPageLayoutInfo.PageHeaderAppearance.FontData.Bold = pageHeaderFontBold
        e.DefaultLogicalPageLayoutInfo.PageHeaderAppearance.TextHAlign = Infragistics.Win.HAlign.Center
        e.DefaultLogicalPageLayoutInfo.PageHeaderAppearance.FontData.SizeInPoints = pageHeaderFontSize
        e.DefaultLogicalPageLayoutInfo.FitWidthToPages = fitWidthToPages
    End Sub
#End Region

    Public Sub New()

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
