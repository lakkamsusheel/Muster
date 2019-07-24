Imports System
Imports System.IO
Imports System.Management
Imports Microsoft.Win32.Registry


Public Class FacPicFileUpload
    Inherits System.Windows.Forms.UserControl

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal facID As Integer, ByVal inspDate As DateTime, ByVal thisModuleID As UIUtilsGen.ModuleID)
        MyBase.New()



        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Me.facility_id = facID
        Me.dtInspDate = inspDate
        Me.moduleID = thisModuleID


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
    Friend WithEvents lblDrive As System.Windows.Forms.Label
    Friend WithEvents lvlDrive As System.Windows.Forms.Label
    Friend WithEvents btnScanDoc As System.Windows.Forms.Button
    Friend WithEvents btnLoadPic As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblDrive = New System.Windows.Forms.Label
        Me.lvlDrive = New System.Windows.Forms.Label
        Me.btnScanDoc = New System.Windows.Forms.Button
        Me.btnLoadPic = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'lblDrive
        '
        Me.lblDrive.Location = New System.Drawing.Point(8, 8)
        Me.lblDrive.Name = "lblDrive"
        Me.lblDrive.Size = New System.Drawing.Size(376, 16)
        Me.lblDrive.TabIndex = 1
        Me.lblDrive.Text = "Click on Drive Info After Plugging it in"
        '
        'lvlDrive
        '
        Me.lvlDrive.BackColor = System.Drawing.Color.Beige
        Me.lvlDrive.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lvlDrive.Enabled = False
        Me.lvlDrive.Location = New System.Drawing.Point(16, 24)
        Me.lvlDrive.Name = "lvlDrive"
        Me.lvlDrive.Size = New System.Drawing.Size(536, 24)
        Me.lvlDrive.TabIndex = 2
        Me.lvlDrive.Text = "Connect Camera or USB drive to load Pictures"
        '
        'btnScanDoc
        '
        Me.btnScanDoc.Location = New System.Drawing.Point(712, 24)
        Me.btnScanDoc.Name = "btnScanDoc"
        Me.btnScanDoc.Size = New System.Drawing.Size(128, 24)
        Me.btnScanDoc.TabIndex = 3
        Me.btnScanDoc.Text = "Load Scanned PDF"
        '
        'btnLoadPic
        '
        Me.btnLoadPic.Location = New System.Drawing.Point(560, 24)
        Me.btnLoadPic.Name = "btnLoadPic"
        Me.btnLoadPic.Size = New System.Drawing.Size(144, 24)
        Me.btnLoadPic.TabIndex = 4
        Me.btnLoadPic.Text = "Load Picture from Device"
        '
        'FacPicFileUpload
        '
        Me.Controls.Add(Me.btnLoadPic)
        Me.Controls.Add(Me.btnScanDoc)
        Me.Controls.Add(Me.lvlDrive)
        Me.Controls.Add(Me.lblDrive)
        Me.Name = "FacPicFileUpload"
        Me.Size = New System.Drawing.Size(848, 56)
        Me.ResumeLayout(False)

    End Sub

#End Region




    Private WithEvents m_MediaConnectWatcher As ManagementEventWatcher
    Public USBDriveName As String
    Public USBDriveLetter As String
    Private facility_id As Integer
    Private moduleID As Integer
    Private fdDriverFiles As OpenFileDialog

    Private dtInspDate As DateTime

    Public Event FilesDone(ByVal sender As Object, ByVal e As EventArgs)


    Public Sub StartDetection()
        ' __InstanceOperationEvent will trap both Creation and Deletion of class instances


        Dim query As ObjectQuery
        Dim query2 As WqlEventQuery
        Dim a As ManagementObjectSearcher

        Try
            query = New ObjectQuery("Select * from Win32_DiskDrive")
            query2 = New WqlEventQuery("SELECT * FROM __InstanceOperationEvent WITHIN 1 " _
      & "WHERE TargetInstance ISA 'Win32_DiskDrive'")

            m_MediaConnectWatcher = New ManagementEventWatcher

            m_MediaConnectWatcher.Query = query2

            m_MediaConnectWatcher.Start()



            a = New System.Management.ManagementObjectSearcher(query)

            For Each item As ManagementObject In a.Get
                GetDriveInfo(item, True)
            Next

        Catch ex As Exception
            Throw ex
        Finally
            If Not a Is Nothing Then
                a.Dispose()
            End If

            If Not query Is Nothing Then
                query = Nothing
            End If

            If Not query2 Is Nothing Then
                query2 = Nothing
            End If


        End Try





    End Sub


    Sub GetDriveInfo(ByVal mbo As ManagementBaseObject, Optional ByVal direct As Boolean = False)

        Dim obj As ManagementBaseObject
        Dim str, drive As String
        Dim path As String


        ' next we need a copy of the instance that was either created or deleted
        If Not direct Then
            obj = CType(mbo("TargetInstance"), ManagementBaseObject)
            path = mbo.ClassPath.ClassName
        Else
            obj = mbo
            path = "__InstanceCreationEvent"
        End If


        Select Case path
            Case "__InstanceCreationEvent"
                If obj("InterfaceType") = "USB" Then
                    drive = GetDriveLetterFromDisk(obj("Name"))
                    str = String.Format("{0} ( Drive Letter : {1} has been plugged in) ", obj("Caption"), drive)

                    Me.USBDriveLetter = drive
                    Me.USBDriveName = str

                    Me.lvlDrive.Text = str
                    Me.lvlDrive.Enabled = True
                End If

            Case "__InstanceDeletionEvent"
                If obj("InterfaceType") = "USB" Then

                    If USBDriveName.ToUpper.StartsWith(obj("Caption").ToString.ToUpper) Then
                        USBDriveLetter = ""
                        USBDriveName = ""
                    End If

                    lvlDrive.Text = "Connect Camera or USB drive to load Pictures"
                    lvlDrive.Enabled = False
                End If
        End Select

    End Sub
    Private Sub Arrived(ByVal sender As Object, ByVal e As System.Management.EventArrivedEventArgs) Handles m_MediaConnectWatcher.EventArrived

        Dim mbo, obj As ManagementBaseObject
        Dim str, drive As String


        ' the first thing we have to do is figure out if this is a creation or deletion event
        mbo = CType(e.NewEvent, ManagementBaseObject)
        ' next we need a copy of the instance that was either created or deleted
        obj = CType(mbo("TargetInstance"), ManagementBaseObject)

        Me.GetDriveInfo(mbo)
    End Sub

    Private Function GetDriveLetterFromDisk(ByVal Name As String) As String
        Dim oq_part, oq_disk As ObjectQuery
        Dim mos_part, mos_disk As ManagementObjectSearcher
        Dim obj_part, obj_disk As ManagementObject
        Dim ans As String = ""

        ' WMI queries use the "\" as an escape charcter
        Name = Replace(Name, "\", "\\")

        ' First we map the Win32_DiskDrive instance with the association called
        ' Win32_DiskDriveToDiskPartition. Then we map the Win23_DiskPartion
        ' instance with the assocation called Win32_LogicalDiskToPartition

        oq_part = New ObjectQuery("ASSOCIATORS OF {Win32_DiskDrive.DeviceID=""" & Name & """} WHERE AssocClass = Win32_DiskDriveToDiskPartition")
        mos_part = New ManagementObjectSearcher(oq_part)
        If Not mos_part Is Nothing Then
            For Each obj_part In mos_part.Get()

                oq_disk = New ObjectQuery("ASSOCIATORS OF {Win32_DiskPartition.DeviceID=""" & obj_part("DeviceID") & """} WHERE AssocClass = Win32_LogicalDiskToPartition")
                mos_disk = New ManagementObjectSearcher(oq_disk)
                For Each obj_disk In mos_disk.Get()
                    ans &= obj_disk("Name") & ","
                Next
            Next
        End If


        Return ans.Trim(","c)
    End Function

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        StartDetection()
    End Sub

    Function ArrayToStr(ByVal arrayItem() As Integer) As String

        Dim s As New Text.StringBuilder
        Dim ret As String = String.Empty

        For index As Integer = arrayItem.GetUpperBound(0) To 0 Step -1
            If arrayItem(index) >= 1 And arrayItem(index) <= 28 Then
                s.Append(Chr(Asc("A") + (arrayItem(index) - 1)))
            End If
        Next

        ret = s.ToString
        s.Length = 0
        Return ret
    End Function

    Sub IncArray(ByVal index As Integer, ByRef arrayItem() As Integer)
        arrayItem(index) += 1
        If arrayItem(index) = 29 Then
            arrayItem(index) = 1
            If index > 0 Then
                IncArray(index - 1, arrayItem)
            Else
                Throw New Exception("All names (AAAA-ZZZZ) has been used for this facility at this date")
            End If
        End If


    End Sub




    Function MoveFile(ByVal filename As String, ByVal PIC_PATH As String, ByVal ext As String) As String


        Dim shortname As String = filename.Substring(filename.LastIndexOf("\"))
        Dim newName As String

        Dim picarray(4) As Integer
        Dim start As String = "F_"

        Try

            If ext.ToUpper = "PDF" Then
                start = String.Format("SYS_SCANNED_{0}_PDF_", UIUtilsGen.GetModuleNameByID(moduleID))
            Else
                start = String.Format("{0}_F_", UIUtilsGen.GetModuleNameByID(moduleID))
            End If

            newName = String.Format("{0}{1}", PIC_PATH, String.Format("{4}{0}_{1}_{2}.{3}", facility_id, dtInspDate.ToString("MM_dd_yyyy"), ArrayToStr(picarray), ext, start))

            While File.Exists(newName)
                IncArray(4, picarray)
                newName = String.Format("{0}{1}", PIC_PATH, String.Format("{4}{0}_{1}_{2}.{3}", facility_id, dtInspDate.ToString("MM_dd_yyyy"), ArrayToStr(picarray), ext, start))
            End While


            File.Copy(filename, newName)

            Return newName

        Catch ex As Exception
            Throw ex
        End Try




    End Function

    Private Sub btnScanDoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnScanDoc.Click

        fdDriverFiles = New OpenFileDialog

        Dim PIC_PATH As String = String.Format("{0}\", MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_SystemArchive).ProfileValue)
        Dim newName As String
        Dim id As Integer
        Dim strModuleName = UIUtilsGen.GetModuleNameByID(moduleID)
        Dim strPrintedPath = String.Format("{0}{1}\{2}\", PIC_PATH, strModuleName, CStr(Format(Now, "yyyy")))

        Dim doNotes As Boolean = False


        With fdDriverFiles
            If Not USBDriveLetter Is Nothing AndAlso Me.USBDriveLetter.Length > 0 Then
                .InitialDirectory = String.Format("{0}\", Me.USBDriveLetter)
            Else
                .InitialDirectory = "C:\"
            End If

            .Filter = "PDF files (*.pdf) | *.pdf"
            .FilterIndex = 2
            .RestoreDirectory = True
            .Multiselect = True
            .Title = "PDF Files to load as scanned documents"



            If .ShowDialog() = System.Windows.Forms.DialogResult.OK AndAlso Not .FileNames Is Nothing AndAlso .FileNames.GetUpperBound(0) >= 0 Then
                Try

                    doNotes = (MsgBox("Do you wish to enter some notation for each Scanned Document", MsgBoxStyle.YesNo, "Document (PFD) notation") = MsgBoxResult.Yes)

                    For Each item As String In .FileNames

                        Dim notes As String = String.Empty

                        newName = MoveFile(item, strPrintedPath, "PDF")

                        If doNotes Then
                            notes = InputBox(String.Format("Enter Notes for {0}", item), "Notation on PDF")
                            If notes.Length > 0 Then notes = String.Format("({0})", notes).Trim
                        End If

                        id = UIUtilsGen.SaveDocument(facility_id, UIUtilsGen.EntityTypes.Facility, newName.Substring(newName.LastIndexOf("\") + 1), _
                                      String.Format("Scanned {0} Document", strModuleName), strPrintedPath, String.Format("Scanned {0} Document for {1} {2}", UIUtilsGen.GetModuleNameByID(moduleID), facility_id, notes), moduleID, 0, 0, 0)

                        'complete archiving of database record
                        MusterContainer.pLetter.UpdatePrintedStatus(id, strPrintedPath, Today.ToShortDateString)

                    Next


                    MsgBox("File(s) has been copied and inserted into this facility")
                    RaiseEvent FilesDone(Me, New EventArgs)

                Catch Ex As Exception
                    MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
                Finally
                    ' Check this again, since we need to make sure we didn't throw an exception on open.
                    .Dispose()
                End Try
            End If

        End With


    End Sub

    Private Sub btnLoadPic_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoadPic.Click


        Dim PIC_PATH As String = String.Format("{0}\", MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_FacImages).ProfileValue)
        fdDriverFiles = New OpenFileDialog
        With fdDriverFiles

            If Not USBDriveLetter Is Nothing AndAlso Me.USBDriveLetter.Length > 0 Then
                .InitialDirectory = String.Format("{0}\", Me.USBDriveLetter)
            Else
                .InitialDirectory = "C:\"
            End If
            .Filter = "JPEG files (*.jpg) | *.jpg"

            .FilterIndex = 2
            .RestoreDirectory = True
            .Multiselect = True
            .Title = "JPEG Loader for Inspection Site Pictures"


            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Try
                    For Each item As String In .FileNames
                        MoveFile(item, PIC_PATH, "JPG")
                    Next

                    MsgBox("File(s) has been copied and inserted into this facility")
                    RaiseEvent FilesDone(Me, New EventArgs)

                Catch Ex As Exception
                    MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
                Finally
                    ' Check this again, since we need to make sure we didn't throw an exception on open.
                    .Dispose()
                End Try
            End If

        End With

    End Sub
End Class
