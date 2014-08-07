Imports System.IO

Public Class Main_Screen
    Inherits System.Windows.Forms.Form

    Dim WithEvents Worker1 As Worker

    Private workerbusy As Boolean = False
    Private steps As Integer = 0
    'steps 0: process not launched
    'steps 1: line count

    Public Delegate Sub WorkerComplete_h()
    Public Delegate Sub WorkerError_h(ByVal Message As Exception)
    Public Delegate Sub WorkerStatusMessage_h(ByVal message As String, ByVal statustag As Integer)

    Private application_exit As Boolean = False
    Private shutting_down As Boolean = False
    Private splash_loader As Splash_Screen
    Public dataloaded As Boolean = False
    Private error_reporting_level

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerCompleteHandler
        AddHandler Worker1.WorkerError, AddressOf WorkerErrorHandler
        AddHandler Worker1.WorkerStatusMessage, AddressOf WorkerStatusMessageHandler
    End Sub

    Public Sub New(ByVal splash As Splash_Screen)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        splash_loader = splash
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerCompleteHandler
        AddHandler Worker1.WorkerError, AddressOf WorkerErrorHandler
        AddHandler Worker1.WorkerStatusMessage, AddressOf WorkerStatusMessageHandler
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
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents ButtonOperationLaunch As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents ServiceName As System.Windows.Forms.Label
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ProgressLabel As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtStatus As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Main_Screen))
        Me.Label3 = New System.Windows.Forms.Label
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ButtonOperationLaunch = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.ProgressLabel = New System.Windows.Forms.Label
        Me.txtStatus = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.ServiceName = New System.Windows.Forms.Label
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(24, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(312, 16)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "The following Windows services will now be installed:"
        '
        'ButtonOperationLaunch
        '
        Me.ButtonOperationLaunch.BackColor = System.Drawing.Color.LightSteelBlue
        Me.ButtonOperationLaunch.Enabled = False
        Me.ButtonOperationLaunch.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonOperationLaunch.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonOperationLaunch.Location = New System.Drawing.Point(376, 216)
        Me.ButtonOperationLaunch.Name = "ButtonOperationLaunch"
        Me.ButtonOperationLaunch.Size = New System.Drawing.Size(88, 20)
        Me.ButtonOperationLaunch.TabIndex = 10
        Me.ButtonOperationLaunch.Text = "Install"
        Me.ToolTip1.SetToolTip(Me.ButtonOperationLaunch, "Launches Windows Service Installer Operation")
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.DimGray
        Me.Label1.Location = New System.Drawing.Point(-8, 248)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(512, 40)
        Me.Label1.TabIndex = 21
        Me.Label1.Text = "              BUILD 20060110.1"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolTip1.SetToolTip(Me.Label1, "BUILD 20060109.1")
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.ProgressLabel)
        Me.Panel1.Location = New System.Drawing.Point(24, 216)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(200, 16)
        Me.Panel1.TabIndex = 25
        Me.ToolTip1.SetToolTip(Me.Panel1, "Job Progress")
        '
        'ProgressLabel
        '
        Me.ProgressLabel.BackColor = System.Drawing.Color.SkyBlue
        Me.ProgressLabel.Location = New System.Drawing.Point(0, 0)
        Me.ProgressLabel.Name = "ProgressLabel"
        Me.ProgressLabel.Size = New System.Drawing.Size(0, 16)
        Me.ProgressLabel.TabIndex = 0
        '
        'txtStatus
        '
        Me.txtStatus.BackColor = System.Drawing.Color.White
        Me.txtStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStatus.ForeColor = System.Drawing.Color.SteelBlue
        Me.txtStatus.Location = New System.Drawing.Point(120, 256)
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.Size = New System.Drawing.Size(344, 24)
        Me.txtStatus.TabIndex = 27
        Me.txtStatus.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.txtStatus, "BUILD 20060109.1")
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.White
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(-8, -8)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(504, 64)
        Me.PictureBox1.TabIndex = 22
        Me.PictureBox1.TabStop = False
        '
        'ServiceName
        '
        Me.ServiceName.BackColor = System.Drawing.Color.White
        Me.ServiceName.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ServiceName.Location = New System.Drawing.Point(24, 6)
        Me.ServiceName.Name = "ServiceName"
        Me.ServiceName.Size = New System.Drawing.Size(408, 40)
        Me.ServiceName.TabIndex = 1
        Me.ServiceName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ListBox1
        '
        Me.ListBox1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ListBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ListBox1.ForeColor = System.Drawing.Color.Black
        Me.ListBox1.Location = New System.Drawing.Point(24, 88)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(440, 80)
        Me.ListBox1.TabIndex = 24
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(24, 176)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(448, 32)
        Me.Label2.TabIndex = 26
        Me.Label2.Text = "To proceed with the installation please click on the 'Install' button below. Note" & _
        " that you cannot pause nor interrupt the installation once it has begun."
        '
        'Main_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.AliceBlue
        Me.ClientSize = New System.Drawing.Size(495, 280)
        Me.Controls.Add(Me.txtStatus)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.ServiceName)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.ButtonOperationLaunch)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.ForeColor = System.Drawing.Color.Black
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(-1, -1)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(503, 314)
        Me.MinimumSize = New System.Drawing.Size(503, 314)
        Me.Name = "Main_Screen"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Windows Service Installer"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg & ":" & ex.ToString)
                filewriter.Flush()
                filewriter.Close()
            End If
        Catch exc As Exception
            MsgBox("An error occurred in Windows Service Installer's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim display As Display_Message = New Display_Message(identifier_msg & ":" & ex.ToString)
                display.ShowDialog()
            End If
        Catch exc As Exception
            MsgBox("An error occurred in Windows Service Installer's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub Activity_Logger(ByVal message As String)
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & Format(Now(), "yyyyMMdd") & "_Activity_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & message)
            filewriter.Flush()
            filewriter.Close()
        Catch ex As Exception
            Error_Handler(ex, "Activity Logger")
        End Try
    End Sub


    Private Sub SendMessage(ByVal labelname As String, ByVal message As String)
        Try
            Dim controllist As ControlCollection = Me.Controls
            Dim cont As Control

            For Each cont In controllist
                If cont.Name = labelname Then
                    cont.Text = message
                    cont.Refresh()
                    Exit For
                End If
            Next
        Catch ex As Exception
            If (Not Worker1.WorkerThread.ThreadState.ToString.IndexOf("Aborted") > -1) And (Not Worker1.WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1) Then
                Error_Handler(ex, "SendMessage")
            End If
        End Try
    End Sub

    Private Sub ButtonOperationLaunch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonOperationLaunch.Click
        ButtonOperationLaunch.Enabled = False
        ProgressLabel.Width = 0
        ProgressLabel.Refresh()
        SendMessage("txtStatus", "")
        Worker1.filequeue1 = ListBox1
        steps = 1
        workerbusy = True
        If ButtonOperationLaunch.Text = "Install" Then
            Activity_Logger("BEGIN INSTALL OPERATION")
            Worker1.ChooseThreads(1)
        End If
        If ButtonOperationLaunch.Text = "Uninstall" Then
            Activity_Logger("BEGIN UNINSTALL OPERATION")
            Worker1.ChooseThreads(2)
        End If
        If ButtonOperationLaunch.Text = "Exit" Then
            Me.Close()
        End If

    End Sub

    Public Sub WorkerStatusMessageHandler(ByVal message As String, ByVal statustag As Integer)
        Try
            ProgressLabel.BackColor() = Color.SkyBlue
            ProgressLabel.Width = ProgressLabel.Width + System.Math.Round(((200 / ListBox1.Items.Count) / 8))
            ProgressLabel.Refresh()
            Activity_Logger(message)
            SendMessage("txtStatus", message)

            If message = "SERVICE ALREADY EXISTS ON SYSTEM" And ButtonOperationLaunch.Text = "Install" Then
                ProgressLabel.BackColor() = Color.Orange
                ProgressLabel.Width = 200
                ProgressLabel.Refresh()
                MsgBox("It has been discovered that the service to be installed already exists on the system. You will need to uninstall this service before you can continue with this installation. The offending service can be found at: " & Worker1.current_service, MsgBoxStyle.Information, "Installation Aborted")
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerErrorHandler(ByVal Message As Exception)
        Try
            Error_Handler(Message)
        Catch ex As Exception
            Error_Handler(ex, "WorkerErrorHandler")
        End Try
    End Sub


    Public Sub WorkerCompleteHandler(ByVal queue As Integer)
        Try
            Dim eventhandled As Boolean = False
            ProgressLabel.Width = 200
            ProgressLabel.Refresh()
            Dim changed As Boolean = False
            If txtStatus.Text = "SERVICE ALREADY EXISTS ON SYSTEM" Then
                ButtonOperationLaunch.Text = "Uninstall"
                changed = True
            End If
            If txtStatus.Text = "SERVICE UNINSTALLED AND REMOVED FROM SYSTEM" Then
                ButtonOperationLaunch.Text = "Install"
                changed = True
            End If
            If changed = False Then
                ButtonOperationLaunch.Text = "Exit"
            End If

            workerbusy = False
            ButtonOperationLaunch.Enabled = True
            eventhandled = True

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub



    Private Sub Main_Screen_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            Worker1.Dispose()
        Catch ex As Exception
            Error_Handler(ex)
        End Try

    End Sub


    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Try
                Dim finfo As FileInfo = New FileInfo((Application.StartupPath & "\inputs.txt").Replace("\\", "\"))
                If finfo.Exists = True Then
                    Dim filereader As StreamReader = New StreamReader((Application.StartupPath & "\inputs.txt").Replace("\\", "\"))
                    If filereader.Peek > -1 Then
                        ServiceName.Text = filereader.ReadLine.Replace("ServiceName=", "") & " Installer"
                    Else
                        ServiceName.Text = "Service Installer"
                    End If
                Else
                    ServiceName.Text = "Service Installer"
                End If
                finfo = Nothing
            Catch ex As Exception
                Error_Handler(ex, "Screen Load")
            End Try
            Try
                ListBox1.Items.Clear()
                Dim dinfo As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\Install").Replace("\\", "\"))
                If dinfo.Exists = True Then
                    Dim files As FileInfo
                    For Each files In dinfo.GetFiles("*.exe")
                        ListBox1.Items.Add(files.Name)
                    Next
                    files = Nothing
                End If
                dinfo = Nothing
                If ListBox1.Items.Count > 0 Then
                    ButtonOperationLaunch.Enabled = True
                End If
            Catch ex As Exception
                Error_Handler(ex, "Screen Load")
            End Try
            SendMessage("txtStatus", "INSTALLER INITIALISED")
            ListBox1.Select()
            ListBox1.Focus()
            dataloaded = True
            splash_loader.Visible = False
        Catch ex As Exception
            Error_Handler(ex, "Screen Load")
        End Try
    End Sub




End Class
