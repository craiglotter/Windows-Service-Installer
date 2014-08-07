Imports System.IO
Imports System.Text


Public Class Worker

    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.

    Public WorkerThread As System.Threading.Thread

    Public filequeue1 As ListBox
    Public current_service As String = ""
  
    Public Event WorkerStatusMessage(ByVal message As String, ByVal statustag As Integer)
    Public Event WorkerError(ByVal Message As Exception)
    Public Event WorkerComplete(ByVal queue As Integer)


#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)

    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        
    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub Error_Handler(ByVal message As Exception)
        Try
            If (Not WorkerThread.ThreadState.ToString.IndexOf("Aborted") > -1) And (Not WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1) Then
                RaiseEvent WorkerError(message)
            End If
        Catch ex As Exception
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
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            ' Determines which thread to start based on the value it receives.
            Select Case threadNumber
                Case 1
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerExecute_Routine)
                    WorkerThread.Start()
                Case 2
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerUninstall_Routine)
                    WorkerThread.Start()
            End Select
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub WorkerExecute_Routine()


        Try
            Dim runner As Integer = 0
            For runner = 0 To filequeue1.Items.Count - 1
                RaiseEvent WorkerStatusMessage("CHECKING FOR EXISTING INSTALLATION", 1)
                Dim destpath As String = (System.Environment.SystemDirectory & "\").Replace("\\", "\") & filequeue1.Items.Item(runner)
                Dim sourcepath As String = (Application.StartupPath & "\").Replace("\\", "\") & "Install\" & filequeue1.Items.Item(runner)
                Dim installer As String = (Application.StartupPath & "\").Replace("\\", "\") & "InstallUtil.exe"
                current_service = destpath
                Activity_Logger("Source Path: " & sourcepath)
                Activity_Logger("Destination Path: " & destpath)
                Dim finfo As FileInfo = New FileInfo(destpath)
                If finfo.Exists = False Then
                    RaiseEvent WorkerStatusMessage("SERVICE NOT FOUND ON SYSTEM", 1)
                    Dim tomove As FileInfo = New FileInfo(sourcepath)
                    If tomove.Exists = True Then
                        RaiseEvent WorkerStatusMessage("COPYING SERVICE TO SYSTEM", 1)
                        tomove.CopyTo(destpath, True)
                    
                        Try
                            Dim testdir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Extra\")
                        If testdir.Exists Then
                            Dim extrainfo As FileInfo
                            For Each extrainfo In testdir.GetFiles()
                                If File.Exists((System.Environment.SystemDirectory & "\" & extrainfo.Name).Replace("\\", "\")) = False Then
                                    extrainfo.CopyTo((System.Environment.SystemDirectory & "\" & extrainfo.Name).Replace("\\", "\"), False)
                                        Activity_Logger("EXTRA: Creating " & (System.Environment.SystemDirectory & "\" & extrainfo.Name).Replace("\\", "\"))
                                    End If
                                Next
                            extrainfo = Nothing
                        End If
                        testdir = Nothing
                        Catch ex As Exception
                            Error_Handler(ex)
                        End Try
                        RaiseEvent WorkerStatusMessage("SERVICE COPIED TO SYSTEM", 1)

                        RaiseEvent WorkerStatusMessage("INSTALLING SERVICE", 1)
                        Activity_Logger("""" & installer & """ """ & destpath & """")
                        DosShellCommand("""" & installer & """ """ & destpath & """")
                        RaiseEvent WorkerStatusMessage("SERVICE INSTALLED", 1)
                        RaiseEvent WorkerStatusMessage("STARTING SERVICE", 1)
                        Activity_Logger("net start """ & filequeue1.Items.Item(runner).ToString().Substring(0, filequeue1.Items.Item(runner).ToString().Length - 4) & """")
                        DosShellCommand("net start """ & filequeue1.Items.Item(runner).ToString().Substring(0, filequeue1.Items.Item(runner).ToString().Length - 4) & """")
                        RaiseEvent WorkerStatusMessage("SERVICE INSTALLED AND STARTED", 1)
                    End If
                Else
                    RaiseEvent WorkerStatusMessage("SERVICE ALREADY EXISTS ON SYSTEM", 1)
                    Exit For
                End If
            Next


        Catch ex As Exception
            Error_Handler(ex)
        End Try

        '        RaiseEvent WorkerStatusMessage("INSTALLATION COMPLETE", 1)
        RaiseEvent WorkerComplete(0)
    End Sub

    Private Sub WorkerUninstall_Routine()
        Try
            Dim runner As Integer = 0
            For runner = 0 To filequeue1.Items.Count - 1
                RaiseEvent WorkerStatusMessage("CHECKING FOR EXISTING INSTALLATION", 1)
                Dim destpath As String = (System.Environment.SystemDirectory & "\").Replace("\\", "\") & filequeue1.Items.Item(runner)
                Dim sourcepath As String = (Application.StartupPath & "\").Replace("\\", "\") & "Install\" & filequeue1.Items.Item(runner)
                Dim installer As String = (Application.StartupPath & "\").Replace("\\", "\") & "InstallUtil.exe"
                current_service = destpath
                Activity_Logger("Source Path: " & sourcepath)
                Activity_Logger("Destination Path: " & destpath)
                Dim finfo As FileInfo = New FileInfo(destpath)
                If finfo.Exists = False Then
                    RaiseEvent WorkerStatusMessage("SERVICE NOT FOUND ON SYSTEM", 1)
                Else
                    RaiseEvent WorkerStatusMessage("SERVICE ALREADY EXISTS ON SYSTEM", 1)

                    Dim tomove As FileInfo = New FileInfo(destpath)
                    If tomove.Exists = True Then
                        RaiseEvent WorkerStatusMessage("STOPPING SERVICE", 1)
                        Activity_Logger("net stop """ & filequeue1.Items.Item(runner).ToString().Substring(0, filequeue1.Items.Item(runner).ToString().Length - 4) & """")
                        DosShellCommand("net stop """ & filequeue1.Items.Item(runner).ToString().Substring(0, filequeue1.Items.Item(runner).ToString().Length - 4) & """")
                        RaiseEvent WorkerStatusMessage("SERVICE STOPPED", 1)
                        RaiseEvent WorkerStatusMessage("UNINSTALLING SERVICE", 1)
                        Activity_Logger("""" & installer & """ /u """ & destpath & """")
                        DosShellCommand("""" & installer & """ /u """ & destpath & """")
                        RaiseEvent WorkerStatusMessage("SERVICE UNINSTALLED", 1)
                        RaiseEvent WorkerStatusMessage("REMOVING SERVICE FROM SYSTEM", 1)
                        tomove.Delete()

                        Try
                            Dim testdir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Extra\")
                            If testdir.Exists Then
                                Dim extrainfo As FileInfo
                                For Each extrainfo In testdir.GetFiles()
                                    If File.Exists((System.Environment.SystemDirectory & "\" & extrainfo.Name).Replace("\\", "\")) = True Then
                                        File.Delete((System.Environment.SystemDirectory & "\" & extrainfo.Name).Replace("\\", "\"))
                                        Activity_Logger("EXTRA: Removing " & (System.Environment.SystemDirectory & "\" & extrainfo.Name).Replace("\\", "\"))
                                    End If
                                Next
                                extrainfo = Nothing
                            End If
                            testdir = Nothing
                        Catch ex As Exception
                            Error_Handler(ex)
                        End Try

                        RaiseEvent WorkerStatusMessage("SERVICE UNINSTALLED AND REMOVED FROM SYSTEM", 1)
                    End If
                End If
            Next
        Catch ex As Exception
            Error_Handler(ex)
        End Try

        '        RaiseEvent WorkerStatusMessage("INSTALLATION COMPLETE", 1)
        RaiseEvent WorkerComplete(0)
    End Sub

    Private Function DosShellCommand(ByVal AppToRun As String) As String
        Dim s As String = ""
        Try
            Dim myProcess As Process = New Process

            myProcess.StartInfo.FileName = "cmd.exe"
            myProcess.StartInfo.UseShellExecute = False
            myProcess.StartInfo.CreateNoWindow = True
            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.RedirectStandardError = True
            myProcess.Start()
            Dim sIn As StreamWriter = myProcess.StandardInput
            sIn.AutoFlush = True

            Dim sOut As StreamReader = myProcess.StandardOutput
            Dim sErr As StreamReader = myProcess.StandardError
            sIn.Write(AppToRun & _
               System.Environment.NewLine)
            sIn.Write("exit" & System.Environment.NewLine)
            s = sOut.ReadToEnd()
            If Not myProcess.HasExited Then
                myProcess.Kill()
            End If

            'MessageBox.Show("The 'dir' command window was closed at: " & myProcess.ExitTime & "." & System.Environment.NewLine & "Exit Code: " & myProcess.ExitCode)

            sIn.Close()
            sOut.Close()
            sErr.Close()
            myProcess.Close()
            'MessageBox.Show(s)
        Catch ex As Exception
            Error_Handler(ex)
        End Try
        Return s
    End Function

End Class
