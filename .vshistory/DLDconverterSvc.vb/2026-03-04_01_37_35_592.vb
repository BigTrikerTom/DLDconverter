Imports System
Imports System.IO
Imports System.ServiceProcess
Imports System.Timers

Public Class DLDconverterSvc

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
    End Sub

End Class


Public Class FileWatcherService
    Inherits ServiceBase

    Private WithEvents watcher As FileSystemWatcher
    Private logPath As String = "C:\Logs\FileWatcherService.log"
    Private sourceDirectory As String = "C:\Watch"
    Private targetDirectory As String = "C:\Output"

    Public Sub New()
        ServiceName = "DLDFileWatcherService"
        DisplayName = "DLD File Watcher Service"
        CanStop = True
        CanPauseAndContinue = False
        AutoLog = True
    End Sub

    Public Property DisplayName As String


    Protected Overrides Sub OnStart(args() As String)
        Try
            InitializeWatcher()
            WriteLog("Service started successfully.")
        Catch ex As Exception
            WriteLog($"Error starting service: {ex.Message}")
        End Try
    End Sub

    Protected Overrides Sub OnStop()
        Try
            If watcher IsNot Nothing Then
                watcher.EnableRaisingEvents = False
                watcher.Dispose()
            End If
            WriteLog("Service stopped successfully.")
        Catch ex As Exception
            WriteLog($"Error stopping service: {ex.Message}")
        End Try
    End Sub

    Private Sub InitializeWatcher()
        watcher = New FileSystemWatcher()
        watcher.Path = sourceDirectory
        watcher.Filter = "*.dld"
        watcher.NotifyFilter = NotifyFilters.FileName Or NotifyFilters.LastWrite
        watcher.EnableRaisingEvents = True
    End Sub

    Private Sub Watcher_Created(sender As Object, e As FileSystemEventArgs) Handles watcher.Created
        Try
            System.Threading.Thread.Sleep(500) ' Warte bis Datei vollständig geschrieben ist

            Dim sourceFile As String = e.FullPath
            Dim fileName As String = Path.GetFileName(sourceFile)
            Dim targetFile As String = Path.Combine(targetDirectory, fileName)

            If File.Exists(sourceFile) Then
                ' Stelle sicher, dass Zielverzeichnis existiert
                If Not Directory.Exists(targetDirectory) Then
                    Directory.CreateDirectory(targetDirectory)
                End If

                ' Kopiere die Datei
                File.Copy(sourceFile, targetFile, True)
                WriteLog($"File copied: {fileName} to {targetDirectory}")

                ' Lösche Originaldatei
                File.Delete(sourceFile)
                WriteLog($"Source file deleted: {fileName}")
            End If

        Catch ex As Exception
            WriteLog($"Error processing file {e.Name}: {ex.Message}")
        End Try
    End Sub

    Private Sub WriteLog(message As String)
        Try
            Dim logDirectory As String = Path.GetDirectoryName(logPath)
            If Not Directory.Exists(logDirectory) Then
                Directory.CreateDirectory(logDirectory)
            End If

            Dim logEntry As String = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}"
            File.AppendAllText(logPath, logEntry & Environment.NewLine)
        Catch
            ' Fehler beim Schreiben des Logs ignorieren
        End Try
    End Sub

    Public Shared Sub Main()
        Dim servicesToRun() As ServiceBase = {New FileWatcherService()}
        ServiceBase.Run(servicesToRun)
    End Sub
End Class

