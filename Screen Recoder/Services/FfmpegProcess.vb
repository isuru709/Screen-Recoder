Imports System.Diagnostics
Imports System.Text
Imports System.Threading

Namespace Services
  Public Class FfmpegProcess
    Implements IDisposable

    Private _proc As Process
    Private _cts As CancellationTokenSource

    Public Event OutputData(line As String)
    Public Event ErrorData(line As String)
    Public Event Progress(key As String, value As String)
    Public Event Exited(exitCode As Integer)

    Public ReadOnly Property IsRunning As Boolean
      Get
        Return _proc IsNot Nothing AndAlso Not _proc.HasExited
      End Get
    End Property

    Public Sub Start(ffmpegPath As String, args As String)
      If String.IsNullOrWhiteSpace(ffmpegPath) OrElse Not IO.File.Exists(ffmpegPath) Then
        Throw New IO.FileNotFoundException("FFmpeg not found.", ffmpegPath)
      End If

      _cts = New CancellationTokenSource()
      _proc = New Process() With {
        .StartInfo = New ProcessStartInfo With {
          .FileName = ffmpegPath,
          .Arguments = args,
          .UseShellExecute = False,
          .RedirectStandardOutput = True,
          .RedirectStandardError = True,
          .RedirectStandardInput = True,
          .CreateNoWindow = True,
          .StandardOutputEncoding = Encoding.UTF8,
          .StandardErrorEncoding = Encoding.UTF8
        },
        .EnableRaisingEvents = True
      }

      AddHandler _proc.OutputDataReceived, AddressOf OnStdOut
      AddHandler _proc.ErrorDataReceived, AddressOf OnStdErr
      AddHandler _proc.Exited, Sub() RaiseEvent Exited(_proc.ExitCode)

      If Not _proc.Start() Then
        Throw New Exception("Failed to start FFmpeg.")
      End If

      _proc.BeginOutputReadLine()
      _proc.BeginErrorReadLine()
    End Sub

    Private Sub OnStdOut(sender As Object, e As DataReceivedEventArgs)
      If String.IsNullOrEmpty(e.Data) Then Return
      RaiseEvent OutputData(e.Data)

      ' Parse -progress pipe:1 key=value pairs
      Dim parts = e.Data.Split("="c)
      If parts.Length = 2 Then
        RaiseEvent Progress(parts(0).Trim(), parts(1).Trim())
      End If
    End Sub

    Private Sub OnStdErr(sender As Object, e As DataReceivedEventArgs)
      If String.IsNullOrEmpty(e.Data) Then Return
      RaiseEvent ErrorData(e.Data)
    End Sub

    Public Sub StopGracefully(Optional timeoutMs As Integer = 4000)
      Try
        If IsRunning Then
          ' Send "q" to ask FFmpeg to quit
          _proc.StandardInput.WriteLine("q")
          If Not _proc.WaitForExit(timeoutMs) Then
            _proc.Kill(True)
          End If
        End If
      Catch
      End Try
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
      Try
        If IsRunning Then
          StopGracefully()
        End If
        _proc?.Dispose()
        _cts?.Dispose()
      Catch
      End Try
    End Sub
  End Class
End Namespace