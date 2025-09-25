Imports System.Collections.Generic
Imports System.Text
Imports ScreenRecoder.Models
Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices

Namespace Services
    Public Class RecorderService
        Implements IDisposable

        Private ReadOnly _proc As New FfmpegProcess()

        Public Event Log(line As String)
        Public Event StatusChanged(status As String)
        Public Event RecordingStopped(exitCode As Integer)
        Public Event Progress(duration As TimeSpan, fps As Double, speed As Double)

        ' Probed device flags (cached)
        Private _devicesProbed As Boolean
        Private _hasDdaGrab As Boolean
        Private _hasGdiGrab As Boolean
        Private _hasD3D11Grab As Boolean
        Private _hasWasapi As Boolean
        Private _hasDshow As Boolean

        ' Available encoders (cached)
        Private ReadOnly _encodersAvailable As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        ' Enumerated device names (cached)
        Private _dshowAudioDevices As New List(Of String)
        Private _dshowMicCandidates As New List(Of String)
        Private _dshowSystemCandidates As New List(Of String)

        Public Sub New()
            AddHandler _proc.OutputData, Sub(l) RaiseEvent Log("[OUT] " & l)
            AddHandler _proc.ErrorData, Sub(l) RaiseEvent Log("[ERR] " & l)
            AddHandler _proc.Exited, Sub(code) RaiseEvent RecordingStopped(code)
            AddHandler _proc.Progress, AddressOf OnProgress
        End Sub

        Private Sub OnProgress(key As String, value As String)
            Static d As Long = 0
            Static f As Double = 0
            Static s As Double = 0
            Select Case key
                Case "out_time_ms" : Long.TryParse(value, d)
                Case "fps" : Double.TryParse(value, f)
                Case "speed"
                    If value.EndsWith("x") Then value = value.Substring(0, value.Length - 1)
                    Double.TryParse(value, s)
            End Select
            If d > 0 Then
                RaiseEvent Progress(TimeSpan.FromMilliseconds(d / 1000.0), f, s)
            End If
        End Sub

        Public Function StartRecording(s As RecorderSettings) As String
            Dim ffmpegPath = GetFfmpegPath()
            Dim ffprobePath = GetFfprobePath()

            If Not File.Exists(ffmpegPath) Then
                Throw New FileNotFoundException("ffmpeg.exe not found next to the application. Please place ffmpeg.exe alongside the program.", ffmpegPath)
            End If
            If Not File.Exists(ffprobePath) Then
                RaiseEvent Log("[WARN] ffprobe.exe not found next to the application. Device probing will use ffmpeg only.")
            End If

            ' Preflight: if capturing window, ensure it exists and is not minimized
            If s.CaptureSource = CaptureSource.WindowTitle Then
                Dim wr As WindowClientRect
                If Not TryGetWindowClientRect(s.WindowTitle, wr) Then
                    Throw New ArgumentException($"Selected window not found or not visible: '{s.WindowTitle}'. Restore the window, refresh the list and pick again.")
                End If
            End If

            Dim output = s.ResolveOutputPath()
            Directory.CreateDirectory(Path.GetDirectoryName(output))
            Dim args = BuildArguments(s, output, ffmpegPath)
            RaiseEvent Log("FFmpeg args: " & args)
            _proc.Start(ffmpegPath, args)
            RaiseEvent StatusChanged("Recording…")
            Return output
        End Function

        Public Sub StopRecording()
            _proc.StopGracefully()
            RaiseEvent StatusChanged("Stopped.")
        End Sub

        Public Sub PauseRecording()
            _proc.Pause()
            RaiseEvent StatusChanged("Paused.")
        End Sub

        Public Sub ResumeRecording()
            _proc.ResumeProcess()
            RaiseEvent StatusChanged("Recording…")
        End Sub

        Private Shared Function GetToolPath(toolExe As String) As String
            Dim baseDir = AppContext.BaseDirectory
            Return Path.Combine(baseDir, toolExe)
        End Function

        Private Shared Function GetFfmpegPath() As String
            Return GetToolPath("ffmpeg.exe")
        End Function

        Private Shared Function GetFfprobePath() As String
            Return GetToolPath("ffprobe.exe")
        End Function

        Private Sub EnsureDevicesProbed()
            If _devicesProbed Then Return
            _devicesProbed = True
            Dim ffmpegPath = GetFfmpegPath()

            Try
                Dim psi As New ProcessStartInfo With {
                    .FileName = ffmpegPath,
                    .Arguments = "-hide_banner -loglevel error -devices",
                    .UseShellExecute = False,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True,
                    .CreateNoWindow = True
                }
                Using p = Process.Start(psi)
                    Dim stdOut = p.StandardOutput.ReadToEnd()
                    Dim stdErr = p.StandardError.ReadToEnd()
                    p.WaitForExit(3000)
                    Dim all = (stdOut & Environment.NewLine & stdErr)
                    _hasDdaGrab = all.IndexOf("ddagrab", StringComparison.OrdinalIgnoreCase) >= 0
                    _hasGdiGrab = all.IndexOf("gdigrab", StringComparison.OrdinalIgnoreCase) >= 0
                    _hasD3D11Grab = all.IndexOf("d3d11grab", StringComparison.OrdinalIgnoreCase) >= 0
                    _hasWasapi = all.IndexOf("wasapi", StringComparison.OrdinalIgnoreCase) >= 0
                    _hasDshow = all.IndexOf("dshow", StringComparison.OrdinalIgnoreCase) >= 0
                End Using

                Dim vids As New List(Of String)
                If _hasDdaGrab Then vids.Add("ddagrab")
                If _hasD3D11Grab Then vids.Add("d3d11grab")
                If _hasGdiGrab Then vids.Add("gdigrab")
                RaiseEvent Log("Detected FFmpeg desktop grab devices: " & If(vids.Count = 0, "(none)", String.Join(", ", vids)))

                Dim auds As New List(Of String)
                If _hasWasapi Then auds.Add("wasapi")
                If _hasDshow Then auds.Add("dshow")
                RaiseEvent Log("Detected FFmpeg audio devices: " & If(auds.Count = 0, "(none)", String.Join(", ", auds)))
            Catch ex As Exception
                RaiseEvent Log("[WARN] Unable to probe FFmpeg devices: " & ex.Message)
            End Try

            Try
                Dim psiEnc As New ProcessStartInfo With {
                    .FileName = ffmpegPath,
                    .Arguments = "-hide_banner -encoders",
                    .UseShellExecute = False,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True,
                    .CreateNoWindow = True
                }
                Using p2 = Process.Start(psiEnc)
                    Dim so = p2.StandardOutput.ReadToEnd()
                    Dim se = p2.StandardError.ReadToEnd()
                    p2.WaitForExit(3000)
                    Dim allEnc = (so & Environment.NewLine & se)

                    Dim known As String() = {
                        "libxvid", "mpeg4",
                        "libx264", "libx265", "libvpx-vp9", "libaom-av1", "libsvtav1", "libvvenc",
                        "h264_nvenc", "hevc_nvenc", "av1_nvenc",
                        "h264_amf", "hevc_amf", "av1_amf",
                        "h264_qsv", "hevc_qsv", "av1_qsv"
                    }
                    For Each name In known
                        If allEnc.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            _encodersAvailable.Add(name)
                        End If
                    Next

                    RaiseEvent Log("Detected FFmpeg video encoders: " & If(_encodersAvailable.Count = 0, "(none)", String.Join(", ", _encodersAvailable)))
                End Using
            Catch ex As Exception
                RaiseEvent Log("[WARN] Unable to probe FFmpeg encoders: " & ex.Message)
            End Try

            If _hasDshow Then
                ProbeDshowDevices()
            End If
        End Sub

        Private Sub ProbeDshowDevices()
            Try
                Dim ffmpegPath = GetFfmpegPath()
                Dim psi As New ProcessStartInfo With {
                    .FileName = ffmpegPath,
                    .Arguments = "-hide_banner -loglevel info -list_devices true -f dshow -i dummy",
                    .UseShellExecute = False,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True,
                    .CreateNoWindow = True
                }
                Using p = Process.Start(psi)
                    Dim stdout = p.StandardOutput.ReadToEnd()
                    Dim stderr = p.StandardError.ReadToEnd()
                    p.WaitForExit(3000)
                    Dim all = stdout & Environment.NewLine & stderr
                    _dshowAudioDevices.Clear()
                    _dshowMicCandidates.Clear()
                    _dshowSystemCandidates.Clear()

                    For Each line In all.Split({Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
                        If line.IndexOf("Alternative name", StringComparison.OrdinalIgnoreCase) >= 0 Then Continue For
                        Dim q1 = line.IndexOf(""""c)
                        Dim q2 = line.LastIndexOf(""""c)
                        If q1 >= 0 AndAlso q2 > q1 Then
                            Dim name = line.Substring(q1 + 1, q2 - q1 - 1)
                            If name.Length > 0 Then
                                _dshowAudioDevices.Add(name)
                            End If
                        End If
                    Next

                    For Each dev In _dshowAudioDevices
                        Dim n = dev.ToLowerInvariant()
                        Dim isSystem As Boolean =
                            n.Contains("stereo mix") OrElse n.Contains("what u hear") OrElse n.Contains("loopback") OrElse
                            n.Contains("mix") OrElse n.Contains("vb-audio") OrElse n.Contains("cable output") OrElse
                            n.Contains("speakers") OrElse n.Contains("output")
                        Dim isMic As Boolean =
                            n.Contains("mic") OrElse n.Contains("microphone") OrElse n.Contains("line in") OrElse
                            n.Contains("headset") OrElse n.Contains("array") OrElse n.Contains("webcam")

                        If isSystem Then _dshowSystemCandidates.Add(dev)
                        If isMic Then _dshowMicCandidates.Add(dev)
                    Next

                    If _dshowAudioDevices.Count > 0 Then
                        RaiseEvent Log("DShow audio devices: " & String.Join("; ", _dshowAudioDevices))
                    End If
                    If _dshowSystemCandidates.Count > 0 Then
                        RaiseEvent Log("DShow system-audio candidates: " & String.Join("; ", _dshowSystemCandidates))
                    Else
                        RaiseEvent Log("[INFO] No obvious system-audio device via DirectShow. If WASAPI is unavailable, install a virtual cable or enable 'Stereo Mix'.")
                    End If
                    If _dshowMicCandidates.Count > 0 Then
                        RaiseEvent Log("DShow microphone candidates: " & String.Join("; ", _dshowMicCandidates))
                    Else
                        RaiseEvent Log("[INFO] No obvious microphone device found via DirectShow.")
                    End If
                End Using
            Catch ex As Exception
                RaiseEvent Log("[WARN] Unable to enumerate DirectShow devices: " & ex.Message)
            End Try
        End Sub

        Private Function GetBestDesktopGrab() As String
            EnsureDevicesProbed()
            If _hasDdaGrab Then Return "ddagrab"
            If _hasD3D11Grab Then Return "d3d11grab"
            If _hasGdiGrab Then Return "gdigrab"
            RaiseEvent Log("[WARN] No desktop grab device detected. Falling back to gdigrab.")
            Return "gdigrab"
        End Function

        Private Function GetDefaultSystemAudioDshow() As String
            EnsureDevicesProbed()
            If _dshowSystemCandidates.Count > 0 Then Return _dshowSystemCandidates(0)
            If _dshowAudioDevices.Count = 1 Then Return _dshowAudioDevices(0)
            Return Nothing
        End Function

        Private Function GetDefaultMicrophoneDshow() As String
            EnsureDevicesProbed()
            If _dshowMicCandidates.Count > 0 Then Return _dshowMicCandidates(0)
            For Each dev In _dshowAudioDevices
                Dim n = dev.ToLowerInvariant()
                Dim looksSystem = n.Contains("stereo mix") OrElse n.Contains("what u hear") OrElse n.Contains("loopback") OrElse
                                  n.Contains("vb-audio") OrElse n.Contains("cable output") OrElse n.Contains("speakers") OrElse n.Contains("output")
                If Not looksSystem Then Return dev
            Next
            Return Nothing
        End Function

        Private Function HasEncoder(name As String) As Boolean
            EnsureDevicesProbed()
            Return _encodersAvailable.Contains(name)
        End Function

        Private Shared Function ResolveCodecId(display As String) As String
            If String.IsNullOrWhiteSpace(display) Then Return "libx264"
            Dim d = display.ToLowerInvariant()
            If d.Contains("xvid") OrElse d.Contains("mpeg") Then Return "libxvid"
            If d.Contains("libx264") Then Return "libx264"
            If d.Contains("libx265") Then Return "libx265"
            If d.Contains("vp9") Then Return "libvpx-vp9"
            If d.Contains("libaom") OrElse d.Contains("aom-av1") Then Return "libaom-av1"
            If d.Contains("libsvt") OrElse d.Contains("svt-av1") Then Return "libsvtav1"
            If d.Contains("vvenc") OrElse d.Contains("vvc") Then Return "libvvenc"
            If d.Contains("h264_nvenc") Then Return "h264_nvenc"
            If d.Contains("hevc_nvenc") Then Return "hevc_nvenc"
            If d.Contains("av1_nvenc") Then Return "av1_nvenc"
            If d.Contains("h264_amf") Then Return "h264_amf"
            If d.Contains("hevc_amf") Then Return "hevc_amf"
            If d.Contains("av1_amf") Then Return "av1_amf"
            If d.Contains("h264_qsv") Then Return "h264_qsv"
            If d.Contains("hevc_qsv") Then Return "hevc_qsv"
            If d.Contains("av1_qsv") Then Return "av1_qsv"
            Return "libx264"
        End Function

        ' Extend to support "Ultra Low" or "Performance"
        Private Shared Function GetQualityToken(s As RecorderSettings) As String
            Dim level = If(String.IsNullOrWhiteSpace(s.QualityLevel), "Medium", s.QualityLevel).Trim().ToLowerInvariant()
            If level.Contains("ultra") OrElse level.Contains("perf") Then Return "ultralow"
            If level.StartsWith("low") Then Return "low"
            If level.StartsWith("high") Then Return "high"
            Return "medium"
        End Function

        Private Function BuildArguments(s As RecorderSettings, output As String, ffmpegPath As String) As String
            EnsureDevicesProbed()

            Dim a As New List(Of String)
            a.Add("-hide_banner")
            a.Add("-y")
            a.Add("-fflags +genpts") ' stable PTS

            Dim inputCount As Integer = 0
            Dim vIndex As Integer = -1
            Dim sysAudIndex As Integer = -1
            Dim micIndex As Integer = -1

            Dim quality = GetQualityToken(s) ' "ultralow" | "low" | "medium" | "high"

            ' Optional crop for WindowTitle
            Dim cropForWindow As String = Nothing
            If s.CaptureSource = CaptureSource.WindowTitle Then
                Dim wr As WindowClientRect
                If Not TryGetWindowClientRect(s.WindowTitle, wr) Then
                    Throw New ArgumentException($"Selected window not found or not visible: '{s.WindowTitle}'.")
                End If
                Dim vs = GetVirtualScreen()
                Dim cropX = wr.Left - vs.Left
                Dim cropY = wr.Top - vs.Top
                If wr.Width <= 0 OrElse wr.Height <= 0 Then
                    Throw New ArgumentException("Target window has invalid size.")
                End If
                cropForWindow = $"crop={wr.Width}:{wr.Height}:{Math.Max(0, cropX)}:{Math.Max(0, cropY)}"
            End If

            ' VIDEO input (queue + timestamp from wall clock)
            If s.CaptureSource = CaptureSource.EntireDesktop OrElse s.CaptureSource = CaptureSource.Monitor OrElse s.CaptureSource = CaptureSource.Region Then
                Dim grab = GetBestDesktopGrab()
                vIndex = inputCount
                a.Add("-thread_queue_size 1024")
                a.Add($"-f {grab}")
                a.Add($"-framerate {s.FPS}")
                a.Add("-use_wallclock_as_timestamps 1")
                a.Add(If(s.IncludeCursor, "-draw_mouse 1", "-draw_mouse 0"))
                If s.CaptureSource = CaptureSource.Region Then
                    a.Add($"-offset_x {s.RegionX}")
                    a.Add($"-offset_y {s.RegionY}")
                    a.Add($"-video_size {s.RegionWidth}x{s.RegionHeight}")
                End If
                a.Add("-i desktop")
                inputCount += 1
            ElseIf s.CaptureSource = CaptureSource.WindowTitle Then
                Dim grab = GetBestDesktopGrab()
                vIndex = inputCount
                a.Add("-thread_queue_size 1024")
                a.Add($"-f {grab}")
                a.Add($"-framerate {s.FPS}")
                a.Add("-use_wallclock_as_timestamps 1")
                a.Add(If(s.IncludeCursor, "-draw_mouse 1", "-draw_mouse 0"))
                a.Add("-i desktop")
                inputCount += 1
            End If

            ' AUDIO inputs (system/mic with wall clock)
            Dim systemAudioUsed As Boolean = False
            Dim microphoneUsed As Boolean = False

            If s.RecordSystemAudio Then
                If _hasWasapi Then
                    sysAudIndex = inputCount
                    a.Add("-thread_queue_size 1024")
                    a.Add("-f wasapi")
                    a.Add("-use_wallclock_as_timestamps 1")
                    a.Add("-i default")
                    inputCount += 1
                    systemAudioUsed = True
                ElseIf _hasDshow Then
                    Dim sysDev = GetDefaultSystemAudioDshow()
                    If Not String.IsNullOrWhiteSpace(sysDev) Then
                        sysAudIndex = inputCount
                        a.Add("-thread_queue_size 1024")
                        a.Add("-f dshow")
                        a.Add("-use_wallclock_as_timestamps 1")
                        a.Add($"-i audio={Quote(sysDev)}")
                        inputCount += 1
                        systemAudioUsed = True
                        RaiseEvent Log("Using DirectShow system audio: " & sysDev)
                    Else
                        RaiseEvent Log("[WARN] No system-audio device found via DirectShow. System audio will be skipped.")
                    End If
                Else
                    RaiseEvent Log("[WARN] WASAPI/DShow not available. System audio will be skipped.")
                End If
            End If

            If s.RecordMicrophone Then
                Dim micName As String = s.MicrophoneDevice
                If String.IsNullOrWhiteSpace(micName) AndAlso _hasDshow Then
                    micName = GetDefaultMicrophoneDshow()
                    If Not String.IsNullOrWhiteSpace(micName) Then
                        RaiseEvent Log("Auto-selected microphone (DirectShow): " & micName)
                    End If
                End If
                If Not String.IsNullOrWhiteSpace(micName) Then
                    If _hasDshow Then
                        micIndex = inputCount
                        a.Add("-thread_queue_size 1024")
                        a.Add("-f dshow")
                        a.Add("-use_wallclock_as_timestamps 1")
                        a.Add($"-i audio={Quote(micName)}")
                        inputCount += 1
                        microphoneUsed = True
                    Else
                        RaiseEvent Log("[WARN] DirectShow not available; cannot capture microphone.")
                    End If
                Else
                    RaiseEvent Log("[INFO] Microphone not selected/found; mic will be skipped.")
                End If
            End If

            If vIndex < 0 Then
                Throw New InvalidOperationException("No video input configured. Check capture source settings.")
            End If

            ' VIDEO filters (crop/scale + strict frame pacing)
            Dim vfFilters As New List(Of String)
            If s.CaptureSource = CaptureSource.Monitor AndAlso s.SelectedMonitor IsNot Nothing Then
                Dim b = s.SelectedMonitor.Bounds
                Dim vs = GetVirtualScreen()
                Dim mx = b.X - vs.Left
                Dim my = b.Y - vs.Top
                vfFilters.Add($"crop={b.Width}:{b.Height}:{mx}:{my}")
            End If
            If s.CaptureSource = CaptureSource.WindowTitle AndAlso Not String.IsNullOrEmpty(cropForWindow) Then
                vfFilters.Add(cropForWindow)
            End If
            If quality = "ultralow" Then
                vfFilters.Add("scale=-2:720")
            End If
            ' Strict CFR pacing: generate exactly F frames per second and re-time them
            vfFilters.Add($"fps={s.FPS},setpts=N/({s.FPS}*TB)")

            ' AUDIO mapping
            Dim filterComplex As String = Nothing
            Dim mapArgs As New List(Of String)
            If systemAudioUsed AndAlso microphoneUsed Then
                filterComplex = $"[{sysAudIndex}:a][{micIndex}:a]amix=inputs=2:duration=longest:dropout_transition=2[aout]"
                mapArgs.Add($"-map {vIndex}:v")
                mapArgs.Add("-map [aout]")
            ElseIf systemAudioUsed Then
                mapArgs.Add($"-map {vIndex}:v")
                mapArgs.Add($"-map {sysAudIndex}:a")
            ElseIf microphoneUsed Then
                mapArgs.Add($"-map {vIndex}:v")
                mapArgs.Add($"-map {micIndex}:a")
            Else
                mapArgs.Add($"-map {vIndex}:v")
            End If

            If Not String.IsNullOrEmpty(filterComplex) Then a.Add("-filter_complex " & Quote(filterComplex))
            If vfFilters.Count > 0 Then a.Add("-vf " & Quote(String.Join(",", vfFilters)))

            ' Encoder selection (as you have)
            Dim codecId = ResolveCodecId(s.VideoEncoder)
            If Not HasEncoder(codecId) Then
                RaiseEvent Log($"[WARN] Encoder '{codecId}' not found. Falling back to libx264.")
                codecId = If(HasEncoder("libx264"), "libx264", codecId)
            End If

            Dim gop As Integer = Math.Max(1, s.FPS * 2)

            Select Case codecId.ToLowerInvariant()
                Case "libxvid"
                    Select Case quality
                        Case "ultralow" : a.Add("-c:v libxvid -qscale:v 6 -pix_fmt yuv420p")
                        Case "low" : a.Add("-c:v libxvid -qscale:v 5 -pix_fmt yuv420p")
                        Case "high" : a.Add("-c:v libxvid -qscale:v 2 -pix_fmt yuv420p")
                        Case Else : a.Add("-c:v libxvid -qscale:v 3 -pix_fmt yuv420p")
                    End Select
                    a.Add($"-g {gop}") ' VP8/9/AV1 ignore -bf; Xvid uses I/P

                Case "libx264"
                    Select Case quality
                        Case "ultralow" : a.Add("-c:v libx264 -preset ultrafast -tune zerolatency -crf 32 -pix_fmt yuv420p")
                        Case "low" : a.Add("-c:v libx264 -preset faster -crf 28 -pix_fmt yuv420p")
                        Case "high" : a.Add("-c:v libx264 -preset slow -crf 18 -pix_fmt yuv420p")
                        Case Else : a.Add("-c:v libx264 -preset veryfast -crf 23 -pix_fmt yuv420p")
                    End Select
                    a.Add($"-g {gop} -bf 0")

                Case "libx265"
                    Select Case quality
                        Case "ultralow" : a.Add("-c:v libx265 -preset ultrafast -crf 36 -pix_fmt yuv420p")
                        Case "low" : a.Add("-c:v libx265 -preset medium -crf 32 -pix_fmt yuv420p")
                        Case "high" : a.Add("-c:v libx265 -preset slow -crf 23 -pix_fmt yuv420p")
                        Case Else : a.Add("-c:v libx265 -preset medium -crf 28 -pix_fmt yuv420p")
                    End Select
                    a.Add($"-g {gop} -bf 0")

                Case "libvpx-vp9"
                    Select Case quality
                        Case "ultralow" : a.Add("-c:v libvpx-vp9 -b:v 0 -crf 42 -row-mt 1 -cpu-used 6 -pix_fmt yuv420p")
                        Case "low" : a.Add("-c:v libvpx-vp9 -b:v 0 -crf 38 -row-mt 1 -cpu-used 4 -pix_fmt yuv420p")
                        Case "high" : a.Add("-c:v libvpx-vp9 -b:v 0 -crf 28 -row-mt 1 -cpu-used 1 -pix_fmt yuv420p")
                        Case Else : a.Add("-c:v libvpx-vp9 -b:v 0 -crf 32 -row-mt 1 -cpu-used 2 -pix_fmt yuv420p")
                    End Select
                    a.Add($"-g {gop}") ' no -bf

                Case "libaom-av1"
                    Select Case quality
                        Case "ultralow" : a.Add("-c:v libaom-av1 -b:v 0 -crf 50 -cpu-used 10 -pix_fmt yuv420p")
                        Case "low" : a.Add("-c:v libaom-av1 -b:v 0 -crf 40 -cpu-used 8 -pix_fmt yuv420p")
                        Case "high" : a.Add("-c:v libaom-av1 -b:v 0 -crf 28 -cpu-used 4 -pix_fmt yuv420p")
                        Case Else : a.Add("-c:v libaom-av1 -b:v 0 -crf 35 -cpu-used 6 -pix_fmt yuv420p")
                    End Select
                    a.Add($"-g {gop}") ' no -bf

                Case "libsvtav1"
                    Select Case quality
                        Case "ultralow" : a.Add("-c:v libsvtav1 -crf 40 -preset 10 -pix_fmt yuv420p")
                        Case "low" : a.Add("-c:v libsvtav1 -crf 36 -preset 8 -pix_fmt yuv420p")
                        Case "high" : a.Add("-c:v libsvtav1 -crf 24 -preset 4 -pix_fmt yuv420p")
                        Case Else : a.Add("-c:v libsvtav1 -crf 30 -preset 6 -pix_fmt yuv420p")
                    End Select
                    a.Add($"-g {gop}") ' no -bf

                Case "libvvenc"
                    Select Case quality
                        Case "ultralow" : a.Add("-c:v libvvenc -crf 36 -preset faster -pix_fmt yuv420p")
                        Case "low" : a.Add("-c:v libvvenc -crf 34 -preset faster -pix_fmt yuv420p")
                        Case "high" : a.Add("-c:v libvvenc -crf 22 -preset slow -pix_fmt yuv420p")
                        Case Else : a.Add("-c:v libvvenc -crf 28 -preset medium -pix_fmt yuv420p")
                    End Select
                    a.Add($"-g {gop}") ' no -bf

                Case "h264_nvenc"
                    Select Case quality
                        Case "ultralow" : a.Add("-c:v h264_nvenc -preset fast -rc vbr -cq 32 -bf 0 -b:v 3M -maxrate 6M")
                        Case "low" : a.Add("-c:v h264_nvenc -preset fast -rc vbr -cq 28 -b:v 4M -maxrate 8M")
                        Case "high" : a.Add("-c:v h264_nvenc -preset slow -rc vbr -cq 18 -b:v 10M -maxrate 20M")
                        Case Else : a.Add("-c:v h264_nvenc -preset medium -rc vbr -cq 22 -b:v 6M -maxrate 12M")
                    End Select
                    a.Add($"-g {gop} -bf 0")

                Case "hevc_nvenc"
                    Select Case quality
                        Case "ultralow" : a.Add("-c:v hevc_nvenc -preset fast -rc vbr -cq 32 -bf 0 -b:v 4M -maxrate 8M")
                        Case "low" : a.Add("-c:v hevc_nvenc -preset fast -rc vbr -cq 28 -b:v 6M -maxrate 12M")
                        Case "high" : a.Add("-c:v hevc_nvenc -preset slow -rc vbr -cq 18 -b:v 16M -maxrate 32M")
                        Case Else : a.Add("-c:v hevc_nvenc -preset medium -rc vbr -cq 22 -b:v 10M -maxrate 20M")
                    End Select
                    a.Add($"-g {gop} -bf 0")

                Case "av1_nvenc"
                    Select Case quality
                        Case "ultralow" : a.Add("-c:v av1_nvenc -preset fast -rc vbr -cq 32 -bf 0 -b:v 8M -maxrate 16M")
                        Case "low" : a.Add("-c:v av1_nvenc -preset fast -rc vbr -cq 28 -b:v 8M -maxrate 16M")
                        Case "high" : a.Add("-c:v av1_nvenc -preset slow -rc vbr -cq 20 -b:v 22M -maxrate 44M")
                        Case Else : a.Add("-c:v av1_nvenc -preset medium -rc vbr -cq 24 -b:v 14M -maxrate 28M")
                    End Select
                    a.Add($"-g {gop} -bf 0")

                Case "h264_amf"
                    Select Case quality
                        Case "ultralow" : a.Add("-c:v h264_amf -quality speed -rc vbr -qvbr_quality 32")
                        Case "low" : a.Add("-c:v h264_amf -quality quality -rc vbr -qvbr_quality 28")
                        Case "high" : a.Add("-c:v h264_amf -quality quality -rc vbr -qvbr_quality 18")
                        Case Else : a.Add("-c:v h264_amf -quality quality -rc vbr -qvbr_quality 22")
                    End Select
                    a.Add($"-g {gop} -bf 0")

                Case "hevc_amf"
                    Select Case quality
                        Case "ultralow" : a.Add("-c:v hevc_amf -quality speed -rc vbr -qvbr_quality 34")
                        Case "low" : a.Add("-c:v hevc_amf -quality quality -rc vbr -qvbr_quality 30")
                        Case "high" : a.Add("-c:v hevc_amf -quality quality -rc vbr -qvbr_quality 20")
                        Case Else : a.Add("-c:v hevc_amf -quality quality -rc vbr -qvbr_quality 24")
                    End Select
                    a.Add($"-g {gop} -bf 0")

                Case "h264_qsv"
                    Select Case quality
                        Case "ultralow" : a.Add("-c:v h264_qsv -preset veryfast -global_quality 32 -look_ahead 0")
                        Case "low" : a.Add("-c:v h264_qsv -preset medium -global_quality 28 -look_ahead 1")
                        Case "high" : a.Add("-c:v h264_qsv -preset medium -global_quality 18 -look_ahead 1")
                        Case Else : a.Add("-c:v h264_qsv -preset medium -global_quality 22 -look_ahead 1")
                    End Select
                    a.Add($"-g {gop} -bf 0")

                Case "hevc_qsv"
                    Select Case quality
                        Case "ultralow" : a.Add("-c:v hevc_qsv -preset veryfast -global_quality 34 -look_ahead 0")
                        Case "low" : a.Add("-c:v hevc_qsv -preset medium -global_quality 30 -look_ahead 1")
                        Case "high" : a.Add("-c:v hevc_qsv -preset medium -global_quality 20 -look_ahead 1")
                        Case Else : a.Add("-c:v hevc_qsv -preset medium -global_quality 24 -look_ahead 1")
                    End Select
                    a.Add($"-g {gop} -bf 0")

                Case "av1_qsv"
                    Select Case quality
                        Case "ultralow" : a.Add("-c:v av1_qsv -preset 10 -global_quality 36")
                        Case "low" : a.Add("-c:v av1_qsv -preset 8 -global_quality 32")
                        Case "high" : a.Add("-c:v av1_qsv -preset 4 -global_quality 24")
                        Case Else : a.Add("-c:v av1_qsv -preset 6 -global_quality 28")
                    End Select
                    a.Add($"-g {gop}") ' leave BF as default for AV1 QSV
                Case Else
                    a.Add("-c:v libx264 -preset veryfast -crf 23 -pix_fmt yuv420p")
                    a.Add($"-g {gop} -bf 0")
            End Select

            ' Audio encode: standard 48 kHz + resample async to smooth A/V drift
            If (systemAudioUsed OrElse microphoneUsed) Then
                a.Add("-c:a aac -b:a 160k -ac 2 -ar 48000")
                a.Add("-af " & Quote("aresample=async=1000:first_pts=0"))
            End If

            ' Force CFR at muxer and set output frame rate
            a.Add("-vsync cfr")
            a.Add($"-r {s.FPS}")

            a.Add("-movflags +faststart")
            a.Add("-progress pipe:1")
            a.Add("-nostats")

            a.AddRange(mapArgs)
            a.Add(Quote(output))
            Return String.Join(" ", a)
        End Function

        Private Shared Function Quote(s As String) As String
            If s.Contains(" "c) OrElse s.Contains("""") Then
                Return """" & s.Replace("""", "\""") & """"
            End If
            Return s
        End Function

        ' Visible top-level window exists?
        Private Shared Function WindowExists(title As String) As Boolean
            Dim wr As WindowClientRect
            Return TryGetWindowClientRect(title, wr)
        End Function

        ' --- Win32 helpers for window/client rect and virtual screen ---
        <StructLayout(LayoutKind.Sequential)>
        Private Structure RECT
            Public Left As Integer
            Public Top As Integer
            Public Right As Integer
            Public Bottom As Integer
        End Structure

        Private Structure WindowClientRect
            Public Left As Integer
            Public Top As Integer
            Public Width As Integer
            Public Height As Integer
        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Private Structure POINT
            Public X As Integer
            Public Y As Integer
        End Structure

        Private Shared Function TryGetWindowClientRect(title As String, ByRef res As WindowClientRect) As Boolean
            If String.IsNullOrWhiteSpace(title) Then Return False
            Dim found As Boolean = False
            Dim hFound As IntPtr = IntPtr.Zero

            EnumWindows(Function(hwnd, lParam)
                            If Not IsWindowVisible(hwnd) Then Return True
                            Dim len = GetWindowTextLength(hwnd)
                            If len <= 0 Then Return True
                            Dim sb As New System.Text.StringBuilder(len + 1)
                            GetWindowText(hwnd, sb, sb.Capacity)
                            Dim t = sb.ToString().Trim()
                            If String.Equals(t, title, StringComparison.Ordinal) Then
                                hFound = hwnd
                                found = True
                                Return False
                            End If
                            Return True
                        End Function, IntPtr.Zero)

            If Not found OrElse hFound = IntPtr.Zero Then Return False
            If IsIconic(hFound) Then Return False

            Dim cr As RECT
            If Not GetClientRect(hFound, cr) Then Return False

            Dim pt As New POINT With {.X = 0, .Y = 0}
            If Not ClientToScreen(hFound, pt) Then Return False

            Dim widthPx = Math.Max(0, cr.Right - cr.Left)
            Dim heightPx = Math.Max(0, cr.Bottom - cr.Top)
            If widthPx = 0 OrElse heightPx = 0 Then Return False

            res = New WindowClientRect With {.Left = pt.X, .Top = pt.Y, .Width = widthPx, .Height = heightPx}
            Return True
        End Function

        Private Structure VirtualScreenRect
            Public Left As Integer
            Public Top As Integer
            Public Width As Integer
            Public Height As Integer
        End Structure

        Private Shared Function GetVirtualScreen() As VirtualScreenRect
            Const SM_XVIRTUALSCREEN As Integer = 76
            Const SM_YVIRTUALSCREEN As Integer = 77
            Const SM_CXVIRTUALSCREEN As Integer = 78
            Const SM_CYVIRTUALSCREEN As Integer = 79
            Dim left = GetSystemMetrics(SM_XVIRTUALSCREEN)
            Dim top = GetSystemMetrics(SM_YVIRTUALSCREEN)
            Dim w = GetSystemMetrics(SM_CXVIRTUALSCREEN)
            Dim h = GetSystemMetrics(SM_CYVIRTUALSCREEN)
            Return New VirtualScreenRect With {.Left = left, .Top = top, .Width = w, .Height = h}
        End Function

        Private Delegate Function EnumWindowsProc(hWnd As IntPtr, lParam As IntPtr) As Boolean
        <DllImport("user32.dll")>
        Private Shared Function EnumWindows(lpEnumFunc As EnumWindowsProc, lParam As IntPtr) As Boolean
        End Function
        <DllImport("user32.dll")>
        Private Shared Function IsWindowVisible(hWnd As IntPtr) As Boolean
        End Function
        <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
        Private Shared Function GetWindowText(hWnd As IntPtr, lpString As System.Text.StringBuilder, nMaxCount As Integer) As Integer
        End Function
        <DllImport("user32.dll")>
        Private Shared Function GetWindowTextLength(hWnd As IntPtr) As Integer
        End Function
        <DllImport("user32.dll")>
        Private Shared Function GetClientRect(hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
        End Function
        <DllImport("user32.dll")>
        Private Shared Function ClientToScreen(hWnd As IntPtr, ByRef lpPoint As POINT) As Boolean
        End Function
        <DllImport("user32.dll")>
        Private Shared Function IsIconic(hWnd As IntPtr) As Boolean
        End Function
        <DllImport("user32.dll")>
        Private Shared Function GetSystemMetrics(nIndex As Integer) As Integer
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            _proc?.Dispose()
        End Sub
    End Class
End Namespace