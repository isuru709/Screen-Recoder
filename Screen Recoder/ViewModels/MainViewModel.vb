Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports ScreenRecoder.Models
Imports ScreenRecoder.Services

Namespace ViewModels
    Public Class MainViewModel
        Implements INotifyPropertyChanged, IDisposable

        Public Property Settings As New RecorderSettings()

        Private ReadOnly _captureSources As ObservableCollection(Of Models.CaptureSource)
        Public ReadOnly Property CaptureSources As ObservableCollection(Of Models.CaptureSource)
            Get
                Return _captureSources
            End Get
        End Property

        Private ReadOnly _monitors As ObservableCollection(Of Models.MonitorInfo)
        Public ReadOnly Property Monitors As ObservableCollection(Of Models.MonitorInfo)
            Get
                Return _monitors
            End Get
        End Property

        ' Bind the Source combo to this so we can raise visibility updates
        Public Property SelectedCaptureSource As CaptureSource
            Get
                Return Settings.CaptureSource
            End Get
            Set(value As CaptureSource)
                If Settings.CaptureSource <> value Then
                    Settings.CaptureSource = value
                    OnPropertyChanged()
                    OnPropertyChanged(NameOf(IsWindowTitleSource))
                    OnPropertyChanged(NameOf(IsRegionSource))
                    If value = CaptureSource.WindowTitle Then
                        LoadWindows()
                    End If
                End If
            End Set
        End Property

        Public ReadOnly Property IsWindowTitleSource As Boolean
            Get
                Return SelectedCaptureSource = CaptureSource.WindowTitle
            End Get
        End Property

        Public ReadOnly Property IsRegionSource As Boolean
            Get
                Return SelectedCaptureSource = CaptureSource.Region
            End Get
        End Property

        Public Property SelectedMonitor As Models.MonitorInfo
            Get
                Return Settings.SelectedMonitor
            End Get
            Set(value As Models.MonitorInfo)
                Settings.SelectedMonitor = value
                OnPropertyChanged()
            End Set
        End Property

        Private ReadOnly _encoders As ObservableCollection(Of String)
        Public ReadOnly Property Encoders As ObservableCollection(Of String)
            Get
                Return _encoders
            End Get
        End Property

        Private ReadOnly _qualityLevels As ObservableCollection(Of String)
        Public ReadOnly Property QualityLevels As ObservableCollection(Of String)
            Get
                Return _qualityLevels
            End Get
        End Property

        ' Running windows for WindowTitle capture
        Private ReadOnly _runningWindows As New ObservableCollection(Of String)()
        Public ReadOnly Property RunningWindows As ObservableCollection(Of String)
            Get
                Return _runningWindows
            End Get
        End Property

        Private ReadOnly _rec As New RecorderService()

        Private _status As String = "Idle."
        Public Property Status As String
            Get
                Return _status
            End Get
            Set(value As String)
                _status = value
                OnPropertyChanged()
            End Set
        End Property

        Private _log As String = ""
        Public Property Log As String
            Get
                Return _log
            End Get
            Set(value As String)
                _log = value
                OnPropertyChanged()
            End Set
        End Property

        Public ReadOnly Property CanStart As Boolean
            Get
                Return Not IsRecording
            End Get
        End Property

        Private _isRecording As Boolean
        Public Property IsRecording As Boolean
            Get
                Return _isRecording
            End Get
            Private Set(value As Boolean)
                _isRecording = value
                OnPropertyChanged()
                OnPropertyChanged(NameOf(CanStart))
            End Set
        End Property

        Private _isPaused As Boolean
        Public Property IsPaused As Boolean
            Get
                Return _isPaused
            End Get
            Private Set(value As Boolean)
                _isPaused = value
                OnPropertyChanged()
            End Set
        End Property

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Public Sub New()
            _captureSources = New ObservableCollection(Of CaptureSource) From {
                CaptureSource.EntireDesktop, CaptureSource.Monitor, CaptureSource.Region, CaptureSource.WindowTitle
            }

            _encoders = New ObservableCollection(Of String) From {
                "Software (Xvid/MPEG)",
                "Software (libx264)",
                "Software (libx265)",
                "Software (libvpx-vp9)",
                "Software (libaom-av1)",
                "Software (libsvt-av1)",
                "Software (VVC/vvenc)",
                "NVIDIA H.264 (h264_nvenc)",
                "NVIDIA HEVC (hevc_nvenc)",
                "NVIDIA AV1 (av1_nvenc)",
                "AMD H.264 (h264_amf)",
                "AMD HEVC (hevc_amf)",
                "AMD AV1 (av1_amf)",
                "Intel H.264 (h264_qsv)",
                "Intel HEVC (hevc_qsv)",
                "Intel AV1 (av1_qsv)"
            }

            _qualityLevels = New ObservableCollection(Of String) From {"Ultra Low", "Low", "Medium", "High"}

            If String.IsNullOrWhiteSpace(Settings.VideoEncoder) Then
                Settings.VideoEncoder = "Software (libx264)"
            End If
            If String.IsNullOrWhiteSpace(Settings.QualityLevel) Then
                Settings.QualityLevel = "Medium"
            End If

            _monitors = New ObservableCollection(Of Models.MonitorInfo)()
            LoadMonitors()
            LoadWindows()

            AddHandler _rec.Log, Sub(line) AppendLog(line)
            AddHandler _rec.StatusChanged, Sub(s) Status = s
            AddHandler _rec.RecordingStopped, Sub(code)
                                                  IsRecording = False
                                                  IsPaused = False
                                              End Sub
            AddHandler _rec.Progress, Sub(t, fps, spd) Status = $"Recording… {t:hh\:mm\:ss} | {fps:0.0} fps | {spd:0.00}x"
        End Sub

        Private Sub LoadMonitors()
            Monitors.Clear()
            Dim idx = 0
            For Each sc In System.Windows.Forms.Screen.AllScreens
                Monitors.Add(New Models.MonitorInfo With {.Index = idx, .Bounds = sc.Bounds})
                idx += 1
            Next
            If Monitors.Count > 0 Then
                SelectedMonitor = Monitors(0)
            End If
        End Sub

        Public Sub LoadWindows()
            RunningWindows.Clear()
            For Each title In EnumerateTopLevelWindowTitles()
                RunningWindows.Add(title)
            Next
            If String.IsNullOrWhiteSpace(Settings.WindowTitle) AndAlso RunningWindows.Count > 0 Then
                Settings.WindowTitle = RunningWindows(0)
                OnPropertyChanged(NameOf(Settings))
            End If
        End Sub

        Private Sub AppendLog(line As String)
            Log &= line & Environment.NewLine
        End Sub

        Private _browseOutputCommand As RelayCommand
        Public ReadOnly Property BrowseOutputCommand As RelayCommand
            Get
                If _browseOutputCommand Is Nothing Then
                    _browseOutputCommand = New RelayCommand(
                        Sub()
                            Dim dlg As New System.Windows.Forms.FolderBrowserDialog()
                            If dlg.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                                Settings.OutputFolder = dlg.SelectedPath
                                OnPropertyChanged(NameOf(Settings))
                            End If
                        End Sub)
                End If
                Return _browseOutputCommand
            End Get
        End Property

        Private _refreshWindowsCommand As RelayCommand
        Public ReadOnly Property RefreshWindowsCommand As RelayCommand
            Get
                If _refreshWindowsCommand Is Nothing Then
                    _refreshWindowsCommand = New RelayCommand(Sub() LoadWindows())
                End If
                Return _refreshWindowsCommand
            End Get
        End Property

        Private _startCommand As RelayCommand
        Public ReadOnly Property StartCommand As RelayCommand
            Get
                If _startCommand Is Nothing Then
                    _startCommand = New RelayCommand(
                        Sub()
                            Try
                                Dim path = _rec.StartRecording(Settings)
                                IsRecording = True
                                IsPaused = False
                                AppendLog("Recording file: " & path)
                            Catch ex As Exception
                                AppendLog("Start failed: " & ex.Message)
                                Status = "Failed to start."
                            End Try
                        End Sub,
                        Function() CanStart)
                End If
                Return _startCommand
            End Get
        End Property

        Private _pauseCommand As RelayCommand
        Public ReadOnly Property PauseCommand As RelayCommand
            Get
                If _pauseCommand Is Nothing Then
                    _pauseCommand = New RelayCommand(
                        Sub()
                            Try
                                _rec.PauseRecording()
                                IsPaused = True
                                Status = "Paused."
                            Catch ex As Exception
                                AppendLog("Pause failed: " & ex.Message)
                            End Try
                        End Sub,
                        Function() IsRecording AndAlso Not IsPaused)
                End If
                Return _pauseCommand
            End Get
        End Property

        Private _resumeCommand As RelayCommand
        Public ReadOnly Property ResumeCommand As RelayCommand
            Get
                If _resumeCommand Is Nothing Then
                    _resumeCommand = New RelayCommand(
                        Sub()
                            Try
                                _rec.ResumeRecording()
                                IsPaused = False
                                Status = "Recording…"
                            Catch ex As Exception
                                AppendLog("Resume failed: " & ex.Message)
                            End Try
                        End Sub,
                        Function() IsPaused)
                End If
                Return _resumeCommand
            End Get
        End Property

        Private _stopCommand As RelayCommand
        Public ReadOnly Property StopCommand As RelayCommand
            Get
                If _stopCommand Is Nothing Then
                    _stopCommand = New RelayCommand(
                        Sub()
                            _rec.StopRecording()
                            IsPaused = False
                        End Sub,
                        Function() IsRecording)
                End If
                Return _stopCommand
            End Get
        End Property

        Protected Overridable Sub OnPropertyChanged(<CallerMemberName> Optional name As String = Nothing)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(name))
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            _rec?.Dispose()
        End Sub

        ' --- Window enumeration (User32) ---
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

        Private Shared Function EnumerateTopLevelWindowTitles() As IEnumerable(Of String)
            Dim results As New List(Of String)
            EnumWindows(Function(hwnd, lParam)
                            If Not IsWindowVisible(hwnd) Then Return True
                            Dim len = GetWindowTextLength(hwnd)
                            If len <= 0 Then Return True
                            Dim sb As New System.Text.StringBuilder(len + 1)
                            GetWindowText(hwnd, sb, sb.Capacity)
                            Dim title = sb.ToString().Trim()
                            If Not String.IsNullOrWhiteSpace(title) Then
                                results.Add(title)
                            End If
                            Return True
                        End Function, IntPtr.Zero)
            Return results.Distinct().OrderBy(Function(t) t)
        End Function
    End Class
End Namespace