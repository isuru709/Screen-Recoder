Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.CompilerServices
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

            _qualityLevels = New ObservableCollection(Of String) From {"Low", "Medium", "High"}

            If String.IsNullOrWhiteSpace(Settings.VideoEncoder) Then
                Settings.VideoEncoder = "Software (libx264)"
            End If
            If String.IsNullOrWhiteSpace(Settings.QualityLevel) Then
                Settings.QualityLevel = "Medium"
            End If

            _monitors = New ObservableCollection(Of Models.MonitorInfo)()
            LoadMonitors()

            AddHandler _rec.Log, Sub(line) AppendLog(line)
            AddHandler _rec.StatusChanged, Sub(s) Status = s
            AddHandler _rec.RecordingStopped, Sub(code) IsRecording = False
            AddHandler _rec.Progress, Sub(t, fps, spd) Status = $"Recording… {t:hh\:mm\:ss} | {fps:0.0} fps | {spd:0.00}x"
        End Sub

        Private Sub LoadMonitors()
            Monitors.Clear()
            Dim idx = 0
            For Each sc In System.Windows.Forms.Screen.AllScreens
                Monitors.Add(New Models.MonitorInfo With {
                    .Index = idx,
                    .Bounds = sc.Bounds
                })
                idx += 1
            Next
            If Monitors.Count > 0 Then
                SelectedMonitor = Monitors(0)
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

        Private _startCommand As RelayCommand
        Public ReadOnly Property StartCommand As RelayCommand
            Get
                If _startCommand Is Nothing Then
                    _startCommand = New RelayCommand(
                        Sub()
                            Try
                                Dim path = _rec.StartRecording(Settings)
                                IsRecording = True
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

        Private _stopCommand As RelayCommand
        Public ReadOnly Property StopCommand As RelayCommand
            Get
                If _stopCommand Is Nothing Then
                    _stopCommand = New RelayCommand(
                        Sub() _rec.StopRecording(),
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
    End Class
End Namespace