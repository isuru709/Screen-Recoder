Namespace Models
  Public Enum CaptureSource
    EntireDesktop
    Monitor
    Region
    WindowTitle
  End Enum

  Public Class MonitorInfo
    Public Property Index As Integer
        Public Property Bounds As System.Drawing.Rectangle
        Public ReadOnly Property DisplayName As String
            Get
                Return $"Monitor {Index + 1} ({Bounds.Width}x{Bounds.Height} @ {Bounds.X},{Bounds.Y})"
            End Get
        End Property
    End Class

    Public Class RecorderSettings
        Public Property OutputFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.MyVideos)
        Public Property OutputFileNamePattern As String = "Recording_{yyyyMMdd_HHmmss}.mp4"

        Public Property CaptureSource As CaptureSource = CaptureSource.EntireDesktop
        Public Property IncludeCursor As Boolean = True
        Public Property FPS As Integer = 60

        Public Property SelectedMonitor As MonitorInfo
        Public Property RegionX As Integer = 0
        Public Property RegionY As Integer = 0
        Public Property RegionWidth As Integer = 1280
        Public Property RegionHeight As Integer = 720
        Public Property WindowTitle As String = ""

        ' Audio
        Public Property RecordSystemAudio As Boolean = True
        Public Property RecordMicrophone As Boolean = False
        Public Property MicrophoneDevice As String = ""

        ' Kept for backward compatibility but no longer used for routing
        Public Property MixMicAndSystem As Boolean = True

        ' Video encoder
        Public Property VideoEncoder As String = "Software (libx264)"
        Public Property QualityLevel As String = "Medium"

        ' Legacy (unused now)
        Public Property QualityPreset As String = ""

        Public Function ResolveOutputPath() As String
            Dim fileName = OutputFileNamePattern
            fileName = fileName.Replace("{yyyyMMdd_HHmmss}", DateTime.Now.ToString("yyyyMMdd_HHmmss"))
            Return IO.Path.Combine(OutputFolder, fileName)
        End Function
    End Class
End Namespace