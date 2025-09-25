Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Windows.Interop

Namespace Screen_Recorder
    Partial Public Class MainWindow
        Inherits System.Windows.Window

        Public Sub New()
            InitializeComponent()
            DataContext = New ViewModels.MainViewModel()
            AddHandler Me.Closing, AddressOf OnClosingHandler
            AddHandler Me.SourceInitialized, AddressOf OnSourceInitializedHandler
        End Sub

        Private Sub OnClosingHandler(sender As Object, e As CancelEventArgs)
            Dim vm = TryCast(DataContext, ViewModels.MainViewModel)
            vm?.Dispose()
        End Sub

        ' Disable border drag-resize, but keep Maximize enabled
        Private Sub OnSourceInitializedHandler(sender As Object, e As EventArgs)
            Dim hwnd = New WindowInteropHelper(Me).Handle
            If hwnd = IntPtr.Zero Then Return
            Const GWL_STYLE As Integer = -16
            Const WS_THICKFRAME As Integer = &H40000 ' resizable frame

            Dim style = GetWindowLong(hwnd, GWL_STYLE)
            style = style And Not WS_THICKFRAME
            SetWindowLong(hwnd, GWL_STYLE, style)
        End Sub

        <DllImport("user32.dll")>
        Private Shared Function GetWindowLong(hWnd As IntPtr, nIndex As Integer) As Integer
        End Function

        <DllImport("user32.dll")>
        Private Shared Function SetWindowLong(hWnd As IntPtr, nIndex As Integer, dwNewLong As Integer) As Integer
        End Function
    End Class
End Namespace