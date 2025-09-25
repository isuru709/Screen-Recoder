Imports System.ComponentModel

Namespace Screen_Recorder
    Partial Public Class MainWindow
        Inherits System.Windows.Window

        Public Sub New()
            InitializeComponent()
            DataContext = New ViewModels.MainViewModel()
            AddHandler Me.Closing, AddressOf OnClosingHandler
        End Sub

        Private Sub OnClosingHandler(sender As Object, e As CancelEventArgs)
            Dim vm = TryCast(DataContext, ViewModels.MainViewModel)
            vm?.Dispose()
        End Sub
    End Class
End Namespace