Imports System.Windows.Input

Namespace ViewModels
  Public Class RelayCommand
    Implements ICommand

    Private ReadOnly _execute As Action
    Private ReadOnly _canExecute As Func(Of Boolean)

    Public Sub New(execute As Action, Optional canExecute As Func(Of Boolean) = Nothing)
            If execute Is Nothing Then Throw New ArgumentNullException(NameOf(execute))
            _execute = execute
            _canExecute = canExecute
        End Sub

    Public Custom Event CanExecuteChanged As EventHandler Implements ICommand.CanExecuteChanged
      AddHandler(value As EventHandler)
                AddHandler CommandManager.RequerySuggested, value
            End AddHandler
            RemoveHandler(value As EventHandler)
                RemoveHandler CommandManager.RequerySuggested, value
            End RemoveHandler
            RaiseEvent(sender As Object, e As EventArgs)
                CommandManager.InvalidateRequerySuggested()
            End RaiseEvent
        End Event

        Public Function CanExecute(parameter As Object) As Boolean Implements ICommand.CanExecute
      If _canExecute Is Nothing Then Return True
      Return _canExecute()
    End Function

    Public Sub Execute(parameter As Object) Implements ICommand.Execute
      _execute()
    End Sub
  End Class
End Namespace