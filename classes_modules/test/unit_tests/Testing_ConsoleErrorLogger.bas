Attribute VB_Name = "Testing_ConsoleErrorLogger"
Option Explicit


Private Property Get standardErrorConsoleLogger() As ErrorLoggers.ConsoleErrorLogger

    Set standardErrorConsoleLogger = ErrorLoggers.Factory.CreateConsoleErrorLogger()
    
End Property


Public Function NoIssuesWithErrorConsoleLogger() As Boolean
' @LogErrorInformation
On Error GoTo ErrorHandling
    
    Debug.Print 5 / 0
    
    Exit Function
    
ErrorHandling:

    standardErrorConsoleLogger.LogErrorInformation Err
    
    NoIssuesWithErrorConsoleLogger = True
    
End Function
