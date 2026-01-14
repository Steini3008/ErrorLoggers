Attribute VB_Name = "Testing_FileErrorLogger"
Option Explicit


Private Property Get pathStandardErrorLoggingFile() As String
    
    With Singleton.Fso
        
        pathStandardErrorLoggingFile = .BuildPath(.GetFolder(ThisWorkbook.Path).SubFolders("out"), "ErrorLoggingFile.txt")
        
    End With
    
End Property


Private Property Get standardErrorFileLogger() As ErrorLoggers.FileErrorLogger
    
    Set standardErrorFileLogger = ErrorLoggers.Factory.CreateFileErrorLoggerWith(pathStandardErrorLoggingFile)
    
End Property


Public Function NoIssuesWithErrorFileLogger() As Boolean
' @LogErrorInformation
On Error GoTo ErrorHandling
    
    Debug.Print 5 / 0
    
    Exit Function
    
ErrorHandling:

    standardErrorFileLogger.LogErrorInformation Err
    
    NoIssuesWithErrorFileLogger = Singleton.Fso.FileExists(pathStandardErrorLoggingFile)
    
End Function

