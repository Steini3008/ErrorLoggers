Attribute VB_Name = "Factory"
Option Explicit


' ----------------------------------------------------------------------------------------------------------------------------------------------------------------
' Module "Factory" exposes the creation of several types of this VBA-Project to other VBA-Projects
'
' Dependencies:
' (1) Global Libraries
'   None
'
' (2) Private Libraries
'   None
'
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------


' Public API

Public Function CreateConsoleErrorLogger() As ConsoleErrorLogger

    Set CreateConsoleErrorLogger = New ConsoleErrorLogger
    
End Function


Public Function CreateFileErrorLogger() As FileErrorLogger

    Set CreateFileErrorLogger = New FileErrorLogger
    
End Function


Public Function CreateFileErrorLoggerWith(fullPathErrorLoggingFile As String) As FileErrorLogger

    Set CreateFileErrorLoggerWith = New FileErrorLogger
    
    With CreateFileErrorLoggerWith
    
        .FullPathLoggingFile = fullPathErrorLoggingFile
        
    End With
    
End Function
