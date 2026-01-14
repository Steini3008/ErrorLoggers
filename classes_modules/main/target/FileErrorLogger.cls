VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileErrorLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


' ----------------------------------------------------------------------------------------------------------------------------------------------------------------
' Class "FileErrorLogger" is an implementation of "Error_Handling.ErrorLogger" for logging information about errors directly into a provided file
'
' Dependencies:
' (1) Global Libraries
'   - "Microsoft Scripting Runtime"
'
' (2) Private Libraries
'   None
'
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------


' Implementations

Implements Error_Handling.ErrorLogger


' Fields

Private pFullPathLoggingFile As String

Private pLoggingFileStream As Scripting.TextStream


' Constructor

' Getters and Setters

Public Property Let FullPathLoggingFile(nFullPathLoggingFile As String)
    
    pFullPathLoggingFile = nFullPathLoggingFile
    
End Property


' Public API

Public Sub LogErrorInformation(errObj As ErrObject)
    
    Set pLoggingFileStream = Singleton.Fso.CreateTextFile(pFullPathLoggingFile, True)
    
    writeErrorInformation errObj
    
    pLoggingFileStream.Close
    
End Sub


Private Sub ErrorLogger_LogErrorInformation(errObj As ErrObject)
    
    LogErrorInformation errObj
    
End Sub


' Helpers

Private Sub writeErrorInformation(errObj As ErrObject)

    With pLoggingFileStream
    
        .WriteLine "Error Code: " & errObj.Number
        
        .WriteLine "Error Reason: " & errObj.Source
        
        .WriteLine "Error Description: " & errObj.Description
        
    End With
    
End Sub

