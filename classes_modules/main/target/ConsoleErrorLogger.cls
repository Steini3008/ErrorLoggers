VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConsoleErrorLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


' ----------------------------------------------------------------------------------------------------------------------------------------------------------------
' Class "ConsoleErrorLogger" is an implementation of "Error_Handling.ErrorLogger" for logging information about errors directly into the console
'
' Dependencies:
' (1) Global Libraries
'   None
'
' (2) Private Libraries
'   None
'
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------


' Implementations

Implements Error_Handling.ErrorLogger


' Fields

' Constructor

' Getters and Setters

' Public API

Public Sub LogErrorInformation(errObj As ErrObject)
    
    With errObj
        Debug.Print .Number, .Source, .Description
    End With
    
End Sub


Private Sub ErrorLogger_LogErrorInformation(errObj As ErrObject)
    
    LogErrorInformation errObj
    
End Sub
