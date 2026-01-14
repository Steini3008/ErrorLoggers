Attribute VB_Name = "Client"
Option Explicit


Public Sub Main()
On Error GoTo ErrorHandling

    Dim errLogger As Error_Handling.ErrorLogger
    
    Set errLogger = Factory.CreateConsoleErrorLogger
    
    Debug.Print 5 / 0
    
    Exit Sub
    
ErrorHandling:
    
    errLogger.LogErrorInformation Err
    
End Sub
