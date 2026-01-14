Attribute VB_Name = "Testing"
Option Explicit


Public Sub Main()

    Dim unitTester As VBA_Unit_Testing.VBAUnitTesting
    
    Set unitTester = VBA_Unit_Testing.Factory.CreateVBAUnitTesting
    
    With unitTester
    
        Set .VBProjectToTest = Workbooks("ErrorLoggers.xlam").VBProject
        
        Set .VBProjectWithTests = ThisWorkbook.VBProject
            
        .AddTestComponents "ConsoleErrorLogger", "FileErrorLogger"
        
        .TestAll
        
        .PrintAllTestResults
        
    End With
        
End Sub
