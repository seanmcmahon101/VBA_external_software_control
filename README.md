# VBA External Program Automation Guide

## Overview

This guide demonstrates how to automate external programs using VBA, specifically using a Fourth Shift journal entry automation as a practical example. The techniques shown can be adapted for automating any Windows application that accepts keyboard input.

## Table of Contents

- [Core Concepts](#core-concepts)
- [Window Management](#window-management)
- [Keyboard Automation](#keyboard-automation)
- [Data Integration](#data-integration)
- [Error Handling & Safety](#error-handling--safety)
- [Timing & Synchronization](#timing--synchronization)
- [Best Practices](#best-practices)
- [Complete Example](#complete-example)

## Core Concepts

### What is External Program Automation?

External program automation involves controlling another application from VBA by:
- Activating windows
- Sending keystrokes
- Reading data from Excel
- Managing timing and synchronization

### Key VBA Functions Used

```vba
AppActivate()           ' Activate external windows
Application.SendKeys    ' Send keyboard input
Sleep()                 ' Add delays for timing
DoEvents               ' Process pending events
```

## Window Management

### Activating External Applications

```vba
Sub ActivateExternalProgram()
    ' Basic window activation
    AppActivate ("FSUK - Fourth Shift - System Control")
  
    ' With error handling
    On Error GoTo WindowError
    AppActivate ("Program Window Title")
    On Error GoTo 0
    Exit Sub
  
WindowError:
    MsgBox "Could not find the program window. Please ensure it's running.", vbCritical
End Sub
```

### Finding Window Titles

To find the exact window title:
1. Open the target application
2. Use Alt+Tab or Task Manager to see the full window title
3. Use partial matches if the title changes (e.g., includes document names)

```vba
' Examples of window titles
AppActivate ("Microsoft Excel")           ' Excel
AppActivate ("FSUK - Fourth Shift")      ' Fourth Shift (partial match)
AppActivate ("Notepad")                   ' Notepad
```

## Keyboard Automation

### SendKeys Syntax

```vba
' Basic text entry
Application.SendKeys "Hello World"

' Special keys
Application.SendKeys "~"           ' Enter key
Application.SendKeys "{TAB}"       ' Tab key
Application.SendKeys "{DELETE}"    ' Delete key
Application.SendKeys "^{HOME}"     ' Ctrl+Home
Application.SendKeys "+{TAB}"      ' Shift+Tab (reverse tab)

' Function keys
Application.SendKeys "{F1}"        ' F1 key
Application.SendKeys "{F6}"        ' F6 key

' Multiple keys
Application.SendKeys "{DOWN 3}"    ' Press Down arrow 3 times
```

### Navigation Patterns

```vba
Sub NavigateToField()
    ' Navigate to search box and enter module name
    Application.SendKeys "^{HOME}"     ' Go to top
    Sleep 500
    Application.SendKeys "{TAB}"       ' Move to search field
    Sleep 500
    Application.SendKeys "GLJE"        ' Type module name
    Sleep 3000
    Application.SendKeys "~"           ' Press Enter
End Sub
```

### Field-by-Field Data Entry

```vba
Sub EnterJournalLine(glCode As String, description As String, amount As Double)
    ' Enter GL Code
    Application.SendKeys glCode
    Sleep 500
  
    ' Move to next field
    Application.SendKeys "{TAB}"
    Sleep 500
  
    ' Enter Description
    Application.SendKeys description
    Sleep 500
  
    ' Move to amount field
    Application.SendKeys "{TAB}"
    Sleep 500
  
    ' Enter formatted amount
    Application.SendKeys Format(amount, "0.00")
    Sleep 500
  
    ' Confirm entry
    Application.SendKeys "~"
End Sub
```

## Data Integration

### Reading from Excel Cells

```vba
Sub ReadExcelData()
    Dim ws As Worksheet
    Dim batchNumber As String
    Dim lastRow As Long
  
    ' Get active worksheet
    Set ws = ActiveSheet
  
    ' Read single cell
    batchNumber = CStr(ws.Range("J10").Value)
  
    ' Find last row with data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
  
    ' Loop through data rows
    For i = 2 To lastRow
        Dim glCode As String
        Dim description As String
        Dim debitAmount As Double
      
        glCode = Trim(CStr(ws.Cells(i, 1).Value))        ' Column A
        description = Trim(CStr(ws.Cells(i, 2).Value))   ' Column B
        debitAmount = CDbl(ws.Cells(i, 3).Value)         ' Column C
      
        ' Process the data
        Call ProcessDataRow(glCode, description, debitAmount)
    Next i
End Sub
```

### Handling Data Types

```vba
Sub SafeDataReading()
    Dim ws As Worksheet
    Set ws = ActiveSheet
  
    ' Safe string conversion
    Dim textValue As String
    textValue = Trim(CStr(ws.Cells(1, 1).Value))
  
    ' Safe numeric conversion with error handling
    Dim numericValue As Double
    On Error Resume Next
    numericValue = CDbl(ws.Cells(1, 2).Value)
    If Err.Number <> 0 Then
        numericValue = 0  ' Default value if conversion fails
        Debug.Print "Warning: Could not convert cell value to number"
    End If
    On Error GoTo 0
End Sub
```

## Error Handling & Safety

### Kill Switch Implementation

```vba
Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public killSwitchActivated As Boolean

Sub CheckKillSwitch()
    ' Check for ESC key (VK_ESCAPE = 27)
    If GetAsyncKeyState(27) <> 0 Then
        killSwitchActivated = True
        MsgBox "*** ESC KEY DETECTED! Process terminated! ***", vbCritical
        Exit Sub
    End If
    DoEvents  ' Process pending events
End Sub

Sub SafeAutomationLoop()
    killSwitchActivated = False
  
    For i = 1 To 100
        ' Check kill switch before each iteration
        Call CheckKillSwitch
        If killSwitchActivated Then Exit Sub
      
        ' Your automation code here
        Application.SendKeys "Some data"
        Sleep 1000
    Next i
End Sub
```

### Comprehensive Error Handling

```vba
Sub RobustAutomation()
    On Error GoTo ErrorHandler
  
    ' Enable ESC key handling
    Application.EnableCancelKey = xlErrorHandler
  
    ' Your automation code here
    Call AutomateExternalProgram()
  
    Exit Sub
  
ErrorHandler:
    If Err.Number = 18 Then  ' User pressed Ctrl+Break or ESC
        MsgBox "Process cancelled by user", vbInformation
    Else
        MsgBox "Error: " & Err.Description & " (Code: " & Err.Number & ")", vbCritical
    End If
  
    ' Reset error handling
    Application.EnableCancelKey = xlInterrupt
End Sub
```

## Timing & Synchronization

### Strategic Sleep Usage

```vba
Sub ProperTiming()
    ' Short delays for UI updates
    Application.SendKeys "{TAB}"
    Sleep 300  ' Allow field to activate
  
    ' Medium delays for data entry
    Application.SendKeys "Important Data"
    Sleep 500  ' Allow typing to complete
  
    ' Long delays for system processing
    Application.SendKeys "~"  ' Submit form
    Sleep 3000  ' Wait for system to process
  
    ' Very long delays for module changes
    Application.SendKeys "MODULE_NAME"
    Application.SendKeys "~"
    Sleep 5000  ' Wait for module to load
End Sub
```

### Adaptive Timing

```vba
Sub WaitForSystemResponse()
    ' Send command
    Application.SendKeys "SEARCH_COMMAND~"
  
    ' Wait with timeout
    Dim timeout As Integer
    timeout = 0
    Do While timeout < 30  ' 30 second maximum wait
        Sleep 1000
        timeout = timeout + 1
      
        ' Check if system is ready (implement your own logic)
        If SystemIsReady() Then Exit Do
      
        ' Check kill switch
        Call CheckKillSwitch
        If killSwitchActivated Then Exit Sub
    Loop
  
    If timeout >= 30 Then
        MsgBox "System did not respond within 30 seconds", vbExclamation
    End If
End Sub
```

## Best Practices

### 1. Modular Design

```vba
' Break automation into logical functions
Sub MainAutomation()
    Call ActivateProgram()
    Call NavigateToModule()
    Call EnterBatchNumber()
    Call ProcessDataRows()
    Call ConfirmAndExit()
End Sub
```

### 2. Extensive Debugging

```vba
Sub DebuggedAutomation()
    Debug.Print "=== Starting automation at " & Now() & " ==="
  
    Debug.Print "Step 1: Activating window..."
    AppActivate ("Target Program")
    Debug.Print "✓ Window activated"
  
    Debug.Print "Step 2: Entering data..."
    Application.SendKeys "Test Data"
    Debug.Print "✓ Data entered"
  
    Debug.Print "=== Automation completed ==="
End Sub
```

### 3. User Feedback

```vba
Sub UserFriendlyAutomation()
    ' Initial confirmation
    If MsgBox("Start automation process?", vbYesNo) = vbNo Then Exit Sub
  
    ' Progress updates
    MsgBox "Step 1 of 3: Activating program...", vbInformation
    Call Step1()
  
    MsgBox "Step 2 of 3: Processing data...", vbInformation
    Call Step2()
  
    MsgBox "Step 3 of 3: Finalizing...", vbInformation
    Call Step3()
  
    ' Completion message
    MsgBox "Automation completed successfully!", vbInformation
End Sub
```

### 4. Data Validation

```vba
Sub ValidateBeforeAutomation()
    Dim ws As Worksheet
    Set ws = ActiveSheet
  
    ' Check required fields
    If ws.Range("J10").Value = "" Then
        MsgBox "Batch number is required in cell J10", vbCritical
        Exit Sub
    End If
  
    ' Check data range
    If ws.Cells(2, 1).Value = "" Then
        MsgBox "No data found starting at row 2", vbCritical
        Exit Sub
    End If
  
    ' Proceed with automation
    Call ProcessAutomation()
End Sub
```

## Complete Example

Here's a simplified but complete example based on the journal entry automation:

```vba
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public killSwitchActivated As Boolean

Sub AutomateJournalEntry()
    ' Initialize
    killSwitchActivated = False
    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo ErrorHandler
  
    ' Read batch number from Excel
    Dim batchNumber As String
    batchNumber = CStr(ActiveSheet.Range("J10").Value)
    If batchNumber = "" Then
        MsgBox "No batch number in J10", vbCritical
        Exit Sub
    End If
  
    ' Activate external program
    AppActivate ("FSUK - Fourth Shift - System Control")
    Sleep 500
  
    ' Navigate and enter batch
    Call NavigateToJournalModule()
    Call EnterBatchNumber(batchNumber)
  
    ' Process data rows
    Dim lastRow As Long
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
  
    For i = 2 To lastRow
        Call CheckKillSwitch()
        If killSwitchActivated Then Exit Sub
      
        ' Read row data
        Dim glCode As String, description As String, amount As Double
        glCode = CStr(ActiveSheet.Cells(i, 1).Value)
        description = CStr(ActiveSheet.Cells(i, 2).Value)
        amount = CDbl(ActiveSheet.Cells(i, 3).Value)
      
        ' Enter journal line
        Call EnterJournalLine(glCode, description, amount)
    Next i
  
    MsgBox "Automation completed!", vbInformation
    Exit Sub
  
ErrorHandler:
    If Err.Number = 18 Then
        MsgBox "Process cancelled", vbInformation
    Else
        MsgBox "Error: " & Err.Description, vbCritical
    End If
End Sub

Sub CheckKillSwitch()
    If GetAsyncKeyState(27) <> 0 Then  ' ESC key
        killSwitchActivated = True
        MsgBox "Kill switch activated!", vbCritical
    End If
    DoEvents
End Sub

Sub NavigateToJournalModule()
    Application.SendKeys "^{HOME}"  ' Go to top
    Sleep 500
    Application.SendKeys "{TAB}"    ' Navigate to search
    Sleep 500
    Application.SendKeys "GLJE"     ' Enter module name
    Sleep 3000
    Application.SendKeys "~"        ' Press Enter
    Sleep 3000
End Sub

Sub EnterBatchNumber(batchNumber As String)
    Application.SendKeys batchNumber
    Sleep 2000
    Application.SendKeys "~"
    Sleep 1000
End Sub

Sub EnterJournalLine(glCode As String, description As String, amount As Double)
    Sleep 800
    Application.SendKeys glCode
    Sleep 500
    Application.SendKeys "{TAB}"
    Sleep 500
    Application.SendKeys description
    Sleep 500
    Application.SendKeys "{TAB}"
    Sleep 500
    Application.SendKeys Format(amount, "0.00")
    Sleep 500
    Application.SendKeys "~"
End Sub
```

## Common Pitfalls

1. **Insufficient Delays**: Always add appropriate Sleep() calls <---- You'll need to alter the Sleep() call to perfect it, ik its not great
2. **No Error Handling**: External programs can be unpredictable
3. **No Kill Switch**: Always provide a way to stop the automation
4. **Hard-coded Values**: Make your code flexible with variables
5. **No Data Validation**: Validate inputs before starting automation
6. **Missing Debug Output**: Use Debug.Print liberally during development

## Conclusion

VBA external program automation is powerful but requires careful attention to timing, error handling, and user safety. Start simple, add extensive debugging, and always include kill switch functionality. The techniques shown here can be adapted to automate virtually any Windows application that accepts keyboard input.
