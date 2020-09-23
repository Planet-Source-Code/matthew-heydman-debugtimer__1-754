<div align="center">

## DebugTimer


</div>

### Description

Have you ever been asked: Which part of the routine is taking so long? or did you ever wonder what function was bogging your app down, or did you ever just want to time a particular statement or function? Welcome to DebugTimer. It's not a resource hog and uses no active-x controls... just the built-in Timer function in VB. This is a very easily implemented class module that allows you to time any line(s) of code or functions or whatever. You can even use multiple timers or nest them. I wrote this to determine the length of time it took to perform various stored procedures, and it worked great. If you

have a similar need, I'm sure this will do the trick.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Heydman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-heydman.md)
**Level**          |Unknown
**User Rating**    |6.0 (629 globes from 105 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Debugging and Error Handling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/debugging-and-error-handling__1-26.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-heydman-debugtimer__1-754/archive/master.zip)

### API Declarations

```
Add a new class module to your project, and name it clsDebugTimer. Paste the following code into it:
' METHODS:
'
' Begin(nTimerIndex, nTimerDescription)
'   - Starts/resets a new timer. Both parameters are optional.
'
'   nTimerIndex should be a number from 0 to 9 to specify
'   which timer is to be used. Omitting this param is the same
'   as passing a zero as this parameter.
'
'   nTimerDescription is a description which can be anything you
'   like, but should probably describe what it is you are timing.
'   Omitting this param will set the description to "Timer 1" (or
'   whatever time index you are using instead of 1)
'
'
' ShowElapsed(nOutputType, nTimerIndex)
'   -Displays the elasped time for the timer specified in nTimerIndex
'   since the Begin method was called. Both parameters are optional.
'
'   nOutputType should be either 1 or 2, and you can use the constants
'   outImmediateWindow and outMsgBox, repectively. This param
'   determines where the output will go- either the immediate window or
'   a message box. The description will be displayed along with the
'   elpased time. If this param is omitted, the output goes to the immediate
'   window.
'
'   nTimerIndex is used to specify which timer you want to display the
'   elapsed time for. (See the description in the Begin method, above).
'   If omitted, timer number 0 (zero) is used.
'
'
'PROPERTIES:
'
'Elapsed(nTimerIndex)
'   -Returns the number of seconds that have elapsed since the Begin
'   method was called for the specified timer. If nTimerIndex is omitted,
'   timer 0 (zero) is assumed.
'
'
Option Explicit
Public Enum OutputTypes
   outImmediateWindow = 1
   outMsgBox = 2
End Enum
Dim nBegin(10) As Single
Dim sDesc(10) As String
Public Sub Begin(Optional nTimerIndex As Integer, Optional sTimerDescription As String)
   If (nTimerIndex < 0 Or nTimerIndex > 9) Then Exit Sub
   If sTimerDescription = "" Then sTimerDescription = "Timer " & Trim(Str(nTimerIndex))
   nBegin(nTimerIndex) = Timer
   sDesc(nTimerIndex) = sTimerDescription
End Sub
Public Property Get Elapsed(Optional nTimerIndex As Integer) As Single
   If (nTimerIndex < 0 Or nTimerIndex > 9) Then Exit Property
   Elapsed = Val(Format(Timer - nBegin(nTimerIndex), "####.##"))
End Property
Public Sub ShowElapsed(Optional nOutputType As OutputTypes, Optional nTimerIndex As Integer)
   If nOutputType = 0 Then nOutputType = outImmediateWindow
   If nOutputType < outImmediateWindow Or nOutputType > outMsgBox Then Exit Sub
   If nOutputType = outImmediateWindow Then
      Debug.Print sDesc(nTimerIndex) & ": " & Elapsed(nTimerIndex) & " seconds"
      Exit Sub
   End If
   If nOutputType = outMsgBox Then
      MsgBox sDesc(nTimerIndex) & ": " & Elapsed(nTimerIndex) & " seconds", vbOKOnly, "Debug Timer"
      Exit Sub
   End If
End Sub
```


### Source Code

```
'Add a new Form to your project, and add 3 command buttons to the
'form (named Command1, Command2, and Command3). Then just
'paste the following code into the form:
Option Explicit
Dim i As Integer
Dim dbg As New clsDebugTimer
Private Sub Command1_Click()
   Me.MousePointer = vbHourglass
   'EXAMPLE 1 - VERY BASIC USAGE
   ' Start the timer
   dbg.Begin
   'Do something that will take a little time
   For i = 0 To 25000: DoEvents: Next
   'By default, calling the ShowElapsed method
   'will display the elapsed time in the immediate window
   dbg.ShowElapsed
   Me.MousePointer = vbDefault
End Sub
Private Sub Command2_Click()
   Me.MousePointer = vbHourglass
   'EXAMPLE 2 - USING THE PARAMETERS
   'Start the timer, this time passing a
   'timer index and a description
   dbg.Begin 0, "Loop from 0 to 25000"
   'Do something that takes time
   For i = 0 To 25000: DoEvents: Next
   'Display the elapsed time for timer index 0 in a message box
   dbg.ShowElapsed outMsgBox, 0
   Me.MousePointer = vbDefault
End Sub
Private Sub Command3_Click()
   Me.MousePointer = vbHourglass
   'EXAMPLE 3 - USING MULTIPLE TIMERS
   'Start the first timer- we'll use an index of 1
   'timer index and a description
   dbg.Begin 1, "Total Time"
      'Start a second timer- (index 2)
      'timer index and a description
      dbg.Begin 2, "Count from 0 to 25000"
      'Do something that takes time
      For i = 0 To 25000: DoEvents: Next
      'Display the elapsed time for the second timer
      dbg.ShowElapsed outImmediateWindow, 2
      'perform another loop like the one we just did above
      dbg.Begin 2, "Count from 0 to 24999"
      'Do something that takes time
      For i = 0 To 24999: DoEvents: Next
      'Display the elapsed time for the second timer
      dbg.ShowElapsed outImmediateWindow, 2
   'Now display the elapsed time for the first timer
   dbg.ShowElapsed outImmediateWindow, 1
   Me.MousePointer = vbDefault
End Sub
```

