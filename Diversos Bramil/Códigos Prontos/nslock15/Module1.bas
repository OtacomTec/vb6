Attribute VB_Name = "Module1"
Option Explicit

' This sample project requires ActiveLock control
' You may download ActiveLock free of charge at
' http://www.insite.com.br/~nferraz/activelock

Sub Main()

  Load frmRegister
  Load Calculator
  
  With frmRegister.ActiveLock1
    ' If the user have registered, then shows the main form
    If .RegisteredUser Then
      Calculator.Show
    Else
      ' If he/she haven't registered yet, check if
      ' the user tried to fool ActiveLock by changing
      ' the time settings
      If .LastRunDate > Now Then
        MsgBox "ActiveLock has detected that you've changed the clock backwards!"
      End If

      ' Check the evaluation period
      If .UsedDays < 21 Then
        frmRegister.Show 1
        Calculator.Show
      Else
        ' If the evaluation period has expired...
        MsgBox "Your evaluation period has expired!"
        Unload frmRegister
        Unload Calculator
      End If
    End If
  End With

End Sub
