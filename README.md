<div align="center">

## Change the Keyboard Speed and Delay


</div>

### Description

This code allows you to change the Keyboard speed (0 to 31) and delay (0 to 3).. useful in games or as an alternative to using the keyboard control panel.
 
### More Info
 
In your own code use SetFastKeyboard() to change to your new speed and RestoreKeyboard() to change back to thier previous defaults.

On my computer when my old settings are restored they are restored to speed 0 and delay 0.. after much research i determined that the value being read from the API is 0 and 0.. or my variables can't hold the values being returned.. Anyone with more info on the API or the syntax of SPI_GETKEYBOARDDELAY and SPI_GETKEYBOARDSPEED please email me.. :)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Joshua Moore](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/joshua-moore.md)
**Level**          |Unknown
**User Rating**    |4.2 (161 globes from 38 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Games](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/games__1-38.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/joshua-moore-change-the-keyboard-speed-and-delay__1-1674/archive/master.zip)

### API Declarations

```
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
 Public Const SPI_GETKEYBOARDDELAY = 22
 Public Const SPI_GETKEYBOARDSPEED = 10
 Public Const SPI_SETKEYBOARDDELAY = 23
 Public Const SPI_SETKEYBOARDSPEED = 11
 Public Const SPIF_SENDCHANGE = 1
 Public OldKeySpeed As Long
 Public OldKeyDelay As Long
```


### Source Code

```
Sub SetFastKeyboard()
 Dim Retcode As Long
 Dim FastKeySpeed As Long
 Dim FastKeyDelay As Long
 Dim dummy As Long
 FastKeySpeed = 31
 FastKeyDelay = 0
 dummy = 0
 Retcode = SystemParametersInfo(SPI_GETKEYBOARDSPEED, 0, OldKeySpeed, 0)
 Retcode = SystemParametersInfo(SPI_GETKEYBOARDDELAY, 0, OldKeyDelay, 0)
 Retcode = SystemParametersInfo(SPI_SETKEYBOARDSPEED, FastKeySpeed, dummy, SPIF_SENDCHANGE)
 Retcode = SystemParametersInfo(SPI_SETKEYBOARDDELAY, FastKeyDelay, dummy, SPIF_SENDCHANGE)
End Sub
Sub RestoreKeyboard()
 Dim Retcode As Long
 Dim dummy As Long
 dummy = 0
 Retcode = SystemParametersInfo(SPI_SETKEYBOARDSPEED, OldKeySpeed, dummy, SPIF_SENDCHANGE)
 Retcode = SystemParametersInfo(SPI_SETKEYBOARDDELAY, OldKeyDelay, dummy, SPIF_SENDCHANGE)
End Sub
```

