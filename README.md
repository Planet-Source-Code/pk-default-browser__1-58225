<div align="center">

## Default Browser \-


</div>

### Description

This code will determine your default browser.
 
### More Info
 
Copyright 2005 Paul Kurczaba

The function returns the location of your default browser.

None :)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[pk](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/pk.md)
**Level**          |Intermediate
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script
**Category**       |[Registry](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/registry__1-36.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/pk-default-browser__1-58225/archive/master.zip)





### Source Code

```
Public Function DefaultBrowser()
On Error Resume Next
Dim Regentry As String
Set TheReg = CreateObject("Wscript.Shell")
Regentry = TheReg.RegRead("HKEY_CLASSES_ROOT\HTTP\shell\open\command\")
Regentry = Replace(Regentry, Chr(34), "")
Regentry = Mid(Regentry, 1, InStr(1, LCase(Regentry), ".exe") + 3)
DefaultBrowser = Regentry
End Function
```

