<div align="center">

## ProperCase


</div>

### Description

If you've used ASP for a while you'll notice that VBScript doesn't support StrConv so you can't proper case your code. Here is small function to include in your ASP that will do that for you.
 
### More Info
 
sString, your input

Your formatted string.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Raine Lightner](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/raine-lightner.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Strings](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/strings__4-26.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/raine-lightner-propercase__4-6009/archive/master.zip)





### Source Code

```
Function ProperCase(sString)
  Dim lTemp
  Dim sTemp, sTemp2
  Dim x
  sString = LCase(sString)
  If Len(sString) Then
    sTemp = Split(sString, " ")
    lTemp = UBound(sTemp)
    For x = 0 To lTemp
      sTemp2 = sTemp2 & UCase(Left(sTemp(x), 1)) & Mid(sTemp(x), 2) & " "
    Next
    ProperCase = trim(sTemp2)
  Else
    ProperCase = sString
  End If
End Function
```

