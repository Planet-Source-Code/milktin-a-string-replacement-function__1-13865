<div align="center">

## A "String Replacement" Function


</div>

### Description

I know there is the

replace(text1.text,"Jack","Jill")

in VB6 which would find all the words Jack and replace them with Jill in text1, but how can I do this in VB5?

I want to be able to put symbols in general sentences, and replace the symbols with specific data. such as:

Thats a great pass from #!

He passes to # who sets up a shot!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[MilkTin](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/milktin.md)
**Level**          |Beginner
**User Rating**    |4.4 (44 globes from 10 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/milktin-a-string-replacement-function__1-13865/archive/master.zip)





### Source Code

```
'Explainaion - http://go.to/cyberprogrammer
Private Sub cmdReplace_Click()
 Text1.Text = pReplace(Text1.Text, txtFind, txtReplace)
End Sub
Public Function pReplace(strExpression As String, strFind As String, strReplace As String)
 Dim intX As Integer
 If (Len(strExpression) - Len(strFind)) >= 0 Then
  For intX = 1 To Len(strExpression)
    If Mid(strExpression, intX, Len(strFind)) = strFind Then
      strExpression = Left(strExpression, (intX - 1)) + strReplace + Mid(strExpression, intX + Len(strFind), Len(strExpression))
    End If
  Next
 End If
 pReplace = strExpression
End Function
```

