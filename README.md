<div align="center">

## Simple Hex Encode / Decode


</div>

### Description

Two functions: One to turn an ASCII string into a HEX string, and one to turn a HEX string into an ASCII string.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[syntax\.](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/syntax.md)
**Level**          |Beginner
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, VBA MS Access, VBA MS Excel
**Category**       |[Encryption](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/encryption__1-48.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/syntax-simple-hex-encode-decode__1-55602/archive/master.zip)





### Source Code

```
'Encodes a string as hex
Public Function sHexEncode(sData As String) As String
 Dim iChar As Integer
 Dim sOutString As String
 Dim sTmpChar As String
 For iChar = 1 To Len(sData)
  sTmpChar = Hex$(Asc(Mid(sData, iChar, 1)))
  If Len(sTmpChar) = 1 Then sTmpChar = "0" & sTmpChar
  sOutString = sOutString & sTmpChar
 Next iChar
 sHexEncode = sOutString
End Function
'Decodes a string from hex
Public Function sHexDecode(sData As String) As String
 Dim iChar As Integer
 Dim sOutString As String
 Dim sTmpChar As String
 For iChar = 1 To Len(sData) Step 2
  sTmpChar = Chr("&H" & Mid(sData, iChar, 2))
  sOutString = sOutString & sTmpChar
 Next iChar
 sHexDecode = sOutString
End Function
```

