<div align="center">

## Font List For Each Language


</div>

### Description

With this code you will be able to select one or more language from a list , and then to get all the fonts that was installed on the operating system which support the languages that was chosen.

So if you want to develop a word prossesor maybe you will find it useful - to change the font list when the user had change the language. some part of it i had found on the net ( I dont know who is the Auther) and more help from the MSDN.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-07-29 12:18:26
**By**             |[Gil Shabthai](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/gil-shabthai.md)
**Level**          |Advanced
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Font List 237067292001\.zip](https://github.com/Planet-Source-Code/gil-shabthai-font-list-for-each-language__1-25596/archive/master.zip)

### API Declarations

```
Public Declare Function GetLocaleInfoA Lib "kernel32" (ByVal LCID As Long, ByVal LCType As Long, ByVal lpData As String, ByVal cchData As Integer) As Integer '*
Public Declare Function TranslateCharsetInfo Lib "gdi32" (lpSrc As Long, lpcs As CHARSETINFO, ByVal dwFlags As Long) As Long '*
Public Declare Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExA" (ByVal hDC As Long, lpLogFont As LOGFONT, ByVal lpEnumFontProc As Long, ByVal LParam As Long, ByVal dw As Long) As Long
```





