Attribute VB_Name = "mdlFont"
'--------------------------------------------------------------------------------
' Project    :       Project1
' Procedure  :       EnumFontFamProc
' Description:       collect all fonts list that support one of the languges
'                    that was chosen by the user.
' Created by :       GIL ( basic function was found on the net)
' Machine    :       GILPC
' Date-Time  :       28/07/2001-11:57:01
'
' Parameters :       lpNLF
'                    lpNTM
'                    FontType , TrueType Font = 4
'                    LParam
'--------------------------------------------------------------------------------
Function EnumFontFamProc(lpNLF As LOGFONT, _
                         lpNTM As NEWTEXTMETRIC, _
                         ByVal FontType As Long, _
                         LParam As Long) As Long
    
    Dim FaceName As String
        
On Error GoTo ErrHandler
        
    '--- get the name of the font ---
    FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
    
    '--- add font name to the List ---
    frmFont.lstFont.AddItem Left(FaceName, InStr(FaceName, vbNullChar) - 1)
    frmFont.lblCount.Caption = CInt(frmFont.lblCount.Caption) + 1
    
    '--- continue counting ---
    EnumFontFamProc = 1
    
Exit Function
ErrHandler:
    MsgBox Err.Number & vbNewLine & Err.Description
End Function
'--------------------------------------------------------------------------------
' Project    :       Project1
' Procedure  :       GetCharsetFromLocaleID
' Description:       get the Charset ID from LCID number
' Created by :       function was found on the net
' Machine    :       GILPC
' Date-Time  :       28/07/2001-11:57:01
'
' Parameters :       LCID
'--------------------------------------------------------------------------------
Public Function GetCharsetFromLocaleID(Optional LCID As Long = 1024) As CharsetList '*
    Dim cpg         As Long
    Dim stBuffer    As String
    Dim cs          As CHARSETINFO
    Dim OEM         As Long
    
On Error GoTo ErrHandler
    
    '--- retrieves the current ANSI code-page identifier for the system
    stBuffer = GetLocaleInfoC(LCID, LOCALE_IDEFAULTANSICODEPAGE)
     
    If Len(stBuffer) > 0 Then
        cpg = stBuffer
        '--- translates based on the specified character set, code page,
        '--- or font signature value, setting all members of the
        '--- destination structure to appropriate values.
        '--- Return Values - 0 indicates failure
        OEM = TranslateCharsetInfo(ByVal cpg, cs, TCI_SRCCODEPAGE)
        
        If (OEM <> 0) Then
            '--- return the charset value
            GetCharsetFromLocaleID = cs.ciCharset
        End If
    End If

Exit Function
ErrHandler:
    MsgBox Err.Number & vbNewLine & Err.Description
End Function

'--------------------------------------------------------------------------------
' Project    :       Project1
' Procedure  :       GetLocaleInfoC
' Description:       retrieves the current ANSI code-page identifier for the system
' Created by :       function was found on the net
' Machine    :       GILPC
' Date-Time  :       28/07/2001-11:57:01
'
' Parameters :       Locale
'                    LCType
'--------------------------------------------------------------------------------

Public Function GetLocaleInfoC(Locale As Long, LCType As Long) As String '*
    Dim LCID As Long
    Dim stBuff As String
    Dim rc As Long

On Error GoTo ErrHandler

    stBuff = String$(255, vbNullChar)

    rc = GetLocaleInfoA(Locale, LCType, ByVal stBuff, Len(stBuff))

    If (rc > 0) Then
        '--- get value without NULL
        GetLocaleInfoC = StringNoNull(stBuff)
    End If

Exit Function
ErrHandler:
    MsgBox Err.Number & vbNewLine & Err.Description
End Function

'--------------------------------------------------------------------------------
' Project    :       Project1
' Procedure  :       StFromSz
' Description:       get value of string without NULL
' Created by :       Gil
' Machine    :       GILPC
' Date-Time  :       28/07/2001-11:57:01
'
' Parameters :       szTmp
'--------------------------------------------------------------------------------

Public Function StringNoNull(StringWithNull As String) As String
    
    Dim FirstNullPlace As Integer
    
On Error GoTo ErrHandler
    
    FirstNullPlace = InStr(1, StringWithNull, vbNullChar, vbBinaryCompare)
    If FirstNullPlace Then
        StringNoNull = Left$(StringWithNull, FirstNullPlace - 1)
    Else
        StringNoNull = StringWithNull
    End If

Exit Function
ErrHandler:
    MsgBox Err.Number & vbNewLine & Err.Description
End Function

