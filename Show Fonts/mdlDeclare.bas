Attribute VB_Name = "mdlDeclare"
Option Explicit


Public Const NTM_REGULAR = &H40&
Public Const NTM_BOLD = &H20&
Public Const NTM_ITALIC = &H1&
Public Const TMPF_FIXED_PITCH = &H1
Public Const TMPF_VECTOR = &H2
Public Const TMPF_DEVICE = &H8
Public Const TMPF_TRUETYPE = &H4
Public Const ELF_VERSION = 0
Public Const ELF_CULTURE_LATIN = 0
Public Const RASTER_FONTTYPE = &H1
Public Const DEVICE_FONTTYPE = &H2
Public Const TRUETYPE_FONTTYPE = &H4
Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64


Public Const LOCALE_IDEFAULTCODEPAGE = &HB& '*
Public Const LOCALE_IDEFAULTANSICODEPAGE = &H1004& '*
Public Const TCI_SRCCODEPAGE = 2 '*

Public Enum CharsetList
    ANSI_CHARSET = 0
    DEFAULT_CHARSET = 1
    SYMBOL_CHARSET = 2
    SHIFTJIS_CHARSET = 128
    HANGEUL_CHARSET = 129
    GB2312_CHARSET = 134
    CHINESEBIG5_CHARSET = 136
    OEM_CHARSET = 255
    JOHAB_CHARSET = 130
    HEBREW_CHARSET = 177
    ARABIC_CHARSET = 178
    GREEK_CHARSET = 161
    TURKISH_CHARSET = 162
    VIETNAMESE_CHARSET = 163
    THAI_CHARSET = 222
    EASTEUROPE_CHARSET = 238
    RUSSIAN_CHARSET = 204
    MAC_CHARSET = 77
    BALTIC_CHARSET = 186
End Enum

Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Type NEWTEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
    ntmFlags As Long
    ntmSizeEM As Long
    ntmCellHeight As Long
    ntmAveWidth As Long
End Type

Public Declare Function GetLocaleInfoA Lib "kernel32" (ByVal LCID As Long, ByVal LCType As Long, ByVal lpData As String, ByVal cchData As Integer) As Integer '*
Public Declare Function TranslateCharsetInfo Lib "gdi32" (lpSrc As Long, lpcs As CHARSETINFO, ByVal dwFlags As Long) As Long '*
Public Declare Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExA" (ByVal hDC As Long, lpLogFont As LOGFONT, ByVal lpEnumFontProc As Long, ByVal LParam As Long, ByVal dw As Long) As Long


Public Type FONTSIGNATURE ' *
        fsUsb(4) As Long
        fsCsb(2) As Long
End Type

Public Type CHARSETINFO '*
        ciCharset As Long
        ciACP As Long
        fs As FONTSIGNATURE
End Type

