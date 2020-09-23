VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFont 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   Icon            =   "frmFont.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstv 
      Height          =   3480
      Left            =   2520
      TabIndex        =   2
      Top             =   720
      Width           =   5652
      _ExtentX        =   9975
      _ExtentY        =   6138
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ListBox lstFont 
      Height          =   3180
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   5652
   End
   Begin VB.Label lblMessage 
      Caption         =   $"frmFont.frx":0742
      Height          =   492
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   5652
   End
   Begin VB.Image Image1 
      Height          =   4932
      Left            =   120
      Picture         =   "frmFont.frx":07DC
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2292
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      DrawMode        =   4  'Mask Not Pen
      X1              =   8400
      X2              =   2160
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label lblCount 
      Caption         =   "Label1"
      Height          =   492
      Left            =   1680
      TabIndex        =   1
      Top             =   -120
      Visible         =   0   'False
      Width           =   1212
   End
End
Attribute VB_Name = "frmFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tLogFont    As LOGFONT
Dim lLocaleID   As Long
Dim sLocaleID   As String

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       Project1
' Procedure  :       Form_Load
' Description:       [type_description_here]
' Created by :       Gil
' Machine    :       GILPC
' Date-Time  :       28/07/2001-11:57:01
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Load()

On Error GoTo ErrHandler

    Call Init

Exit Sub
ErrHandler:
    MsgBox Err.Number & vbNewLine & Err.Description
End Sub

'--------------------------------------------------------------------------------
' Project    :       Project1
' Procedure  :       Init
' Description:
' Created by :       Gil
' Machine    :       GILPC
' Date-Time  :       28/07/2001-11:57:01
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Init()

On Error GoTo ErrHandler
    '--- fill the collection of languages list and their LCID (local ID number
    '--- for each languages .
    Call FillCol
    
    '--- fill the list view with data
    Call FillListView(lstv)
    
    '--- set lable message
    lblMessage = "Select one or more languages on the list below , then click on the next button to get the installed  font or fonts that support your chosen  language."

Exit Sub
ErrHandler:
    MsgBox Err.Number & vbNewLine & Err.Description
End Sub

'--------------------------------------------------------------------------------
' Project    :       Project1
' Procedure  :       cmdNext_Click
' Description:
' Created by :       Gil
' Machine    :       GILPC
' Date-Time  :       28/07/2001-11:57:01
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdNext_Click()

Dim i               As Integer
Dim j               As Integer
Dim colCharset      As Collection
Dim bAdd            As Boolean
Dim iCharSet        As Long
Dim sLanguages()    As String
Dim sCaption        As String
    
On Error GoTo ErrHandler
    
sLocaleID = ""
Set colCharset = New Collection

    '--- get a list of all languages (LCID) that was chosen
    '--- by the user , we will get long string separate with "|"
    For i = 1 To lstv.ListItems.Count
        If lstv.ListItems.Item(i).Checked = True Then
            lLocaleID = lstv.ListItems.Item(i).SubItems(1)
            
            '--- get the charset ID for each local ID (LCID)
            iCharSet = GetCharsetFromLocaleID(lLocaleID)
            
            sLocaleID = lstv.ListItems(i).Text & "#" & iCharSet & "|" & sLocaleID
            
            '--- ADD EACH CHARSET TO A COLLECTION
            '=================================================
            '--- if this is the first time to add value to the
            '--- collection there will be no need to start with
            '--- the check action - this case we will go to the
            '--- ELSE statement.
            If colCharset.Count > 0 Then
                '--- if there is at least 1 value in the collection
                '--- we have to check if the new one is exist
                '--- if exist flag bAdd = False
                For j = 1 To colCharset.Count
                    If colCharset.Item(j) = CStr(iCharSet) Then
                        bAdd = False
                    Else
                        bAdd = True
                    End If
                Next j
                
                '--- bAdd = True mean that we can add the new value
                '---to the collection
                If bAdd = True Then
                    colCharset.Add iCharSet, CStr(iCharSet)
                End If
            Else
                 colCharset.Add iCharSet, CStr(iCharSet)
            End If
        End If
    Next i
     
    If colCharset.Count = 0 Then
        iCharSet = GetCharsetFromLocaleID(1024)
        colCharset.Add iCharSet, CStr(iCharSet)
    End If
    
    lblCount.Caption = CInt(0)
    
    sLanguages = Split(sLocaleID, "|")
    
    '--- get the charset ID for each local ID (LCID)
    For j = 1 To colCharset.Count
        sCaption = ""
        tLogFont.lfCharSet = colCharset(j)
        '--- set caption for each group of charset
        For i = LBound(sLanguages) To UBound(sLanguages) - 1
            If CLng((Mid(sLanguages(i), InStr(1, sLanguages(i), "#") + 1)) = CLng(colCharset(j))) Then
                sCaption = Mid(sLanguages(i), 1, InStr(1, sLanguages(i), "#") - 1) & " - " & sCaption
            End If
        Next i
        
        If Len(sCaption) = 0 Then
            sCaption = "----------- All Font List ----------"
        End If
        
        lstFont.AddItem ""
        lstFont.AddItem "----------- " & sCaption & " -----------"
        lstFont.AddItem "---------------------------------------------------------------------------------------"
        '--- get font list that will support this charset
        EnumFontFamiliesEx _
                Me.hDC, _
                tLogFont, _
                AddressOf EnumFontFamProc, _
                ByVal 0&, _
                0
    Next j
        
    lblMessage.Caption = "Here is the list of installed fonts in your operating system that support the languages that you choose"
    lstv.Visible = False
    cmdNext.Enabled = False

Exit Sub
ErrHandler:
    MsgBox Err.Number & vbNewLine & Err.Description
End Sub

'--------------------------------------------------------------------------------
' Project    :       Project1
' Procedure  :       cmdBack_Click
' Description:       [type_description_here]
' Created by :       Gil
' Machine    :       GILPC
' Date-Time  :       28/07/2001-11:57:01
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdBack_Click()

On Error GoTo ErrHandler

    lblMessage = "Select one or more languages on the list below , then click on the next button to get the installed  font or fonts that support your chosen  language."
    lstv.Visible = True
    lstFont.Clear
    cmdNext.Enabled = True
Exit Sub
ErrHandler:
    MsgBox Err.Number & vbNewLine & Err.Description
End Sub

'--------------------------------------------------------------------------------
' Project    :       Project1
' Procedure  :       cmdExit_Click
' Description:
' Created by :       Gil
' Machine    :       GILPC
' Date-Time  :       28/07/2001-11:57:01
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdExit_Click()
    
On Error GoTo ErrHandler
    
    Set colLCID = Nothing
    End

Exit Sub
ErrHandler:
    MsgBox Err.Number & vbNewLine & Err.Description
End Sub
