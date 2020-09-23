Attribute VB_Name = "mdlGeneral"
Option Explicit

Public colLCID As Collection
Private i As Integer

'--------------------------------------------------------------------------------
' Project    :       Project1
' Procedure  :       FillListView
' Description:       make headers and fill the ListView
' Created by :       Gil
' Machine    :       GILPC
' Date-Time  :       28/07/2001-11:57:01
'
' Parameters :       lstv
'--------------------------------------------------------------------------------
Public Function FillListView(lstv As ListView)

On Error GoTo ErrHandler

    Call InitListView(lstv)
    
    lstv.ColumnHeaders.Add , , "Languages List"
    lstv.ColumnHeaders.Add , , "Languages LCID"
    
    lstv.ColumnHeaders.Item(1).Width = 3500
    
    For i = 1 To colLCID.Count
        lstv.ListItems.Add , , Mid(colLCID(i), 1, InStr(1, colLCID(i), "|") - 1)
        lstv.ListItems.Item(i).ListSubItems.Add , , Mid(colLCID(i), InStr(1, colLCID(i), "|") + 1)
    Next i
   
Exit Function
ErrHandler:
    MsgBox Err.Number & vbNewLine & Err.Description
End Function
'--------------------------------------------------------------------------------
' Project    :       Project1
' Procedure  :       InitListView
' Description:       set some properties for the ListView
' Created by :       Gil
' Machine    :       GILPC
' Date-Time  :       28/07/2001-11:57:01
'
' Parameters :       lstv
'--------------------------------------------------------------------------------
Private Sub InitListView(lstv As ListView)

On Error GoTo ErrHandler

    lstv.Arrange = lvwAutoTop
    lstv.GridLines = False
    
    lstv.View = lvwReport
    lstv.Appearance = cc3D
    lstv.Checkboxes = True
    lstv.FullRowSelect = True
    lstv.HideColumnHeaders = False
    lstv.AllowColumnReorder = False
Exit Sub
ErrHandler:
    MsgBox Err.Number & vbNewLine & Err.Description
End Sub
'--------------------------------------------------------------------------------
' Project    :       Project1
' Procedure  :       FillCol
' Description:       make a collection for all languages and their LCID
' Created by :       Gil
' Machine    :       GILPC
' Date-Time  :       28/07/2001-11:57:01
'
' Parameters :
'--------------------------------------------------------------------------------
Public Sub FillCol()

On Error GoTo ErrHandler

Set colLCID = New Collection

With colLCID
    .Add "Afrikaans|1078", "1078" '(&H436)
    .Add "Albanian|1052", "1052"  '(&H41C)"
    .Add "Arabic|1025", "1025" '(&H401)"
    .Add "Arabic Algeria|5121", "5121" '(&H1401)"
    .Add "Arabic Bahrain|15361", "15361" ' (&H3C01)
    .Add "Arabic Egypt|3073", "3073" '(&HC01)"
    .Add "Arabic Iraq|2049", "2049" '(&H801)"
    .Add "Arabic Jordan|11265", "11265" '(&H2C01)"
    .Add "Arabic Kuwait|13313", "13313" '(&H3401)"
    .Add "Arabic Lebanon|12289", "12289" ' (&H3001)"
    .Add "Arabic Libya|4097", "4097" '(&H1001)"
    .Add "Arabic Morocco|6145", "6145" '(&H1801)"
    .Add "Arabic Oman|8193", "9193" '(&H2001)"
    .Add "Arabic Qatar|16385", "16385" '(&H4001)"
    .Add "Arabic Saudi Arabia|1025B", "1025B" '(&H401)"
    .Add "Arabic Syria|10241", "10241" '(&H2801)"
    .Add "Arabic Tunisia|7169", "7169" '(&H1C01)"
    .Add "Arabic U.A.E|14337", "14337" '(&H3801"
    .Add "Arabic Yemen|9217", "9217" '(&H2401)"
    .Add "Armenian|1067", "1067" '(&H42B)"
    .Add "Assamese|1101", "1101" '(&H44D)"
    .Add "Azeri Cyrillic|2092", "2092" ' (&H82C)"
    .Add "Azeri Latin|1068", "1068" '(&H42C)"
    .Add "Basque|1069", "1069" '(&H42D)"
    .Add "Belgian Dutch|2067", "2067" ' (&H813)"
    .Add "Belgian French|2060", "2060" ' (&H80C)"
    .Add "Bengali|1093", "1093" '(&H445)"
    .Add "Brazilian Portuguese|1046", "1046" ' (&H416)"
    .Add "Bulgarian|1026", "1026" ' (&H402)"
    .Add "Burmese|1109", "1109" ' (&H455)"
    .Add "Byelorussian (Belarusian)|1059", "1059" '(&H423)"
    .Add "Catalan|1027", "1027" '(&H403)"
    .Add "Chinese Hong Kong SAR|3076", "3076" ' (&HC04)"
    .Add "Chinese Macau SAR|5124", "5124" '(&H1404)"
    .Add "Chinese Simplified|2052", "2052" ' (&H804)"
    .Add "Chinese Singapore|4100", "4100" '(&H1004)"
    .Add "Chinese Traditional|1028", "1028" ' (&H404)"
    .Add "Croatian|1050", "1050" '(&H41A)"
    .Add "Czech|1029", "1029" '(&H405)"
    .Add "Danish|1030", "1030" '(&H406)"
    .Add "Dutch|1043", "1043" '(&H413)"
    .Add "English Australia|3081", "3081" ' (&HC09)"
    .Add "English Belize|10249", "10249" '(&H2809)"
    .Add "English Canadian|4105", "4105" '(&H1009)"
    .Add "English Caribbean|9225", "9225" '(&H2409)"
    .Add "English Ireland|6163", "6163" '(&H1809)"
    .Add "English Jamaica|8201", "8201" '(&H2009)"
    .Add "English New Zealand|5129", "5129" '(&H1409)"
    .Add "English Philippines|13321", "13321" ' (&H3409)"
    .Add "English South Africa|7177", "7177" ' (&H1C09)"
    .Add "English Trinidad|11273", "11273" '(&H2C09)"
    .Add "English U.K.|2057", "2057" '(&H809)"
    .Add "English U.S.|1033", "1033" ' (&H409)"
    .Add "English Zimbabwe|12297", "12297" '(&H3009)"
    .Add "Estonian|1061", "1061" '(&H425)"
    .Add "Faeroese|1080", "1080" ' (&H438)"
    .Add "Farsi|1065", "1065" '(&H429)"
    .Add "Finnish|1035", "1035" '(&H40B)"
    .Add "French|1036", "1036" '(&H40C)"
    .Add "French Cameroon|11276", "11276" ' (&H2C0C)"
    .Add "French Canadian|3084", "3084" '(&HC0C)"
    .Add "French Cote d'Ivoire|12300", "12300" '(&H300C)"
    .Add "French Luxembourg|5132", "5132" '(&H140C)"
    .Add "French Mali|13324", "13324" '(&H340C)"
    .Add "French Monaco|6156", "6156" ' (&H180C)"
    .Add "French Reunion|8204", "8204" ' (&H200C)"
    .Add "French Senegal|10252", "10252" '(&H280C)"
    .Add "French West Indies|7180", "7180" ' (&H1C0C)"
    .Add "Congo (DRC)|9228", "9228" '(&H240C)"
    .Add "Frisian Netherlands|1122", "1122" '(&H462)"
    .Add "Gaelic Ireland|2108", "2108" '(&H83C)"
    .Add "Gaelic Scotland|1084", "1084" '(&H43C)"
    .Add "Galician|1110", "1110" '(&H456)"
    .Add "Georgian|1079", "1079" ' (&H437)"
    .Add "German|1031", "1031" '(&H407)"
    .Add "German Austria|3079", "3079" '(&HC07)"
    .Add "German Liechtenstein|5127", "5127" '(&H1407)"
    .Add "German Luxembourg|4103", "4103" '(&H1007)"
    .Add "Greek|1032", "1032" '(&H408)"
    .Add "Gujarati|1095", "1095" '(&H447)"
    .Add "Hebrew|1037", "1037" '(&H40D)"
    .Add "Hindi|1081", "1081" '(&H439)"
    .Add "Hungarian|1038", "1038" '(&H40E)"
    .Add "Icelandic|1039", "1039" '(&H40F)"
    .Add "Indonesian|1057", "1057" '(&H421)"
    .Add "Italian|1040", "1040" '(&H410)"
    .Add "Japanese|1041", "1041" ' (&H411)"
    .Add "Kannada|1099", "1099" '(&H44B)"
    .Add "Kashmiri|1120", "1120" ' (&H460)"
    .Add "Kazakh|1087", "1087" '(&H43F)"
    .Add "Khmer|1107", "1107" '(&H453)"
    .Add "Kirghiz|1088", "1088" '(&H440)"
    .Add "Konkani|1111", "1111" '(&H457)"
    .Add "Korean|1042", "1042" '(&H412)"
    .Add "Lao|1108", "1108" '(&H454)"
    .Add "Latvian|1062", "1062" '(&H426)"
    .Add "Lithuanian|1063", "1063" '(&H427)"
    .Add "FYRO Macedonian|1071", "1071" ' (&H42F)"
    .Add "Malayalam|1100", "1100" ' (&H44C)"
    .Add "Malay Brunei Darussalam|2110", "2110" '(&H83E)"
    .Add "Malaysian|1086", "1086" ' (&H43E)"
    .Add "Maltese|1082", "1082" '(&H43A)"
    .Add "Manipuri|1112", "1112" ' (&H458)"
    .Add "Marathi|1102", "1102" '(&H44E)"
    .Add "Mongolian|1104", "1104" ' (&H450)"
    .Add "Nepali|1121", "1121" '(&H461)"
    .Add "Norwegian Bokmol|1044", "1044" '(&H414)"
    .Add "Norwegian Nynorsk|2068", "2068" ' (&H814)"
    .Add "Oriya|1096", "1096" '(&H448)"
    .Add "Polish|1045", "1045" ' (&H415)"
    .Add "Portuguese|2070", "2070" '(&H816)"
    .Add "Punjabi|1094", "1094" '(&H446)"
    .Add "Rhaeto -Romanic|1047", "1047" '(&H417)"
    .Add "Romanian|1048", "1048" '(&H418)"
    .Add "Romanian Moldova|2072", "2072" '(&H818)"
    .Add "Russian|1049", "1049" '(&H419)"
    .Add "Russian Moldova|2073", "2073" ' (&H819)"
    .Add "Sami Lappish|1083", "1083" '(&H43B)"
    .Add "Sanskrit|1103", "1103" '(&H44F)"
    .Add "Serbian Cyrillic|3098", "3098" '(&HC1A)"
    .Add "Serbian Latin|2074", "2074" '(&H81A)"
    .Add "Sesotho|1072", "1072" '(&H430)"
    .Add "Sindhi|1113", "1113" '(&H459)"
    .Add "Slovak|1051", "1051" '(&H41B)"
    .Add "Slovenian|1060", "1060" '(&H424)"
    .Add "Sorbian|1070", "1070" '(&H42E)"
    .Add "Spanish (Traditional)|1034", "1034" ' (&H40A)"
    .Add "Spanish Argentina|11274", "11274" '(&H2C0A)"
    .Add "Spanish Bolivia|16394", "16394" '(&H400A)"
    .Add "Spanish Chile|13322", "13322" '(&H340A)"
    .Add "Spanish Colombia|9226", "9226" '(&H240A)"
    .Add "Spanish Costa Rica|5130", "5130" '(&H140A)"
    .Add "Spanish Dominican Republic|7178", "7178" '(&H1C0A)"
    .Add "Spanish Ecuador|12298", "12298" '(&H300A)"
    .Add "Spanish El Salvador|17418", "17418" ' (&H440A)"
    .Add "Spanish Guatemala|4106", "4106" '(&H100A)"
    .Add "Spanish Honduras|18442", "18442" '(&H480A)"
    .Add "Spanish Nicaragua|19466", "19466" ' (&H4C0A)"
    .Add "Spanish Panama|6154", "6154" '(&H180A)"
    .Add "Spanish Paraguay|15370", "15370" '(H3C0A)"
    .Add "Spanish Peru|10250", "10250" '(&H280A)"
    .Add "Spanish Puerto Rico|20490", "20490" '(&H500A)"
    .Add "Spanish Spain (Modern Sort)|3082", "3082" '(&HC0A)"
    .Add "Spanish Uruguay|14346", "14346" '(&H380A)"
    .Add "Spanish Venezuela|8202", "8202" ' (&H200A)"
    .Add "Sutu|1072B", "1072B" '(&H430)"
    .Add "Swahili|1089", "1089" '(&H441)"
    .Add "Swedish|1053", "1053" '(&H41D)"
    .Add "Swedish Finland|2077", "2077" '(&H81D)"
    .Add "Swiss French|4108", "4108" '(&H100C)"
    .Add "Swiss German|2055", "2055" '(&H807)"
    .Add "Swiss Italian|2064", "2064" ' (&H810)"
    .Add "Tajik|1064", "1064" '(&H428)"
    .Add "Tamil|1097", "1097" '(&H449)"
    .Add "Tatar|1092", "1092" '(&H444)"
    .Add "Telugu|1098", "1098" '(&H44A)"
    .Add "Thai|1054", "1054" '(&H41E)"
    .Add "Tibetan|1105", "1105" '(&H451)"
    .Add "Tsonga|1073", "1073" '(&H431)"
    .Add "Tswana|1074", "1074" '(&H432)"
    .Add "Turkish|1055", "1055" ' (&H41F)"
    .Add "Turkmen|1090", "1090" ' (&H442)"
    .Add "Ukrainian|1058", "1058" '(&H422)"
    .Add "Urdu|1056", "1056" '(&H420)"
    .Add "Uzbek Cyrillic|2115", "2115" '(&H843)"
    .Add "Uzbek Latin|1091", "1091" '(&H443)"
    .Add "Venda|1075", "1075" '(&H433)"
    .Add "Vietnamese|1066", "1066" '(&H42A)"
    .Add "Welsh|1106", "1106" '(&H452)"
    .Add "Xhosa|1076", "1076" '(&H434)"
    .Add "Zulu|1077", "1077" '(&H435)"
End With


Exit Sub
ErrHandler:
    MsgBox Err.Number & vbNewLine & Err.Description
End Sub
