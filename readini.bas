Attribute VB_Name = "writereadini"
'#################################################################################
'*** Public Variablen
'#################################################################################
Public lastProjPath As String, SchenckPath As String, XoneINI As String
Public PathXoneINI As String, Xonesec As String, Xoneins As String
Public ExportPath As String, DestroyPath As String
Public strSourceName As String, strDestroyName, strAuswertDB As String
Public strFileName As String, INIPath, DateStr As String
Public strDatabasepath As String, strDatabasename As String
Public tmpdbPath As String, tmpdb As String, tmpdbkmpl As String
Public TNr As String, Versuch As String, Pruefling As String
Public TestNr As String, Pruefstand As String, CheckBoxVis As String
Public CSVFileCounts As Integer, ColsNames() As String
Public ButtonCount As Integer, Button01 As String, Button02 As String, Button03 As String
Public Al00, ColsRead, ColsCnt, ChangeSeps, DataCnt As Integer

Public init As Integer, i As Integer, a As Long
Public fieldnames(9, 999)
Public fieldValue(0, 999)
Public LineCount As Long

'#################################################################################
'### For read the ini-files
'#################################################################################
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
                  "GetPrivateProfileStringA" ( _
                  ByVal lpApplicationName As String, _
                  ByVal lpKeyName As Any, _
                  ByVal lpDefault As String, _
                  ByVal lpReturnedString As String, _
                  ByVal nSize As Long, _
                  ByVal lpFileName As String) As Long
 
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias _
                  "WritePrivateProfileStringA" ( _
                  ByVal lpApplicationName As String, _
                  ByVal lpKeyName As Any, _
                  ByVal lpString As Any, _
                  ByVal lpFileName As String) As Long
                  
Public Property Get sPath() As String
    sPath = INIPath
End Property
 
Public Property Let sPath(ByVal NewValue As String)
    INIPath = NewValue
End Property
 
Public Sub WriteString(ByVal Section As String, ByVal Key As String, ByVal sValue As String)
    WritePrivateProfileString Section, Key, sValue, INIPath
End Sub
 
Public Sub WriteValue(ByVal Section As String, ByVal Key As String, ByVal vValue As Variant)
    WriteString Section, Key, CStr(vValue)
End Sub
 
Public Function GetIniString(ByVal Section As String, ByVal Key As String, _
        Optional ByVal Default As String = "") As String
    
    Dim sTemp As String
 
    sTemp = String(256, 0)
    GetPrivateProfileString Section, Key, "", sTemp, Len(sTemp), INIPath
    If InStr(sTemp, Chr$(0)) Then
        sTemp = Left$(sTemp, InStr(sTemp, vbNullChar$) - 1)
    Else
        sTemp = Default
    End If
    
    GetIniString = sTemp
End Function
 
Public Function GetIniLong(ByVal Section As String, ByVal Key As String, Optional ByVal Default As _
        Long = -1) As Long
Dim sTemp As String
 
    sTemp = GetIniString(Section, Key, CStr(Default))
    If IsNumeric(sTemp) Then
        GetIniLong = CInt(sTemp)
    'Else
        'Evtl. Fehlermeldung ausgeben
    End If
End Function
 
Public Function GetIniBool(ByVal Section As String, ByVal Key As String, Optional ByVal Default As _
        Boolean = False) As Boolean
    GetIniBool = CBool(GetIniLong(Section, Key, CInt(Default)))
End Function
 
Sub IniTal()

Dim ININame As String, LastProj1 As String

' Deklaration der Application ini datei:
' Debug.Print App.Path -> Normalerweise nach Erstellung der Anwendung
   INIPath = App.Path '"C:\"
   ININame = "\dbfdata.ini" ' "\dbfdata.ini"
   INIPath = INIPath & ININame
   
   
'DateStr
' Application INI abarbeiten und Globale Variablen fuehlen:
' Hier wird die Xone.ini lokalisiert und darin,
' Path zu Xone auf dem Rechner:
  LastProj1 = GetIniString("Path", "Schenck", INIPath)
  SchenckPath = LastProj1

' Ziel Path fuer copy
  LastProj1 = GetIniString("Path", "Dest", INIPath)
  DestroyPath = LastProj1
  frmdbview.Text1.Text = LastProj1
   
  LastProj1 = GetIniString("Path", "tmpdbf", INIPath)
  tmpdb = LastProj1
   
  LastProj1 = GetIniString("Path", "Export", INIPath)
  ExportPath = "\" & LastProj1
   
  LastProj1 = GetIniString("Files", "dbfData", INIPath)
  tmpdbkmpl = Right(LastProj1, 4)
   
  LastProj1 = GetIniString("Path", "tmpdbPath", INIPath)
  tmpdbPath = LastProj1
   
' See what you want
  LastProj1 = GetIniString("Visual", "CheckVisu", INIPath)
  CheckBoxVis = LastProj1
   
  LastProj1 = GetIniString("Visual", "Button1", INIPath)
  Button01 = LastProj1
   
  LastProj1 = GetIniString("Visual", "Button2", INIPath)
  Button02 = LastProj1
   
  LastProj1 = GetIniString("Visual", "Button3", INIPath)
  Button03 = LastProj1
  
  LastProj1 = GetIniString("Visual", "ChangeSeps", INIPath)
  ChangeSeps = LastProj1
  
  ' Read default Columns Names for Databases:
  DateStr = GetIniString("Colnam", "DateStr", INIPath)
  DataCnt = CInt(DateStr)
  
  LastProj1 = GetIniString("Colnam", "Colscnt", INIPath)
  
  If Len(LastProj1) <> 0 Then
    ColsCnt = CInt(LastProj1)
    ReDim Preserve ColsNames(ColsCnt)
   
    For i = 1 To ColsCnt
        LastProj1 = GetIniString("ColNam", "Cols" & i, INIPath)
        ColsNames(i) = LastProj1
    Next
  End If
  
  DateStr = UCase(ColsNames(DataCnt))
  
' Name der ini datei im obigen Path:
  LastProj1 = GetIniString("Path", "INIFile", INIPath)
  XoneINI = LastProj1

' Section und Insert in der xone.ini delklarieren:
  LastProj1 = GetIniString("Path", "Xonesec", INIPath)
  Xonesec = LastProj1
  LastProj1 = GetIniString("Path", "Xoneins", INIPath)
  Xoneins = LastProj1

' Path zur Xone.ini zusammenstellen:
  INIPath = SchenckPath & XoneINI
   
'Lesen des letzten Projektes auf dem Prüfstand:
  LastProj1 = GetIniString(Xonesec, Xoneins, INIPath)  '***
  lastProjPath = LastProj1 & ExportPath
  frmdbview.Text2.Text = LastProj1 & ExportPath

End Sub

