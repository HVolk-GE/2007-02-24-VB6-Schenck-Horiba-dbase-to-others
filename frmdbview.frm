VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmdbview 
   Caption         =   "Horiba [Schenck] dbf-to-country-csv-file"
   ClientHeight    =   2355
   ClientLeft      =   2640
   ClientTop       =   2985
   ClientWidth     =   7500
   Icon            =   "frmdbview.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   7500
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox ColsCheck2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   240
      TabIndex        =   47
      Top             =   2040
      Width           =   255
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2520
      TabIndex        =   45
      Top             =   3960
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Schenck-file -> sql-file"
      Height          =   345
      Left            =   2400
      TabIndex        =   44
      Top             =   2880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   855
      Left            =   600
      TabIndex        =   43
      Top             =   3240
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "copy file to des.\*.dbf"
      Height          =   345
      Left            =   480
      TabIndex        =   42
      Top             =   2880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   37
      Text            =   "Text2"
      Top             =   960
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   1560
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Schenck-file -> csv-file"
      Height          =   345
      Left            =   5640
      TabIndex        =   35
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   345
      Left            =   5640
      TabIndex        =   34
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "A&ktualisieren"
      Height          =   300
      Left            =   9720
      TabIndex        =   33
      Top             =   4860
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Neu laden"
      Height          =   300
      Left            =   8640
      TabIndex        =   32
      Top             =   4860
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Löschen"
      Height          =   300
      Left            =   7560
      TabIndex        =   31
      Top             =   4860
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Hinzufügen"
      Height          =   300
      Left            =   6480
      TabIndex        =   30
      Top             =   4860
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Data Data1 
      Align           =   2  'Unten ausrichten
      Connect         =   "Dbase IV;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'Standard-Cursor
      DefaultType     =   2  'ODBC verwenden
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Tabelle
      RecordSource    =   "AUSWERT"
      Top             =   2010
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PRUEFSTAND"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   14
      Left            =   8400
      MaxLength       =   20
      TabIndex        =   29
      Top             =   4520
      Width           =   4500
   End
   Begin VB.TextBox txtFields 
      DataField       =   "VERSUCHSNR"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   13
      Left            =   8400
      MaxLength       =   20
      TabIndex        =   27
      Top             =   4200
      Width           =   4500
   End
   Begin VB.TextBox txtFields 
      DataField       =   "EXPORTFORM"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   12
      Left            =   8400
      MaxLength       =   254
      TabIndex        =   25
      Top             =   3880
      Width           =   4500
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BESCHREIBU"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   11
      Left            =   8400
      MaxLength       =   254
      TabIndex        =   23
      Top             =   3560
      Width           =   4500
   End
   Begin VB.TextBox txtFields 
      DataField       =   "FORMAT"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   10
      Left            =   8400
      MaxLength       =   254
      TabIndex        =   21
      Top             =   3240
      Width           =   4500
   End
   Begin VB.TextBox txtFields 
      DataField       =   "FILENAME"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   9
      Left            =   8400
      MaxLength       =   254
      TabIndex        =   19
      Top             =   2920
      Width           =   4500
   End
   Begin VB.TextBox txtFields 
      DataField       =   "VERSUCHSID"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   8
      Left            =   8400
      TabIndex        =   17
      Top             =   2600
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "REVNUMMER"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   7
      Left            =   8400
      TabIndex        =   15
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "REVDATE"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   6
      Left            =   8400
      MaxLength       =   254
      TabIndex        =   13
      Top             =   1960
      Width           =   4500
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CREADATE"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   5
      Left            =   8400
      MaxLength       =   254
      TabIndex        =   11
      Top             =   1640
      Width           =   4500
   End
   Begin VB.TextBox txtFields 
      DataField       =   "SCHLUESSEL"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   4
      Left            =   8400
      MaxLength       =   20
      TabIndex        =   9
      Top             =   1320
      Width           =   4500
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PRUEFLING"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   3
      Left            =   8400
      MaxLength       =   20
      TabIndex        =   7
      Top             =   1000
      Width           =   4500
   End
   Begin VB.TextBox txtFields 
      DataField       =   "VERS_TYP"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   2
      Left            =   8400
      MaxLength       =   254
      TabIndex        =   5
      Top             =   680
      Width           =   4500
   End
   Begin VB.TextBox txtFields 
      DataField       =   "VERSUCH"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   8400
      MaxLength       =   20
      TabIndex        =   3
      Top             =   360
      Width           =   4500
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DATENTYP"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   8400
      MaxLength       =   3
      TabIndex        =   1
      Top             =   40
      Width           =   4500
   End
   Begin VB.Label Label6 
      Caption         =   "Create a file with ini Columnsnames"
      Height          =   255
      Left            =   600
      TabIndex        =   48
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   615
      Left            =   1560
      TabIndex        =   46
      Top             =   120
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   120
      Picture         =   "frmdbview.frx":014A
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label4 
      Caption         =   "Little tools farm® H. Volk"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   41
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Werte aus der INI Datei bitte dort aendern"
      Height          =   735
      Left            =   4680
      TabIndex        =   40
      Top             =   3840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Last Project :"
      Height          =   255
      Left            =   360
      TabIndex        =   39
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Destination Path :"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRUEFSTAND:"
      Height          =   255
      Index           =   14
      Left            =   6480
      TabIndex        =   28
      Top             =   4545
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "VERSUCHSNR:"
      Height          =   255
      Index           =   13
      Left            =   6480
      TabIndex        =   26
      Top             =   4215
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "EXPORTFORM:"
      Height          =   255
      Index           =   12
      Left            =   6480
      TabIndex        =   24
      Top             =   3900
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BESCHREIBU:"
      Height          =   255
      Index           =   11
      Left            =   6480
      TabIndex        =   22
      Top             =   3585
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "FORMAT:"
      Height          =   255
      Index           =   10
      Left            =   6480
      TabIndex        =   20
      Top             =   3255
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "FILENAME:"
      Height          =   255
      Index           =   9
      Left            =   6480
      TabIndex        =   18
      Top             =   2940
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "VERSUCHSID:"
      Height          =   255
      Index           =   8
      Left            =   6480
      TabIndex        =   16
      Top             =   2625
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "REVNUMMER:"
      Height          =   255
      Index           =   7
      Left            =   6480
      TabIndex        =   14
      Top             =   2295
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "REVDATE:"
      Height          =   255
      Index           =   6
      Left            =   6480
      TabIndex        =   12
      Top             =   1980
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CREADATE:"
      Height          =   255
      Index           =   5
      Left            =   6480
      TabIndex        =   10
      Top             =   1665
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "SCHLUESSEL:"
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   8
      Top             =   1335
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRUEFLING:"
      Height          =   255
      Index           =   3
      Left            =   6480
      TabIndex        =   6
      Top             =   1020
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "VERS_TYP:"
      Height          =   255
      Index           =   2
      Left            =   6480
      TabIndex        =   4
      Top             =   705
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "VERSUCH:"
      Height          =   255
      Index           =   1
      Left            =   6480
      TabIndex        =   2
      Top             =   375
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DATENTYP:"
      Height          =   255
      Index           =   0
      Left            =   6480
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmdbview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
  Data1.Recordset.AddNew
End Sub

Private Sub cmdDelete_Click()
  ' Hierdurch kann ein Fehler hervorgerufen werden, wenn der
  ' gelöschte Datensatz der letzte oder der einzige innerhalb
  ' der Datensatzgruppe ist.
  Data1.Recordset.Delete
  Data1.Recordset.MoveNext
End Sub

Private Sub cmdRefresh_Click()
  ' Dies wird ausschließlich für Mehrbenutzeranwendungen verwendet.
  Data1.Refresh
End Sub

Private Sub cmdUpdate_Click()
  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
End Sub

Private Sub cmdClose_Click()
  Unload Me
  Unload frmdbfResData
End Sub

Private Sub Command2_Click()
Dim inta As Integer, f As Integer, g As Integer

Me.Command1.Enabled = False
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.cmdClose.Enabled = False

strDatabasepath = Data1.DatabaseName & "\"
strDatabasename = Data1.RecordSource & ".dbf"

strAuswertDB = strDatabasepath & strDatabasename
g = 0

For i = 1 To Data1.Recordset.RecordCount
     
     If Me.txtFields(0) = "REP" Then
        strSourceName = Me.txtFields(9)
        TNr = Me.txtFields(4)
        
      For f = 1 To Len(strSourceName)
        strTemp0 = Left(strSourceName, f)
        strTemp1 = Mid(strSourceName, f, 1)
        
        If strTemp1 = "\" Then g = g + 1
        
        If strTemp1 = "\" And g > 5 Then
           strFileName = Mid(strSourceName, f + 1, Len(strSourceName) - f)
           inta = 1
           Exit For
        End If
      Next f
     End If
     
     If inta = 1 Then
        CopyFiles
        inta = 0
     End If
     Data1.Recordset.MoveNext
Next i
  
  MsgBox "Have done ... you can find the files in : " & _
         Me.Text1.Text, vbInformation, "Copy have done..."

Me.cmdClose.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Command3.Enabled = True
  
MsgBox "Copy files have done ....!", vbInformation, "OK, have done !"
  
Unload Me
Unload frmdbfResData
  
End Sub

Private Sub Command1_Click()

Me.Command1.Enabled = False
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.cmdClose.Enabled = False

If frmdbview.ColsCheck2.Value = 1 Then
   ColsRead = 1
   IniTal
End If

ButtonCount = 1
CreateCSVFromDBF
ButtonCount = 0

Me.cmdClose.Enabled = True
Me.Command3.Enabled = True
Me.Command2.Enabled = True
Me.Command1.Enabled = True

MsgBox "CSV file done ....!", vbInformation, "OK, have done !"

Unload Me
Unload frmdbfResData

End Sub

Private Sub Command3_Click()
Me.Command1.Enabled = False
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.cmdClose.Enabled = False

ButtonCount = 2
CreateCSVFromDBF
'CreateSQLFromDBF
ButtonCount = 0

Me.cmdClose.Enabled = True
Me.Command3.Enabled = True
Me.Command2.Enabled = True
Me.Command1.Enabled = True

MsgBox "SQL file done ....!", vbInformation, "OK, have done !"
Unload Me
Unload frmdbfResData

End Sub

Private Sub Data1_Error(DataErr As Integer, Response As Integer)
  ' Hier sollte der Code zur Fehlerbehandlung eingefügt werden.
  ' Falls die Fehler ignoriert werden sollen, kommentieren Sie die nächste Zeile aus.
  ' Falls die Fehler abgefangen werden sollen,
  ' fügen Sie hier den Code für ihre Behandlung ein.
  MsgBox "Datenfehler-Ereignis ausgelöst. Fehler:" & Error$(DataErr)
  Response = 0  ' Ignorieren des Fehlers.
End Sub

Private Sub Data1_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  ' Anzeigen der Position des aktuellen Datensatzes
  ' für Dynasets und Snapshots
  Data1.Caption = "Datensatz: " & (Data1.Recordset.AbsolutePosition + 1)
  ' Die Index-Eigenschaft muß für das Tabellenobjekt festgelegt werden, wenn
  ' die Datensatzgruppe erstellt wird. Dies geschieht mit der folgenden Zeile.
  'Data1.Caption = "Datensatz: " & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
  ' Hier sollte der Code für die Überprüfung der Daten eingefügt werden.
  ' Dieses Ereignis wird ausgelöst, wenn die folgenden Aktionen stattfinden.
  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
  End Select
  Screen.MousePointer = vbHourglass
End Sub

Private Sub Form_Load()

On Error GoTo marke:

ColsRead = 0

frmdbview.ColsCheck2.Value = ColsRead

lstTrennCountry = GetEntry(&HC)

frmdbview.Data1.DatabaseName = App.Path
frmdbview.Data1.RecordSource = "AUSWERT"
frmdbview.Data1.Refresh
frmdbview.Data1.UpdateRecord

frmdbfResData.Data2.DatabaseName = App.Path
frmdbfResData.Data2.RecordSource = "AUSWERT"
frmdbfResData.Data2.Refresh
frmdbfResData.Data2.UpdateRecord

Me.Label5.Caption = "Export Schenck Data to csv-file" & vbCr & _
                    "Tool supported country special" & vbCr & _
                    "List-Cut chars."
                    
Label3.Caption = "Value from the Application-ini file !" & vbCr & _
                 " Please, if you change that in this file !"

IniTal

If Button01 = "True" Then
   Me.Command2.Enabled = True
Else
   Me.Command2.Enabled = False
End If
   
If Button02 = "True" Then
   Me.Command3.Enabled = True
Else
   Me.Command3.Enabled = False
End If
   
If Button03 = "True" Then
   Me.Command1.Enabled = True
Else
   Me.Command1.Enabled = False
End If
   
If CheckBoxVis = "True" Then
   Me.Check1.Enabled = True
   frmdbview.Check1.Caption = "Create MySQL File ? (Yes=Activate and Line-end->Chr(27)ESC) !"
Else
   Me.Check1.Enabled = False
   frmdbview.Check1.Caption = "Create MySQL File ? (Yes=Activate and Line-end->Chr(27)ESC) !"
End If

If lastProjPath = "" Then
 MsgBox "Please, process failed ... can not read the last project in local" & _
 "Xone.ini-file, check this !", vbCritical
 Unload Me
End If

frmdbview.Data1.DatabaseName = lastProjPath
frmdbview.Data1.RecordSource = "AUSWERT"

frmdbview.Data1.Refresh
frmdbview.Data1.UpdateRecord

Exit Sub

marke:
MsgBox "Sorry, one error by loading." & Chr(10) & _
       "Please check follow : " & Chr(10) & _
       "1.) Have you, on these Computer a directory x:\xone ? " & Chr(10) & _
       "2.) Have you, on these Computer a Xone.ini file on directory for 1. ? " & Chr(10) & _
       "3.) Have you in the ini-File a Project on selection [data] ? " & Chr(10) & _
       "You found not Errors with past steps, then contact me under : " & Chr(10) & _
       "EMail : support@little-tools-farm.de", vbCritical, "Not currect config"
End
'Resume Next

End Sub

Private Sub Form_Terminate()
 Unload Me
 Unload frmdbfResData
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unload Me
 Unload frmdbfResData
End Sub
