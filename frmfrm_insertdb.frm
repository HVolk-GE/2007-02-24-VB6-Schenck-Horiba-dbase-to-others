VERSION 5.00
Begin VB.Form frmfrm_insertdb 
   Caption         =   "crack"
   ClientHeight    =   2385
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   3210
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   3210
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Schließen"
      Height          =   300
      Left            =   4440
      TabIndex        =   182
      Top             =   28540
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "A&ktualisieren"
      Height          =   300
      Left            =   3360
      TabIndex        =   181
      Top             =   28540
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Neu laden"
      Height          =   300
      Left            =   2280
      TabIndex        =   180
      Top             =   28540
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Löschen"
      Height          =   300
      Left            =   1200
      TabIndex        =   179
      Top             =   28540
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Hinzufügen"
      Height          =   300
      Left            =   120
      TabIndex        =   178
      Top             =   28540
      Width           =   975
   End
   Begin VB.Data Data1 
      Align           =   2  'Unten ausrichten
      Connect         =   "ODBC;DATABASE=dynoresults;DESCRIPTION=dynoresults;DSN=dynoresults;OPTION=0;PWD=091266;PORT=0;SERVER=localhost;UID=root"
      DatabaseName    =   ""
      DefaultCursorType=   0  'Standard-Cursor
      DefaultType     =   2  'ODBC verwenden
      Exclusive       =   0   'False
      Height          =   300
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "crack"
      Top             =   2085
      Width           =   3210
   End
   Begin VB.TextBox txtFields 
      DataField       =   "VSO_PRO"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   88
      Left            =   2040
      TabIndex        =   177
      Top             =   28200
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Index"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   87
      Left            =   2040
      MaxLength       =   255
      TabIndex        =   175
      Top             =   27880
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "VMIN_SB"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   86
      Left            =   2040
      TabIndex        =   173
      Top             =   27560
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "VERSUCH"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   85
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   171
      Top             =   27240
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "V_START"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   84
      Left            =   2040
      TabIndex        =   169
      Top             =   26920
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "V_END"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   83
      Left            =   2040
      TabIndex        =   167
      Top             =   26600
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "V"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   82
      Left            =   2040
      TabIndex        =   165
      Top             =   26280
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TS2_6"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   81
      Left            =   2040
      TabIndex        =   163
      Top             =   25960
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TS2_5"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   80
      Left            =   2040
      TabIndex        =   161
      Top             =   25640
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TS2_4"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   79
      Left            =   2040
      TabIndex        =   159
      Top             =   25320
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TS2_3"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   78
      Left            =   2040
      TabIndex        =   157
      Top             =   25000
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TS2_2"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   77
      Left            =   2040
      TabIndex        =   155
      Top             =   24680
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TS2_1"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   76
      Left            =   2040
      TabIndex        =   153
      Top             =   24360
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TS1_6"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   75
      Left            =   2040
      TabIndex        =   151
      Top             =   24040
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TS1_5"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   74
      Left            =   2040
      TabIndex        =   149
      Top             =   23720
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TS1_4"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   73
      Left            =   2040
      TabIndex        =   147
      Top             =   23400
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TS1_3"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   72
      Left            =   2040
      TabIndex        =   145
      Top             =   23080
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TS1_2"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   71
      Left            =   2040
      TabIndex        =   143
      Top             =   22760
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TS1_1"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   70
      Left            =   2040
      TabIndex        =   141
      Top             =   22440
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TRQ1_2"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   69
      Left            =   2040
      TabIndex        =   139
      Top             =   22120
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TRQ1_1"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   68
      Left            =   2040
      TabIndex        =   137
      Top             =   21800
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TRQ_MAX"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   67
      Left            =   2040
      TabIndex        =   135
      Top             =   21480
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TRQ_AVG"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   66
      Left            =   2040
      TabIndex        =   133
      Top             =   21160
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Time"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   65
      Left            =   2040
      TabIndex        =   131
      Top             =   20840
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "T2_MAX"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   64
      Left            =   2040
      TabIndex        =   129
      Top             =   20520
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "T2_AVGST"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   63
      Left            =   2040
      TabIndex        =   127
      Top             =   20200
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "T2_AVGEN"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   62
      Left            =   2040
      TabIndex        =   125
      Top             =   19880
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "T2_6"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   61
      Left            =   2040
      TabIndex        =   123
      Top             =   19560
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "T2_5"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   60
      Left            =   2040
      TabIndex        =   121
      Top             =   19240
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "T2_4"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   59
      Left            =   2040
      TabIndex        =   119
      Top             =   18920
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "T2_3"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   58
      Left            =   2040
      TabIndex        =   117
      Top             =   18600
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "T2_2"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   57
      Left            =   2040
      TabIndex        =   115
      Top             =   18280
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "T2_1"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   56
      Left            =   2040
      TabIndex        =   113
      Top             =   17960
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "T1_2_AVGEN"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   55
      Left            =   2040
      TabIndex        =   111
      Top             =   17640
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "T1_2AVGST"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   54
      Left            =   2040
      TabIndex        =   109
      Top             =   17320
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "T1_MAX"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   53
      Left            =   2040
      TabIndex        =   107
      Top             =   17000
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "T1_AVGST"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   52
      Left            =   2040
      TabIndex        =   105
      Top             =   16680
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "T1_AVGEN"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   51
      Left            =   2040
      TabIndex        =   103
      Top             =   16360
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "T1_AVG"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   50
      Left            =   2040
      TabIndex        =   101
      Top             =   16040
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "T1_6"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   49
      Left            =   2040
      TabIndex        =   99
      Top             =   15720
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "T1_5"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   48
      Left            =   2040
      TabIndex        =   97
      Top             =   15400
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "T1_4"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   47
      Left            =   2040
      TabIndex        =   95
      Top             =   15080
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "T1_3"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   46
      Left            =   2040
      TabIndex        =   93
      Top             =   14760
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "T1_2"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   45
      Left            =   2040
      TabIndex        =   91
      Top             =   14440
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "T1_1"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   44
      Left            =   2040
      TabIndex        =   89
      Top             =   14120
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "STUFNR"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   43
      Left            =   2040
      TabIndex        =   87
      Top             =   13800
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "STR2_MAX"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   42
      Left            =   2040
      TabIndex        =   85
      Top             =   13480
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "STR2_1"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   41
      Left            =   2040
      TabIndex        =   83
      Top             =   13160
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "STR1_MAX"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   40
      Left            =   2040
      TabIndex        =   81
      Top             =   12840
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "STR1_1"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   39
      Left            =   2040
      TabIndex        =   79
      Top             =   12520
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "STEPNR"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   38
      Left            =   2040
      TabIndex        =   77
      Top             =   12200
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "SEQUENCE"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   37
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   75
      Top             =   11880
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "SCHLUESSEL"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   36
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   73
      Top             =   11560
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PS2_1SET"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   35
      Left            =   2040
      TabIndex        =   71
      Top             =   11240
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PS2_1MAX"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   34
      Left            =   2040
      TabIndex        =   69
      Top             =   10920
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PS2"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   33
      Left            =   2040
      TabIndex        =   67
      Top             =   10600
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PS1_1SET"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   32
      Left            =   2040
      TabIndex        =   65
      Top             =   10280
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PS1_1MAX"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   31
      Left            =   2040
      TabIndex        =   63
      Top             =   9960
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PS1"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   30
      Left            =   2040
      TabIndex        =   61
      Top             =   9640
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PRUEFSTAND"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   29
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   59
      Top             =   9320
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PRUEFLING"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   28
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   57
      Top             =   9000
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "P2_MIT"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   27
      Left            =   2040
      TabIndex        =   55
      Top             =   8680
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "P1_MIT"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   26
      Left            =   2040
      TabIndex        =   53
      Top             =   8360
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "MUSCH2"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   25
      Left            =   2040
      TabIndex        =   51
      Top             =   8040
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "MUSCH1"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   24
      Left            =   2040
      TabIndex        =   49
      Top             =   7720
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "MS1_PRO"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   23
      Left            =   2040
      TabIndex        =   47
      Top             =   7400
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ModulName"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   22
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   45
      Top             =   7080
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "MFDDMS2"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   21
      Left            =   2040
      TabIndex        =   43
      Top             =   6760
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "MFDD2KNM"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   20
      Left            =   2040
      TabIndex        =   41
      Top             =   6440
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "MFDD1KNM"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   19
      Left            =   2040
      TabIndex        =   39
      Top             =   6120
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "MD2_IST"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   18
      Left            =   2040
      TabIndex        =   37
      Top             =   5800
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "MD1_IST"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   17
      Left            =   2040
      TabIndex        =   35
      Top             =   5480
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "MD_SUM"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   16
      Left            =   2040
      TabIndex        =   33
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "LOOP2"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   15
      Left            =   2040
      TabIndex        =   31
      Top             =   4840
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "LOOP1"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   14
      Left            =   2040
      TabIndex        =   29
      Top             =   4520
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "LOOP_1"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   13
      Left            =   2040
      TabIndex        =   27
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "LFZEIT"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   12
      Left            =   2040
      TabIndex        =   25
      Top             =   3880
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "KU_ATIM"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   11
      Left            =   2040
      TabIndex        =   23
      Top             =   3560
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Frictioncoefficient"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   10
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   21
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "F1_2"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   19
      Top             =   2920
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DYNWHEEL"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   17
      Top             =   2600
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DREHRI"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   7
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   15
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "MFDATETIME"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   6
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   13
      Top             =   1960
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DATETIME"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   1640
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BRAKENR"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BR_TIME"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1000
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BR_DIST"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   680
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BETRIEB"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   3
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "B_MASS1"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   40
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      Caption         =   "VSO_PRO:"
      Height          =   255
      Index           =   88
      Left            =   120
      TabIndex        =   176
      Top             =   28220
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Index:"
      Height          =   255
      Index           =   87
      Left            =   120
      TabIndex        =   174
      Top             =   27900
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "VMIN_SB:"
      Height          =   255
      Index           =   86
      Left            =   120
      TabIndex        =   172
      Top             =   27580
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "VERSUCH:"
      Height          =   255
      Index           =   85
      Left            =   120
      TabIndex        =   170
      Top             =   27260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "V_START:"
      Height          =   255
      Index           =   84
      Left            =   120
      TabIndex        =   168
      Top             =   26940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "V_END:"
      Height          =   255
      Index           =   83
      Left            =   120
      TabIndex        =   166
      Top             =   26620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "V:"
      Height          =   255
      Index           =   82
      Left            =   120
      TabIndex        =   164
      Top             =   26300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TS2_6:"
      Height          =   255
      Index           =   81
      Left            =   120
      TabIndex        =   162
      Top             =   25980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TS2_5:"
      Height          =   255
      Index           =   80
      Left            =   120
      TabIndex        =   160
      Top             =   25660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TS2_4:"
      Height          =   255
      Index           =   79
      Left            =   120
      TabIndex        =   158
      Top             =   25340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TS2_3:"
      Height          =   255
      Index           =   78
      Left            =   120
      TabIndex        =   156
      Top             =   25020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TS2_2:"
      Height          =   255
      Index           =   77
      Left            =   120
      TabIndex        =   154
      Top             =   24700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TS2_1:"
      Height          =   255
      Index           =   76
      Left            =   120
      TabIndex        =   152
      Top             =   24380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TS1_6:"
      Height          =   255
      Index           =   75
      Left            =   120
      TabIndex        =   150
      Top             =   24060
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TS1_5:"
      Height          =   255
      Index           =   74
      Left            =   120
      TabIndex        =   148
      Top             =   23740
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TS1_4:"
      Height          =   255
      Index           =   73
      Left            =   120
      TabIndex        =   146
      Top             =   23420
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TS1_3:"
      Height          =   255
      Index           =   72
      Left            =   120
      TabIndex        =   144
      Top             =   23100
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TS1_2:"
      Height          =   255
      Index           =   71
      Left            =   120
      TabIndex        =   142
      Top             =   22780
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TS1_1:"
      Height          =   255
      Index           =   70
      Left            =   120
      TabIndex        =   140
      Top             =   22460
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TRQ1_2:"
      Height          =   255
      Index           =   69
      Left            =   120
      TabIndex        =   138
      Top             =   22140
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TRQ1_1:"
      Height          =   255
      Index           =   68
      Left            =   120
      TabIndex        =   136
      Top             =   21820
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TRQ_MAX:"
      Height          =   255
      Index           =   67
      Left            =   120
      TabIndex        =   134
      Top             =   21500
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TRQ_AVG:"
      Height          =   255
      Index           =   66
      Left            =   120
      TabIndex        =   132
      Top             =   21180
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Time:"
      Height          =   255
      Index           =   65
      Left            =   120
      TabIndex        =   130
      Top             =   20860
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "T2_MAX:"
      Height          =   255
      Index           =   64
      Left            =   120
      TabIndex        =   128
      Top             =   20540
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "T2_AVGST:"
      Height          =   255
      Index           =   63
      Left            =   120
      TabIndex        =   126
      Top             =   20220
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "T2_AVGEN:"
      Height          =   255
      Index           =   62
      Left            =   120
      TabIndex        =   124
      Top             =   19900
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "T2_6:"
      Height          =   255
      Index           =   61
      Left            =   120
      TabIndex        =   122
      Top             =   19580
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "T2_5:"
      Height          =   255
      Index           =   60
      Left            =   120
      TabIndex        =   120
      Top             =   19260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "T2_4:"
      Height          =   255
      Index           =   59
      Left            =   120
      TabIndex        =   118
      Top             =   18940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "T2_3:"
      Height          =   255
      Index           =   58
      Left            =   120
      TabIndex        =   116
      Top             =   18620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "T2_2:"
      Height          =   255
      Index           =   57
      Left            =   120
      TabIndex        =   114
      Top             =   18300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "T2_1:"
      Height          =   255
      Index           =   56
      Left            =   120
      TabIndex        =   112
      Top             =   17980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "T1_2_AVGEN:"
      Height          =   255
      Index           =   55
      Left            =   120
      TabIndex        =   110
      Top             =   17660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "T1_2AVGST:"
      Height          =   255
      Index           =   54
      Left            =   120
      TabIndex        =   108
      Top             =   17340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "T1_MAX:"
      Height          =   255
      Index           =   53
      Left            =   120
      TabIndex        =   106
      Top             =   17020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "T1_AVGST:"
      Height          =   255
      Index           =   52
      Left            =   120
      TabIndex        =   104
      Top             =   16700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "T1_AVGEN:"
      Height          =   255
      Index           =   51
      Left            =   120
      TabIndex        =   102
      Top             =   16380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "T1_AVG:"
      Height          =   255
      Index           =   50
      Left            =   120
      TabIndex        =   100
      Top             =   16060
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "T1_6:"
      Height          =   255
      Index           =   49
      Left            =   120
      TabIndex        =   98
      Top             =   15740
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "T1_5:"
      Height          =   255
      Index           =   48
      Left            =   120
      TabIndex        =   96
      Top             =   15420
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "T1_4:"
      Height          =   255
      Index           =   47
      Left            =   120
      TabIndex        =   94
      Top             =   15100
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "T1_3:"
      Height          =   255
      Index           =   46
      Left            =   120
      TabIndex        =   92
      Top             =   14780
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "T1_2:"
      Height          =   255
      Index           =   45
      Left            =   120
      TabIndex        =   90
      Top             =   14460
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "T1_1:"
      Height          =   255
      Index           =   44
      Left            =   120
      TabIndex        =   88
      Top             =   14140
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "STUFNR:"
      Height          =   255
      Index           =   43
      Left            =   120
      TabIndex        =   86
      Top             =   13820
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "STR2_MAX:"
      Height          =   255
      Index           =   42
      Left            =   120
      TabIndex        =   84
      Top             =   13500
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "STR2_1:"
      Height          =   255
      Index           =   41
      Left            =   120
      TabIndex        =   82
      Top             =   13180
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "STR1_MAX:"
      Height          =   255
      Index           =   40
      Left            =   120
      TabIndex        =   80
      Top             =   12860
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "STR1_1:"
      Height          =   255
      Index           =   39
      Left            =   120
      TabIndex        =   78
      Top             =   12540
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "STEPNR:"
      Height          =   255
      Index           =   38
      Left            =   120
      TabIndex        =   76
      Top             =   12220
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "SEQUENCE:"
      Height          =   255
      Index           =   37
      Left            =   120
      TabIndex        =   74
      Top             =   11900
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "SCHLUESSEL:"
      Height          =   255
      Index           =   36
      Left            =   120
      TabIndex        =   72
      Top             =   11580
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PS2_1SET:"
      Height          =   255
      Index           =   35
      Left            =   120
      TabIndex        =   70
      Top             =   11260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PS2_1MAX:"
      Height          =   255
      Index           =   34
      Left            =   120
      TabIndex        =   68
      Top             =   10940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PS2:"
      Height          =   255
      Index           =   33
      Left            =   120
      TabIndex        =   66
      Top             =   10620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PS1_1SET:"
      Height          =   255
      Index           =   32
      Left            =   120
      TabIndex        =   64
      Top             =   10300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PS1_1MAX:"
      Height          =   255
      Index           =   31
      Left            =   120
      TabIndex        =   62
      Top             =   9980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PS1:"
      Height          =   255
      Index           =   30
      Left            =   120
      TabIndex        =   60
      Top             =   9660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRUEFSTAND:"
      Height          =   255
      Index           =   29
      Left            =   120
      TabIndex        =   58
      Top             =   9340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "PRUEFLING:"
      Height          =   255
      Index           =   28
      Left            =   120
      TabIndex        =   56
      Top             =   9020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "P2_MIT:"
      Height          =   255
      Index           =   27
      Left            =   120
      TabIndex        =   54
      Top             =   8700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "P1_MIT:"
      Height          =   255
      Index           =   26
      Left            =   120
      TabIndex        =   52
      Top             =   8380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "MUSCH2:"
      Height          =   255
      Index           =   25
      Left            =   120
      TabIndex        =   50
      Top             =   8060
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "MUSCH1:"
      Height          =   255
      Index           =   24
      Left            =   120
      TabIndex        =   48
      Top             =   7740
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "MS1_PRO:"
      Height          =   255
      Index           =   23
      Left            =   120
      TabIndex        =   46
      Top             =   7420
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ModulName:"
      Height          =   255
      Index           =   22
      Left            =   120
      TabIndex        =   44
      Top             =   7100
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "MFDDMS2:"
      Height          =   255
      Index           =   21
      Left            =   120
      TabIndex        =   42
      Top             =   6780
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "MFDD2KNM:"
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   40
      Top             =   6460
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "MFDD1KNM:"
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   38
      Top             =   6140
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "MD2_IST:"
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   36
      Top             =   5820
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "MD1_IST:"
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   34
      Top             =   5500
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "MD_SUM:"
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   32
      Top             =   5180
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "LOOP2:"
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   30
      Top             =   4860
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "LOOP1:"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   28
      Top             =   4540
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "LOOP_1:"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   26
      Top             =   4220
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "LFZEIT:"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   24
      Top             =   3900
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "KU_ATIM:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   22
      Top             =   3580
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Frictioncoefficient:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   3260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "F1_2:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Top             =   2940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DYNWHEEL:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   2620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DREHRI:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   2300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "MFDATETIME:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "DATETIME:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BRAKENR:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BR_TIME:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BR_DIST:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "BETRIEB:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "B_MASS1:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmfrm_insertdb"
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

