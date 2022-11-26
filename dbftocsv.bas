Attribute VB_Name = "tofile"
'**********************************************************************************
'*** Read position of AUSWERT.DBF and read this for location the data-dbf file
'*** in Xone directory.
'*** Here read the AUSWERT.DBF from last Project, that exported by XBrake
'*** was. Read that and create 1 by 1 a csv-file for that.
'*** Read the AUSWERT.DBF for:
'*** 1.) Schedule Number
'*** 2.) Projectdata Number
'*** 3.) Test Number
'*** 4.) Dynamometername
'**********************************************************************************

Sub CreateCSVFromDBF()

'********************************************************************************************
'*** Here read this tool, the Schenck "AUSWERT - dbf" file, in currectly path,
'*** to find this "AUSWERT.dbf file, read this tool the original Schenck - xone.ini
'*** file, locate this in the harddisc and set this information in the appl.-ini file.
'*** The next search this tool, in the AUSWERT.dbf the Path for all datarows with "REP".
'*** Now read all information for create the files, from this file !
'********************************************************************************************

CSVFileCounts = 0
Al00 = 0

For ini = 1 To frmdbview.Data1.Recordset.RecordCount

g = 0

     StrRepName = frmdbview.txtFields(0)
     If StrRepName = "REP" Then
     
        strSourceName = frmdbview.txtFields(9) 'Quellen Path + datei
        
        'Versuch = Programm Nr. -> frmdbview.txtFields(1)
        'Pruefling = Technische Daten -> frmdbview.txtFields(3)
        'Schuessel = Testnummer -> frmdbview.txtFields(4)
        'Pruefstand = Pruefstand -> frmdbview.txtFields(14)
        
        Versuch = Left(frmdbview.txtFields(1), 8)
        Pruefling = frmdbview.txtFields(3)
        TestNr = frmdbview.txtFields(4)           ' Testnummer
        Pruefstand = frmdbview.txtFields(14)
        
      For f = 1 To Len(strSourceName)
        strTemp0 = Left(strSourceName, f)
        strTemp1 = Mid(strSourceName, f, 1)
        
        If strTemp1 = "\" Then g = g + 1
        
        If strTemp1 = "\" And g > 5 Then
           strFileName = Mid(strSourceName, f + 1, Len(strSourceName) - f) 'Nur Dateinamen
           inta = 1
           Exit For
        End If
      Next f
     'End If
     
     strFileName = Left(strFileName, Len(strFileName) - 4)
     
     'If inta > 0 Then
        CSVFileCounts = CSVFileCounts + 1
        
        frmdbfResData.Data2.DatabaseName = lastProjPath
        frmdbfResData.Data2.RecordSource = strFileName '"AUSWERT"
        
        frmdbfResData.Data2.Refresh
        frmdbfResData.Data2.UpdateRecord
        If ButtonCount = 1 Then
           CreateCSVFromCSV
        ElseIf ButtonCount = 2 Then
           CreateSQLFromDBF
        End If
        inta = 0
     End If
     frmdbview.Data1.Recordset.MoveNext
Next ini

End Sub

Sub CreateCSVFromCSV()
Dim Expcnt() As Integer
Dim ExpZaehl, ExpZaehlb As Integer

'*****************************************************************************************
'*** Read the dbf-File now:
'*** in this step, read the tool all rows in the schenck data - dbf file, make format
'*** for standard csv-file, by location.
'*** then create the new files in target directory add the information from AUSWERT.dbf
'*** and make form all this informations one csv-file with standard seperators and
'*** as you can import this files to anyone others (databases, excel and and ...)
'*****************************************************************************************

frmdbfResData.Data2.Recordset.MoveFirst
  
LineCount = 0
  
While Not frmdbfResData.Data2.Recordset.EOF
   LineCount = LineCount + 1
   frmdbfResData.Data2.Recordset.MoveNext
Wend

frmdbfResData.Data2.Recordset.MoveFirst

For i = 0 To frmdbfResData.Data2.Recordset.Fields.Count - 1
 
    init = i
 
    If init < frmdbfResData.Data2.Recordset.Fields.Count - 1 Then
          fieldnames(0, i) = frmdbfResData.Data2.Recordset.Fields(i).Name & lstTrennCountry
          If UCase(Left(frmdbfResData.Data2.Recordset.Fields(i).Name, 3)) = "EXP" Then
             ReDim Preserve Expcnt(a)
             Expcnt(a) = i
             a = a + 1
          End If
    Else
       fieldnames(0, i) = frmdbfResData.Data2.Recordset.Fields(i).Name
       Exit For
    End If
 
    frmdbfResData.Data2.Recordset.MoveNext

Next i

ExpZaehl = a - 1

writeline = ""

For i = 0 To init
    If UCase(Left(fieldnames(0, i), 3)) <> "EXP" Then
        If writeline = "" Then
            writeline = fieldnames(0, i)
        Else
            writeline = writeline & fieldnames(0, i)
        End If
    End If
Next i

If TestNr <> "" Then
   StrSoucreFile3 = DestroyPath & TestNr & ".csv"
ElseIf TestNr = "" Then
   StrSoucreFile3 = DestroyPath & "SCHDATA" & "_" & CSVFileCounts & ".csv"
End If

frmdbfResData.Data2.Recordset.MoveFirst

Open StrSoucreFile3 For Output As 1

'*****************************************************************************************
'*** Here add the columns descriptions, for information from AUSWERT.dbf
'*****************************************************************************************

If ColsRead = 0 Then

     writeline0 = lstTrennCountry & "VERSUCH" & _
                  lstTrennCountry & "PRUEFLING" & _
                  lstTrennCountry & "SCHLUESSEL" & _
                  lstTrennCountry & "PRUEFSTAND"

'*****************************************************************************************
'*** Here a specialy:
'*** If the checkbox checked=True then make this tool a csv-file for MySQL database
'*** LOAD FILE, with line end char = Chr 27 (ESC)
'*****************************************************************************************
    
    If frmdbview.Check1.Value = 1 Then
        Print #1, writeline & writeline0 & Chr(27)
    Else
        Print #1, writeline & writeline0
    End If

ElseIf ColsRead = 1 Then
    If Al00 <> 1 Then
       writeline = ""
       writeline0 = ""
       For i = 1 To ColsCnt
        If writeline = "" Then
           writeline = ColsNames(i) & lstTrennCountry
        Else
           writeline = writeline & ColsNames(i) & lstTrennCountry
        End If
       Next
    End If
    
    If frmdbview.Check1.Value = 1 Then
        Print #1, writeline & Chr(27)
        Al00 = 1
    Else
        Print #1, writeline
        Al00 = 1
    End If

End If
    
    writeline = ""
    writeline0 = ""
    ExpZaehlb = 0
    
For a = 1 To LineCount '- 1
    
    For i = 0 To init - 1
        If UCase(Left(frmdbfResData.Data2.Recordset.Fields(i).Name, 3)) <> "EXP" _
           And UCase(Left(frmdbfResData.Data2.Recordset.Fields(i).Name, 3)) <> "DAT" Then
            If i = 0 Then
                fieldValue(0, i) = frmdbfResData.Data2.Recordset.Fields(i).Value & lstTrennCountry
            Else
                fieldValue(0, i) = frmdbfResData.Data2.Recordset.Fields(i).Value & lstTrennCountry
            End If
        End If
        If UCase(Left(frmdbfResData.Data2.Recordset.Fields(i).Name, 4)) = Left(DateStr, 4) Then
           If i = 0 Then
                fieldValue(0, i) = CDate(frmdbfResData.Data2.Recordset.Fields(i).Value) & lstTrennCountry
              Else
                fieldValue(0, i) = CDate(frmdbfResData.Data2.Recordset.Fields(i).Value) & lstTrennCountry
           End If
           Else
            If i = 0 Then
                fieldValue(0, i) = frmdbfResData.Data2.Recordset.Fields(i).Value & lstTrennCountry
            Else
                fieldValue(0, i) = frmdbfResData.Data2.Recordset.Fields(i).Value & lstTrennCountry
            End If
        End If
    Next i
    
    If UCase(Left(frmdbfResData.Data2.Recordset.Fields(i).Name, 3)) <> "EXP" Then
       fieldValue(0, i) = frmdbfResData.Data2.Recordset.Fields(i).Value
    End If

    For i = 0 To init ' - 1
        If UCase(Left(frmdbfResData.Data2.Recordset.Fields(i).Name, 3)) <> "EXP" Then
           If i > 0 Then
              writeline = writeline & fieldValue(0, i)
              Else
              writeline = fieldValue(0, i)
           End If
        End If
    Next i
    
    writeline0 = lstTrennCountry & Versuch & _
                 lstTrennCountry & Pruefling & _
                 lstTrennCountry & TestNr & _
                 lstTrennCountry & Pruefstand
                  
    If frmdbview.Check1.Value = 1 Then
       Print #1, writeline & writeline0 & Chr(27)
    Else
       Print #1, writeline & writeline0
    End If
    
    writeline = ""
    writeline0 = ""
    
    frmdbfResData.Data2.Recordset.MoveNext
Next a

Close #1

If ChangeSeps = 1 Then ReplacedCharaters

End Sub

Sub ReplacedCharaters()
Dim sZeilen3() As String
Dim i As Long

If TestNr <> "" Then
   StrSoucreFile3 = DestroyPath & TestNr & ".csv"
ElseIf TestNr = "" Then
   StrSoucreFile3 = DestroyPath & "SCHDATA" & "_" & CSVFileCounts & ".csv"
End If

sOldDecimal = ","
sNewDecimal = "."
sOldSeparator = ";"
sNewSeparator = ","

Open StrSoucreFile3 For Input As 1
a = 0
 While Not EOF(1)
       ReDim Preserve sZeilen3(a) As String
       Line Input #1, sZeilen3(a)
       sZeilen3(a) = Replace(sZeilen3(a), sOldDecimal, sNewDecimal)
       sZeilen3(a) = Replace(sZeilen3(a), sOldSeparator, sNewSeparator)
       a = a + 1
Wend
       
Close #1

i = a - 1
         
Open StrSoucreFile3 For Output As 1
        
For a = 0 To i
   Print #1, sZeilen3(a)
Next
        
Close #1

End Sub

Sub CreateSQLFromDBF()

'********************************************************************************************
'*** Here create this tool, a *.sql file from Schenck - dbf file !
'*** This can import to MySQL database with webinterface or other,
'*** in moment without check "TABLE IF EXITS" !
'********************************************************************************************

frmdbfResData.Data2.Recordset.MoveFirst
  
LineCount = 0
  
While Not frmdbfResData.Data2.Recordset.EOF
   LineCount = LineCount + 1
   frmdbfResData.Data2.Recordset.MoveNext
Wend

frmdbfResData.Data2.Recordset.MoveFirst

If TestNr <> "" Then
   StrSoucreFile3 = DestroyPath & TestNr & ".sql"
ElseIf TestNr = "" Then
   StrSoucreFile3 = DestroyPath & "SCHDATA" & CSVFileCounts & ".sql"
End If

Open StrSoucreFile3 For Output As 1

fieldnames(0, 0) = "CREATE TABLE " & Versuch & " ("
Print #1, fieldnames(0, 0)

For i = 1 To frmdbfResData.Data2.Recordset.Fields.Count - 1
 init = i
    If init < frmdbfResData.Data2.Recordset.Fields.Count - 1 Then
        fieldnames(0, i) = "  " & frmdbfResData.Data2.Recordset.Fields(i).Name & " varchar(255) default NULL,"
        fieldnames(1, i) = frmdbfResData.Data2.Recordset.Fields(i).Name
        Print #1, fieldnames(0, i)
    Else
        fieldnames(0, i) = "  " & frmdbfResData.Data2.Recordset.Fields(i).Name & " varchar(255) default NULL,"
        fieldnames(1, i) = frmdbfResData.Data2.Recordset.Fields(i).Name
        Print #1, fieldnames(0, i)
        Exit For
    End If
 
    frmdbfResData.Data2.Recordset.MoveNext

Next i

For i = 1 To init
    If fieldnames0 = "" Then
        fieldnames0 = fieldnames(1, i) & ", "
    Else
        fieldnames0 = fieldnames0 & fieldnames(1, i) & ", "
    End If
Next

fieldnames0 = fieldnames0 & "VERSUCH, PRUEFLING, SCHLUESSEL, PRUEFSTAND"

fieldnames0 = "INSERT INTO " & Versuch & " (" & fieldnames0 & ") VALUES ("

writeline = ""

For i = 1 To init
    If writeline = "" Then
       writeline = fieldnames(0, i)
    Else
       writeline = writeline & fieldnames(0, i)
    End If
Next i

 frmdbfResData.Data2.Recordset.MoveFirst

     'writeline0 =
     Print #1, "  VERSUCH" & " varchar(255) default NULL,"
     Print #1, "  PRUEFLING" & " varchar(255) default NULL,"
     Print #1, "  SCHLUESSEL" & " varchar(255) default NULL,"
     Print #1, "  PRUEFSTAND" & " varchar(255) default NULL"
     Print #1, ") TYPE=MyISAM;"
     Print #1, ""

    writeline = ""
    writeline0 = ""
    
For a = 1 To LineCount '-1
    For i = 1 To init - 1
        If i = 1 Then
            fieldValue(0, i) = "('" & frmdbfResData.Data2.Recordset.Fields(i).Value & "'"
            fieldnames1 = "'" & frmdbfResData.Data2.Recordset.Fields(i).Value & "',"
        Else
            fieldValue(0, i) = "'" & frmdbfResData.Data2.Recordset.Fields(i).Value & "'"
            fieldnames1 = fieldnames1 & "'" & frmdbfResData.Data2.Recordset.Fields(i).Value & "',"
        End If
    Next i
    
    fieldnames1 = fieldnames1 & "'" & frmdbfResData.Data2.Recordset.Fields(i).Value & "',"
    fieldValue(0, i) = frmdbfResData.Data2.Recordset.Fields(i).Value & "',"
    
    For i = 1 To init
           If i > 1 Then
              writeline = writeline & fieldValue(0, i)
              Else
              writeline = fieldValue(0, i)
           End If
    Next i
    
    writeline = fieldnames1
    
    writeline0 = "'" & Versuch & "'," & _
                  "'" & Pruefling & "'," & _
                  "'" & TestNr & "'," & _
                  "'" & Pruefstand & "');"
                  
    If frmdbview.Check1.Value = 1 Then
       Print #1, writeline & writeline0 & Chr(27)
    Else
       Print #1, fieldnames0 & writeline & writeline0
    End If
    
    writeline = ""
    writeline0 = ""
    
    frmdbfResData.Data2.Recordset.MoveNext
Next a

Close #1

End Sub

'*************************************************************************************
'*** Here copy any tool the both dbf-files to target directory:
'*** copy AUSWERT.dbf and the Schenck-data-dbf-files in one directory
'*************************************************************************************

Sub CopyFiles()
   
   For i = 1 To Len(strFileName) - 1
       tmpstr = Mid(strFileName, i, 1)
       If tmpstr = "." Then
          tmpstr0 = Left(strFileName, i - 1)
       Exit For
       End If
   Next i
   
   strTemp000 = strFileName
   
   If TNr <> "" Then
      strFileName = TNr & tmpdbkmpl
   ElseIf TNr = "" Then
      strFileName = tmpstr0 & tmpdbkmpl
   End If
      
   FileCopy strSourceName, DestroyPath & strFileName
   
   frmdbview.ProgressBar1.Value = 50
   
   'Create Auswert.dbf on this directory
   
   frmdbview.Data1.DatabaseName = App.Path 'tmpdbPath '
   frmdbview.Data1.RecordSource = tmpdb
   
   frmdbview.Data1.Refresh
   frmdbview.Data1.UpdateRecord
   
   FileCopy lastProjPath & "\" & tmpdb & tmpdbkmpl, DestroyPath & tmpdb & "_" & TNr & tmpdbkmpl
   
   frmdbview.ProgressBar1.Value = 100
   
   frmdbview.Data1.DatabaseName = lastProjPath
   frmdbview.Data1.RecordSource = tmpdb
   frmdbview.Data1.Refresh
   frmdbview.Data1.UpdateRecord
   
   strFileName = strTemp000
   
End Sub

Sub CreateInToMySQLDB()
Dim conn As New ADODB.Connection
Dim rec As New ADODB.Recordset
'frmfrm_insertdb.Data1.RecordSource = Tablename


End Sub
