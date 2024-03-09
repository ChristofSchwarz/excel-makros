Option Explicit

Const sucheInAccount = "office@digirentalwien.at"  ' Outlook Account
Const sucheInFolder = "*\Posteingang"    ' Outlook Ordner
    
Const ShapeId_Draft = "Flussdiagramm: Verbinder 232"
Const ShapeId_Angeb = "Flussdiagramm: Verbinder 235"
Const ShapeId_Miete = "Flussdiagramm: Verbinder 236"
Const ShapeId_Rechn = "Flussdiagramm: Verbinder 237"
Const ShapeId_VKV = "Rectangle 1"
Const ShapeId_SL = "Rectangle 2"
Const ShapeId_USR = "Rectangle 3"
Const ShapeId_HDW = "Rectangle 4"
Const ShapeId_VKB = "Rectangle 5"
Const ShapeId_ZA = "Rectangle 6"


Dim global_AbsprungZelle

Const PreviewCalendarInAnfrage_FirstCol = "I"

' Angaben zur Inventarliste (InvLi) unten im "Buchung" Sheet
Const InvLi_StartZeile = 94
Const InvLi_Col_WebBezeichnung = "T"
Const InvLi_Col_SystemBezeichnung = "S"
'Const InvLi_Col_GeraeteId = "Q"
Const InvLi_Col_VirtuelleId = "P"

' Angaben zum Buchungen Excel (BuXls)

'Const BuXls_Datei = "C:\Users\chris\develop\digirental\Buchungen_2021\Buchungen_Video_2021.xlsm"
'Const BuXls_Datei = "Y:\3.00_VIDEO_VK_2021\Buchungen_2021\Buchungen_Video_2021.xlsm"
Const BuXls_Sheet = "ab_16_04_2021"
Const BuXls_StartRow = 34
Const BuXls_EndRow = 5000
Const BuXls_Col_VirtuelleId = "B"
Const BuXls_Col_GeraeteId = "D"
Const BuXls_Col_Bezeichn = "F"

' Farben
Const bgColorIndex_Header = 15 ' ColorIndex der Hintergrundfarbe der hinzuzufügenden Kopfzeilen
Const fontColorIndex_Header = 56
Const bgColorIndex_ChangeFound = 46
Const bgColorIndex_BookingConflict = 3


Function fn_UnmergeCellsAusser(Auswahl)
    Dim row, mergedRows, m, mergedContent
    Dim FirstRow, lastRow
    Dim AuswahlVon, AuswahlBis, mergedCellsCounter
    
    mergedCellsCounter = 0
    FirstRow = 1
    lastRow = Range("A1048576").End(xlUp).row
    ' Als erstes merged cells, die über Spalten (nebeneinander) gehen einfach trennen
    For row = FirstRow To lastRow
        If Range("A" & row).MergeCells Then
            mergedCellsCounter = mergedCellsCounter + 1
            If Range("A" & row).MergeArea.Columns.Count > 1 Then
                Range("A" & row).UnMerge
            End If
        End If
        If Range("A" & row).Value Like Auswahl Then
            mergedCellsCounter = mergedCellsCounter - 1
            AuswahlVon = row
            AuswahlBis = row + Range("A" & row).MergeArea.Rows.Count - 1
        End If
    Next row
    
    If mergedCellsCounter > 0 Then
    
        ' Jetzt alle Merged Cells, die über Zeilen gehen entfernen (außer Zelle "Auswahl"), dazu
        ' werden die mehrfach-Zeilen in "B" neben der gemergten Zelle "A" mit CHR(10) getrennt als
        ' ein Text zusammengefasst
        
        row = FirstRow
        
        While row <= lastRow
            If Range("A" & row).MergeCells Then
                mergedRows = Range("A" & row).MergeArea.Rows.Count
                If Range("A" & row) Like Auswahl Then
                    ' Auswahl-Zelle nicht mergen
                    row = row + mergedRows - 1
                ElseIf mergedRows > 1 Then
                    mergedContent = Range("B" & row).Text
                    ' Kann auch sein, dass auch B gemerged ist. Dann einfach unmergen
                    Range("B" & row).UnMerge
                    For m = row + 1 To row + mergedRows - 1
                        mergedContent = mergedContent & Chr(10) & Range("B" & m).Text
                        Range("B" & m).Clear
                    Next m
                    Range("A" & row).UnMerge
                    Range("B" & row).Value = mergedContent
                    ' Content nach oben nachrücken
                    ' Um eine Rückfrage zu vermeiden wird "Auswahl" vorrübergehend ge-unmerged und nach dem hochrücken wieder gemerged
                    Range("A" & AuswahlVon).UnMerge
                    Range("A" & (row + mergedRows) & ":B" & lastRow).Cut Destination:=Range("A" & (row + 1) & ":B" & (lastRow - mergedRows + 1))
                    lastRow = lastRow - mergedRows + 1
                    AuswahlVon = AuswahlVon - mergedRows + 1
                    AuswahlBis = AuswahlBis - mergedRows + 1
                    Range("A" & AuswahlVon & ":A" & AuswahlBis).Merge
                End If
            End If
            row = row + 1
        Wend
    End If
    
    fn_UnmergeCellsAusser = mergedCellsCounter
    
End Function


Sub Anfrage_BuchungErstellen()

' Christof Schwarz, 13-Nov-2016
'
' Dieses Makro befüllt in diesem Excel das Sheet "Buchungen" mit jenen Zeilen, die in
' Spalte G mit "x" markiert sind.

    Dim sh_Anfrage, sh_Buchung
    Dim Auswahl_StartZeile, Auswahl_EndZeile
    Dim row, sh_Bu_SourceRow
    Dim ErsteFreieZeile
    Dim Lookup_SystemId
    Dim Lookup_VirtualId
    Dim ret ' Retourwert der MsgBox Funktion
    Dim CountX
    
    If ActiveSheet.Name <> "Anfrage" Then fn_FatalError ("Du bist nicht auf dem 'Anfrage' Sheet")
    Set sh_Anfrage = ActiveSheet
    Set sh_Buchung = Sheets("Buchung")
    
    Auswahl_StartZeile = fn_GetAuswahlStartZeile 'fn_LookupArray(ActiveSheet, "A", 2, 99, "Auswahl*", , 2)
    Auswahl_EndZeile = fn_GetAuswahlEndZeile(Auswahl_StartZeile) 'fn_LookupArray(ActiveSheet, "A", 2, 99, "Bemerkung*", , 2) - 1

    CountX = WorksheetFunction.CountIf(ActiveSheet.Range("G" & Auswahl_StartZeile & ":G" & Auswahl_EndZeile), "x")
    
    
    If WorksheetFunction.CountIf(sh_Buchung.Range("B9:B79"), "") < 71 Then
        ret = MsgBox("Mietvertrag (Buchungs-Sheet) ist schon ausgefüllt. " & _
        Chr(10) & "Soll der Inhalt des Mietvertrags geleer werden?" _
        , vbOKCancel, CountX & " Eintragungen stehen an")
        If ret = vbCancel Then End
        If ret = vbOK Then sh_Buchung.Range("9:79").Clear

    End If
    
    
    For row = Auswahl_StartZeile To Auswahl_EndZeile
    
        If Range("G" & row).Value Like "x" Then
        
            Lookup_SystemId = Range("E" & row).Value
            Lookup_VirtualId = Range("D" & row).Value
            If Lookup_SystemId = "" Then fn_FatalError ("Keine SystemId in E" & row & " vorhanden!")
            
            ' Finde die passende Zeile im Buchung-Sheet entweder nach der SystemID in Spalte O
            ' oder nach der VirtualId in Spalte N
            
            sh_Bu_SourceRow = fn_LookupArray(sh_Buchung, "Q", InvLi_StartZeile, 10000, Lookup_SystemId)
            If sh_Bu_SourceRow = "#NV" Then
                sh_Bu_SourceRow = fn_LookupArray(sh_Buchung, "P", InvLi_StartZeile, 10000, Lookup_VirtualId, , 2)
            End If
            Call fn_AddRowToBuchung(sh_Bu_SourceRow)
               
            'MsgBox "Change in " & Row & ": " & InventoryRow(UBound(InventoryRow))
        End If
    Next row
        
    'If MsgBox("Einträge fertig. " _
    '& Chr(10) & "Sollen die " & CountX & " 'x'-Markierungen der Spalte 'Mietvertrag' nun entfernt werden?" _
    ', vbYesNo) = vbYes Then
    '    With sh_Anfrage.Range("G" & Auswahl_StartZeile & ":G" & Auswahl_EndZeile)
    '        .Clear
    '        .Interior.ColorIndex = bgColorIndex_Header
    '    End With
    'End If
    sh_Buchung.Activate
    sh_Buchung.Range("B8").Activate
End Sub


Sub Anfrage_FindeAenderungen()

    
'    Dim buExcelOpen ' Flag of Buchungs-Excel bereits ge_ffnet ist
    Dim sh_AnfrageCopy ' Arbeitsblatt, das eine Kopie von "Anfrage" darstellt, die zusammen mit der letzte Prüfung angelegt wurde
    Dim CompareRow ' passende Zeile im sh_AnfrageCopy (da wo die Spalte "F" übereinstimmt)
    Dim StartColNo, EndColNo
    Dim row, col, i, Rows
    Dim selections ' zusammengesetzter String, Liste aller Zellen im Anfrage Sheet, in denen —nderungen seit der Kopie gemacht wurden
'    Dim wb_BuXls  ' Excel Workbook "Buchungsfile"
'    Dim sh_BuXls  ' Excel Worksheet "Buchungsliste" im Buchungsfile
    Dim sh_Anfrage ' "Anfrage" Sheet
    Dim ZeilenIndex ' Zeilen-Nummer in Spalte E
    Dim buSh_CheckRow ' Zeilen-Nummer im Buchungs-Sheet (Buchungsfile)
    Dim buSh_CheckCol ' Spalten-ID im Buchungs-Sheet (Buchungsfile)
    Dim DisplayAlerts ' Flag, Setting in Excel, ob Warnungen angezeigt werden sollen
    Dim systemId
    Dim myRange       ' string, h_lt einen ganzen Zell-Bereich wenn eine —nderung gefunden wurde
    Dim Auswahl_StartZeile, Auswahl_EndZeile
    Dim SysId, doppelteSysId
    
    ' Los gehts ...
    
    If ActiveSheet.Name <> "Anfrage" Then fn_FatalError ("Du bist nicht auf dem 'Anfrage' Sheet")
      
    
'    buExcelOpen = False
    
    Set sh_Anfrage = ActiveSheet
    Set sh_AnfrageCopy = Sheets(ActiveSheet.Index + 1)
    If Not sh_AnfrageCopy.Name Like "Anfrage (*)" Then
        fn_FatalError ("Es gibt noch kein Vergleichs Sheet namens 'Anfrage (2)'. Zuerst Prüfen!")
    End If
    
    StartColNo = Columns(PreviewCalendarInAnfrage_FirstCol).Column + 1
    EndColNo = WorksheetFunction.Max(StartColNo, Range(PreviewCalendarInAnfrage_FirstCol & "12").End(xlToRight).Column - 1)
    'MsgBox fn_ColNo2Txt(StartColNo) & "-" & fn_ColNo2Txt(EndColNo)
    
    Auswahl_StartZeile = fn_GetAuswahlStartZeile()
    Auswahl_EndZeile = fn_GetAuswahlEndZeile(Auswahl_StartZeile)
    'startRow = fn_LookupArray(sh_Anfrage, "A", 40, 200, "Auswahl*", , 2)
    'Row = StartRow
    selections = ""
    SysId = Chr(8) ' Liste der verwendeten System Ids, mit Tab-Zeichen getrennt, beginnen
    doppelteSysId = Chr(8) ' Liste der doppelt verwendeten System Ids beginnen
    
    ' Hintergrundfarbe von F und G zurücksetzen auf Grau
    Range("F" & Auswahl_StartZeile & ":G" & Auswahl_EndZeile).Interior.ColorIndex = bgColorIndex_Header
    
    'While Not sh_Anfrage.Range("A" & Row).Value Like "Bemerkung*" And Row < Auswahl_EndZeile
    For row = Auswahl_StartZeile To Auswahl_EndZeile
    
        'ZeilenIndex = sh_Anfrage.Range("F" & Row).Value
        systemId = sh_Anfrage.Range("E" & row).Value
        CompareRow = fn_LookupArray(sh_AnfrageCopy, "E", Auswahl_StartZeile, 200, systemId, , 2)
        
        For col = StartColNo To EndColNo
            
            If sh_Anfrage.Cells(row, col).Value <> sh_AnfrageCopy.Cells(CompareRow, col).Value Then
            
                ' Hier wurde eine —nderung gefunden
                'Selections = Selections & fn_ColNo2Txt(Col) & Row
                myRange = fn_ColNo2Txt(StartColNo) & row
                If EndColNo <> StartColNo Then myRange = myRange & ":" & fn_ColNo2Txt(EndColNo) & row
                If Len(Replace(selections, myRange, "")) = Len(selections) Then
                    ' string Selections enthält noch nicht diesen Bereich, also füge ihn hinzu
                    If Len(selections) > 0 Then selections = selections & ","
                    selections = selections & myRange
                End If
                
                Range("F" & row).Value = "x"
                'sh_Anfrage.Range("F" & Row).Interior.ColorIndex = bgColorIndex_ChangeFound
                Range("G" & row).Value = "x"
            End If
        Next
        
        If Range("F" & row).Value = "x" Then
            If InStr(1, SysId, Chr(8) & Range("E" & row).Value & Chr(8)) > 0 Then
                doppelteSysId = doppelteSysId & Range("E" & row).Value & Chr(8)
            End If
            SysId = SysId & Range("E" & row).Value & Chr(8)
        End If
    Next
   
    ' Markiere SystemIds die doppelt vergeben wurden
    If Len(doppelteSysId) > 1 Then
        ' System Ids wurden doppelt verwendet
        doppelteSysId = Split(Mid(doppelteSysId, 2), Chr(8))
        For i = LBound(doppelteSysId) To UBound(doppelteSysId) - 1
            Rows = fn_LookupArrayAll(ActiveSheet, "E", Auswahl_StartZeile, Auswahl_EndZeile, doppelteSysId(i))
            Rows = Split(Rows, "|")
            For row = LBound(Rows) To UBound(Rows)
                Range("F" & Rows(row)).Interior.Color = 7444471 ' Rote markierung bei Konflikten
            Next row
        Next i
        MsgBox "System IDs " & Join(doppelteSysId, " ") & "wurden doppelt benutzt!", vbExclamation, "Achtung"
    End If
    
    If Len(selections) = 0 Then fn_FatalError ("Keine Änderungen gefunden.")
    Range(selections).Select
    
End Sub


Sub Anfrage_Hinzufuegen()

'###############################################################################################


' Prozedur funktioniert auf zwei Arten:
'  - entweder ist der Cursor gerade im Anfrage-Sheet, dann wechselt das Makro
'    auf das Buchungs-Sheet in die Ger_teliste
'  - oder man ist in der Geräteliste am Buchungs-Sheet, dann wird die aktuelle
'    Zeile im Anfrage-Sheet eingefügt, wo man zuletzt den Cursor hatte
    
' Änderung 17.01.2024: Kann auch in ganz leerer Anfrage +Hinzu fügen ...

    Dim sh_Anfrage ' Arbeitsblatt "Anfrage" in diesem Excel
    Dim sh_Buchung ' Arbeitsblatt "Buchung" in diesem Excel
    Dim SearchText As String ' Suchtext, vom Benutzer in Inputbox eingegeben
    Dim Spalte As String
    Dim gefundenZeile  ' Zeile im sh_Buchung, wo der Suchtext gefinden wurde
    Dim copySysText  ' Text in der Spalte InvLi_Col_SystemBezeichnung aus sh_Buchung, der in sh_Anfrage eingefügt werden soll
    Dim copyWebText  ' Text in der Spalte InvLi_Col_WebBezeichnung aus sh_Buchung, der in sh_Anfrage eingefügt werden soll
    Dim copyId  ' Text in der Spalte InvLi_Col_VirtuelleId aus sh_Buchung, der in sh_Anfrage eingefügt werden soll
    Dim ret  ' return value einer Messagebox (Yes, No, Cancel)
    Dim tmp ' temp variable
    Dim lastRow
    Dim row
    Dim ActiveRow, Zeile_darunter, Auswahl_ErsteZeile, Auswahl_LetzteZeile
    'Dim WertInSpalteA As String
    'Dim WertInSpalteA_darunter As String
    Dim PasteZeile
    
    Set sh_Anfrage = Sheets("Anfrage")
    Set sh_Buchung = Sheets("Buchung")

    Select Case ActiveSheet.Name
    
    Case "Anfrage"
        
        
        If ActiveCell.row < 100 Then  ' Benutzer ist im oberen Teil der Anfrage
           
            ' Prüfe ob du noch im Block stehst, der in Spalte A mit "Auswahl" beginnt

            If Not Range("A" & CStr(ActiveCell.row)).End(xlUp).Text Like "Auswahl*" _
            And Not Range("A" & CStr(ActiveCell.row)).Text Like "Auswahl*" Then
            
                fn_FatalError ("Du befindest dich in einem undefinierten Bereich des Arbeitsblattes. " _
                & Chr(10) & "Die Spalte A sollte 'Auswahl' stehen haben, dann bist du richtig.")
                
            End If
            
            
            Auswahl_ErsteZeile = fn_GetAuswahlStartZeile
            Auswahl_LetzteZeile = fn_GetAuswahlEndZeile(Auswahl_ErsteZeile)
    
           ' Finde erste Zeile von "Auswahl" Block im Anfrage Sheet
            'If Range("A" & CStr(ActiveCell.row)).Text Like "Auswahl*" Then
            '    Auswahl_ErsteZeile = ActiveCell.row
            'ElseIf Range("A" & CStr(ActiveCell.row)).End(xlUp).Text Like "Auswahl*" Then
            '    Auswahl_ErsteZeile = Range("A" & CStr(ActiveCell.row)).End(xlUp).row
            'End If
            
            ' Finde letzte Zeile im "Auswahl" Block
            'If Range("A" & CStr(ActiveCell.row)).Text Like "Auswahl*" _
            'And Range("A" & (ActiveCell.row + 1)).Text <> "" Then
            '    Auswahl_LetzteZeile = Auswahl_ErsteZeile
            'Else
            '    Auswahl_LetzteZeile = Range("A" & ActiveCell.row).End(xlDown).row - 1
            'End If
            
            
            
            While Range("B" & (ActiveCell.row + 1)).Text = "" And ActiveCell.row < Auswahl_LetzteZeile
                Range("B" & (ActiveCell.row + 1)).Select
            Wend
        
            ' Merk dir, wo sp_ter das neue Ger_t einzufügen ist
            
            global_AbsprungZelle = "$B$" & ActiveCell.row
            Range(global_AbsprungZelle).Select
            
            If sh_Buchung.Range(InvLi_Col_WebBezeichnung & InvLi_StartZeile).Text <> "bez_Website" Then
                fn_FatalError ("Inventarliste auf 'Buchung' Sheet nicht an erwarteter Stelle! (" & Chr(10) & _
                InvLi_Col_WebBezeichnung & InvLi_StartZeile & " sollte 'bez_Website' stehen haben.")
            End If
            
            If MsgBox("Position " & global_AbsprungZelle & " gemerkt." & Chr(10) _
                & "Du kannst jetzt mit [Strg]+[A] auf dem Buchung Sheet Zeilen hinzufügen." _
                , vbOKCancel) = vbCancel Then End
                
            
            ' Suchbegriff gefunden, springe in die Zeile auf dem Buchung Blatt
            sh_Buchung.Activate
            sh_Buchung.Range("B" & (InvLi_StartZeile + 1) & ":" & "B" & "10000").Select
            
            
        End If
        
    Case "Buchung"
        
        If global_AbsprungZelle = "" Then fn_FatalError ("Starte das Hochkopieren im Arbeitsblatt 'Anfrage'.")
        
        copySysText = Range(InvLi_Col_SystemBezeichnung & ActiveCell.row)
        copyWebText = Range(InvLi_Col_WebBezeichnung & ActiveCell.row)
        copyId = Range(InvLi_Col_VirtuelleId & ActiveCell.row)
        
        ret = MsgBox("'" & copySysText & "' wird zur Anfrage hinzugefügt. " & Chr(10) _
            & "Willst du danach noch weitere Zeilen einfügen?", vbYesNoCancel)
        
        If ret = vbYes Or ret = vbNo Then
          
            'Hochkopieren (also einfügen an der gemerkten Position)
            sh_Anfrage.Activate
            Range(global_AbsprungZelle).Select
            Application.CutCopyMode = False ' do not fill all cells of the inserted row with some content found in clipboard
            
            'Spezialfälle von Insert-Punkten
            'Zwecks einfacher Lesbarkeit holen wir uns ein paar Werte in Variablen
            Auswahl_ErsteZeile = fn_GetAuswahlStartZeile
            Auswahl_LetzteZeile = fn_GetAuswahlEndZeile(Auswahl_ErsteZeile)
            'Auswahl_ErsteZeile = ActiveCell.row
            'If Range("A" & Auswahl_ErsteZeile).Text = "" Then
            '    Auswahl_ErsteZeile = Range("A" & Auswahl_ErsteZeile).End(xlUp).row
            'End If
            'Auswahl_LetzteZeile = Auswahl_ErsteZeile
            'While Range("A" & (Auswahl_LetzteZeile + 1)).Text <> "Excel-Code I:" And Auswahl_LetzteZeile < 200
            '    Auswahl_LetzteZeile = Auswahl_LetzteZeile + 1
            'Wend
            If Not Range("A" & Auswahl_ErsteZeile).Text Like "Auswahl*" Then ' Or Range("A" & Auswahl_LetzteZeile + 1).Text <> "Excel-Code I:" Then
                MsgBox "Diese Cursor-Position ist nicht im Auswahl Bereich. Kann nicht fortsetzen"
                End
            End If
            'MsgBox "Auswahl Bereich ist von " & Auswahl_ErsteZeile & " bis " & Auswahl_LetzteZeile
    
            ActiveRow = ActiveCell.row
            Zeile_darunter = ActiveCell.row + 1
            'WertInSpalteA = Range("A" & ActiveCell.row).Text
            'WertInSpalteA_darunter = Range("A" & Zeile_darunter).Text
            
            
            If ActiveRow = Auswahl_ErsteZeile And ActiveRow = Auswahl_LetzteZeile And ActiveCell.Text = "" Then
                ' Situation 1: Auswahl ist einzeilig und ganz leer
                PasteZeile = ActiveRow
                            
            ElseIf ActiveRow = Auswahl_LetzteZeile Then ' And ActiveCell.Text <> ""
                'Situation 2: Einfügen am Ende
                Rows(Zeile_darunter & ":" & Zeile_darunter).Insert Shift:=xlDown
                Auswahl_LetzteZeile = Auswahl_LetzteZeile + 1
                Range("A" & Auswahl_ErsteZeile & ":A" & Auswahl_LetzteZeile).Merge
                PasteZeile = Zeile_darunter
                
            Else
                'Situation 3: Einfügen in der Mitte oder erste Zeile
                Rows(Zeile_darunter & ":" & Zeile_darunter).Insert Shift:=xlDown
                Auswahl_LetzteZeile = Auswahl_LetzteZeile + 1
                PasteZeile = Zeile_darunter
                
            End If
            
            
            Range("B" & PasteZeile).Select
            global_AbsprungZelle = "$B$" & PasteZeile
            ' Einfügen der Texte und Ids
            Range("B" & PasteZeile).Value = Trim(copyWebText)
            Range("B" & PasteZeile).Font.ColorIndex = 11
            Range("B" & PasteZeile).AddComment ("hinzugefügt " & CStr(Now) & " von " & Application.UserName)
            Range("B" & PasteZeile).Comment.Visible = False
            Range("C" & PasteZeile).Value = Trim(copySysText)
            Range("D" & PasteZeile).Value = copyId
            Range("D" & PasteZeile).NumberFormat = "General"
            
            Range(Auswahl_ErsteZeile & ":" & Auswahl_LetzteZeile).EntireRow.AutoFit
            
            
            If ret = vbYes Then
                sh_Buchung.Activate
            ElseIf ret = vbNo Then
                global_AbsprungZelle = ""
            End If
            
        End If
        
        ' Gehe zur zuvor gemerkten Zelle
        
    Case Else
        fn_FatalError ("Dieses Makro funktioniert nur am 'Anfrage' oder 'Buchung' Blatt.")
        
    End Select


End Sub



Sub Anfrage_MarkiereRelevanteZellen()


    Dim Auswahl_StartZeile, Auswahl_EndZeile
    Dim StartColNo, EndColNo
    Dim row, col, tmp
    Dim selections ' zusammengesetzter String, Liste aller Zellen im Anfrage Sheet, in denen —nderungen seit der Kopie gemacht wurden
    Dim WriteBackCells
    Dim sh_Anfrage ' "Anfrage" Sheet
    Dim suchString, suchString2, ret
    Dim myRanges(), myRanges_length
    Dim prevRange, thisRange, textVorschlag, textGefunden
    Dim appendToMyRanges ' Boolean ob die aktuelle Adresse zum Array myRanges als neuer Eintrag gemacht werden soll
    
    ' Los gehts ...
    myRanges_length = 0

    
    If ActiveSheet.Name <> "Anfrage" Then fn_FatalError ("Du bist nicht auf dem 'Anfrage' Sheet")
    Set sh_Anfrage = ActiveSheet
    
    suchString = Range("C4").Value
    suchString2 = Split(Range("C4").Value & "/", "/")
    suchString2 = suchString2(LBound(suchString2))
    
    'StartCol = PreviewCalendarInAnfrage_FirstCol
    StartColNo = Columns(PreviewCalendarInAnfrage_FirstCol).Column + 1
    EndColNo = WorksheetFunction.Max(StartColNo, Range(PreviewCalendarInAnfrage_FirstCol & "12").End(xlToRight).Column - 1)
    
    Auswahl_StartZeile = fn_GetAuswahlStartZeile
    Auswahl_EndZeile = fn_GetAuswahlEndZeile(Auswahl_StartZeile)
    'startRow = fn_LookupArray(sh_Anfrage, "A", 2, 20, "Auswahl*", , 2)
    'Row = startRow
    selections = ""
    WriteBackCells = ""
    
    ' Hintergrundfarbe von F und G zurücksetzen auf Grau
    Range("F" & Auswahl_StartZeile & ":G" & Auswahl_EndZeile).Interior.ColorIndex = bgColorIndex_Header
        
    'While Not sh_Anfrage.Range("A" & Row).Value Like "Bemerkung*" And Row < 200
    For row = Auswahl_StartZeile To Auswahl_EndZeile
       
        For col = StartColNo To EndColNo
            
            thisRange = fn_ColNo2Txt(col) & row
            
            If sh_Anfrage.Range(thisRange).Value Like "*" & suchString _
            Or sh_Anfrage.Range(thisRange).Value Like "*" & suchString2 _
            Or sh_Anfrage.Range(thisRange).Value Like "*" & suchString2 & "/*" Then
            
                ' Hier wurde eine —nderung gefunden
                
                If myRanges_length = 0 Then
                    appendToMyRanges = True
                Else
                    appendToMyRanges = False
                    'prevRange = ":" & myRanges(myRanges_length)
                    prevRange = Split(myRanges(myRanges_length), ":")
                    prevRange = prevRange(UBound(prevRange))
                    'MsgBox sh_Anfrage.Range(prevRange).Row
                    'MsgBox sh_Anfrage.Range(prevRange).Column
                
                    If sh_Anfrage.Range(prevRange).row = row And sh_Anfrage.Range(prevRange).Column = col - 1 Then
                        'Nachbarzelle der vorigen Zelle, mach einen Range daraus
                        If myRanges(myRanges_length) Like "*:*" Then
                            ' zelle hat schon ein einen Bereich, l_sche den Teil nach dem ":"
                            myRanges(myRanges_length) = Left(myRanges(myRanges_length), InStr(myRanges(myRanges_length), ":") - 1)
                        End If
                        myRanges(myRanges_length) = myRanges(myRanges_length) & ":" & thisRange
                    Else
                        ' start a next Range in myRanges array
                        appendToMyRanges = True
                    End If
                End If
                
                If appendToMyRanges Then
                    myRanges_length = myRanges_length + 1
                    ReDim Preserve myRanges(1 To myRanges_length)
                    myRanges(myRanges_length) = thisRange
                End If
                
                'If Len(Selections) > 0 Then Selections = Selections & ","
                'Selections = Selections & fn_ColNo2Txt(Col) & Row
                
                If Len(WriteBackCells) = Len(Replace(WriteBackCells, "F" & row, "")) Then
                    If Len(WriteBackCells) > 0 Then WriteBackCells = WriteBackCells & ","
                    WriteBackCells = WriteBackCells & "F" & row
                End If
                
            End If
        Next col
        'Row = Row + 1
    'Wend
    Next row
    
    If myRanges_length = 0 Then
        fn_FatalError ("SuchString '" & suchString & "' wurde nicht gefunden.")
    Else
        selections = Join(myRanges, ",")
        Range(selections).Select
        
        ' Neu 06.Jun 2021: Prüfe ob sich vorgeschlagene Zeile (durch hinzufügen der Angebotsnummer) von den markierten Bereichen unterscheidet
        ' Nimm die erste Zelle aus der ersten gefundenen Range
        textGefunden = Split(myRanges(LBound(myRanges)), ":")
        textGefunden = Range(textGefunden(LBound(textGefunden))).Value
        textGefunden = Split(" " & textGefunden, " ")
        textGefunden = textGefunden(UBound(textGefunden)) ' Text nach letztem Leerzeichen
        If textGefunden <> suchString Then
            ret = MsgBox("Text in Zelle " & textGefunden & " ist nicht wie SuchBegriff " & suchString & ". Jetzt ersetzen?", vbInformation + vbYesNo)
            If ret = vbYes Then
                Selection.Replace What:=textGefunden, Replacement:=suchString, LookAt:=xlPart, SearchOrder:=xlByRows, _
                MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
                
                Call Anfrage_ZurueckschreibenBuchung
            End If
            
        End If
    
        
        'If MsgBox("WriteBack Flag setzen in markierten Zeilen?", vbYesNo) = vbYes Then
        '    Range(WriteBackCells).Value = "x"
        'End If
    End If
End Sub


Sub Anfrage_EmailPruefen()

    Dim table1, table2, kundenEmailTreffer, lookupEmail
    Dim wb_KuXls ' Workbook "Kundenliste", separates Excel
    Dim i, ScanCol, ret1, ret2
    Dim message
    
    If ActiveSheet.Name <> "Anfrage" Then fn_FatalError ("Du bist nicht auf dem 'Anfrage' Sheet")

    ' Kopiere die <table> ... html zelle auf Sheet Buchung als Text
    table1 = fn_LookupArray(Sheets("Anfrage"), "A", 2, 999, "Excel-Code I:", "B")
    table2 = fn_LookupArray(Sheets("Anfrage"), "A", 2, 999, "Excel-Code II:", "B")
    message = ""
    
    If table1 <> "#NV" And table1 <> "" Then
    
        Sheets("Buchung").Activate
        If Range("XFD2").End(xlToLeft).Column = 1 Then
            ret1 = vbYes
        Else
            ret1 = MsgBox("In Zeile 2 auf Sheet ""Buchung"" befinden sich schon Kundeninformationen." _
                & Chr(10) & Chr(10) & "Nochmals überschreiben?", vbYesNoCancel + vbQuestion, "Rückfrage")
            If ret1 = vbCancel Then End
        End If
        
        If ret1 = vbYes Then
            fn_clipboard (table1)
            Range("A2").Select
            ActiveSheet.Paste
            message = message & """Excel-Code I"" Zelle in Buchung-Sheet übertragen." & Chr(10)
        End If
        Sheets("Anfrage").Activate
        
    End If
    
    If table2 <> "#NV" And table2 <> "" Then
    
        Sheets("Buchung").Activate

        If Range("XFD4").End(xlToLeft).Column = 1 Then
            ret2 = vbYes
        Else
            ret2 = MsgBox("In Zeile 4 auf Sheet ""Buchung"" befinden sich schon Buchungsinformationen." _
                & Chr(10) & Chr(10) & "Nochmals überschreiben?", vbYesNoCancel + vbQuestion, "Rückfrage")
            If ret2 = vbCancel Then End
        End If
        
        If ret2 = vbYes Then
            fn_clipboard (table2)
            Range("A4").Select
            ActiveSheet.Paste
            ' setzte numerische Werte als richtiges Format (Datum, Uhrzeit ..)
            Range("R4").NumberFormat = "ddd/ dd.mm.yyyy;@"
            Range("S4").NumberFormat = "hh:mm;@"
            Range("T4").NumberFormat = "ddd/ dd.mm.yyyy;@"
            Range("U4").NumberFormat = "hh:mm;@"
            Range("V4").NumberFormat = "General"
            message = message & """Excel-Code II"" Zelle in Buchung-Sheet übertragen." & Chr(10)
        End If
        Sheets("Anfrage").Activate
        
    End If
    

    lookupEmail = fn_LookupArray(ActiveSheet, "A", 2, 999, "Meine E-Mail:", "B", 2)
    If Trim(lookupEmail) = "" Then
        MsgBox message & "Keine Email gefunden in Spalte B.", vbInformation
        Range("E7").Value = "Keine Email in Spalte B gefunden."
    Else
    
        ' Öffne die Kundenliste (separates Excel) und finde Email wenn möglich
        
        Set wb_KuXls = Workbooks.Open(Range("KundenlisteExcel").Value)
        kundenEmailTreffer = fn_LookupArrayAll(wb_KuXls.Sheets("Verleihkunden"), "G", 27, 20000, lookupEmail)
        If kundenEmailTreffer = "#NV" Then
            kundenEmailTreffer = fn_LookupArrayAll(wb_KuXls.Sheets("Verleihkunden"), "I", 27, 20000, lookupEmail)
            If kundenEmailTreffer = "#NV" Then
                kundenEmailTreffer = fn_LookupArrayAll(wb_KuXls.Sheets("Verleihkunden"), "AJ", 27, 20000, lookupEmail)
                If kundenEmailTreffer = "#NV" Then
                    kundenEmailTreffer = fn_LookupArrayAll(wb_KuXls.Sheets("Verleihkunden"), "AT", 27, 20000, lookupEmail)
                End If
            End If
        End If
        If kundenEmailTreffer = "#NV" Then
            wb_KuXls.Close
            'MsgBox message & "Email nicht gefunden im Kundenliste Excel. Wahrscheinlich Neukunde?", vbInformation
            Range("E7").Value = "Neukunde?"
        ElseIf kundenEmailTreffer Like "*|*" Then
            wb_KuXls.Close
            MsgBox message & "Email " & lookupEmail & " wurde mehrfach gefunden in Kundenliste Excel, Zeilen " & kundenEmailTreffer, vbCritical, "Fehler"
        Else
            i = "<table><tr>"
            For ScanCol = 1 To Columns("BL").Column
                i = i & "<td>" & Cells(kundenEmailTreffer, ScanCol).Text & "</td>"
            Next ScanCol
            i = i & "</tr></table>"
            'wb_KuXls.Sheets("Verleihkunden").Rows(kundenEmailTreffer & ":" & kundenEmailTreffer).Select
            'Selection.Copy
            'wb_KuXls.Close
            fn_clipboard (i)
            wb_KuXls.Close
            Sheets("Buchung").Select
            Range("A2").Select
            ActiveSheet.Paste
            Sheets("Anfrage").Select
            'MsgBox message & "Kunde gefunden im Kundenliste Excel, Zeile " & kundenEmailTreffer, vbInformation
            Range("E7").Value = "Kunde gefunden Zeile " & kundenEmailTreffer
        End If
    End If
        
End Sub


'###############################################################################################

Sub Anfrage_VerfuegbarkeitPruefen(Optional ByVal showMessage = True)
    
'###############################################################################################

' Christof Schwarz, 13.10.2016  03.11.2016  20.04.2021, 04.05.2021

    Dim sh_Anfrage ' Arbeitsblatt "Anfrage" in diesem Excel
    Dim sh_Buchung ' Arbeitsblatt "Buchung" in diesem Excel
    Dim wb_BuXls ' Workbook "Buchungsliste", separates Excel
    Dim sh_BuXls ' Arbeitsblatt auf dem die gültigen Buchungen erfolgen, ist im wb_BuXls
    Dim StartDatum, EndDatum, DrehStartDatum, DrehEndDatum ' Start- u. Enddatum laut Anfrage-Sheet
    Dim ScanCol ' Hilfsvar zum iterieren über Spalten
    Dim row ' Hilfsvar zum iterieren über Spalten
    Dim StartCol_BuXls, EndCol_BuXls  ' Erste/letzte Spalte des BuLi Sheets, das gefunden wurde
    Dim StartCol_Anfr, EndCol_Anfr  ' Erste/letzte Spalte des Anfrage Sheets, in das eingefügt wird
    Dim Auswahl_StartZeile, Auswahl_EndZeile ' Erste und letzte Zeile im Anfrage Sheet, das zum Bereich "Auswahl:" geh_rt (verbundene Zeilen in Spalte A)
    Dim webname_to_look_up As String ' Bezeichnung in Spalte "B" aus dem Anfrage Sheet
    Dim lookup_id_virtuell ' Gefundene id aus der Spalte InvLi_Col_VirtuelleId im sh_Buchung
    Dim lookup_rowno ' Gefundener sysname aus der Spalte InvLi_Col_VirtuelleId im sh_Buchung
    Dim isfirst ' Flag, ob es sich um die erste von evtl. mehreren Zeilen mit gleicher virtualId im sh_BuXls handelt
    Dim DisplayAlerts ' Flag, ob Excel application Warnungen anzeigen soll oder nicht
    Dim Row_BuXls ' Zeile im sh_BuXls, die gefunden wurde bzw. verarbeitet wird
    Dim DoppelteInZeile
    Dim tmpStr ' Tempor_re String Variable
    Dim Menge, MengeDoppelt, Geraet
    Dim ColOffset ' Column Offset (is colum number of PreviewCalendarInAnfrage_FirstCol - 1)
    Dim i, Spalte_Von, Spalte_Bis ' Spalte von/bis wo der Buchungstext Vorschlag steht (spalte 1)
    Dim SpaltenAnz ' Anzahl Zellen, di der Buchungstext Vorschlag überspannt
    Dim findThisVirtualId
    Dim BuXls_Datei, KundenXls
    Dim multiples

    
    ' Los gehts ...
    If ActiveSheet.Name <> "Anfrage" Then fn_FatalError ("Du bist nicht auf dem 'Anfrage' Sheet")
    BuXls_Datei = Replace(Range("BuchungenExcel").Value, ".\", ActiveWorkbook.Path & "\")
    KundenXls = Range("KundenlisteExcel").Value
    Set sh_Anfrage = ActiveSheet
    Set sh_Buchung = Sheets("Buchung")
    i = fn_UnmergeCellsAusser("Auswahl*") ' Gemergte Zellen auflösen, ausser die neben "Auswahl"
    
    Auswahl_StartZeile = fn_GetAuswahlStartZeile()
    Auswahl_EndZeile = fn_GetAuswahlEndZeile(Auswahl_StartZeile)
    

    StartDatum = Range("C10").Value
    DrehStartDatum = Range("D10").Value
    If Len(StartDatum) = 0 Then fn_FatalError ("StartDatum fehlt (C10)")
    EndDatum = Range("C12").Value
    DrehEndDatum = Range("D12").Value
    If Len(EndDatum) = 0 Then fn_FatalError ("EndDatum fehlt (C12)")
    If EndDatum < StartDatum Then fn_FatalError ("EndDatum in C10 ist vor dem StartDatum in C12")
    
    ColOffset = Range(PreviewCalendarInAnfrage_FirstCol & ":" & PreviewCalendarInAnfrage_FirstCol).Column - 1
    
    
    ' Mache Buchungstext Vorschlag (horizontal über Spalten mit Name des Anfragenden)
    If StartDatum = EndDatum Then
        ' Spezialfall: Abholung und Rückgabe am selben Tag
        Cells(10, ColOffset + 2).Value = "A " & Range("C11").Value & " " & Range("C4").Value _
            & " R " & Range("C13").Value & " " & Range("C4").Value
        SpaltenAnz = 1
        Spalte_Von = fn_ColNo2Txt(ColOffset + 2)
        Spalte_Bis = fn_ColNo2Txt(ColOffset + 2)
    Else
        For ScanCol = 1 To 2 * (1 + EndDatum - StartDatum)
            Select Case ScanCol
            Case 1, 2 * (1 + EndDatum - StartDatum): 'Do nothing
            Case 2:
                Cells(10, ColOffset + ScanCol).Value = "A " & Range("C11").Value & " " & Range("C4").Value
                Spalte_Von = ColOffset + ScanCol
            Case 2 * (1 + EndDatum - StartDatum) - 1:
                Cells(10, ColOffset + ScanCol).Value = "R " & Range("C13").Value & " " & Range("C4").Value
                Spalte_Bis = ColOffset + ScanCol
            Case Else
                Cells(10, ColOffset + ScanCol).Value = Range("C4").Value
            End Select
            Cells(10, ColOffset + ScanCol).Font.Bold = False
            ' markiere den bereich, wo der Dreh stattfindet, in fett
            If (StartDatum + ScanCol / 2) > DrehStartDatum And (StartDatum + ScanCol / 2) <= (DrehEndDatum + 1) Then
                Cells(10, ColOffset + ScanCol).Font.Bold = True
            End If
        Next
        SpaltenAnz = Spalte_Bis - Spalte_Von + 1
        Spalte_Von = fn_ColNo2Txt(Spalte_Von)
        Spalte_Bis = fn_ColNo2Txt(Spalte_Bis)
    End If
    
    
    ' Öffne das Buchungen-Workbook (Nicht zu verwechseln mit BuchungsSheet in diesem Excel!)
    Set wb_BuXls = Workbooks.Open(BuXls_Datei)
    ActiveWindow.WindowState = xlMinimized
    ThisWorkbook.Activate
    Set sh_BuXls = wb_BuXls.Sheets(BuXls_Sheet)
    
    ' Suche die Spalten im Buchungs-Excel (Kalender Einträge in Zeile 2)
    StartCol_BuXls = 0
    EndCol_BuXls = 0
    ScanCol = 5
    Do
        If sh_BuXls.Cells(2, ScanCol).Value = StartDatum Then StartCol_BuXls = ScanCol
        If sh_BuXls.Cells(2, ScanCol).Value = EndDatum Then EndCol_BuXls = ScanCol + 1
        ScanCol = ScanCol + 1
    Loop Until EndCol_BuXls > 0 Or ScanCol = 5000
    
    If StartCol_BuXls = 0 Then fn_FatalError ("StartDatum '" & StartDatum & "' nicht gefunden")
    If EndCol_BuXls = 0 Then fn_FatalError ("EndDatum '" & EndDatum & "' nicht gefunden")
 
    StartCol_Anfr = Range(PreviewCalendarInAnfrage_FirstCol & ":" & PreviewCalendarInAnfrage_FirstCol).Column
    EndCol_Anfr = StartCol_Anfr + EndCol_BuXls - StartCol_BuXls
    
    For ScanCol = 0 To EndCol_Anfr - StartCol_Anfr
        Cells(12, ScanCol + StartCol_Anfr).Value = fn_ColNo2Txt(ScanCol + StartCol_BuXls)
        Cells(12, ScanCol + StartCol_Anfr).Interior.ColorIndex = bgColorIndex_Header
        Cells(12, ScanCol + StartCol_Anfr).Font.ColorIndex = fontColorIndex_Header
    Next
    
    ' Clear all cells rechts von der letzten Adresse aus
    Range(fn_ColNo2Txt(EndCol_Anfr + 1) & "12:" & fn_ColNo2Txt(sh_Anfrage.Columns.Count) & "12").Clear
    
    ' Convert all ...Col Variables to Column Names (A-Z, AA-XFD)
    StartCol_BuXls = fn_ColNo2Txt(StartCol_BuXls + 0)
    EndCol_BuXls = fn_ColNo2Txt(EndCol_BuXls + 0)
    StartCol_Anfr = fn_ColNo2Txt(StartCol_Anfr + 0)
    EndCol_Anfr = fn_ColNo2Txt(EndCol_Anfr + 0)
    
    ' Copy/Paste Kalender Zeile aus Buchungsliste in das sh_Anfrage.
    sh_BuXls.Range(StartCol_BuXls & "2:" & EndCol_BuXls & "2").Copy
    sh_Anfrage.Paste Destination:=sh_Anfrage.Range(StartCol_Anfr & "13:" & EndCol_Anfr & "13")
    

    
    ' 1.Loop)
    ' Wir laufen durch alle Zeilen der "Auswahl" des Web-users. Wir dampfen Doppelte Zeile in eine einzelne
    ' Zeile ein und erhöhen dafür den Zähler in der Textspalte, z.B. "(2x) Canon EOS 5D Mark II"
    
    row = Auswahl_StartZeile
    While row <= Auswahl_EndZeile
    
        Do
            If Range("B" & row).Value = "" Then
                DoppelteInZeile = "#NV"
            Else
                Menge = fn_ReturnMenge(Range("B" & row).Value)
                Geraet = fn_ReturnGeraet(Range("B" & row).Value)
            
                ' Prüfe ob es weiter unten in der Liste doppelte gibt (syntax kann auch mit "(Nx) Gerätename" oder
                ' "(NNx) Gerätename" beginnen (N = 0..9)
                        
                DoppelteInZeile = fn_LookupArray(ActiveSheet, "B", row + 1, Auswahl_EndZeile, Geraet)
                If DoppelteInZeile = "#NV" Then DoppelteInZeile = fn_LookupArray(ActiveSheet, "B", row + 1, Auswahl_EndZeile, "(?x) " & Geraet)
                If DoppelteInZeile = "#NV" Then DoppelteInZeile = fn_LookupArray(ActiveSheet, "B", row + 1, Auswahl_EndZeile, "(??x) " & Geraet)
                
                If DoppelteInZeile <> "#NV" Then
                    ' Es gibt doppelte Geräte in der Wunschliste, diese müssen zusammengefasst werden in die erste
                    ' der doppelten Zeilen (Erhöhung der Menge), die zweite wird dann gelöscht.
                    'MsgBox Row & " doppelte in " & DoppelteInZeile
                    MengeDoppelt = fn_ReturnMenge(Range("B" & DoppelteInZeile).Value)
                    Range("B" & row).Value = "(" & (Menge + MengeDoppelt) & "x) " & Geraet
                    Range(DoppelteInZeile & ":" & DoppelteInZeile).Delete Shift:=xlUp
                    Auswahl_EndZeile = Auswahl_EndZeile - 1
                    While Range("B" & DoppelteInZeile).Value = "" And DoppelteInZeile <= Auswahl_EndZeile
                        ' Wenn noch Zeilen ohne Eintrag in "B" folgen, lösche diese auch mit
                        Range(DoppelteInZeile & ":" & DoppelteInZeile).Delete Shift:=xlUp
                        Auswahl_EndZeile = Auswahl_EndZeile - 1
                    Wend
                End If
            End If
        Loop Until DoppelteInZeile = "#NV"
        
        row = row + 1
    Wend
    
    ' 2.Loop)
    ' -------
    ' Wir laufen durch alle Zeilen der "Auswahl" des Web-users. Durch Einfügungen wird die Auswahl_EndZeile
    ' nach unten nachrutschen (daher eine While Schleife, keine For-Schleife)
    
    row = Auswahl_StartZeile
    While row <= Auswahl_EndZeile
        sh_Anfrage.Range("D" & row).Select
                    
        ' Beim wiederholten Prüfen sind VirtualId/SystemId bereits ausgefüllt, wird nicht
        ' nochmal auf sh_Buchung gesucht
        If Range("D" & row).Value = "" Then
        
            ' Zuerst finde die virtuelle Id heraus, dazu wird im Sheet sh_Buchung gesucht (in diesem Excel, nicht in BuLi Excel)
            
            webname_to_look_up = fn_ReturnGeraet(Range("B" & row).Value)
                     
            lookup_id_virtuell = fn_LookupArrayAll(sh_Buchung, InvLi_Col_WebBezeichnung _
                , InvLi_StartZeile + 1, 5000, webname_to_look_up, InvLi_Col_VirtuelleId, 1)  ' ///////////
                
            lookup_rowno = fn_LookupArrayAll(sh_Buchung, InvLi_Col_WebBezeichnung _
                , InvLi_StartZeile + 1, 5000, webname_to_look_up, , 0)
            
            Range("D" & row).Value = lookup_id_virtuell
            Range("D" & row).NumberFormat = "General"
            Range("C" & row).Clear
            
            ' Änderung April 2021: Mehrere Virtuelle IDs in pipe-separated Liste können gefunden werden (array)
            
            If lookup_id_virtuell <> "#NV" Then

                lookup_id_virtuell = Split(lookup_id_virtuell, "|") ' mache array
                lookup_rowno = Split(lookup_rowno, "|") ' mache array
                
                For i = LBound(lookup_id_virtuell) To UBound(lookup_id_virtuell)
                    If i > LBound(lookup_id_virtuell) Then
                        ' ab dem zweiten Array Element muss eine Zeile eingefügt werden
                        row = row + 1
                        Rows(row).Insert Shift:=xlShiftDown
                        Auswahl_EndZeile = Auswahl_EndZeile + 1
                        ' Wenn in der letzten Zeile eingefügt wird, muss die "merged cell" 'Auswahl' in Spalte A auch erweitert werden
                        If row = Auswahl_EndZeile Then Range("A" & Auswahl_StartZeile & ":A" & Auswahl_EndZeile).Merge
                        
                        Range("E" & row & ":ZZ" & row).Interior.Pattern = xlNone ' Clear BgColor
                        Range("E" & row & ":ZZ" & row).Font.ColorIndex = xlAutomatic ' Clear Font Color
                    End If
                    Range("D" & row).Value = lookup_id_virtuell(i)
                    Range("D" & row).NumberFormat = "General"
                    Range("C" & row).Formula = "=Buchung!$" & InvLi_Col_SystemBezeichnung & "$" & lookup_rowno(i)
                Next i
            End If
        End If
        row = row + 1
    Wend
   
    ' 3.Loop)
    ' Wir suchen die SystemId im BuchungsExcel (BuXls)
    
    row = Auswahl_StartZeile
    While row <= Auswahl_EndZeile
        sh_Anfrage.Range("E" & row).Select
        If Range("E" & row).Value = "" Then
        
            findThisVirtualId = Range("D" & row).Value
            Row_BuXls = fn_LookupArrayAll(sh_BuXls, BuXls_Col_VirtuelleId, _
                BuXls_StartRow, BuXls_EndRow, findThisVirtualId, , 1)
                
            If Row_BuXls <> "#NV" Then
                Row_BuXls = Split(Row_BuXls, "|") ' Mehrfach-treffer möglich, mit "|" verkettet
                
                For i = LBound(Row_BuXls) To UBound(Row_BuXls)
                    If i > LBound(Row_BuXls) Then
                        ' ab dem zweiten Array Element muss eine Zeile eingefügt werden
                        row = row + 1
                        Rows(row).Insert Shift:=xlShiftDown
                        Auswahl_EndZeile = Auswahl_EndZeile + 1
                        Range("D" & row).Value = findThisVirtualId
                        ' Wenn in der letzten Zeile eingefügt wird, muss die "merged cell" 'Auswahl'
                        ' in Spalte A auch erweitert werden
                        If row = Auswahl_EndZeile Then Range("A" & Auswahl_StartZeile & ":A" & Auswahl_EndZeile).Merge
                    End If
                    
                    ' Kopiere aus dem Buchungs-Excel
                    ' 1) Kopiere die VirtualId aus sh_BuXls
                    Range("E" & row).Value = sh_BuXls.Range(BuXls_Col_GeraeteId & Row_BuXls(i)).Value
                    Range("E" & row).NumberFormat = "General"
                    ' 2) Kopiere Geräte-Bezeichnung
                    sh_BuXls.Range(BuXls_Col_Bezeichn & Row_BuXls(i)).Copy
                    sh_Anfrage.Paste Destination:=sh_Anfrage.Range("H" & row)
                Next i
            End If
        End If
        row = row + 1
    Wend
    
    ' Loop 4
    ' Kopieren der aktuellen Kalendereinträge aus Buchungen-Excel
    For row = Auswahl_StartZeile To Auswahl_EndZeile
        
        ' Kopiere die Kalenderspalten
        Row_BuXls = fn_LookupArray(sh_BuXls, BuXls_Col_GeraeteId, BuXls_StartRow, BuXls_EndRow, Range("E" & row).Value, , 1)
        sh_BuXls.Range(StartCol_BuXls & Row_BuXls & ":" & EndCol_BuXls & Row_BuXls).Copy
        sh_Anfrage.Paste Destination:=sh_Anfrage.Range(StartCol_Anfr & row & ":" & EndCol_Anfr & row)
        
        ' Markiere Zeilen, in denen über den ganzen angefragten Zeitraum frei ist
        If WorksheetFunction.CountIf(Range(Spalte_Von & row & ":" & Spalte_Bis & row), "") = SpaltenAnz _
        And Range("H" & row).Interior.Color <> 5273700 Then  ' Das ist die "Hands-off grün" Farbe
            Range(Spalte_Von & row & ":" & Spalte_Bis & row).Interior.Color = 10092441
        End If
    Next row
    
    wb_BuXls.Close
    
    If fn_HasRangeDuplicates("E" & Auswahl_StartZeile & ":E" & Auswahl_EndZeile) Then
         For row = Auswahl_StartZeile To Auswahl_EndZeile
            multiples = fn_LookupArrayAll(ActiveSheet, "E", Auswahl_StartZeile, Auswahl_EndZeile, Range("E" & row).Value)
            If multiples Like "*|*" Then
                multiples = Split(multiples, "|")
                Range("E" & multiples(LBound(multiples))).Interior.Color = 65535
                Range("E" & multiples(UBound(multiples))).Interior.Color = 65535
            End If
         Next row
          MsgBox "Achtung, im Bereich E" & Auswahl_StartZeile & ":E" & Auswahl_EndZeile _
           & " kommen SystemIds mehrfach vor. Aufpassen, dass nicht die selbe SystemId doppelt verplant wird!" _
           , vbExclamation, "Achtung"
    End If
    
    ' Wenn Sheet "Anfrage" schon mal kopiert wurde, dann lösche es.
    If Sheets(sh_Anfrage.Index + 1).Name Like "Anfrage (*)" Then
        DisplayAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        Sheets(sh_Anfrage.Index + 1).Delete
        Application.DisplayAlerts = DisplayAlerts
    End If
    
    ' Mach eine Kopie von Sheet "Anfrage"
    sh_Anfrage.Copy After:=sh_Anfrage
    sh_Anfrage.Activate
    
    Range(Spalte_Von & "10:" & Spalte_Bis & "10").Select
    Selection.Copy
    ActiveWindow.ScrollColumn = 1
    If showMessage Then MsgBox "Fertig geprüft. Buchungszellen sind im Clipboard.", vbInformation
    
End Sub



Sub Anfrage_ZurueckschreibenBuchung()


    Dim Auswahl_StartZeile, Auswahl_EndZeile
    Dim Rows_with_x
    Dim wb_BuXls  ' Excel Workbook "Buchungsfile"
    Dim sh_BuXls  ' Excel Worksheet "Buchungsliste" im Buchungsfile
    Dim sh_Anfrage ' "Anfrage" Sheet
    Dim sh_AnfrageCopy
    Dim buSh_CheckCol, buSh_CheckRow ' Vergleichs-Spalte/Zeile im "Buchungen (2)" Sheet
    Dim buExcelOpen ' Flag of Buchungs-Excel bereits ge_ffnet ist
    'Dim StartCol ' Spaltenname der Spalte in der die KalenderSpalten beginnen
    Dim StartColNo, EndColNo ' Nummer der ersten bis letzten Kalenderspalte
    Dim systemId   ' Wert aus der Spalte "E" des Buchungssheets
    Dim CompareRow ' selbe Zeile in der "Anfrage (2)" Kopie
    Dim SourceRange ' Bereich (aus dem Buchungssheet) aus dem kopiert wird
    Dim TargetRange ' Bereich (im Ziel Excel "Buchungsexcel") in das hineinkopiert wird
    Dim row, col ' Integer variablen für For-Next Schleifen über Zeilen und Spalten
    Dim Summary ' Textzusammenfassung aller Copy/Paste Schritte
    Dim DisplayAlerts ' Flag, ob die Applikation Warnungen anzeigt
    Dim BuXls_Datei
    
    buExcelOpen = False
    
    If ActiveSheet.Name <> "Anfrage" Then fn_FatalError ("Du bist nicht auf dem 'Anfrage' Sheet")
    BuXls_Datei = Replace(Range("BuchungenExcel").Text, ".\", ActiveWorkbook.Path & "\")
    
    Set sh_Anfrage = ActiveSheet
    Set sh_AnfrageCopy = Sheets(ActiveSheet.Index + 1)
    If Not sh_AnfrageCopy.Name Like "Anfrage (*)" Then
        fn_FatalError ("Es gibt noch kein Vergleichs Sheet namens 'Anfrage (2)'. Zuerst Prüfen!")
    End If
    
    
    Auswahl_StartZeile = fn_GetAuswahlStartZeile
    Auswahl_EndZeile = fn_GetAuswahlEndZeile(Auswahl_StartZeile)
    If fn_writeBackDuplicateSystemId(Auswahl_StartZeile, Auswahl_EndZeile) <> "" Then End
    
    'StartCol = PreviewCalendarInAnfrage_FirstCol
    StartColNo = sh_Anfrage.Range(PreviewCalendarInAnfrage_FirstCol & "12").Column + 1
    EndColNo = WorksheetFunction.Max(StartColNo, Range(PreviewCalendarInAnfrage_FirstCol & "12").End(xlToRight).Column - 1)
    'EndColNo = sh_Anfrage.Range(PreviewCalendarInAnfrage_FirstCol & "12").End(xlToRight).Column - 1  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
    Rows_with_x = WorksheetFunction.CountIf(ActiveSheet.Range("F" & Auswahl_StartZeile & ":F" & Auswahl_EndZeile), "x")
    
    If Rows_with_x > 0 Then
        If MsgBox("Zurückschreiben von " & Rows_with_x & " markierten Zeilen ins Buchungs-Excel?", vbYesNo) = vbNo Then End
        
        
        ' Äffne das Buchungen-Workbook
        Set wb_BuXls = Workbooks.Open(BuXls_Datei)
        ActiveWindow.WindowState = xlMinimized
        ThisWorkbook.Activate
        Set sh_BuXls = wb_BuXls.Sheets(BuXls_Sheet)
        buExcelOpen = True
    
        For row = Auswahl_StartZeile To Auswahl_EndZeile
            If Range("F" & row).Value = "x" Then
        
                ' Hier wurde eine Zeile als "WriteBack" markiert
                    
                systemId = sh_Anfrage.Range("E" & row).Value
                CompareRow = fn_LookupArray(sh_AnfrageCopy, "E", Auswahl_StartZeile, 200, systemId, , 2)
                
                For col = StartColNo To EndColNo
                                            
                    ' æberprüfe ob das sh_AnfrageCopy noch am gleichen Stand wie das Master BuchungsExcel ist
                    ' Die zust_ndige Col wird aus der in Zeile 12 der Anfrage festgehaltenen Spaltennamen genommen
                    ' Die zust_ndige Row wird nachgeschlagen an der eindeutigen SystemId in Spalte B des Buchungsexcel
                    
                    buSh_CheckCol = sh_AnfrageCopy.Cells(12, col).Value
                    buSh_CheckRow = fn_LookupArray(sh_BuXls, "D", 18, 10000, systemId, , 2)
                    
                    If sh_BuXls.Range(buSh_CheckCol & buSh_CheckRow).Value <> sh_AnfrageCopy.Cells(CompareRow, col).Value Then
                        ' Buchungskonflikt: Der Stand des BuchungsExcel ist nicht mehr wie er war, als das Kopie-Sheet erstellt wurde
                        sh_Anfrage.Range("F" & row).Interior.ColorIndex = bgColorIndex_BookingConflict
                        wb_BuXls.Close
                        sh_Anfrage.Cells(row, col).Select
                        fn_FatalError ("Buchungs-Excel wurde inzwischen ver_ndert in Zelle " & buSh_CheckCol & buSh_CheckRow _
                        & Chr(10) & "Bitte erneut prüfen.")
    
                    End If
                    
                Next
                    
                ' Jetzt wird wirklich zurückgeschrieben (copy/paste) ins Buchungsexcel
                SourceRange = fn_ColNo2Txt(StartColNo) & row & ":" & fn_ColNo2Txt(EndColNo) & row
                TargetRange = sh_AnfrageCopy.Cells(12, StartColNo).Value & buSh_CheckRow & ":" & sh_AnfrageCopy.Cells(12, EndColNo).Value & buSh_CheckRow
                sh_Anfrage.Range(SourceRange).Copy
                sh_BuXls.Paste Destination:=sh_BuXls.Range(TargetRange)
                Summary = Summary & SourceRange & "->" & TargetRange & " , "
            End If
        Next
    Else
        fn_FatalError ("Keine Zeilen mit 'x' in Spalte F markiert")
    End If
    
    wb_BuXls.Save
    wb_BuXls.Close
    
    DisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    sh_AnfrageCopy.Delete
    Application.DisplayAlerts = DisplayAlerts
    
    MsgBox "Fertig zurückgeschrieben. ", vbInformation, "Fertig" ' & Chr(10) & Left(Summary, Len(Summary) - 2)
End Sub



Function fn_Reset(Auswahl_StartZeile, Auswahl_EndZeile)

    Dim DisplayAlerts
    Dim row
    
    ' Funktion wird von Sub Blatt_Zuruecksetzen und Anfrage_Zuruecksetzen aufgerufen

    If ActiveSheet.Name <> "Anfrage" Then fn_FatalError ("Du bist nicht auf dem 'Anfrage' Sheet")
    
    ' Wenn Sheet "Anfrage" schon mal kopiert wurde, dann lösch es.
    If Sheets(ActiveSheet.Index + 1).Name Like "Anfrage (*)" Then
        DisplayAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        Sheets(ActiveSheet.Index + 1).Delete
        Application.DisplayAlerts = DisplayAlerts
    End If
    
    
    Columns("F:" & fn_ColNo2Txt(ActiveSheet.Columns.Count)).Delete Shift:=xlToLeft
    Columns("F:" & fn_ColNo2Txt(ActiveSheet.Columns.Count)).ClearFormats
    Range("C1").Clear
    Range("C14:E1000").Clear
    Range("C14:E1000").ClearFormats
    Range("E7").Clear
    Range("E7").ClearFormats
        

   
' Tabellenüberschriften eintragen
    Range("C" & (Auswahl_StartZeile - 1)).Value = "SystemName"
    Range("D" & (Auswahl_StartZeile - 1)).Value = "VirtualId"
    Range("D" & Auswahl_StartZeile & ":E5000").NumberFormat = "General"
    Columns("D").ColumnWidth = 8
    
    Range("E" & (Auswahl_StartZeile - 1)).Value = "SystemId"
    Columns("E").ColumnWidth = 8
    
    Range("F" & (Auswahl_StartZeile - 1)).Value = "WriteBack"
    Range("F" & (Auswahl_StartZeile - 1)).Orientation = 90
    Columns("F").ColumnWidth = 4
    
    Range("G" & (Auswahl_StartZeile - 1)).Value = "Mietvertrag"
    Range("G" & (Auswahl_StartZeile - 1)).Orientation = 90
    Columns("G").ColumnWidth = 4
    
    Range("H" & (Auswahl_StartZeile - 1)).Value = "Bezeichnung"
    Columns("H").ColumnWidth = 45
    
    Range("C" & (Auswahl_StartZeile - 1) & ":H" & (Auswahl_StartZeile - 1)).Interior.ColorIndex = bgColorIndex_Header
    Range("C" & (Auswahl_StartZeile - 1) & ":H" & (Auswahl_StartZeile - 1)).Font.ColorIndex = fontColorIndex_Header
    Range("F" & (Auswahl_StartZeile) & ":G200").Interior.ColorIndex = bgColorIndex_Header
    Range("F" & (Auswahl_StartZeile) & ":G200").Font.ColorIndex = fontColorIndex_Header
    Range("F" & (Auswahl_StartZeile) & ":G200").HorizontalAlignment = xlCenter
    
End Function




Sub Anfrage_Zuruecksetzen()

    Dim Auswahl_StartZeile
    Dim Auswahl_EndZeile
    Dim row

    
    
    If ActiveSheet.Name <> "Anfrage" Then fn_FatalError ("Du bist nicht auf dem 'Anfrage' Sheet")
    If MsgBox("Willst du wirklich das Sheet zurücksetzen?", vbYesNo) = vbNo Then End
     
    
    Columns("F:" & fn_ColNo2Txt(ActiveSheet.Columns.Count)).Delete Shift:=xlToLeft
    Columns("F:" & fn_ColNo2Txt(ActiveSheet.Columns.Count)).ClearFormats
    Range("C1").Clear
    Range("C14:E1000").Clear
    Range("C14:E1000").ClearFormats
    Range("E7").Clear
    Range("E7").ClearFormats

    Auswahl_StartZeile = fn_GetAuswahlStartZeile
    Auswahl_EndZeile = fn_GetAuswahlEndZeile(Auswahl_StartZeile)
        
    row = Auswahl_StartZeile + 1
    While row <= Auswahl_EndZeile
        If Range("B" & row).Value = "" Then
            Rows(row).Delete Shift:=xlUp
            Auswahl_EndZeile = Auswahl_EndZeile - 1
            row = row - 1
        End If
        row = row + 1
    Wend
   
    Call fn_Reset(Auswahl_StartZeile, Auswahl_EndZeile)
   
End Sub



Sub Blatt_Zuruecksetzen()

' Wird normalerweise nur vom Entwickler gebraucht, um rasch das Anfrage Blatt
' auf leer zu setzen

    Dim Auswahl_StartZeile
    Dim Auswahl_EndZeile
    Dim row
    Dim DisplayAlerts
    
    
    If ActiveSheet.Name <> "Anfrage" Then fn_FatalError ("Du bist nicht auf dem 'Anfrage' Sheet")
    Auswahl_StartZeile = fn_GetAuswahlStartZeile
    Auswahl_EndZeile = fn_GetAuswahlEndZeile(Auswahl_StartZeile)
    
    If MsgBox("Willst du wirklich das Sheet zurücksetzen?" & Chr(10) _
    & "der Bereich B" & Auswahl_StartZeile & ":B" & Auswahl_EndZeile & " wird gelöscht.", vbYesNo) = vbNo Then End
    
    row = Auswahl_StartZeile + 1
    While row <= Auswahl_EndZeile
        
        Rows(row).Delete Shift:=xlUp
        Auswahl_EndZeile = Auswahl_EndZeile - 1
        
    Wend
    Range("B" & Auswahl_StartZeile).Value = ""
    Range("B" & Auswahl_StartZeile).ClearComments
    
    
    Call fn_Reset(Auswahl_StartZeile, Auswahl_EndZeile)
       
End Sub


Function fn_SetColors(ObjectID, FontColor, BackgroundColor, LineColor)

    ' Funktion setzt die Schrift-, Hintergrund- und Linienfarbe des Shapes ("button")
    ' mit der gegebenen ID
    
    Dim shape As Object
    'Set shape = ActiveSheet.Shapes(CInt(ObjectID))
    Set shape = ActiveSheet.Shapes(ObjectID)
    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = FontColor
    shape.Fill.ForeColor.RGB = BackgroundColor
    shape.Line.ForeColor.RGB = LineColor
    
End Function



Function fn_Save(filename As String, DeletePrevFile As Boolean)

    Dim NewName As String
    Dim OldName As String
    Dim BaseDir As String
    Dim IsSameFile As Boolean
    Dim chkNewFile
    Dim deleteInfo As String
    Dim errorcode
    
    BaseDir = Left(ActiveWorkbook.FullName, Len(ActiveWorkbook.FullName) - Len(ActiveWorkbook.Name))
    OldName = ActiveWorkbook.FullName
    
    NewName = filename
    ' Wenn NewName kein absoluten Verzeichnuspfad enth_lt, dann erweitere um das Verzeichnis des aktuellen Excel
    If Not NewName Like "?:*" And Not NewName Like "\\*" Then NewName = BaseDir & NewName
            
    ' Wenn NewName keine Dateierweiterung enth_lt, fÙr sie hinzu
    If Not LCase(NewName) Like "*.xlsm" Then NewName = NewName & ".xlsm"
    
    IsSameFile = (UCase(NewName) = UCase(OldName))
   
    deleteInfo = ""
    errorcode = 0
    If IsSameFile Then
        On Error Resume Next
        ActiveWorkbook.Save
        errorcode = Err()
        On Error GoTo 0
        If errorcode = 0 Then MsgBox "Gespeichert.", vbInformation, "Info"
    Else
        On Error Resume Next
        
        ActiveWorkbook.SaveAs (NewName)
        errorcode = Err()
        On Error GoTo 0
        If errorcode = 0 Then
            If DeletePrevFile Then
                ' L_sche File unter vorigem Namen
                ' PrÙfe sicherheitshalber, dass die neue Datei vorhanden ist, dann l_sche die alte
                chkNewFile = Dir(NewName)
                If ActiveWorkbook.Name = chkNewFile Then
                    Kill OldName
                    deleteInfo = Chr(10) & "und File '" & Mid(OldName, 1 + Len(Range("RootFolder").Value)) _
                    & "' wurde entfernt."
                End If
            End If
            MsgBox "Gespeichert unter '" & Mid(ActiveWorkbook.FullName, 1 + Len(Range("RootFolder").Value)) _
            & "'" & deleteInfo, vbInformation, "Info"
        End If
    End If
    
    If errorcode <> 0 Then
        MsgBox "Speichern als '" & NewName & "' hat nicht geklappt.", vbCritical, "Fehler"
    End If
    fn_Save = errorcode ' Retour-Wert
    
End Function


Sub ResetShapeColors():
    ChkIfOnSheet ("Buchung")
    If MsgBox("Zurücksetzen von allen Status Feldern und Button Farben?", vbOKCancel) = vbCancel Then End
    
    
    ' Reset der Button Farben und Einstellungen. Nur w_hrend Development notwendig
    ' Im Produktivbetrieb wÙrde sonst der Workflow nicht gew_hrleistet sein
    
    Call fn_SetColors(ShapeId_Draft, Range("BC5").Font.Color, Range("BC5").Interior.Color, RGB(120, 120, 120))
    Call fn_SetColors(ShapeId_Angeb, Range("BC6").Font.Color, Range("BC6").Interior.Color, RGB(120, 120, 120))
    Call fn_SetColors(ShapeId_Miete, Range("BC7").Font.Color, Range("BC7").Interior.Color, RGB(120, 120, 120))
    Call fn_SetColors(ShapeId_Rechn, Range("BC8").Font.Color, Range("BC8").Interior.Color, RGB(120, 120, 120))
    Call fn_SetColors(ShapeId_Rechn, Range("BC8").Font.Color, Range("BC8").Interior.Color, RGB(120, 120, 120))
    
    Call fn_SetColors(ShapeId_VKV, RGB(255, 255, 255), RGB(192, 192, 192), RGB(80, 80, 80))
    Call fn_SetColors(ShapeId_SL, RGB(255, 255, 255), RGB(192, 192, 192), RGB(80, 80, 80))
    Call fn_SetColors(ShapeId_USR, RGB(255, 255, 255), RGB(192, 192, 192), RGB(80, 80, 80))
    Call fn_SetColors(ShapeId_HDW, RGB(255, 255, 255), RGB(192, 192, 192), RGB(80, 80, 80))
    Call fn_SetColors(ShapeId_VKB, RGB(255, 255, 255), RGB(192, 192, 192), RGB(80, 80, 80))
    Call fn_SetColors(ShapeId_ZA, RGB(255, 255, 255), RGB(192, 192, 192), RGB(80, 80, 80))
    
    Range("RechnungsTyp") = ""
    Range("UID") = ""
    Range("DocStat") = "Vorlage"
    Range("DraftDat").Value = ""
    Range("AngebNr").Value = ""
    Range("AngebDat").Value = ""
    Range("MieteNr").Value = ""
    Range("MieteDat").Value = ""
    Range("ReNr").Value = ""
    Range("ReDatum").Value = ""
    Range("RootFolder").Value = ""
    Range("DocStat").Select

End Sub


Sub ChkIfOnSheet(expectedSheet)
    If ActiveSheet.Name <> expectedSheet Then
        fn_FatalError ("Du bist nicht am Buchungs-Sheet. Hier geht das Makro nicht.")
    End If
End Sub



Sub MigrateDocToStatus(NewStatus As String, IdCell As String, DateCell As String, FileCell As String, DeletePrevFile As Boolean _
, ShapeId As String, NewColorCell As String, OldColorCell As String)
    
    Dim StatBefore, DatBefore, IdBefore
    Dim SaveStatus
    Dim confirm
                
    '1) Status merken
    StatBefore = Range("DocStat").Value
    DatBefore = Range(DateCell).Value
    IdBefore = Range(IdCell).Value
    '2) Filename vorschlagen aus aktuellem Datum/Uhrzeit
    '2) Filename vorschlagen aus Id und aktueller Datum/Uhrzeit
    If Len(Range(DateCell).Value) = 0 Then Range(DateCell).Value = Now()
    If Len(Range(IdCell).Value) = 0 Then
        Range(IdCell).Value = "'" & fn_NextFreeId(NewStatus, Range("RootFolder") & Range(FileCell))
    End If
    confirm = MsgBox("Zu " & NewStatus & " migrieren und speichern unter " & Range(FileCell) & Chr(10) & "Ordner: " & Range("RootFolder"), vbYesNo)
    
    If confirm = vbYes Then
        '3) Neuen Workflow Status
        Call fn_SetColors(ShapeId, Range(NewColorCell).Font.Color, Range(NewColorCell).Interior.Color, RGB(120, 120, 120))
        Range("DocStat").Value = NewStatus
        If NewStatus = "Draft" Then
            ' Einmalig wird die Verzeichnis-Position der Vorlage als relatives Root gemerkt °°°
            ' Range("RootFolder") = Left(ActiveWorkbook.FullName, Len(ActiveWorkbook.FullName) - Len(ActiveWorkbook.Name))
            Range("UID").Value = fn_GenerateUniqueId()
            Range("ProgVersion").Value = ActiveWorkbook.Name
        End If
        SaveStatus = fn_Save(Range("RootFolder") & Range(FileCell), DeletePrevFile)
    End If
    
    If confirm = vbNo Or SaveStatus <> 0 Then
         '4) Rollback zu Status davor
         Range(DateCell).Value = DatBefore
         Range("DocStat").Value = StatBefore
         Range(IdCell).Value = IdBefore
         If NewStatus = "Draft" Then
            Range("RootFolder").Value = ""
            Range("UID").Value = ""
            Range("ProgVersion").Value = ""
         End If
         Call fn_SetColors(ShapeId, Range(OldColorCell).Font.Color, Range(OldColorCell).Interior.Color, RGB(120, 120, 120))
         End
     End If

End Sub


Sub ExtendedSaveLogic(FileCell As String, Optional keepCopy As Boolean = True)
    Dim confirm, SaveStatus
    
    If keepCopy Then
        SaveStatus = fn_Save(Range("RootFolder") & Range(FileCell), False)
    ElseIf Not UCase(ActiveWorkbook.FullName) Like UCase("*" & Range(FileCell).Value & "*") Then
        confirm = MsgBox("Filename hat sich geändert, wenn Sie jetzt speichern wird es '" & Range(FileCell).Value _
        & "'" & Chr(10) & "Soll das alte File behalten werden?" _
        & Chr(10) & "(Ja = Kopie)", vbYesNoCancel)
        If confirm = vbYes Then SaveStatus = fn_Save(Range("RootFolder") & Range(FileCell), False)
        If confirm = vbNo Then SaveStatus = fn_Save(Range("RootFolder") & Range(FileCell), True)
    Else
        SaveStatus = fn_Save(Range("RootFolder") & Range(FileCell), True)
    End If
    If SaveStatus <> 0 Then
        fn_FatalError ("Speichern unter '" & Range("RootFolder") & Range(FileCell) & "' schlug fehl.")
    End If

End Sub



Sub Click_Shape_Draft()
' ---------------------
    ChkIfOnSheet ("Buchung")
    
    If Range("RechnungsTyp") = "" Then fn_FatalError ("Rechnungstyp ist noch nicht gesetzt (VKV, HDW, USR ...)")
    If Range("RootFolder") = "" Then fn_FatalError ("Root Folder wurde nicht gesetzt. Das passiert zu allererst bei der Wahl des Rechnungstypes.")
         
    If Range("DocStat").Value = "Vorlage" Then    ' Konvertieren Vorlage -> Draft
    
        ' Unter neuem Filename speichern, voriges File (=Vorlage) wird natÙrlich behalten (False)
        Call MigrateDocToStatus("Draft", "DraftNr", "DraftDat", "DraftFile", False, ShapeId_Draft, "BB5", "BC5")

    ElseIf Range("DocStat").Value = "Draft" Then   ' Nur speichern

        ExtendedSaveLogic ("DraftFile")
        
    Else
        fn_FatalError ("Zu Status 'Draft' kann nur aus dem Status 'Vorlage' gewechselt werden.")
    End If
        
End Sub



Sub Click_Shape_Angeb()
' ---------------------
    Dim ret, ret2, nextAngebotNr, BuXls_Datei
    
    ChkIfOnSheet ("Buchung")
    BuXls_Datei = Replace(Range("BuchungenExcel").Text, ".\", ActiveWorkbook.Path & "\")
    If Range("RechnungsTyp") = "" Then fn_FatalError ("Rechnungstyp ist noch nicht gesetzt (VKV, HDW, USR ...)")
    If Range("RootFolder") = "" Then fn_FatalError ("Root Folder wurde nicht gesetzt. Das passiert zu allererst bei der Wahl des Rechnungstypes.")
    If Dir(BuXls_Datei) = "" Then fn_FatalError ("Datei '" & BuXls_Datei & "' nicht gefunden. Wechseln zu 'Angebot' nicht möglich.")
    Const msg = "Soll nun die Angebotsnr ins Buchungsexcel geschrieben werden?"
    ret2 = vbCancel
    
    If Range("DocStat").Value = "Draft" Then    ' Konvertieren Draft -> Angebot
    
        ' Unter neuem Filename speichern, voriges File wird nicht behalten (True)
        Call MigrateDocToStatus("Angebot", "AngebNr", "AngebDat", "AngebFile", True, ShapeId_Angeb, "BB6", "BC6")
        ret2 = vbYes
    ElseIf Range("DocStat").Value = "Angebot" Then ' Nur speichern
        nextAngebotNr = fn_NextFreeId("Angebot", Range("RootFolder") & Range("AngebFile"))
        ret = MsgBox("Als neues Angebot (" & nextAngebotNr & ") speichern?" & Chr(10) & Chr(10) _
            & " [Nein] speichert normal ...", vbYesNoCancel)
        If ret = vbCancel Then End
        If ret = vbYes Then
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            nextAngebotNr = fn_NextFreeId("Angebot", Range("RootFolder") & Range("AngebFile"))
            Range("AngebNr").Value = nextAngebotNr
            Call ExtendedSaveLogic("AngebFile", True)  ' True speichert ohne Rückfragen als neues File und behält das alte
            ret2 = MsgBox(msg, vbYesNo + vbQuestion)
        Else
            Call ExtendedSaveLogic("AngebFile")
            ret2 = MsgBox(msg, vbYesNo + vbQuestion)
        End If
    Else
        fn_FatalError ("Zu Status 'Angebot' kann nur aus dem Status 'Draft' gewechselt werden.")
    End If

    If ret2 = vbYes Then
        Sheets("Anfrage").Select
        Call Anfrage_VerfuegbarkeitPruefen(False)
        Call Anfrage_MarkiereRelevanteZellen
    End If
End Sub



Sub Click_Shape_Miete()
' ---------------------
    ChkIfOnSheet ("Buchung")
    If Range("RechnungsTyp") = "" Then fn_FatalError ("Rechnungstyp ist noch nicht gesetzt (VKV, HDW, USR ...)")
    If Range("RootFolder") = "" Then fn_FatalError ("Root Folder wurde nicht gesetzt. Das passiert zu allererst bei der Wahl des Rechnungstypes.")
             
    If Range("RechnungsTyp") <> "VKV" Then
        fn_FatalError ("Den Status Miete gibt es nur beim RechnungsTyp VKV, nicht bei " & Range("RechnungsTyp"))
        
    ElseIf Range("DocStat").Value = "Draft" Then    ' Konvertiere Draft -> Miete
    
        ' Unter neuem Filename speichern, voriges File wird nicht behalten (True)
        Call MigrateDocToStatus("Miete", "MieteNr", "MieteDat", "MieteFile", True, ShapeId_Miete, "BB7", "BC7")
                
    ElseIf Range("DocStat").Value = "Angebot" Then   ' Konvertiere Angebot -> Miete
    
        ' Unter neuem Filename speichern, voriges File wird aber behalten (False)
        Call MigrateDocToStatus("Miete", "MieteNr", "MieteDat", "MieteFile", False, ShapeId_Miete, "BB7", "BC7")
                        
    ElseIf Range("DocStat").Value = "Miete" Then
        ExtendedSaveLogic ("MieteFile")
        
    Else
        fn_FatalError ("Zu Status 'Miete' kann nur aus dem Status 'Draft' oder 'Angebot' gewechselt werden.")
    End If

End Sub


Sub Click_Shape_Rechn()
' ---------------------
    ChkIfOnSheet ("Buchung")
    If Range("RechnungsTyp") = "" Then fn_FatalError ("Rechnungstyp ist noch nicht gesetzt (VKV, HDW, USR ...)")
    If Range("RootFolder") = "" Then fn_FatalError ("Root Folder wurde nicht gesetzt. Das passiert zu allererst bei der Wahl des Rechnungstypes.")
     
    Dim DocStat, ReNr
    
    DocStat = Range("DocStat").Value
    
    If Range("RechnungsTyp").Value = "VKV" Then
    
        If DocStat = "Miete" Then    ' Konvertieren Miete -> Rechnung

            ' Unter neuem Filename speichern, voriges File wird nicht behalten (True)
            Call MigrateDocToStatus("Rechnung", "ReNr", "ReDatum", "RechnFile", True, ShapeId_Rechn, "BB8", "BC8")
        ElseIf DocStat = "Rechnung" Then
            ExtendedSaveLogic ("RechnFile")
        Else
            fn_FatalError ("Zu Status 'Rechnung' kann bei VKV nur aus dem Status 'Miete' gewechselt werden.")
        End If
    
    Else
        ' Neu seit Jan 2023 ... andere Rechnungstypen können direkt als Rechnung speichern
        
        If DocStat = "Draft" Then ' Draft -> Rechnung
            
            ReNr = fn_NextFreeId("Rechnung", Range("RootFolder") & Range("RechnFile"))
            Call MigrateDocToStatus("Rechnung", "ReNr", "ReDatum", "RechnFile", True, ShapeId_Rechn, "BB8", "BC8")
        
        ElseIf DocStat = "Angebot" Then ' Angebot -> Rechnung
        
                ' Unter neuem Filename speichern, voriges File wird aber behalten (False)
            Call MigrateDocToStatus("Rechnung", "ReNr", "ReDatum", "RechnFile", False, ShapeId_Rechn, "BB8", "BC8")
        
        Else
             fn_FatalError ("Zu Status 'Rechnung' kann aus Status '" & DocStat & "' nicht gewechselt werden.")
        End If
        
    End If
End Sub


Sub Click_StatusDetails()
    ChkIfOnSheet ("Buchung")
    Range("BB5:BE8").Select
End Sub


Sub Click_Zurueck()
    ChkIfOnSheet ("Buchung")
    Range("A3").Select
End Sub







Function fn_ConvertBase(ByVal d As Double, ByVal sNewBaseDigits As String) As String
    Dim S As String, tmp As Double, i As Integer, lastI As Integer
    Dim BaseSize As Integer
    BaseSize = Len(sNewBaseDigits)
    Do While Val(d) <> 0
        tmp = d
        i = 0
        Do While tmp >= BaseSize
            i = i + 1
            tmp = tmp / BaseSize
        Loop
        If i <> lastI - 1 And lastI <> 0 Then S = S & String(lastI - i - 1, Left(sNewBaseDigits, 1)) 'get the zero digits inside the number
        tmp = Int(tmp) 'truncate decimals
        S = S + Mid(sNewBaseDigits, tmp + 1, 1)
        d = d - tmp * (BaseSize ^ i)
        lastI = i
    Loop
    S = S & String(i, Left(sNewBaseDigits, 1)) 'get the zero digits at the end of the number
    fn_ConvertBase = S
End Function


Function fn_GenerateUniqueId() As String
    Const base36 = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim DaysSince2010 As Double
    Dim i, Today, SecondsSinceMidnight
    Dim Part1, Part2, Part3
    i = Now()
    Today = Int(i) - 1 * (Int(i) > i)
    SecondsSinceMidnight = Round((i - Today) * 24 * 3600)
    DaysSince2010 = Today - DateSerial(2010, 1, 1)
    Part1 = Right("000" & fn_ConvertBase(DaysSince2010, base36), 3)
    Part2 = Right("0000" & fn_ConvertBase(SecondsSinceMidnight, base36), 4)
    Part3 = Right("00" & fn_ConvertBase(Round(Rnd() * 36 * 36), base36), 2)
    fn_GenerateUniqueId = Part1 & Part2 & Part3
End Function


Function fn_LookupArray(SheetObj, _
    LookupColumn, _
    StartRow, EndRow, _
    SearchString, _
    Optional ReturnColumn = "#N/A", _
    Optional ErrHandling = 0)

' Funktion ist so eine Art "SVERWEIS()" für VBA,
' Sucht in der Spalte LookupColumn von StartRow bis EndRow nach "LIKE" Vergleich
' (also Wildcards wie * und ? sind erlaubt) nach dem SearchString (case-insensitive)
' Wenn es einen Treffer gibt, h_ngt der Rückgabewert davon ob, ob der
' Parameter ReturnColumn angegegben wurde. Fehlt dieser, wird die Treffer-Zeilennummer
' zurückgeliefert, wird er angeführt, wird der Wert aus der Spalte ReturnColumn der
' Trefferzeile zurückgeliefert.
' Gibt es bis zur EndRow keinen Treffer, dann wird "#NV" zurückgeliefert.
        
' ErrHandling ist optional und kann fehlen oder 0 sein (=kein Error Handling),
' kann 1 sein (=Warnung zeigen aber weiter machen), oder 2 (=Fehlermeldung und
' Abbruch der Verarbeitung)

    Dim row  ' Zeilenz_hler found
    Dim found As Boolean  ' Wahr/Falsch Flag ob Suchbegriff gefunden wurde
    Dim sheetInfo As String ' Info für den User wenn nichts gefunden wurde
    Dim ret ' Return value of a messagebox
    
    ' Return Value Default
    fn_LookupArray = "#NV"
    row = StartRow
    
    
    Do
        found = Trim(UCase(SheetObj.Range(LookupColumn & CStr(row)).Value)) Like Trim(UCase(SearchString))
        If Not found Then
            row = row + 1
        Else
            If ReturnColumn = "#N/A" Then
                fn_LookupArray = row
            Else
                fn_LookupArray = SheetObj.Range(ReturnColumn & CStr(row)).Value
            End If
        End If
    Loop Until row > EndRow Or found
    
    ' Error Handling
    If Not found And ErrHandling > 0 Then
        ' Wenn die Suche auf einem anderen als dem aktuellen Arbeitsblatt durchgeführt wurde, gib an, wo.
        If SheetObj.Name = ActiveSheet.Name And SheetObj.Parent.FullName = ThisWorkbook.FullName Then
            sheetInfo = ""
        Else
            sheetInfo = "File : '" & SheetObj.Parent.Name & "'" & Chr(10) _
            & "Sheet : '" & SheetObj.Name & "'" & Chr(10)
        End If
        ' Gib Fehlermeldung aus
        ret = MsgBox(sheetInfo & "Konnte Suchbegriff '" & SearchString & "' nicht im Bereich " & LookupColumn & StartRow _
        & ":" & LookupColumn & EndRow & " finden. ", vbOKCancel)
        If ErrHandling = 2 Or ret = vbCancel Then End
    End If
End Function





Function fn_LookupArrayAll(SheetObj, _
    LookupColumn, _
    StartRow, EndRow, _
    SearchString, _
    Optional ReturnColumn = "#N/A", _
    Optional ErrHandling = 0)

' Wie Funktion fn_LookupArray aber sucht nach weiteren Treffern und liefert ein Antwort PseudoArray mit "|" verkettet
' Ausserdem wird der Suchbegriff auch innerhalb einer Zelle als Teil eines ";"-verketteten Strings gesucht
' Wildcards erlaubt im Suchbegriff und Groß/Kleinschreibung ist egal

    Dim row  ' Zeilenz_hler found
    Dim found As Boolean '
    Dim hits As Integer ' Zähler, wie oft gefunden
    Dim sheetInfo As String ' Info für den User wenn nichts gefunden wurde
    Dim ret ' Return value of a messagebox
    Dim Arr, i
    Dim results, resultsArr
    
    ' Return Value Default
    fn_LookupArrayAll = "#NV"
    row = StartRow
    hits = 0
    results = ""
    Do
        found = False
        Arr = Split(SheetObj.Range(LookupColumn & CStr(row)).Value, ";")
        For i = LBound(Arr) To UBound(Arr)
            If UCase(Trim(Arr(i))) Like UCase(SearchString) Then
                found = True
                hits = hits + 1
            End If
        Next i
    
        If found Then
            If ReturnColumn = "#N/A" Then
                results = results & "|" & row
            Else
                results = results & "|" & SheetObj.Range(ReturnColumn & CStr(row)).Value
            End If
        End If
        row = row + 1
    Loop Until row > EndRow
    
    ' Error Handling
    If hits = 0 And ErrHandling > 0 Then
        ' Wenn die Suche auf einem anderen als dem aktuellen Arbeitsblatt durchgeführt wurde, gib an, wo.
        If SheetObj.Name = ActiveSheet.Name And SheetObj.Parent.FullName = ThisWorkbook.FullName Then
            sheetInfo = ""
        Else
            sheetInfo = "File : '" & SheetObj.Parent.Name & "'" & Chr(10) _
            & "Sheet : '" & SheetObj.Name & "'" & Chr(10)
        End If
        ' Gib Fehlermeldung aus
        ret = MsgBox(sheetInfo & "Konnte Suchbegriff '" & SearchString & "' nicht im Bereich " & LookupColumn & StartRow _
        & ":" & LookupColumn & EndRow & " finden. ", vbOKCancel)
        If ErrHandling = 2 Or ret = vbCancel Then End
    End If
    
    
    If hits > 0 Then
        resultsArr = Split(Mid(results, 2), "|")
        resultsArr = fn_ArrayRemoveDups(resultsArr)
        fn_LookupArrayAll = Join(resultsArr, "|")
    End If
End Function


Function fn_ArrayRemoveDups(MyArray As Variant) As Variant
 ' from https://www.automateexcel.com/vba/remove-duplicates-array
 
    Dim nFirst As Long, nLast As Long, i As Long
    Dim item As String
    
    Dim arrTemp() As String
    Dim Coll As New Collection

    'Get First and Last Array Positions
    nFirst = LBound(MyArray)
    nLast = UBound(MyArray)
    ReDim arrTemp(nFirst To nLast)

    'Convert Array to String
    For i = nFirst To nLast
        arrTemp(i) = CStr(MyArray(i))
    Next i
    
    'Populate Temporary Collection
    On Error Resume Next
    For i = nFirst To nLast
        Coll.Add arrTemp(i), arrTemp(i)
    Next i
    Err.Clear
    On Error GoTo 0

    'Resize Array
    nLast = Coll.Count + nFirst - 1
    ReDim arrTemp(nFirst To nLast)
    
    'Populate Array
    For i = nFirst To nLast
        arrTemp(i) = Coll(i - nFirst + 1)
    Next i
    
    'Output Array
    fn_ArrayRemoveDups = arrTemp

End Function

'###############################################################################################

Function fn_FatalError(msg)

'###############################################################################################
    
    MsgBox msg, 16, "Fatal Error"
    End

End Function



'###############################################################################################





Function fn_GetAuswahlStartZeile()
    ' in mehreren Subs brauchen wir die Definition, wo die aus der Webanfrage gesendete Auswahlbereich beginnt
    fn_GetAuswahlStartZeile = fn_LookupArray(ActiveSheet, "A", 2, 999, "Auswahl*", , 2)
End Function


Function fn_GetAuswahlEndZeile(Auswahl_StartZeile)
    ' in mehreren Subs brauchen wir die Definition, wo die aus der Webanfrage gesendete Auswahlbereich endet
    fn_GetAuswahlEndZeile = Range("A" & Auswahl_StartZeile).MergeArea.Rows.Count + Auswahl_StartZeile - 1
End Function



Function fn_HasRangeDuplicates(rangeAddr)
    ' Gibt True/False zurück, ob ein Bereich von mehreren Zeilen (nur 1 Spalte!) Duplikate enthält
    ' z.B. HasRangeDuplicates("E41:E58")
    ' Mehrere Leer-Zellen werden ignoriert, also nicht als Duplikate gesehen
    Dim x, hasDupl, myRange, StartRow, EndRow, col
    Set myRange = Range(rangeAddr)
    col = fn_ColNo2Txt(myRange.Column)
    StartRow = myRange.Rows(1).row
    EndRow = myRange.Rows(myRange.Rows.Count).row
    
   'find duplicate values in range ("B5:B10") using the For Loop
    fn_HasRangeDuplicates = False
    For x = StartRow To EndRow
        If Application.WorksheetFunction.CountIf(Range(rangeAddr), Range(col & x)) > 1 Then
            fn_HasRangeDuplicates = True
        End If
    Next x
    
End Function

Function fn_clipboard(StoreText As String) As String
'PURPOSE: Read/Write to Clipboard
'Source: ExcelHero.com (Daniel Ferry)
  Dim x As Variant
'Store as variant for 64-bit VBA support
  x = StoreText
'Create HTMLFile Object
  With CreateObject("htmlfile")
    With .parentWindow.clipboardData
        .setData "text", x
    End With
  End With
End Function


Function fn_ColNo2Txt(ColumnNo As Variant) As String

    fn_ColNo2Txt = Replace(Replace(ActiveSheet.Cells(1, CInt(ColumnNo)).Address, "$1", ""), "$", "")

End Function








Function fn_ReturnMenge(Value) As Integer

    ' Aus einer Zelle der Spalte B im Anfrage-Sheet gibt es Werte, die mit
    ' einer (Klammer) beginnen, in welcher die Menge steht, gefolgt vom Ger_t. Wenn
    ' Klammer steht, ist der Eintrag nur der Ger_te Name und die Menge ist 1.
    ' Diese Funktion liefert die Menge zurück

    If Value Like "(?x) *" And IsNumeric(Mid(Value, 2, 1)) Then
        fn_ReturnMenge = CInt(Mid(Value, 2, 1))
    ElseIf Value Like "(??x) *" And IsNumeric(Mid(Value, 2, 2)) Then
        fn_ReturnMenge = CInt(Mid(Value, 2, 2))
    Else
        fn_ReturnMenge = 1
    End If
End Function

Function fn_ReturnGeraet(Value) As String

    ' Aus einer Zelle der Spalte B im Anfrage-Sheet gibt es Werte, die mit
    ' einer (Klammer) beginnen, in welcher die Menge steht, gefolgt vom Ger_t. Wenn
    ' Klammer steht, ist der Eintrag nur der Ger_te Name. Diese Funktion retourniert
    ' den Ger_tenamen (ignoriert ggf. den Eintrag in der Klammer)

    If Value Like "(?x) *" And IsNumeric(Mid(Value, 2, 1)) Then
        fn_ReturnGeraet = Mid(Value, 6)
    ElseIf Value Like "(??x) *" And IsNumeric(Mid(Value, 2, 2)) Then
        fn_ReturnGeraet = Mid(Value, 7)
    Else
        fn_ReturnGeraet = Value
    End If

End Function




Function fn_AddRowToBuchung(SourceRow)

    ' Funktion arbeitet auf dem Sheet "Buchungen" in diesem Excel und fügt die angegebene Zeile SourceRow
    ' (muss größer als 67 sein) in den oberen Bereich des Sheets hinzu, wo die Buchungen des Mieters stehen
    ' Dabei wird unterschieden, ob es sich um Bulk-Ware (ohne Seriennummer) oder inventarisierte Miet-Geräte
    ' handelt. Letztere erhalten, wenn mehrere von der gleichen Sorte ausgeliehen werden, jeweils eine eigene
    ' Zeile in der Buchung (mit unterschiedlicher S/N), hingegen bei Bulk-Ware wird einfach nur der Zähler
    ' in Spalte A erhöht.
    ' Das Hinzufügen geschieht indem über das Clipboard die ganze SourceRow in die Zielzeile kopiert wird,
    ' also mit allen Angaben über Mietpreis-Staffelungen in den Spalten weiter rechts ...

    Dim SeparateRow As Boolean ' Flag entscheidet, ob beim Mietvertrag mehrfache Geräte als separate Zeile vorkommen (true) oder einfach der Zähler in Spalte A erhöht wird
    Dim TargetRow ' Zeile der Buchung (zwischen 9 und 79), in die eingefügt werden soll
    Dim Title ' Spalte B im Buchung-Sheet
    Dim sh_Bu
    Dim prevAmount ' Z_hler in Spalte A vor der Erh_hung bzw "" bei einer neuen
    
    Set sh_Bu = Sheets("Buchung")
    SeparateRow = (Len(Trim(sh_Bu.Range("R" & SourceRow).Value)) > 0)
    
    Title = sh_Bu.Range("B" & SourceRow).Value
    If Len(Title) = 0 Then fn_FatalError ("Fehler: Die verlinkte Zeile " & SourceRow & " im Buchung Sheet enthält keinen Titel in Spalte B")
    
    If SeparateRow Then
        TargetRow = fn_LookupArray(sh_Bu, "A", 9, 79, "", , 2)
        'MsgBox "(Separate) Insert row " & SourceRow & " (" & Title & ") as separate into row " & TargetRow
        
    Else
        TargetRow = fn_LookupArray(sh_Bu, "B", 9, 79, sh_Bu.Range("B" & SourceRow).Value)
        If TargetRow = "#NV" Then TargetRow = fn_LookupArray(sh_Bu, "A", 9, 79, "", , 2)
    End If
    
    prevAmount = sh_Bu.Range("A" & TargetRow).Value
    
    If prevAmount = "" Then
        sh_Bu.Range(SourceRow & ":" & SourceRow).Copy
        sh_Bu.Paste Destination:=sh_Bu.Range(TargetRow & ":" & TargetRow)
    Else
        sh_Bu.Range("A" & TargetRow).Value = prevAmount + 1
    End If
    
End Function







Sub Mietvertrag()  '

    Const FirstEquipmRow = 96
    Const LastEquipmRow = 2000
    Const FirstContractRow = 9
    Const LastContractRow = 79
    Const CellIdentifyContract = "C8"
    Const CellIdentifyContractText = "Buchung"
    Dim SearchStr As String
    Dim found As Integer
    Dim nRow As Integer
    Dim OldCellAddr As String
    Dim cellAddr

    If Range(CellIdentifyContract).Text = CellIdentifyContractText Then
        OldCellAddr = ActiveCell.Address
        If ActiveCell.row >= FirstEquipmRow Then
            Range(CStr(ActiveCell.row) & ":" & CStr(ActiveCell.row)).Select
            Selection.Copy
            nRow = FirstContractRow
            Do While Range("A" & CStr(nRow)).Text <> "" And nRow <= LastContractRow
                nRow = nRow + 1
            Loop
            If nRow > LastContractRow Then
                MsgBox "Keine Zeile mehr frei im Mietvertrag", vbCritical, "Mietvertrag Makro"
                End
            Else
                Range(CStr(nRow) & ":" & CStr(nRow)).Select
                ActiveSheet.Paste
                'ActiveWindow.ScrollRow = nRow - 7
                If MsgBox("Kopiert in Zeile " & CStr(nRow) & Chr(10) & "Wenn das die letzte Zeile war, click auf Abbrechen", vbOKCancel) = vbOK Then
                    Range(OldCellAddr).Select
                Else
                    'ActiveWindow.ScrollRow = 1
                End If
            End If
    
        ElseIf ActiveCell.Column = 1 Then
    
            SearchStr = ActiveCell.Text
            found = 0
            cellAddr = ""
            For nRow = FirstEquipmRow To LastEquipmRow
                If UCase(Range("A" & CStr(nRow)).Text) = UCase(SearchStr) Then
                    found = found + 1
                    cellAddr = cellAddr & "," & CStr(nRow) & ":" & CStr(nRow)
                End If
            Next
            If found > 0 Then
                Range(Mid(cellAddr, 2)).Select
                Application.CutCopyMode = False
                Selection.Copy
                Range(OldCellAddr).Select
                MsgBox CStr(found) & " Zeilen ins Clipboard kopiert"
                '    ActiveSheet.Paste
            Else
                MsgBox "Suchbegriff " & SearchStr & " nicht gefunden."
            End If
        Else
            MsgBox "Cursor an undefinierter Stelle", vbCritical, "Mietvertrag Makro"
        End If
    Else
        MsgBox "Ooops! Das ist kein Mietvertrag!", vbCritical, "Mietvertrag Makro"
    End If
End Sub

Function fn_NextFreeId(ForStatus As String, fileOrFolder As String) As String

    Dim StrFile As String
    Dim AktId As Double
    Dim LetzteId As Double
    Dim BaseDir
    Dim Arr() As String
    Dim searchPattern, startNumAt
    
    If ForStatus = "Draft" Then
        fn_NextFreeId = (Year(Now()) - 2000) & "xxxx"
        
    ElseIf ForStatus = "Angebot" Then
        ' Funktion sucht im gegebenen Folder nach dem h_chsten
        ' 6-stelligen numerischen Filenamen nach "AN??????_*" . Dann wird die n_chst-h_here Zahl
        ' zurÙckgeliefert, also wenn das File "AN180633_SamuelSaenz_04_06_18.xlsm"
        ' w_re, dann "AN180634"
        ' Es ist okay ein File in dem Folder als "fileOrFolder" zu Ùbergeben, die file-komponente wird
        ' ggf. entfernt. Ein Foldername muss mit "\" enden
        Arr = Split(fileOrFolder, "\")
        If UBound(Arr) <= 1 Then fn_FatalError ("Kann keine Angebotsnummer im Ordner '" & fileOrFolder & "' finden.")
    
        ' wenn es kein Folder ist (file) dann schneide das ab dem letzten "\" ab
        If (Right(fileOrFolder, 1) <> "\") Then
            ReDim Preserve Arr(UBound(Arr) - 1)
            BaseDir = Join(Arr, "\") & "\"
        Else
            BaseDir = fileOrFolder
        End If
        
        LetzteId = 0
        If Range("Rechnungstyp") = "VKV" Then
            searchPattern = "AN??????_*.xls*"
            startNumAt = 3
        Else
            searchPattern = Range("Rechnungstyp") & "_AN??????_*.xls*"
            startNumAt = Len(Range("Rechnungstyp").Text) + 4
        End If
        
        StrFile = Dir(BaseDir & searchPattern)
        Do While Len(StrFile) > 0
    
            On Error Resume Next
            AktId = CDbl(Mid(StrFile, startNumAt, 6))
            'MsgBox StrFile, , Err()
            If Err() = 0 And AktId > LetzteId Then LetzteId = AktId
            StrFile = Dir
        Loop
        On Error GoTo 0
        
        If Range("Rechnungstyp") = "VKV" Then
            fn_NextFreeId = "AN" & Right("00000" & CStr(LetzteId + 1), 6)
        Else
            fn_NextFreeId = Range("Rechnungstyp") & "_AN" & Right("00000" & CStr(LetzteId + 1), 6)
        End If
    
    ElseIf ForStatus = "Miete" Then
        fn_NextFreeId = Range("DraftNr").Value
    
    ElseIf ForStatus = "Rechnung" Then
    
        ' Funktion sucht im gegebenen Folder nach dem h_chsten
        ' 6-stelligen numerischen Filenamen nach "??????_*.xls*" . Dann wird die n_chst-h_here Zahl
        ' zurÙckgeliefert, also wenn das File "180633_SamuelSaenz_04_06_18.xlsm"
        ' w_re, dann "180634"
        ' Es ist okay ein File in dem Folder als "fileOrFolder" zu Ùbergeben, die file-komponente wird
        ' ggf. entfernt. Ein Foldername muss mit "\" enden
        Arr = Split(fileOrFolder, "\")
        If UBound(Arr) <= 1 Then fn_FatalError ("Kann keine Rechnungssnummer im Ordner '" & fileOrFolder & "' finden.")
    
        ' wenn es kein folder ist (file) dann schneide das ab dem letzten "\" ab
        If (Right(fileOrFolder, 1) <> "\") Then
            ReDim Preserve Arr(UBound(Arr) - 1)
            BaseDir = Join(Arr, "\") & "\"
        Else
            BaseDir = fileOrFolder
        End If
        
        LetzteId = 0
         If Range("Rechnungstyp") = "VKV" Then
            searchPattern = "??????_*.xls*"
            startNumAt = 1
        Else
            searchPattern = Range("Rechnungstyp") & "_??????_*.xls*"
            startNumAt = Len(Range("Rechnungstyp").Text) + 2
        End If
        
        StrFile = Dir(BaseDir & searchPattern)
        Do While Len(StrFile) > 0
    
            On Error Resume Next
            AktId = CDbl(Mid(StrFile, startNumAt, 6))
            'MsgBox StrFile, , Err()
            If Err() = 0 And AktId > LetzteId Then LetzteId = AktId
            StrFile = Dir
        Loop
        On Error GoTo 0
        
        If Range("Rechnungstyp") = "VKV" Then
            fn_NextFreeId = Right("00000" & CStr(LetzteId + 1), 6)
        Else
            fn_NextFreeId = Range("Rechnungstyp") & "_" & Right("00000" & CStr(LetzteId + 1), 6)
        End If
             
    
    Else
        fn_FatalError ("Für unbekannten Status '" & ForStatus & "' kann keine neue Id vergeben werden.")
    End If

End Function








Function fn_DateiAuslesen(Dateipfad As String, arg As String) As String
 
    Dim BaseDir As String
    Dim wholefile As String
    Dim arguments() As String
    Dim found As Integer
    Dim i As Integer
    Dim ret As String
    
    If Not Dateipfad Like "*\*" Then
        ' Wenn Dateipfad kein Verzeichnis ist, dann erweitere es um das Verzeichnis des aktuellen Excel
        Dateipfad = Left(ActiveWorkbook.FullName, Len(ActiveWorkbook.FullName) - Len(ActiveWorkbook.Name)) & Dateipfad
    End If
    On Error Resume Next
    
    wholefile = CreateObject("Scripting.FileSystemObject").OpenTextFile(Dateipfad).ReadAll
    If Err() <> 0 Then fn_FatalError ("Fehler " & Err() & " beim Lesen von '" & Dateipfad & "'")
    
    ' Filesyntax: $$variable := "eine gÙltige Excelformel"
    ' Ganzes File landet in einem String, wird nun zu Array gesplittet
    wholefile = Replace(wholefile, "$$", ":=$$")
    arguments = Split(wholefile, ":=")
    
    found = -1
    ' Suche im Array nach dem key "$$[variable]" und liefere den Index zurÙck
    For i = LBound(arguments) To UBound(arguments)
        If Trim(arguments(i)) Like arg Then found = i + 1
    Next
    If found < 0 Then fn_FatalError ("Datei '" & Dateipfad & "' enthielt nicht variable " & arg)
    ret = Trim(arguments(found))
    ' Schneide eventuell CR+LF am ende des Values weg
    If Right(ret, 1) = Chr(10) Then ret = Left(ret, Len(ret) - 1)
    If Right(ret, 1) = Chr(13) Then ret = Left(ret, Len(ret) - 1)
    ' Interpretiere die Formel auf Sheet "Buchung"
    ret = ActiveWorkbook.Sheets("Buchung").Evaluate(ret)
    If TypeName(ret) Like "Error*" Then fn_FatalError ("Variable " & arg & " ist keine gÙltige Excel Formel")
    fn_DateiAuslesen = ret
    
End Function

Function fn_writeBackDuplicateSystemId(StartRow, EndRow)
    Dim row, Arr, Duplicate, Id
    Arr = "|"
    fn_writeBackDuplicateSystemId = "" ' Default return value
    For row = StartRow To EndRow
        If Range("F" & row) = "x" Then
            Id = Range("E" & row).Value
            If Arr Like "*|" & Id & "|*" Then
                fn_writeBackDuplicateSystemId = Id
                MsgBox "Die SystemID " & Id & " würde mehrfach gebucht werden!", vbCritical, "Fehler"
            End If
            Arr = Arr & Id & "|"
        End If
    Next row
    
End Function



Sub getOutlookEntryID()
    ' Nur für Testzwecke gedacht
    
    Dim olApp, olStore, olAccount, olFolder, olRootFolders, olMails, olMail
    Dim ret
    Const folderType = 6  ' 5=Sent Mail, 6=Inbox, 16=Drafts
    Set olApp = GetObject("", "Outlook.Application")
    
    For Each olAccount In olApp.Session.Accounts
        If olAccount = "office@digirentalwien.at" Then
            Set olStore = olAccount.DeliveryStore
            Set olRootFolders = olStore.GetDefaultFolder(folderType).Parent.Folders
            For Each olFolder In olRootFolders
                If olFolder.FolderPath Like "*Posteingang" Then
                    Set olMails = olFolder.Items.Restrict("[Subject] = 'Test Christof'")
                    If olMails.Count > 0 Then
                        For Each olMail In olMails
                            ret = MsgBox(olMails.Count & " E-Mail(s) gefunden." & Chr(10) _
                                & Chr(10) & "Willst Du die ID des Emails von " & olMail.SenderEmailAddress & " vom " & olMail.SentOn & "?" _
                                , vbInformation + vbYesNoCancel)
                            If ret = vbCancel Then End
                            If ret = vbYes Then
                                ret = InputBox("Deine EntryID", "Gefunden", olMail.entryID)
                            End If
                        Next olMail
                    End If
                End If
            Next
        End If
    Next
End Sub



Sub PrintAndSendPDF()
    
    Dim filename As String
    Dim printrange
    Dim olApp, olOrigMail, olReplyMail, entryID
    Dim mailto, subject, body
    Dim tempfolder
    Dim lastRow
    Dim Arr() As String
    
    

    
    ' Ermittle tempor_ren Folder und Export-Dateiname
    tempfolder = Environ("temp")
    If Len(tempfolder) = 0 Then tempfolder = Environ("tmp")
    If Len(tempfolder) = 0 Then fn_FatalError ("Ohne 'tmp' oder 'temp' Umgebungsvariable kann das Makro kein PDF exportieren.")
        
    
    Select Case ActiveSheet.Name
    Case "Angebot_VKV"
        If Len(Sheets("Buchung").Range("AngebNr").Value) Then
            'LastRow = Range("A65535").End(xlUp).Row
            lastRow = Split(Range("Y2").Value, "$")
            lastRow = lastRow(UBound(lastRow))
            printrange = "A1:K" & lastRow
            mailto = ActiveSheet.Range("A12").Value
            subject = Sheets("Buchung").Range("AngebSubj").Value
            body = Sheets("Buchung").Range("AngebMail").Value
            filename = "\" & Sheets("Buchung").Range("AngebFile").Value & ".pdf"
        Else
            If MsgBox("Dieses Dokument ist noch nicht im Angebots-Status. " _
            & Chr(10) & "ZurÙck zum Buchungssheet? ", vbYesNo) = vbYes Then
                Sheets("Buchung").Select
                End
            End If
        End If
                
    Case "Miete_Lieferschein_VKV"
        If Len(Sheets("Buchung").Range("MieteDat").Value) Then
            'LastRow = Range("A65535").End(xlUp).Row
            lastRow = Split(Range("Y2").Value, "$")
            lastRow = lastRow(UBound(lastRow))
            printrange = "A1:K" & lastRow
            mailto = ActiveSheet.Range("A12").Value
            subject = Sheets("Buchung").Range("MieteSubj").Value
            body = Sheets("Buchung").Range("MieteMail").Value
            filename = "\" & Sheets("Buchung").Range("MieteFile").Value & ".pdf"
        Else
            If MsgBox("Dieses Dokument ist noch nicht im Miete-Status. " _
            & Chr(10) & "ZurÙck zum Buchungssheet? ", vbYesNo) = vbYes Then
                Sheets("Buchung").Select
                End
            End If
        End If
        
    Case "HDW", "SL", "VKB", "USR", "ZA", "VKV"
        If Len(Sheets("Buchung").Range("ReNr").Value) Then
            'LastRow = Range("A65535").End(xlUp).Row
            lastRow = Split(Range("Y2").Value, "$")
            lastRow = lastRow(UBound(lastRow))
            printrange = "A1:K" & lastRow
            mailto = Sheets("Buchung").Range("AW8").Value
            subject = Sheets("Buchung").Range("RechnSubj").Value
            body = Sheets("Buchung").Range("RechnMail").Value
            filename = "\" & Sheets("Buchung").Range("RechnFile").Value & ".pdf"
        Else
            If MsgBox("Dieses Dokument ist noch nicht im Rechnungs-Status. " _
            & Chr(10) & "Zurück zum Buchungssheet? ", vbYesNo) = vbYes Then
                Sheets("Buchung").Select
                End
            End If
        End If
        
    Case Else
        fn_FatalError ("Zu diesem Worksheet gibt es keine Send-Funktion.")
    End Select
    
    ' sollte unterverzeichnis im filenamen sein, nur filenamen
    Arr = Split(filename, "\")
    filename = Arr(UBound(Arr))
    filename = tempfolder & "\" & filename

    ' Print To PDF
    ActiveWorkbook.ActiveSheet.Range(printrange).ExportAsFixedFormat _
        Type:=xlTypePDF, filename:=filename, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

    entryID = Sheets("Anfrage").Range("C1").Value
    
    'Set olApp = CreateObject("Outlook.Application")
    Set olApp = GetObject("", "Outlook.Application")
    
    If Len(entryID) > 0 Then
       ' Suche das Email in Outlook heraus, das wir beim Makro "Anfrage_FindeInOutlook" schon gefunden haben
        On Error Resume Next
        Set olOrigMail = olApp.Session.GetItemFromID(entryID)
        If Err() <> 0 Then
            If MsgBox("Outlook EntryID auf Zelle C1 (Anfrage Sheet) gibt es nicht (mehr). Neues Email wird erstellt.", vbOKCancel) = vbCancel Then End
            Set olReplyMail = olApp.CreateItem(0)
        Else
            Set olReplyMail = olOrigMail.Reply
        End If
        On Error GoTo 0
    Else
        Set olReplyMail = olApp.CreateItem(0)
    End If
    
    olReplyMail.To = mailto
    olReplyMail.subject = subject
    olReplyMail.HTMLBody = Replace(body, Chr(10), "<br>") & olReplyMail.HTMLBody
    olReplyMail.Attachments.Add filename
    olReplyMail.Display
    'olReplyMail.Send
    
    Set olReplyMail = Nothing
    Set olApp = Nothing
    Kill filename
    
End Sub



Sub TestReplyEmail()
    Dim olApp, olOrigMail, olReplyMail
    'Const entryID = "000000004B9BD25BB7F53F46A97282C484A2CE380700C3B68E10F77511CEB4CD00AA00BBB6E600000000000C0000D9539C2261A6BB45B9DAB62C7081B3C10100931301000000"
    Dim entryID: entryID = Sheets("Anfrage").Range("C1").Value
    Set olApp = GetObject("", "Outlook.Application")
    On Error Resume Next
    Set olOrigMail = olApp.Session.GetItemFromID(entryID)
    If Err() <> 0 Then
        MsgBox "No such EntryID in Outlook", vbCritical
        Set olReplyMail = olApp.CreateItem(0)
    Else
        Set olReplyMail = olOrigMail.Reply
    End If
    On Error GoTo 0
    olReplyMail.HTMLBody = "Test Email ... bitte ignorieren." & Chr(10) & olReplyMail.HTMLBody
    olReplyMail.subject = "Test Email"
    olReplyMail.To = "christof.schwarz@gmx.at"
    olReplyMail.Display
    'olReplyMail.Send
    'MsgBox "Email Antwort gesendet"
End Sub

'ReminderSet
Sub FindeEmailMitNachverfolgung()
    Dim olApp, olStore, olAccount, olFolder, olRootFolders, olMails, olMail, ret
    Const folderType = 6  ' 5=Sent Mail, 6=Inbox, 16=Drafts
    
    If ActiveSheet.Name <> "Anfrage" Then fn_FatalError ("Du bist nicht auf dem 'Anfrage' Sheet")
    Set olApp = GetObject("", "Outlook.Application")
    
    For Each olAccount In olApp.Session.Accounts
    
        If olAccount = sucheInAccount Then
            Set olStore = olAccount.DeliveryStore
            Set olRootFolders = olStore.GetDefaultFolder(folderType).Parent.Folders
            
            For Each olFolder In olRootFolders
                If olFolder.FolderPath Like sucheInFolder Then
                    ' Los gehts, richtiger Outlook Account, richtiger Folder
                    Set olMails = olFolder.Items.Restrict("@SQL= urn:schemas:httpmail:ReminderSet=true") '("[Subject]='Hallo'") '("[ReminderSet] = true")
                    
                    If olMails.Count > 0 Then
                        For Each olMail In olMails
                            ret = MsgBox(olMails.Count & " E-Mail(s) mit Nachverfolgung gefunden." & Chr(10) _
                                & Chr(10) & "Email von " & olMail.SenderEmailAddress & " am " & olMail.SentOn & "?" _
                                , vbInformation + vbYesNoCancel, olMail.FlagStatus)
                            If ret = vbCancel Then End
                            If ret = vbYes Then
                                Range("C1").Value = "'" & olMail.entryID
                                End
                            End If
                                                        
                        Next olMail
                    
                    Else
                        MsgBox "Keine Emails mit Nachverfolgungs-Flag gefunden.", vbInformation
                    End If
                End If
            Next
        End If
    Next
End Sub


Sub GetAllExistingReminders()
    Dim objReminders 'As Reminders
    Dim objReminder 'As Reminder
    Dim strReminderDetails 'As String
    Dim objNewNote 'As NoteItem
    Dim olApp, olNoteApp
 
    Set olApp = GetObject("", "Outlook.Application")
    Set objReminders = olApp.Reminders

   'Check if any reminders exist
    If objReminders.Count = 0 Then
       MsgBox "There is no existing reminder."
    Else
       For Each objReminder In objReminders
           strReminderDetails = strReminderDetails & objReminder.Caption & " (" & TypeName(objReminder.item) & ") ---- " & objReminder.NextReminderDate & vbCrLf
       Next objReminder
       'Display report in a new note item
       Set objNewNote = olApp.CreateItem(0)

    End If
End Sub


Sub Anfrage_FindeInOutlook()
    Dim olApp, olStore, olAccount, olFolder, olRootFolders, olMails, olMail

    'Dim searchInFolders: searchInFolders = Split(GetMacroSetting("olFolders"), ";")
    Const folderType = 6  ' 5=Sent Mail, 6=Inbox, 16=Drafts
    Const sucheNachSubject = "Anfrage Kameraverleih"
    Dim accountListe: accountListe = ""
    Dim folderListe: folderListe = ""
    Dim ret
    Dim body
    
    If ActiveSheet.Name <> "Anfrage" Then fn_FatalError ("Du bist nicht auf dem 'Anfrage' Sheet")
    Set olApp = GetObject("", "Outlook.Application")
    
    For Each olAccount In olApp.Session.Accounts
    
        accountListe = accountListe & Chr(10) & olAccount
        If olAccount = sucheInAccount Then
            'Set ns = olApp.GetNamespace("MAPI")
            Set olStore = olAccount.DeliveryStore
            'Set olFolder = olStore.GetDefaultFolder(folderType)
            Set olRootFolders = olStore.GetDefaultFolder(folderType).Parent.Folders
            
            For Each olFolder In olRootFolders
                folderListe = folderListe & Chr(10) & olFolder.FolderPath
                
                If olFolder.FolderPath Like sucheInFolder Then
                    ' Los gehts, richtiger Outlook Account, richtiger Folder
                    'Set olMails = olFolder.Items.Restrict("@SQL= urn:schemas:httpmail:subject LIKE '%Rechnung 210403%'")
                    Set olMails = olFolder.Items.Restrict("[UnRead] = true And [Subject] = '" & sucheNachSubject & "'")
                    
                    
                    If olMails.Count > 0 Then
                        For Each olMail In olMails
                            If olMail.subject Like sucheNachSubject And Len(olMail.ReplyRecipientNames) > 0 Then
                                ret = MsgBox(olMails.Count & " neue E-Mail(s) mit Subject '" & sucheNachSubject & "' gefunden." & Chr(10) _
                                    & Chr(10) & "Verarbeite Email von " & olMail.ReplyRecipientNames & " am " & olMail.SentOn & "?" _
                                    , vbInformation + vbYesNoCancel)
                                If ret = vbCancel Then End
                                If ret = vbYes Then
                                    body = fn_clipboard(olMail.HTMLBody)
                                    olMail.unRead = False
                                    Range("C1").Value = "'" & olMail.entryID
                                    Range("A1").Select
                                    ActiveSheet.Paste
                                    Columns("A:B").Font.Size = 9
                                    Call Anfrage_EmailPruefen
                                    End
                                End If
                            End If
                        Next olMail
                    
                    Else
                        MsgBox "Keine Emails mit Subject " & sucheNachSubject & " in " & sucheInFolder & " gefunden.", vbInformation
                        
                    End If
                
                
                End If
            Next
        End If
    Next
    
    If Not accountListe Like ("*" & sucheInAccount & "*") Then
        MsgBox "Mögliche Accounts:" & accountListe, vbCritical, "Account " & sucheInAccount & " nicht gefunden"
    Else
        If Not folderListe Like (sucheInFolder & "*") Then
            MsgBox "Mögliche Folder:" & folderListe, vbCritical, "Folder " & sucheInFolder & " nicht gefunden"
        End If
    End If
        
    
End Sub



Sub Button_VKV()
    If fn_SetzeRechnungsTyp("VKV", ShapeId_VKV) Then Call Click_Shape_Draft
End Sub

Sub Button_SL()
    If fn_SetzeRechnungsTyp("SL", ShapeId_SL) Then Call Click_Shape_Draft
End Sub

Sub Button_USR()
    If fn_SetzeRechnungsTyp("USR", ShapeId_USR) Then Call Click_Shape_Draft
End Sub

Sub Button_HDW()
    If fn_SetzeRechnungsTyp("HDW", ShapeId_HDW) Then Call Click_Shape_Draft
End Sub

Sub Button_VKB()
    If fn_SetzeRechnungsTyp("VKB", ShapeId_VKB) Then Call Click_Shape_Draft
End Sub

Sub Button_ZA()
    If fn_SetzeRechnungsTyp("ZA", ShapeId_ZA) Then Call Click_Shape_Draft
End Sub


Function fn_SetzeRechnungsTyp(RechnungsTyp, shapeName)

    Dim button, zielVerzeichnis
    
    Set button = Sheets("Buchung").Shapes.Range(Array(shapeName))
    
    If Range("Rechnungstyp").Text = "" Then
        
        zielVerzeichnis = fn_FindeZielVerzeichnis(RechnungsTyp)
        If zielVerzeichnis = "" Then
            MsgBox "Für den Rechnungstyp " & RechnungsTyp & " gibt es kein passendes Zielverzeichnis", vbCritical, "Error"
            fn_SetzeRechnungsTyp = False
        Else
        
            Range("RootFolder").Value = zielVerzeichnis
            ' MsgBox "Rechnungstyp und Zielverzeichnis gesetzt: " & zielVerzeichnis, vbInformation, "Passt"
            button.Fill.ForeColor.RGB = RGB(255, 185, 0)
            Range("Rechnungstyp").Value = RechnungsTyp
            fn_SetzeRechnungsTyp = True
        End If
    Else
        MsgBox "Der Rechnungstyp wurde bereits auf " & Range("Rechnungstyp").Text & " gesetzt und kann nicht mehr geändert werden.", vbExclamation, "Fehler"
        fn_SetzeRechnungsTyp = False
    End If
    
End Function


Function fn_FindeZielVerzeichnis(RechnungsTyp)

' sucht im Folder des Workbooks nach Unterordnern, deren Name mit dem RechnungsTyp beginnt, also etwa "VKV_2023"

    Dim folder, zielVerzeichnis
    folder = Dir(ActiveWorkbook.Path & "\*.", vbDirectory)
    zielVerzeichnis = ""
        
    Do While folder <> ""
        If folder Like (RechnungsTyp & "*") Then
            zielVerzeichnis = ActiveWorkbook.Path & "\" & folder & "\"
        End If
        folder = Dir()
    Loop
    fn_FindeZielVerzeichnis = zielVerzeichnis
End Function

Sub Positioniere_Alle_Buttons()
   
    With Sheets("Buchung").Shapes.Range(Array(ShapeId_Draft))
        .Left = Sheets("Buchung").Range("E:E").Left
        .Width = 27
        .Top = (Sheets("Buchung").Range("7:7").Top + Sheets("Buchung").Range("8:8").Top) / 2
        .Height = 27
    End With
    With Sheets("Buchung").Shapes.Range(Array(ShapeId_Angeb))
        .Left = Sheets("Buchung").Range("E:E").Left + 22
        .Width = 27
        .Top = (Sheets("Buchung").Range("7:7").Top + Sheets("Buchung").Range("8:8").Top) / 2
        .Height = 27
    End With
    With Sheets("Buchung").Shapes.Range(Array(ShapeId_Miete))
        .Left = Sheets("Buchung").Range("E:E").Left + 44
        .Width = 27
        .Top = (Sheets("Buchung").Range("7:7").Top + Sheets("Buchung").Range("8:8").Top) / 2
        .Height = 27
    End With
    With Sheets("Buchung").Shapes.Range(Array(ShapeId_Rechn))
        .Left = Sheets("Buchung").Range("E:E").Left + 66
        .Width = 27
        .Top = (Sheets("Buchung").Range("7:7").Top + Sheets("Buchung").Range("8:8").Top) / 2
        .Height = 27
    End With
    
    With Sheets("Buchung").Shapes.Range(Array(ShapeId_VKV))
        .Left = Sheets("Buchung").Range("J:J").Left
        .Width = Sheets("Buchung").Range("J:J").Width - 1
        .Top = Sheets("Buchung").Range("5:5").Top
        .Height = Sheets("Buchung").Range("5:5").Height - 1
    End With
    With Sheets("Buchung").Shapes.Range(Array(ShapeId_SL))
        .Left = Sheets("Buchung").Range("J:J").Left
        .Width = Sheets("Buchung").Range("J:J").Width - 1
        .Top = Sheets("Buchung").Range("6:6").Top
        .Height = Sheets("Buchung").Range("6:6").Height - 1
    End With
    With Sheets("Buchung").Shapes.Range(Array(ShapeId_USR))
        .Left = Sheets("Buchung").Range("J:J").Left
        .Width = Sheets("Buchung").Range("J:J").Width - 1
        .Top = Sheets("Buchung").Range("7:7").Top
        .Height = Sheets("Buchung").Range("7:7").Height - 1
    End With
    With Sheets("Buchung").Shapes.Range(Array(ShapeId_HDW))
        .Left = Sheets("Buchung").Range("K:K").Left
        .Width = Sheets("Buchung").Range("K:K").Width - 1
        .Top = Sheets("Buchung").Range("5:5").Top
        .Height = Sheets("Buchung").Range("5:5").Height - 1
    End With
    With Sheets("Buchung").Shapes.Range(Array(ShapeId_VKB))
        .Left = Sheets("Buchung").Range("K:K").Left
        .Width = Sheets("Buchung").Range("K:K").Width - 1
        .Top = Sheets("Buchung").Range("6:6").Top
        .Height = Sheets("Buchung").Range("6:6").Height - 1
    End With
    With Sheets("Buchung").Shapes.Range(Array(ShapeId_ZA))
        .Left = Sheets("Buchung").Range("K:K").Left
        .Width = Sheets("Buchung").Range("K:K").Width - 1
        .Top = Sheets("Buchung").Range("7:7").Top
        .Height = Sheets("Buchung").Range("7:7").Height - 1
    End With
End Sub




