Sub ICS_Erstellen()
    'Sortiert die Tabelle nach Geburtstagen (chronologische Reihenfolge)
    'EintrŠge ohne Datumsangabe werden dabei ans Ende gestellt
    On Error GoTo Fehler
        Dim varICS_Datei As Variant
        Dim FF As Integer
        Dim Zeitstempel As String
        Dim datum As Date, datum2 As Date
        Dim i As Long       'Zeilenzähler
        Dim k As String
        'Dateiauswahldialog für ICS-Datei abhängig vom Betriebssystem öffnen
        If Left(VBA.Environ("OS"), 7) = "Windows" Then
            'unter Windows-Betriebssystem
            varICS_Datei = Application.GetSaveAsFilename( _
            FileFilter:="ICS-Datei (*.ics),*.ics", _
            Title:="Bitte Dateiname für die ICS-Datei eingeben/auswählen")
        Else
            'etwas anderes - ggf. Parameter FileFilter oder FilterIndex passend für _
            Betriebs-System angeben
            'MAC evtl: FileFilter:=MacID("ICS") oder FileFilter:=MacID("TEXT")
            Dialog_anderes_OS:
            varICS_Datei = Application.GetSaveAsFilename( _
            Title:="Bitte Dateiname für die ICS-Datei eingeben/auswählen", _
            ButtonText:="Auswählen")
        End If
        If varICS_Datei = False Then Exit Sub
        'Daten aufsteigend nach Datum sortieren
        Range("A1").Select
        Selection.Sort Key1:=Range("C1"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
        'Erstellt den Zeitstempels
        'wird benštigt fŸr die UID des Kalendereintrages und fŸr die Felder
        '"erstellt am" --> "DTSTAMP" und "zuletzt geŠndert am" --> "LAST-MODIFIED"
        'einfachere Form für Zeitstempel
        Zeitstempel = Format(Now, "YYYYMMDD""T""hhmmss""Z""")
        'Erstellt die Kalenderdatei
        'Dateiname kann frei gewŠhlt werden
        'Der entsprechende Ordner MUSS vorhanden sein, da sonst ein Fehler auftritt
        'Datendatei erstellen und zum beschreiben öffnen
        FF = FreeFile()
        Open varICS_Datei For Output As FF
        'Schreibt den allgemeinen Teils der Kalenderdatei
        Print #FF, "BEGIN:VCALENDAR"
        Print #FF, "VERSION"
        Print #FF, " :2.0"
        Print #FF, "PRODID"
        Print #FF, " :-//Mozilla.org/NONSGML Mozilla Calendar V1.0//EN"
        'Schleife zur Ermittlung aller EintrŠge
        'Benutzt alle DatensŠtze, die ein Datum enthalten
        i = 1
        While ActiveCell.Offset(i, 2) <> ""
            'Ermittelt die Daten fŸr den Kalendereintrag
            'Person und Geburtstagsdatum
            nachname = ActiveCell.Offset(i, 0)
            nachname = Replace(nachname, "Š", "ae")
            nachname = Replace(nachname, "€", "Ae")
            nachname = Replace(nachname, "š", "oe")
            nachname = Replace(nachname, "…", "Oe")
            nachname = Replace(nachname, "Ÿ", "ue")
            nachname = Replace(nachname, "†", "Ue")
            nachname = Replace(nachname, "§", "ss")
            nachname = Replace(nachname, "Ž", "e")
            'Mein Sunbird hatte Probleme mit Sonderzeichen, aus diesem Grund habe ich die "Wichtigsten"  _
            ersetzt
            vorname = ActiveCell.Offset(i, 1)
            vorname = Replace(vorname, "Š", "ae")
            vorname = Replace(vorname, "€", "Ae")
            vorname = Replace(vorname, "š", "oe")
            vorname = Replace(vorname, "…", "Oe")
            vorname = Replace(vorname, "Ÿ", "ue")
            vorname = Replace(vorname, "†", "Ue")
            vorname = Replace(vorname, "§", "ss")
            vorname = Replace(vorname, "Ž", "e")
            datum = ActiveCell.Offset(i, 2)
            datum2 = datum + 1
            'Die Angaben mit dem Zusatz 2 werden fŸr das Ende des jeweiligen Termins gebraucht
            'Das Ende eines ganztŠgigen Ereignisses ist immer der darauffolgende Tag
            k = Format(i, "0")
            'Schreibt den Kalendereintrag
            'Der Zusatz "-@kuechi-" in der UID kann nach Belieben geŠndert werden
            'k ist ein durchlaufender ZŠhler
            Print #FF, "BEGIN:VEVENT"
            Print #FF, "UID:" & Zeitstempel & "-@kuechi-" & k
            Print #FF, "SUMMARY"                            'Zusammenfassung/Betreff
            Print #FF, " :" & vorname & " " & nachname
            Print #FF, "DESCRIPTION"                        'Beschreibung / Notiz
            Print #FF, " :" & Format(Year(datum), "0000")
            Print #FF, "LOCATION"                           'Ort
            Print #FF, " :" & ""
            Print #FF, "X-MOZILLA-ALARM-DEFAULT-LENGTH"
            Print #FF, " :0"
            Print #FF, "X-MOZILLA-RECUR-DEFAULT-UNITS"      'Wiederholung-EInheit
            Print #FF, (" :years")
            Print #FF, "RRULE"
            Print #FF, " :FREQ=YEARLY;INTERVAL=1"           'Wiederholung-Frequenz/Interval
            Print #FF, "DTSTART"                            'Start - Datum
            Print #FF, " ;VALUE=DATE"
            Print #FF, " :" & Format(datum, "YYYYMMDD")
            Print #FF, "DTEND"                              'Ende - Datum
            Print #FF, " ;VALUE=DATE"
            Print #FF, " :" & Format(datum2, "YYYYMMDD")
            Print #FF, "DTSTAMP"                            'Speicherzeitpunkt
            Print #FF, " :" & Zeitstempel
            Print #FF, "LAST-MODIFIED"                      'Letzte Änderung-zeitpunkt
            Print #FF, " :" & Zeitstempel
            Print #FF, "END:VEVENT"
            i = i + 1
        Wend
        'Ende der Schleife
        'Ende der Kalenderdatei
        Print #FF, ("END:VCALENDAR")
        'Datendatei wieder schließen
        Close #FF
        'Sortiert die ursprŸngliche Tabelle in alphabetischer Reihenfolge
        Range("A1").Select
        Selection.Sort Key1:=Range("B1"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
        Selection.Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    Fehler:
        With Err
            Select Case .Number
            Case 0 'alles OK
            Case Else
                MsgBox "Fehler-Nr.: " & .Number & vbLf & .Description
                Close
            End Select
        End With
End Sub