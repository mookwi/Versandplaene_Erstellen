Sub Versandplaene_Erstellen()
' Hier sollen die Versandpläne für einen laufenden Monat (aber nur bei einem Wochentag)
' erstellt werden. Hierzu gibt man den Monat und das Jahr in 2 Feldern an und diese
' Funktion ersllt dann für jeden Werktag des entsprechenden Monats eine Exceltabelle
' im Format "Abholung JJ-MM-TT", z.B. Abholung 22-03-22.
'
' Temporäres Verzechnis, in die die Excel-Tabellen abgelegt werden
' z.B. "T:\VERSANDPLAN\_Temp\" (siehe Feld B4)
' Monat und Jahr der zu erstellenden Versandpläne
' z.B. "03" für Monat und "2022" für Jahr (siehe Feld B5 und B6)
'
' Erstellt von:     Carsten Graage
' Erstellt am:      07.03.2022
'
Dim sMonat As String, sJahr As String, sWochentag As String, sTag As String
Dim startDatum As String
Dim sUmbenannt As String
Dim aktuellDatum As Date
Dim iMonatsErster, iMonatsLetzter As Integer
Dim x As Integer

sVorlage = Worksheets("Versandplaene").Range("B3").Value & "\" & _
           Worksheets("Versandplaene").Range("B2").Value

sMonat = Worksheets("Versandplaene").Range("B5").Value
sJahr = Worksheets("Versandplaene").Range("B6").Value
startDatum = "01/" & sMonat & "/" & sJahr
' Ermittlung des letzten Tag des gewählten Monats
iMonatsErster = 1
iMonatsLetzter = Letzter_Tag_im_Monat(CDate(startDatum))


' Verzeichnis aus Feld B4 komplett leeren
' Daten_Loeschen

' Jetzt durchlaufen wir eine Schleife vom ersten bis zum letzten Tag des Monats

For x = iMonatsErster To iMonatsLetzter
' Hier werden einstellige Tageswerte um eine NULL (0) ergänzt
If Len(x) <= 1 Then
    aktuellDatum = CDate("0" & x & "/" & sMonat & "/" & sJahr)
Else
    aktuellDatum = CDate(x & "/" & sMonat & "/" & sJahr)
End If

sWochentag = Weekday(aktuellDatum)

' vbSunday    1   Sonntag
' vbMonday    2   Montag
' vbTuesday   3   Dienstag
' vbWednesday 4   Mittwoch
' vbThursday  5   Donnerstag
' vbFriday    6   Freitag
' vbSaturday  7   Samstag
' Von 2 (Montag) bis 6 (Freitag) soll eine Datei angelegt werden, bzw. erst
' kopiert und dann umbenannt werden, Sonntag und Samstag nicht!
' Es werden ebenfalls keine Feiertage berücksichtigt! 

Select Case sWochentag
    Case 2 To 6
    ' Datei Kopieren
    If Vorlage_Copy(Worksheets("Versandplaene").Range("B3").Value & "\" & _
                    Worksheets("Versandplaene").Range("B2").Value) = True Then
      ' Datei Umbenennen
      If Len(x) <= 1 Then
          sTag = "0" & CStr(x)
      Else
          sTag = CStr(x)
      End If
    
      If Len(sMonat) <= 1 Then
          sMonat = "0" & sMonat
      End If
      Call Vorlage_Rename(Worksheets("Versandplaene").Range("B6").Value, sMonat, sTag)
    End If
End Select

MsgBox ("Die Versandpläne wurde erstellt!")

Next x

End Sub

Function Letzter_Tag_im_Monat(startDatum As Date)

' Hier wird der Letzte Tag des Monats ermittelt
Dim lastDay As Byte
Dim iLetzterTag As Integer
lastDay = Day(DateSerial(Year(startDatum), Month(startDatum) + 1, 0))
iLetzterTag = CInt(lastDay)
Letzter_Tag_im_Monat = iLetzterTag

End Function

Public Function Vorlage_Copy(sVorlagendatei As String) As Boolean
Dim SourceFile, DestinationFile
' Vorlagendatei kopieren
SourceFile = sVorlagendatei ' Define source file name.
DestinationFile = Worksheets("Versandplaene").Range("B4").Value & "\" & _
Worksheets("Versandplaene").Range("B2").Value ' dateinamen erstellen
FileCopy SourceFile, DestinationFile ' Datei umbenennen
Vorlage_Copy = True
End Function

Function Vorlage_Rename(sYear As String, sMonth As String, sDay As String)
Dim AlterDateiName As String
Dim NeuerDateiName As String
' Die eben kopierte Vorlagendatei umbenennen
AlterDateiName = Worksheets("Versandplaene").Range("B4").Value & "\" & _
                 Worksheets("Versandplaene").Range("B2").Value
' Wenn Tag einstellig, dann führende "0" vor diesen Tag setzen
If Len(sDay) = 1 Then
    NeuerDateiName = Worksheets("Versandplaene").Range("B4").Value & "\" & _
                     "Abholung " & Right(sYear, 2) & "-" & sMonth & "-0" & sDay & ".xls"
    Name AlterDateiName As NeuerDateiName
Else
    NeuerDateiName = Worksheets("Versandplaene").Range("B4").Value & "\" & _
                     "Abholung " & Right(sYear, 2) & "-" & sMonth & "-" & sDay & ".xls"
    Name AlterDateiName As NeuerDateiName
End If

End Function
