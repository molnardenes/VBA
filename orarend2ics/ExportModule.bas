Attribute VB_Name = "ExportModule"
Option Explicit
Sub exportToIcs(ws As Worksheet)
    Dim courses As Collection
    Dim thisCourse As Course
    Dim finalString As String
    Dim nameOfCourse As String
    Dim dateOfCourse As String
    Dim startTime As String
    Dim endTime As String
    Dim profOfCourse As String
    Dim roomOfCourse As String
    Dim lastRow As Integer
    Dim counter As Integer
    Dim guid As String
    Dim icsFileName As String
    
    'csináljunk egy másolatot a paraméterként kapott munkalapról, hogy ne az eredetit módosítgassuk
    ws.Copy After:=Sheets(Sheets.Count)
    ActiveSheet.name = "backup"
    
    
    'Utolsó használatban lévõ sor meghatározása
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    
    'Hátulról végigiterál a listán, ha az elõzõ megegyezik az aktuális sorral, akkor az aktuális záró idejével
    'felülírja az elõzõ záróidejét és törli az aktuális sort
    For counter = lastRow To 2 Step -1
        If (Range("B" & counter) = Range("B" & counter - 1)) And (Range("E" & counter) = Range("E" & counter - 1)) _
            And (Range("F" & counter) = Range("F" & counter - 1)) Then
            If Len(Trim(Range("D" & counter - 1))) < 11 Then
                If Len(Trim(Range("D" & counter))) = 9 Then
                    Range("D" & counter - 1).Value = Left(Range("D" & counter - 1), 4) & "-" & Right(Range("D" & counter), 4)
                Else
                    Range("D" & counter - 1).Value = Left(Range("D" & counter - 1), 4) & "-" & Right(Range("D" & counter), 5)
                End If
            Else
                Range("D" & counter - 1).Value = Left(Range("D" & counter - 1), 5) & "-" & Right(Range("D" & counter), 5)
            End If
            Range("A" & counter).EntireRow.Delete
        End If
    Next counter
    
    'Utolsó használatban lévõ sor meghatározása újra
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    
    'Létrehozunk egy új gyûjteményt, amelyben a tanórákat tároljuk majd
    Set courses = New Collection
    
    'Végigiterálunk az összes használatban lévõ soron és létrehozunk tanóra példányokat
    'felöltve a property-ket értelemszerûem a megfelelõ oszlopokból, majd hozzáadjuk a
    'gyûjteményhez az aktuális példányt
    For counter = 2 To lastRow
        Set thisCourse = New Course
        thisCourse.className = Range("E" & counter)
        thisCourse.datum = Range("B" & counter)
        'az idõpont trükkösebb: ha az egyszámjegyû idõpontokat (pl: reggel 8) nem nullával kiegészítve adják meg
        'akkor csak 8 vagy 9 karakter hosszú lesz a string, ezért ezt ellenõrizni kell és megfelelõ formára javítani
        Select Case Len(Trim(Range("D" & counter)))
            Case Is = 9
                thisCourse.time = "0" & Left(Range("D" & counter), 4) & ":00-0" & Right(Range("D" & counter), 4) & ":00"
            Case Is = 10
                thisCourse.time = "0" & Left(Range("D" & counter), 4) & ":00-" & Right(Range("D" & counter), 5) & ":00"
            Case Is = 11
                thisCourse.time = Left(Range("D" & counter), 5) & ":00-" & Right(Range("D" & counter), 5) & ":00"
        End Select
        'a megjegyzésnél és a helyszínnél cseréljük a sorvége jelet vesszõre
        thisCourse.professor = Replace(Range("F" & counter), Chr(10), ", ")
        thisCourse.room = Replace(Range("G" & counter), Chr(10), ", ")
        courses.Add thisCourse
    Next counter

    'Az ics fájl tartalmát egy string változóba töltjük
    finalString = "BEGIN:VCALENDAR" & vbCrLf
    finalString = finalString & "PRODID:-//Google Inc//Google Calendar 70.9054//EN" & vbCrLf
    finalString = finalString & "VERSION:2.0" & vbCrLf
    finalString = finalString & "METHOD:PUBLISH" & vbCrLf
    
    'Végigiterálunk a tanóra gyûjteményünkön és  szükséges értékeket hozzáadjuk a kiírandó
    'stringhez
    For Each thisCourse In courses
        nameOfCourse = thisCourse.className
        dateOfCourse = Format(thisCourse.datum, "yyyymmdd")
        startTime = dateOfCourse & "T" & Format(Left(thisCourse.time, 8), "hhmmss")
        endTime = dateOfCourse & "T" & Format(Right(thisCourse.time, 8), "hhmmss")
        profOfCourse = thisCourse.professor
        roomOfCourse = thisCourse.room
        finalString = finalString & "BEGIN:VEVENT" & vbCrLf
        finalString = finalString & "DTSTART:" & startTime & vbCrLf
        finalString = finalString & "DTEND:" & endTime & vbCrLf
        finalString = finalString & "DTSTAMP:" & Format(Date, "yyyymmdd") & "T" & Format(time, "hhmmss") & vbCrLf
        finalString = finalString & "UID:" & Mid$(CreateObject("Scriptlet.TypeLib").guid, 2, 36) & vbCrLf
        finalString = finalString & "DESCRIPTION:" & profOfCourse & vbCrLf
        finalString = finalString & "LOCATION:" & roomOfCourse & vbCrLf
        finalString = finalString & "SEQUENCE:0" & vbCrLf
        finalString = finalString & "STATUS:CONFIRMED" & vbCrLf
        finalString = finalString & "SUMMARY:" & nameOfCourse & vbCrLf
        finalString = finalString & "TRANSP:TRANSPARENT" & vbCrLf
        finalString = finalString & "END:VEVENT" & vbCrLf
    Next
    finalString = finalString & "END:VCALENDAR"
    
    'megadjuk, hogy milyen néven akarjuk a fájlt menteni
    icsFileName = exportUserForm.fileNameTextBox.Text
    
    'létrehozzuk a fájlt és kiírjuk a tartalmát
    createFile icsFileName, finalString
    
    'a backup sheetet törölhetjük nyugodtan, nem is kérünk róla értesítést
    Application.DisplayAlerts = False
    Sheets("backup").Delete
    Application.DisplayAlerts = True
    Sheets("Export").Activate
End Sub
Function createFile(fileName As String, contents As String)
'fájl létrehozása és feltöltése a megadott stringgel
 
Dim streamObject As Object

'létrehozzuk a stream objektumot
Set streamObject = CreateObject("ADODB.Stream")

'text/string típusú az objektumunk, ez a type 2
streamObject.Type = 2

'karakterkódolást beállítjuk
streamObject.Charset = "utf-8"

'fájl megnyitása és adat beírása
streamObject.Open
streamObject.writeText contents

On Error GoTo errorHandler
streamObject.saveToFile fileName

Exit Function
errorHandler:
  MsgBox "Error " & Err.Number & ": " & Err.Description

End Function
