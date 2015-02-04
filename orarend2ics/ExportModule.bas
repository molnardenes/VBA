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
    
    'csin�ljunk egy m�solatot a param�terk�nt kapott munkalapr�l, hogy ne az eredetit m�dos�tgassuk
    ws.Copy After:=Sheets(Sheets.Count)
    ActiveSheet.name = "backup"
    
    
    'Utols� haszn�latban l�v� sor meghat�roz�sa
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    
    'H�tulr�l v�gigiter�l a list�n, ha az el�z� megegyezik az aktu�lis sorral, akkor az aktu�lis z�r� idej�vel
    'fel�l�rja az el�z� z�r�idej�t �s t�rli az aktu�lis sort
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
    
    'Utols� haszn�latban l�v� sor meghat�roz�sa �jra
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    
    'L�trehozunk egy �j gy�jtem�nyt, amelyben a tan�r�kat t�roljuk majd
    Set courses = New Collection
    
    'V�gigiter�lunk az �sszes haszn�latban l�v� soron �s l�trehozunk tan�ra p�ld�nyokat
    'fel�ltve a property-ket �rtelemszer�em a megfelel� oszlopokb�l, majd hozz�adjuk a
    'gy�jtem�nyhez az aktu�lis p�ld�nyt
    For counter = 2 To lastRow
        Set thisCourse = New Course
        thisCourse.className = Range("E" & counter)
        thisCourse.datum = Range("B" & counter)
        'az id�pont tr�kk�sebb: ha az egysz�mjegy� id�pontokat (pl: reggel 8) nem null�val kieg�sz�tve adj�k meg
        'akkor csak 8 vagy 9 karakter hossz� lesz a string, ez�rt ezt ellen�rizni kell �s megfelel� form�ra jav�tani
        Select Case Len(Trim(Range("D" & counter)))
            Case Is = 9
                thisCourse.time = "0" & Left(Range("D" & counter), 4) & ":00-0" & Right(Range("D" & counter), 4) & ":00"
            Case Is = 10
                thisCourse.time = "0" & Left(Range("D" & counter), 4) & ":00-" & Right(Range("D" & counter), 5) & ":00"
            Case Is = 11
                thisCourse.time = Left(Range("D" & counter), 5) & ":00-" & Right(Range("D" & counter), 5) & ":00"
        End Select
        'a megjegyz�sn�l �s a helysz�nn�l cser�lj�k a sorv�ge jelet vessz�re
        thisCourse.professor = Replace(Range("F" & counter), Chr(10), ", ")
        thisCourse.room = Replace(Range("G" & counter), Chr(10), ", ")
        courses.Add thisCourse
    Next counter

    'Az ics f�jl tartalm�t egy string v�ltoz�ba t�ltj�k
    finalString = "BEGIN:VCALENDAR" & vbCrLf
    finalString = finalString & "PRODID:-//Google Inc//Google Calendar 70.9054//EN" & vbCrLf
    finalString = finalString & "VERSION:2.0" & vbCrLf
    finalString = finalString & "METHOD:PUBLISH" & vbCrLf
    
    'V�gigiter�lunk a tan�ra gy�jtem�ny�nk�n �s  sz�ks�ges �rt�keket hozz�adjuk a ki�rand�
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
    
    'megadjuk, hogy milyen n�ven akarjuk a f�jlt menteni
    icsFileName = exportUserForm.fileNameTextBox.Text
    
    'l�trehozzuk a f�jlt �s ki�rjuk a tartalm�t
    createFile icsFileName, finalString
    
    'a backup sheetet t�r�lhetj�k nyugodtan, nem is k�r�nk r�la �rtes�t�st
    Application.DisplayAlerts = False
    Sheets("backup").Delete
    Application.DisplayAlerts = True
    Sheets("Export").Activate
End Sub
Function createFile(fileName As String, contents As String)
'f�jl l�trehoz�sa �s felt�lt�se a megadott stringgel
 
Dim streamObject As Object

'l�trehozzuk a stream objektumot
Set streamObject = CreateObject("ADODB.Stream")

'text/string t�pus� az objektumunk, ez a type 2
streamObject.Type = 2

'karakterk�dol�st be�ll�tjuk
streamObject.Charset = "utf-8"

'f�jl megnyit�sa �s adat be�r�sa
streamObject.Open
streamObject.writeText contents

On Error GoTo errorHandler
streamObject.saveToFile fileName

Exit Function
errorHandler:
  MsgBox "Error " & Err.Number & ": " & Err.Description

End Function
