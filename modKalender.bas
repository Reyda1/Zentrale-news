Attribute VB_Name = "modKalender"
Option Explicit

Private Const BUB_LONDON As Long = 1
Private Const BUB_NEW_YORK As Long = 2
Private Const BUB_PM As Long = 4
Private Const BUB_ALLE As Long = BUB_LONDON Or BUB_NEW_YORK Or BUB_PM
Private Const BUB_MAX_ZEILENHOEHE As Double = 405
Private Const BUB_DETAILS_ZEILE As Long = 17
Private Const BUB_DETAILS_TITEL As String = "BUB-Details"


'==========================================================
' HAUPTMAKRO
'==========================================================

Public Sub KalenderNeuAktualisieren()

    Dim wsKal As Worksheet
    Dim newsDict As Object
    Dim regeln As Collection

    Dim strategie As String
    Dim istBub As Boolean
    Dim anzahlRegeln As Long
    Dim monat As Long
    Dim jahr As Long

    Dim calcAlt As XlCalculation
    Dim eventsAlt As Boolean
    Dim screenAlt As Boolean

    On Error GoTo Fehlerbehandlung

    Set wsKal = ThisWorkbook.Worksheets("Kalender")

    '------------------------------------------------------
    ' Einstellungen prüfen
    '------------------------------------------------------

    strategie = Trim(CStr(wsKal.Range("B1").Value))

    If strategie = "" Then
        MsgBox "Bitte in B1 eine Strategie auswählen.", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(wsKal.Range("B2").Value) Then
        MsgBox "Bitte in B2 die Anzahl der Regeln auswählen.", vbExclamation
        Exit Sub
    End If

    anzahlRegeln = CLng(wsKal.Range("B2").Value)

    If anzahlRegeln < 1 Then
        MsgBox "Die Anzahl der Regeln muss mindestens 1 sein.", vbExclamation
        Exit Sub
    End If

    monat = KalenderMonatNummer(Trim(CStr(wsKal.Range("B3").Value)))

    If monat = 0 Then
        MsgBox "Der Monat in B3 wurde nicht erkannt.", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(wsKal.Range("B4").Value) Then
        MsgBox "Bitte in B4 ein gültiges Jahr eintragen.", vbExclamation
        Exit Sub
    End If

    jahr = CLng(wsKal.Range("B4").Value)
    istBub = (StrComp(Left$(strategie & " ", 4), "BUB ", vbTextCompare) = 0)


    '------------------------------------------------------
    ' Excel-Einstellungen merken
    '------------------------------------------------------

    calcAlt = Application.Calculation
    eventsAlt = Application.EnableEvents
    screenAlt = Application.ScreenUpdating

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual


    '------------------------------------------------------
    ' News und Regeln laden
    '------------------------------------------------------

    Set newsDict = BaueNewsDictionary()

    Set regeln = KalenderRegelnLaden( _
        strategie, _
        anzahlRegeln, _
        istBub)


    If regeln.Count = 0 Then

        MsgBox _
            "Für '" & strategie & _
            "' wurden im Blatt Optimierung keine Regeln gefunden.", _
            vbExclamation

        GoTo SauberBeenden

    End If


    '------------------------------------------------------
    ' Kalender zeichnen
    '------------------------------------------------------

    KalenderMonatZeichnen _
        wsKal, _
        monat, _
        jahr, _
        regeln, _
        newsDict, _
        istBub


SauberBeenden:

    Application.Calculation = calcAlt
    Application.EnableEvents = eventsAlt
    Application.ScreenUpdating = screenAlt

    Exit Sub


Fehlerbehandlung:

    Dim fehlerNummer As Long
    Dim fehlerText As String

    fehlerNummer = Err.Number
    fehlerText = Err.Description

    On Error Resume Next

    Application.Calculation = calcAlt
    Application.EnableEvents = eventsAlt
    Application.ScreenUpdating = screenAlt

    MsgBox _
        "Fehler " & fehlerNummer & ":" & vbCrLf & _
        fehlerText, _
        vbCritical, _
        "Kalender"

End Sub



'==========================================================
' TOP-X-REGELN AUS OPTIMIERUNG LADEN
'==========================================================

Private Function KalenderRegelnLaden( _
    ByVal strategie As String, _
    ByVal anzahlRegeln As Long, _
    ByVal istBub As Boolean) As Collection

    Dim wsOpt As Worksheet
    Dim regeln As New Collection

    Dim letzteZeile As Long
    Dim i As Long
    Dim gefunden As Long

    Dim regel As Variant
    Dim newsName As String
    Dim abstand As Long
    Dim sessions As Long

    Set wsOpt = ThisWorkbook.Worksheets("Optimierung")

    letzteZeile = _
        wsOpt.Cells(wsOpt.Rows.Count, "A").End(xlUp).Row


    'Zeile 2 = Basis
    'Regeln beginnen ab Zeile 3

    For i = 3 To letzteZeile

        If StrComp( _
            Trim(CStr(wsOpt.Cells(i, "B").Value)), _
            strategie, _
            vbTextCompare) = 0 Then


            newsName = Trim(CStr(wsOpt.Cells(i, "C").Value))


            If newsName <> "" _
            And UCase(newsName) <> "BASIS" Then


                If IsNumeric(wsOpt.Cells(i, "D").Value) Then
                    abstand = CLng(wsOpt.Cells(i, "D").Value)
                Else
                    abstand = 0
                End If


                '0 = Newsname
                '1 = Abstand
                If istBub Then
                    sessions = KalenderBubRegelSessions( _
                        CStr(wsOpt.Cells(i, "N").Value), newsName, abstand)

                    If sessions = 0 Then
                        Err.Raise vbObjectError + 100, "KalenderRegelnLaden", _
                            "In Optimierung!N" & i & " fehlt eine gültige BUB-Session für " & newsName & "."
                    End If

                    regel = Array(newsName, abstand, sessions)
                Else
                    regel = Array(newsName, abstand)
                End If

                regeln.Add regel

                gefunden = gefunden + 1

                If gefunden >= anzahlRegeln Then
                    Exit For
                End If

            End If

        End If

    Next i


    Set KalenderRegelnLaden = regeln

End Function



'==========================================================
' MONAT ZEICHNEN UND TAGE BEWERTEN
'==========================================================

Private Sub KalenderMonatZeichnen( _
    ByVal ws As Worksheet, _
    ByVal monat As Long, _
    ByVal jahr As Long, _
    ByVal regeln As Collection, _
    ByVal newsDict As Object, _
    ByVal istBub As Boolean)

    Dim ersterTag As Date
    Dim letzterTag As Date
    Dim aktuellesDatum As Date

    Dim tag As Long
    Dim zeile As Long
    Dim spalte As Long

    Dim ersterWochentag As Long
    Dim gruende As String
    Dim bubGruende() As String


    ersterTag = DateSerial(jahr, monat, 1)
    letzterTag = DateSerial(jahr, monat + 1, 0)


    '------------------------------------------------------
    ' Alten Monat löschen
    '------------------------------------------------------

    KalenderBubDetailsLoeschen ws

    If Not istBub And IstBubKalenderAnzeige(ws) Then
        ' Die Wochenhöhen der bestehenden Monatsansicht wiederherstellen.
        ws.Rows("7:12").RowHeight = 69.95
        ws.Rows(9).RowHeight = 57
        ws.Range("A7:G12").WrapText = False
    End If

    With ws.Range("A7:G12")

        .ClearContents
        .Interior.Pattern = xlNone

    End With


    'Monatsüberschrift
    ws.Range("A5").Value = _
        KalenderMonatName(monat) & " " & jahr


    '1 = Montag ... 7 = Sonntag
    ersterWochentag = Weekday(ersterTag, vbMonday)

    zeile = 7
    spalte = ersterWochentag


    '------------------------------------------------------
    ' Tage eintragen
    '------------------------------------------------------

    For tag = 1 To Day(letzterTag)

        aktuellesDatum = DateSerial(jahr, monat, tag)

        ws.Cells(zeile, spalte).Value = tag
        With ws.Cells(zeile, spalte)

    .Font.Size = 11
    .Font.Bold = False
    .VerticalAlignment = xlTop
    .HorizontalAlignment = xlLeft

End With


        '--------------------------------------------------
        ' Wochenende
        '--------------------------------------------------

        If spalte = 6 Or spalte = 7 Then

            ws.Cells(zeile, spalte).Interior.Color = _
                RGB(217, 217, 217)


        ElseIf istBub Then

            KalenderBubTagGruende aktuellesDatum, regeln, newsDict, bubGruende
            KalenderBubTagZeichnen ws.Cells(zeile, spalte), tag, bubGruende

        Else

            '----------------------------------------------
            ' News-Regeln prüfen
            '----------------------------------------------

            gruende = KalenderTagGruende( _
                aktuellesDatum, _
                regeln, _
                newsDict)

If gruende <> "" Then

    ws.Cells(zeile, spalte).Interior.Color = _
        RGB(255, 199, 206)

    ws.Cells(zeile, spalte).Value = _
        tag & vbLf & KalenderGrundKurz(gruende)

    ws.Cells(zeile, spalte).WrapText = True
    ws.Cells(zeile, spalte).VerticalAlignment = xlTop
    ws.Cells(zeile, spalte).Value = _
    tag & vbLf & KalenderGrundKurz(gruende)

With ws.Cells(zeile, spalte)

    .WrapText = True
    .VerticalAlignment = xlTop
    .HorizontalAlignment = xlLeft

    .Font.Size = 9
    .Font.Bold = False

End With
With ws.Cells(zeile, spalte)

    .Characters(1, Len(CStr(tag))).Font.Bold = True
    .Characters(1, Len(CStr(tag))).Font.Size = 11

    If Len(.Value) > Len(CStr(tag)) Then

        .Characters(Len(CStr(tag)) + 1).Font.Size = 9

    End If

End With
Else

    ws.Cells(zeile, spalte).Interior.Color = _
        RGB(198, 239, 206)

    ws.Cells(zeile, spalte).Value = tag
    With ws.Cells(zeile, spalte)

    .Font.Size = 11
    .Font.Bold = False
    .VerticalAlignment = xlTop
    .HorizontalAlignment = xlLeft

End With

End If

        End If


        'Nächster Kalendertag
        spalte = spalte + 1

        If spalte > 7 Then
            spalte = 1
            zeile = zeile + 1
        End If

    Next tag

    If istBub Then KalenderBubZeilenAnpassen ws, monat, jahr

End Sub



'==========================================================
' PRÜFEN, WELCHE REGELN EINEN TAG SPERREN
'==========================================================

Private Function KalenderTagGruende( _
    ByVal kalenderDatum As Date, _
    ByVal regeln As Collection, _
    ByVal newsDict As Object) As String

    Dim regel As Variant

    Dim newsName As String
    Dim abstand As Long

    Dim newsDatum As Date
    Dim gruende As String


    For Each regel In regeln

        newsName = CStr(regel(0))
        abstand = CLng(regel(1))


        '--------------------------------------------------
        ' Beispiel:
        '
        ' Regel -1:
        ' Kalendertag liegt 1 Tag VOR der News.
        ' Also News am Kalendertag +1 suchen.
        '
        ' Deshalb:
        ' Newsdatum = Kalendertag - Abstand
        '--------------------------------------------------

        newsDatum = kalenderDatum - abstand


        If KalenderNewsVorhanden( _
            newsDict, _
            newsDatum, _
            newsName) Then


            If gruende <> "" Then
                gruende = gruende & " | "
            End If


            gruende = gruende & _
                newsName & _
                " [" & _
                IIf(abstand > 0, "+", "") & _
                abstand & "]"

        End If

    Next regel


    KalenderTagGruende = gruende

End Function



'==========================================================
' EINE BESTIMMTE NEWS AN EINEM DATUM SUCHEN
'==========================================================

Private Function KalenderNewsVorhanden( _
    ByVal newsDict As Object, _
    ByVal datum As Date, _
    ByVal gesuchteNews As String) As Boolean

    Dim newsText As String

    Dim newsEintraege As Variant
    Dim newsEintrag As Variant
    Dim teile As Variant

    Dim vorhandeneNews As String


    KalenderNewsVorhanden = False


    newsText = HoleNewsFuerDatum( _
        newsDict, _
        datum, _
        0)


    If Len(newsText) = 0 Then
        Exit Function
    End If


    newsEintraege = Split(newsText, " || ")


    For Each newsEintrag In newsEintraege

        teile = Split( _
            CStr(newsEintrag), _
            " | ")


        If UBound(teile) >= 3 Then

            vorhandeneNews = _
                Trim(CStr(teile(3)))


            If StrComp( _
                vorhandeneNews, _
                gesuchteNews, _
                vbTextCompare) = 0 Then

                KalenderNewsVorhanden = True

                Exit Function

            End If

        End If

    Next newsEintrag

End Function



'==========================================================
' MONATSNAME -> NUMMER
'==========================================================

Private Function KalenderMonatNummer( _
    ByVal monatText As String) As Long

    Select Case LCase(Trim(monatText))

        Case "januar"
            KalenderMonatNummer = 1

        Case "februar"
            KalenderMonatNummer = 2

        Case "märz", "maerz"
            KalenderMonatNummer = 3

        Case "april"
            KalenderMonatNummer = 4

        Case "mai"
            KalenderMonatNummer = 5

        Case "juni"
            KalenderMonatNummer = 6

        Case "juli"
            KalenderMonatNummer = 7

        Case "august"
            KalenderMonatNummer = 8

        Case "september"
            KalenderMonatNummer = 9

        Case "oktober"
            KalenderMonatNummer = 10

        Case "november"
            KalenderMonatNummer = 11

        Case "dezember"
            KalenderMonatNummer = 12

        Case Else
            KalenderMonatNummer = 0

    End Select

End Function



'==========================================================
' NUMMER -> DEUTSCHER MONATSNAME
'==========================================================

Private Function KalenderMonatName( _
    ByVal monat As Long) As String

    Select Case monat

        Case 1: KalenderMonatName = "Januar"
        Case 2: KalenderMonatName = "Februar"
        Case 3: KalenderMonatName = "März"
        Case 4: KalenderMonatName = "April"
        Case 5: KalenderMonatName = "Mai"
        Case 6: KalenderMonatName = "Juni"
        Case 7: KalenderMonatName = "Juli"
        Case 8: KalenderMonatName = "August"
        Case 9: KalenderMonatName = "September"
        Case 10: KalenderMonatName = "Oktober"
        Case 11: KalenderMonatName = "November"
        Case 12: KalenderMonatName = "Dezember"

    End Select

End Function

Private Function KalenderGrundKurz(ByVal gruende As String) As String

    Dim teile As Variant
    Dim ersterGrund As String
    Dim anzahl As Long
    Dim pos As Long

    teile = Split(gruende, " | ")

    anzahl = UBound(teile) + 1
    ersterGrund = Trim(CStr(teile(0)))

    'Abstand wie [0], [-1], [+2] entfernen
    pos = InStrRev(ersterGrund, " [")

    If pos > 0 Then
        ersterGrund = Left(ersterGrund, pos - 1)
    End If

    'Zu lange Namen kürzen
    If Len(ersterGrund) > 22 Then
        ersterGrund = Left(ersterGrund, 19) & "..."
    End If

    If anzahl > 1 Then

        KalenderGrundKurz = _
            ersterGrund & vbLf & _
            "+" & (anzahl - 1) & " weitere"

    Else

        KalenderGrundKurz = ersterGrund

    End If

End Function


'==========================================================
' BUB: SESSION-AUSSCHLÜSSE AUS DER KUMULIERTEN REGELLISTE
'==========================================================

Private Function KalenderBubRegelSessions( _
    ByVal aktiveAusschluesse As String, _
    ByVal newsName As String, _
    ByVal abstand As Long) As Long

    Dim eintraege As Variant
    Dim eintrag As Variant
    Dim text As String
    Dim regelText As String
    Dim posSession As Long
    Dim posAbstand As Long
    Dim abstandText As String

    aktiveAusschluesse = Trim$(aktiveAusschluesse)
    If Len(aktiveAusschluesse) = 0 Then Exit Function
    eintraege = Split(aktiveAusschluesse, " | ")

    ' Ein Eintrag kann auch nur aus der Sessionangabe in Klammern bestehen.
    If UBound(eintraege) = 0 Then
        text = Trim$(CStr(eintraege(0)))
        If Left$(text, 1) = "[" And Right$(text, 1) = "]" Then
            If InStr(2, text, "[") = 0 Then
                KalenderBubRegelSessions = KalenderBubSessionMaske(Mid$(text, 2, Len(text) - 2))
                Exit Function
            End If
        End If
    End If

    For Each eintrag In eintraege
        text = Trim$(CStr(eintrag))
        posSession = InStrRev(text, " [")
        If posSession > 0 And Right$(text, 1) = "]" Then
            regelText = Left$(text, posSession - 1)
            posAbstand = InStrRev(regelText, " [")

            If posAbstand > 0 And Right$(regelText, 1) = "]" Then
                abstandText = Trim$(Mid$(regelText, posAbstand + 2, Len(regelText) - posAbstand - 2))
                If Left$(abstandText, 1) = "+" Then abstandText = Mid$(abstandText, 2)

                If abstandText = CStr(abstand) And _
                    StrComp(Trim$(Left$(regelText, posAbstand - 1)), newsName, vbTextCompare) = 0 Then
                    KalenderBubRegelSessions = KalenderBubRegelSessions Or _
                        KalenderBubSessionMaske(Mid$(text, posSession + 2, Len(text) - posSession - 2))
                End If
            End If
        End If
    Next eintrag

End Function

Private Function KalenderBubSessionMaske(ByVal sessionText As String) As Long

    Dim session As Variant

    For Each session In Split(sessionText, "+")
        Select Case UCase$(Trim$(CStr(session)))
            Case "ALLE"
                KalenderBubSessionMaske = KalenderBubSessionMaske Or BUB_ALLE
            Case "03:00-06:00"
                KalenderBubSessionMaske = KalenderBubSessionMaske Or BUB_LONDON
            Case "09:50-12:10"
                KalenderBubSessionMaske = KalenderBubSessionMaske Or BUB_NEW_YORK
            Case "13:00-15:50"
                KalenderBubSessionMaske = KalenderBubSessionMaske Or BUB_PM
            Case Else
                KalenderBubSessionMaske = 0
                Exit Function
        End Select
    Next session

End Function

Private Sub KalenderBubTagGruende( _
    ByVal kalenderDatum As Date, _
    ByVal regeln As Collection, _
    ByVal newsDict As Object, _
    ByRef gruende() As String)

    Dim regel As Variant
    Dim newsDatum As Date
    Dim grund As String
    Dim sessionMasken As Variant
    Dim i As Long

    sessionMasken = Array(BUB_LONDON, BUB_NEW_YORK, BUB_PM)
    ReDim gruende(0 To UBound(sessionMasken))

    For Each regel In regeln
        newsDatum = kalenderDatum - CLng(regel(1))

        If KalenderNewsVorhanden(newsDict, newsDatum, CStr(regel(0))) Then
            grund = CStr(regel(0)) & " – " & KalenderBubAbstandText(kalenderDatum, newsDatum)

            For i = LBound(sessionMasken) To UBound(sessionMasken)
                If (CLng(regel(2)) And CLng(sessionMasken(i))) <> 0 Then
                    If InStr(1, vbLf & gruende(i), vbLf & grund & vbLf, vbTextCompare) = 0 Then
                        gruende(i) = gruende(i) & grund & vbLf
                    End If
                End If
            Next i
        End If
    Next regel

End Sub

Private Function KalenderBubAbstandText(ByVal kalenderDatum As Date, ByVal newsDatum As Date) As String

    Dim tage As Long
    tage = DateDiff("d", kalenderDatum, newsDatum)

    Select Case tage
        Case Is < -1: KalenderBubAbstandText = "war vor " & Abs(tage) & " Tagen"
        Case -1: KalenderBubAbstandText = "war vor 1 Tag"
        Case 0: KalenderBubAbstandText = "ist am selben Tag"
        Case 1: KalenderBubAbstandText = "kommt in 1 Tag"
        Case Else: KalenderBubAbstandText = "kommt in " & tage & " Tagen"
    End Select

End Function

Private Sub KalenderBubTagZeichnen(ByVal zelle As Range, ByVal tag As Long, ByRef gruende() As String)

    Dim sessions As Variant
    Dim punktPositionen(0 To 2) As Long
    Dim text As String
    Dim i As Long

    sessions = Array("London", "New York", "PM")
    text = CStr(tag) & vbLf

    For i = LBound(sessions) To UBound(sessions)
        text = text & CStr(sessions(i)) & " "
        punktPositionen(i) = Len(text) + 1
        text = text & ChrW(&H25CF) & vbLf & gruende(i)
    Next i

    With zelle
        .Value = Left$(text, Len(text) - 1)
        .WrapText = True
        .Font.Size = 9
        .Characters(1, Len(CStr(tag))).Font.Bold = True
        .Characters(1, Len(CStr(tag))).Font.Size = 11

        For i = LBound(sessions) To UBound(sessions)
            If Len(gruende(i)) = 0 Then
                .Characters(punktPositionen(i), 1).Font.Color = RGB(0, 128, 0)
            Else
                .Characters(punktPositionen(i), 1).Font.Color = RGB(192, 0, 0)
            End If
        Next i
    End With

End Sub

Private Function IstBubKalenderAnzeige(ByVal ws As Worksheet) As Boolean

    Dim zelle As Range

    For Each zelle In ws.Range("A7:E12")
        If VarType(zelle.Value2) = vbString Then
            If InStr(1, CStr(zelle.Value2), vbLf & "London " & ChrW(&H25CF), vbBinaryCompare) > 0 Then
                IstBubKalenderAnzeige = True
                Exit Function
            End If
        End If
    Next zelle

End Function

Private Sub KalenderBubDetailsLoeschen(ByVal ws As Worksheet)

    Dim titel As Variant
    Dim ueberschrift As Variant
    Dim anzahl As Variant
    Dim ende As Long

    titel = ws.Cells(BUB_DETAILS_ZEILE, "A").Value2
    ueberschrift = ws.Cells(BUB_DETAILS_ZEILE + 1, "C").Value2
    If VarType(titel) <> vbString Or VarType(ueberschrift) <> vbString Then Exit Sub
    If titel <> BUB_DETAILS_TITEL Or ueberschrift <> "News und Abstand" Then Exit Sub
    anzahl = ws.Cells(BUB_DETAILS_ZEILE, "C").Value2
    If Not IsNumeric(anzahl) Then Exit Sub
    If anzahl < 1 Or anzahl > ws.Rows.Count - BUB_DETAILS_ZEILE - 1 Then Exit Sub

    ende = BUB_DETAILS_ZEILE + 1 + CLng(anzahl)
    ws.Range(ws.Cells(BUB_DETAILS_ZEILE, "A"), ws.Cells(ende, "C")).Clear
    ws.Rows(BUB_DETAILS_ZEILE & ":" & ende).RowHeight = ws.StandardHeight

End Sub

Private Sub KalenderBubZeilenAnpassen(ByVal ws As Worksheet, ByVal monat As Long, ByVal jahr As Long)

    Dim zeile As Long
    Dim spalte As Long
    Dim tag As Long
    Dim text As String
    Dim zeilen As Variant
    Dim eintrag As Variant
    Dim session As String
    Dim details As New Collection
    Dim daten() As Variant
    Dim i As Long
    Dim hoehe As Double
    Dim bedarf As Double

    ws.Range("A7:G12").Rows.AutoFit

    For zeile = 7 To 12
        hoehe = ws.Rows(zeile).RowHeight
        If hoehe < 69.95 Then hoehe = 69.95
        For spalte = 1 To 5
            bedarf = KalenderBubZeilenbedarf(ws.Cells(zeile, spalte))
            If bedarf > hoehe Then hoehe = bedarf
        Next spalte

        If hoehe < BUB_MAX_ZEILENHOEHE Then
            ws.Rows(zeile).RowHeight = hoehe
        Else
            ws.Rows(zeile).RowHeight = BUB_MAX_ZEILENHOEHE
            ' Excel begrenzt die Zeilenhöhe. Lange Wochen zusätzlich unter dem Raster ausgeben.
            For spalte = 1 To 5
                text = CStr(ws.Cells(zeile, spalte).Value2)
                tag = Val(text)
                If tag > 0 Then
                    zeilen = Split(Mid$(text, InStr(text, vbLf) + 1), vbLf)
                    For Each eintrag In zeilen
                        If Right$(CStr(eintrag), 1) = ChrW(&H25CF) Then
                            session = Left$(CStr(eintrag), Len(eintrag) - 2)
                        Else
                            details.Add Array(DateSerial(jahr, monat, tag), session, CStr(eintrag))
                        End If
                    Next eintrag
                End If
            Next spalte
        End If
    Next zeile

    If details.Count = 0 Then Exit Sub

    With ws.Cells(BUB_DETAILS_ZEILE, "A").Resize(details.Count + 2, 3)
        If Application.WorksheetFunction.CountA(.Cells) > 0 Then
            MsgBox "Die BUB-Details passen nicht in das Monatsraster. " & _
                "Der Zusatzbereich ab A" & BUB_DETAILS_ZEILE & " ist bereits belegt.", vbExclamation
            Exit Sub
        End If

        .Cells(1, 1).Value = BUB_DETAILS_TITEL
        .Cells(1, 2).Value = "Anzahl Gründe"
        .Cells(1, 3).Value = details.Count
        .Cells(2, 1).Value = "Handelstag"
        .Cells(2, 2).Value = "Session"
        .Cells(2, 3).Value = "News und Abstand"
        .Cells(1, 1).Resize(2, 3).Font.Bold = True
        ReDim daten(1 To details.Count, 1 To 3)
        For i = 1 To details.Count
            eintrag = details(i)
            daten(i, 1) = eintrag(0)
            daten(i, 2) = eintrag(1)
            daten(i, 3) = eintrag(2)
        Next i
        With .Cells(3, 1).Resize(details.Count, 3)
            .Value = daten
            .Columns(1).NumberFormat = "dd.mm.yyyy"
            .Columns(3).WrapText = True
            .Columns(3).Font.Size = 9
            .Rows.AutoFit
            For i = 1 To details.Count
                bedarf = KalenderBubZeilenbedarf(.Cells(i, 3))
                If bedarf > .Rows(i).RowHeight Then .Rows(i).RowHeight = bedarf
            Next i
        End With
    End With

End Sub

Private Function KalenderBubZeilenbedarf(ByVal zelle As Range) As Double

    Dim textzeile As Variant
    Dim text As String
    Dim zeichenProZeile As Long
    Dim zeilen As Long

    text = CStr(zelle.Value2)
    If Len(text) = 0 Then Exit Function

    ' Sicherheitsabstand für umgebrochene Texte, falls AutoFit Rich Text unterschätzt.
    zeichenProZeile = Int((zelle.Width - 5) / 6)
    If zeichenProZeile < 1 Then zeichenProZeile = 1

    For Each textzeile In Split(text, vbLf)
        zeilen = zeilen + 1
        If Len(textzeile) > 0 Then zeilen = zeilen + (Len(textzeile) - 1) \ zeichenProZeile
    Next textzeile

    KalenderBubZeilenbedarf = 5 + 13 * zeilen

End Function
