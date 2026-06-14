' =====================================================================
'  SDS_match_Callout  --  VBScript-Portierung des Python-Tools
' =====================================================================
'
'  Was macht dieses Skript?
'  ------------------------
'  1. Liest eine Excel-Datei ein (Spalte A = Datum, B = Uhrzeit, C = Text).
'  2. Klassifiziert jede Zeile als "callout" oder "sds".
'  3. Wendet Suchmuster (Regex) an und zieht strukturierte Felder heraus
'     (PLZ, Ort, Straße, Hausnummer, Objekt, Schlagwort, ... ).
'  4. Vergleicht alle Datensätze paarweise: stimmen Felder überein, gibt es
'     pro Feld Punkte (Gewichte). Übersteigt die Summe einen Schwellwert und
'     liegen die Datensätze im selben Zeitfenster, gehören sie zum selben
'     "Fall".
'  5. Gruppiert die Datensätze zu Fällen (Union-Find) und gibt Fälle +
'     eine Trefferstatistik aus.
'
'  Aufruf (immer mit CSCRIPT, nicht WSCRIPT!):
'     cscript //nologo sds_match.vbs <datei.xlsx> [Optionen]
'
'  Optionen:
'     /sheet:<name|index>   Tabellenblatt (Index 0 = erstes Blatt, Default 0)
'     /only:a,b,c           Nur diese Suchmuster aktivieren
'     /skip:a,b             Diese Suchmuster deaktivieren
'     /list                 Verfügbare Suchmuster auflisten und beenden
'     /show-content         Original-Text mit ausgeben
'     /threshold:0.5        Score-Schwellwert (überschreibt Default)
'     /time-window:120      Zeitfenster in Minuten (überschreibt Default)
'     /only-multi           Nur Fälle mit >= 2 Datensätzen anzeigen
'
'  WICHTIG zur Datei-Kodierung:
'     Die Umlaute in den Regex-Mustern werden im Code bewusst mit ChrW(...)
'     gebaut (siehe InitGlobals). Dadurch funktioniert das MATCHING
'     unabhängig davon, in welcher Kodierung die Datei gespeichert ist.
'     Die deutschen Texte in der Konsolenausgabe sind hingegen direkt als
'     Umlaut geschrieben -- erscheinen sie dort als Buchstabensalat, dann ist
'     nur die Anzeige betroffen, nicht das Ergebnis. Speichere die Datei in
'     diesem Fall als ANSI (Windows-1252) oder UTF-16, dann passt auch das.
'
'  Erläuterung jedes einzelnen Regex: siehe README.md in diesem Ordner.
' =====================================================================

Option Explicit

' ---------------------------------------------------------------------
'  KONFIGURATION  (im Python-Original lag das in config.yaml)
' ---------------------------------------------------------------------
Const CALLOUT_PREFIX     = "IncomingCallout:"  ' Zeilen, die so beginnen = Callout, sonst SDS
Const TIME_WINDOW_MIN    = 120                  ' Zeitfenster in Minuten
Const SCORE_THRESHOLD    = 0.5                  ' ab dieser Punktzahl gilt ein Paar als "zusammengehörig"

' Gewichte je Suchmuster werden in BuildExtractors() gesetzt (Spalte "Gewicht").
' Ein Gewicht von 0 bedeutet: zählt nicht fürs Matching (nur Anzeige).

' ---------------------------------------------------------------------
'  GLOBALE VARIABLEN (werden in InitGlobals / InitRegex befüllt)
' ---------------------------------------------------------------------
Dim g_ae, g_oe, g_ue, g_Ae, g_Oe, g_Ue, g_sz      ' Umlaut-Zeichen (per ChrW)
Dim g_umlLower, g_umlAll, g_umlSchlag             ' fertige Zeichenklassen-Teile
Dim g_STRASSE                                     ' das Ersetzungswort "straße"

Dim g_reSdsAddr, g_reSdsOt, g_reSdsObj
Dim g_reCoFull, g_reCoPlzOrt, g_reCoObjekt
Dim g_reSchlagCo, g_reSchlagSds
Dim g_reStichCo, g_reStichSds
Dim g_rePipeBlocks, g_reSchlagToken
Dim g_reStr, g_reStrasse, g_reWs

Dim g_addrCache   ' Cache für ParseAddress (entspricht @lru_cache in Python)


' =====================================================================
'  EINSTIEGSPUNKT
' =====================================================================
Call Main()


Sub Main()
    CheckHost()
    InitGlobals()
    InitRegex()
    Set g_addrCache = CreateObject("Scripting.Dictionary")

    ' --- Argumente lesen ---
    Dim xlsxPath : xlsxPath = ""
    If WScript.Arguments.Unnamed.Count > 0 Then xlsxPath = WScript.Arguments.Unnamed(0)

    Dim named : Set named = WScript.Arguments.Named
    Dim doList      : doList      = named.Exists("list")
    Dim showContent : showContent = named.Exists("show-content")
    Dim onlyMulti   : onlyMulti   = named.Exists("only-multi")
    Dim onlyArg  : onlyArg  = "" : If named.Exists("only")  Then onlyArg  = named("only")
    Dim skipArg  : skipArg  = "" : If named.Exists("skip")  Then skipArg  = named("skip")
    Dim sheetArg : sheetArg = "" : If named.Exists("sheet") Then sheetArg = named("sheet")

    Dim threshold : threshold = SCORE_THRESHOLD
    If named.Exists("threshold") Then threshold = ToDouble(named("threshold"))
    Dim window : window = TIME_WINDOW_MIN
    If named.Exists("time-window") Then window = CLng(named("time-window"))

    ' --- Suchmuster aufbauen ---
    Dim extractors : extractors = BuildExtractors()

    If doList Then
        PrintList extractors
        Exit Sub
    End If

    If xlsxPath = "" Then
        Die "Pfad zur Excel-Datei fehlt (oder /list verwenden)."
    End If
    If Not FileExists(xlsxPath) Then
        Die "Datei nicht gefunden: " & xlsxPath
    End If

    ' --- only/skip anwenden ---
    ConfigureExtractors extractors, onlyArg, skipArg
    Dim active : active = FilterActive(extractors)
    If NumElems(active) = 0 Then
        Die "Keine Suchmuster aktiv -- nichts zu tun."
    End If

    ' --- Gewichte der aktiven Muster sammeln ---
    Dim weights : Set weights = CreateObject("Scripting.Dictionary")
    Dim ex
    For Each ex In active
        If ex.Weight > 0 Then weights.Add ex.Name, ex.Weight
    Next

    ' --- Excel einlesen ---
    Dim recs : recs = LoadRecords(xlsxPath, sheetArg, active)
    Dim n : n = NumElems(recs)

    Dim n_co, n_sds : n_co = 0 : n_sds = 0
    Dim r
    For Each r In recs
        If r.Kind = "callout" Then n_co = n_co + 1 Else n_sds = n_sds + 1
    Next

    WScript.Echo "Gelesen: " & n & " Datensätze (" & n_co & " Callouts, " & n_sds & " SDS)"
    WScript.Echo "Aktive Suchmuster: " & JoinNames(active)
    WScript.Echo "Zeitfenster: " & window & " min   Schwellwert: " & Fmt2(threshold)

    ' --- Fälle bilden ---
    Dim cases : cases = FindCases(recs, active, weights, window, threshold)

    Dim multi, single : multi = 0 : single = 0
    Dim c
    For Each c In cases
        If NumElems(c.Records) >= 2 Then multi = multi + 1 Else single = single + 1
    Next
    WScript.Echo "Fälle: " & NumElems(cases) & " (" & multi & " mit mehreren Records, " & single & " Singletons)"

    ' --- Ausgabe ---
    PrintCases cases, active, showContent, onlyMulti
    PrintStats recs, active
End Sub


' =====================================================================
'  INITIALISIERUNG
' =====================================================================

' Stellt sicher, dass das Skript unter CSCRIPT (Konsole) läuft -- unter
' WSCRIPT würde jede Ausgabe ein eigenes Fenster öffnen.
Sub CheckHost()
    If InStr(1, WScript.FullName, "cscript", vbTextCompare) = 0 Then
        WScript.Echo "Bitte mit cscript starten:" & vbCrLf & _
                     "  cscript //nologo sds_match.vbs <datei.xlsx>"
        WScript.Quit 1
    End If
End Sub

' Baut die Umlaut-Zeichen kodierungsunabhängig über ChrW (Unicode-Codepoint).
Sub InitGlobals()
    g_ae = ChrW(228) : g_oe = ChrW(246) : g_ue = ChrW(252)   ' ä ö ü
    g_Ae = ChrW(196) : g_Oe = ChrW(214) : g_Ue = ChrW(220)   ' Ä Ö Ü
    g_sz = ChrW(223)                                          ' ß

    g_umlLower  = g_ae & g_oe & g_ue & g_sz                              ' äöüß
    g_umlAll    = g_Ae & g_Oe & g_Ue & g_ae & g_oe & g_ue & g_sz        ' ÄÖÜäöüß
    g_umlSchlag = g_Oe & g_oe & g_Ae & g_ae & g_sz                      ' ÖöÄäß

    g_STRASSE = "stra" & g_sz & "e"                                     ' "straße"
End Sub

' Erzeugt ein RegExp-Objekt. ic = IgnoreCase, gl = Global (alle Treffer).
Function NewRegex(pat, ic, gl)
    Dim re : Set re = New RegExp
    re.Pattern    = pat
    re.IgnoreCase = ic
    re.Global     = gl
    Set NewRegex = re
End Function

' Legt alle Regex-Objekte an. Die "logische" Form (mit echten Umlauten) steht
' jeweils im Kommentar; die README erklärt jedes Muster Zeichen für Zeichen.
Sub InitRegex()
    ' Hausnummer-Bauteil (kein eigener Regex, wird in CO_FULL eingebettet):
    '   logisch:  \d+\s?[a-zäöüß]{0,2}(?:\s?[/\-]\s?\d+\s?[a-zäöüß]{0,2})?
    Dim hausnr
    hausnr = "\d+\s?[a-z" & g_umlLower & "]{0,2}" & _
             "(?:\s?[/\-]\s?\d+\s?[a-z" & g_umlLower & "]{0,2})?"

    ' --- SDS-Adresse ---  ||EO: 82515 Wolfratshausen; Margeritenstraße; 22a||
    '   logisch:  \|\|\s*EO:\s*(\d{5})\s+([^;|]+?)\s*;\s*([^;|]+?)\s*;\s*([^|]+?)\s*\|\|
    '   Gruppen:  1=PLZ  2=Ort  3=Straße  4=Hausnummer
    Set g_reSdsAddr = NewRegex( _
        "\|\|\s*EO:\s*(\d{5})\s+([^;|]+?)\s*;\s*([^;|]+?)\s*;\s*([^|]+?)\s*\|\|", _
        True, False)

    ' --- SDS-Ortsteil ---  ||OT: Wolfratshausen||   (Gruppe 1 = Ort)
    Set g_reSdsOt  = NewRegex("\|\|\s*OT:\s*([^|]+?)\s*\|\|",  True, False)
    ' --- SDS-Objekt ---    ||OBJ: Kirche||          (Gruppe 1 = Objekt)
    Set g_reSdsObj = NewRegex("\|\|\s*OBJ:\s*([^|]+?)\s*\|\|", True, False)

    ' --- Callout-Volladresse ---  ||82515 Wolfratshausen - Wolfratshausen||Margeritenstraße 22a  ...||
    '   logisch:  \|\|(\d{5})\s+([^|]+?)\|\|([A-Za-zÄÖÜäöüß][A-Za-zÄÖÜäöüß.\-]*(?:[ \-][A-Za-zÄÖÜäöüß.\-]+)*?)\s+(HAUSNR)\b
    '   Gruppen:  1=PLZ  2=Ort_roh  3=Straße  4=Hausnummer
    Set g_reCoFull = NewRegex( _
        "\|\|(\d{5})\s+([^|]+?)\|\|" & _
        "([A-Za-z" & g_umlAll & "][A-Za-z" & g_umlAll & ".\-]*" & _
        "(?:[ \-][A-Za-z" & g_umlAll & ".\-]+)*?)" & _
        "\s+(" & hausnr & ")\b", _
        False, False)

    ' --- Callout nur PLZ+Ort (Fallback) ---  ||82515 Wolfratshausen||
    Set g_reCoPlzOrt = NewRegex("\|\|(\d{5})\s+([^|]+?)\|\|", False, False)

    ' --- Callout-Objekt direkt nach der Straße ---  ...||Kirche||...
    '   logisch (am Anfang des Reststrings verankert):
    '     ^[^|]*\|\|([A-Za-zÄÖÜäöüß][A-Za-zÄÖÜäöüß .\-]*?)\|\|
    Set g_reCoObjekt = NewRegex( _
        "^[^|]*\|\|([A-Za-z" & g_umlAll & "][A-Za-z" & g_umlAll & " .\-]*?)\|\|", _
        False, False)

    ' --- Schlagwort ---  Callout: |#T2410#...|   SDS: |+SW: #T2410#...|+
    Set g_reSchlagCo  = NewRegex("\|(#[TBIR]\d{4}[a-zA-Z"     & g_umlSchlag & "#\- ]*)\|+",  False, False)
    Set g_reSchlagSds = NewRegex("\|+SW: (#[BTIR]\d{4}[a-zA-Z" & g_umlSchlag & "#\- ]*)\|+", False, False)

    ' --- Stichwort ---  Platzhalter, matcht (noch) nichts. "(?!x)x" = unmögliches Muster.
    Set g_reStichCo  = NewRegex("(?!x)x", False, False)   ' TODO: Callout-Stichwort-Regex eintragen
    Set g_reStichSds = NewRegex("(?!x)x", False, False)   ' TODO: SDS-Stichwort-Regex eintragen

    ' --- Einsatzmittel ---  alle ||...||-Blöcke (überlappend via Lookahead)
    Set g_rePipeBlocks   = NewRegex("\|\|\s*([^|]+?)\s*(?=\|\|)", False, True)
    Set g_reSchlagToken  = NewRegex("^#[TBIR]\d{4}", False, False)

    ' --- Normalisierung ---  "str." / "strasse" -> "straße"; Mehrfach-Leerzeichen -> eins
    Set g_reStr     = NewRegex("\bstr\.?\b",  True, True)
    Set g_reStrasse = NewRegex("\bstrasse\b", True, True)
    Set g_reWs      = NewRegex("\s+",         False, True)
End Sub


' =====================================================================
'  KLASSEN  (entsprechen den @dataclass / Klassen in Python)
' =====================================================================

Class Extractor
    Public Name
    Public Func          ' Funktionsreferenz (per GetRef)
    Public Description
    Public SortOrder
    Public DisplayMode   ' "inline" oder "extra_line"
    Public Enabled
    Public Weight
End Class

Class Record
    Public Timestamp     ' Date oder Empty
    Public Kind          ' "callout" | "sds"
    Public Content
    Public Fields        ' Dictionary: Mustername -> extrahierter Wert ("" = kein Treffer)
End Class

Class CaseRec
    Public Records       ' Array von Record
    Public ScoreMax
    Public ScoreAvg
End Class

' Union-Find: gruppiert Indizes, die "zusammengehören", in Cluster.
Class UnionFind
    Private parent()
    Public Sub Init(n)
        ReDim parent(n - 1)
        Dim i : For i = 0 To n - 1 : parent(i) = i : Next
    End Sub
    Public Function Find(x)
        Do While parent(x) <> x
            parent(x) = parent(parent(x))   ' Pfad-Halbierung
            x = parent(x)
        Loop
        Find = x
    End Function
    Public Sub Union(a, b)
        Dim ra, rb : ra = Find(a) : rb = Find(b)
        If ra <> rb Then parent(ra) = rb
    End Sub
End Class


' =====================================================================
'  SUCHMUSTER-DEFINITION
' =====================================================================

Function MakeEx(nm, fnName, desc, ordr, dmode, wt)
    Dim e : Set e = New Extractor
    e.Name = nm
    Set e.Func = GetRef(fnName)
    e.Description = desc
    e.SortOrder = ordr
    e.DisplayMode = dmode
    e.Enabled = True
    e.Weight = wt
    Set MakeEx = e
End Function

' Baut die Liste aller Suchmuster (entspricht patterns/*.py + den Gewichten
' aus config.yaml). Sortiert nach SortOrder, dann Name.
Function BuildExtractors()
    Dim arr(7)
    Set arr(0) = MakeEx("plz",          "ExtractPlz",          "PLZ -- SDS aus EO:-Block, Callout aus ||PLZ Ort||-Block",                     10, "inline",     0.25)
    Set arr(1) = MakeEx("street",       "ExtractStreet",       "Straße ohne Hausnummer -- SDS aus EO:-Block, Callout aus ||Straße Hausnr||",  20, "inline",     0.15)
    Set arr(2) = MakeEx("house_number", "ExtractHouseNumber",  "Hausnummer (Zahl + optional Buchstaben/Bruch)",                              25, "inline",     0.10)
    Set arr(3) = MakeEx("city",         "ExtractCity",         "Ort -- SDS aus ||OT:|| bzw. ||EO:||, Callout aus ||PLZ Ort||",               30, "inline",     0.05)
    Set arr(4) = MakeEx("objekt",       "ExtractObjekt",       "Objekt -- SDS aus ||OBJ:||, Callout aus Block direkt nach der Straße",       35, "inline",     0.10)
    Set arr(5) = MakeEx("schlagwort",   "ExtractSchlagwort",   "Schlagwort zwischen Pipes (B#/T#/I#/R# + 4 Ziffern)",                        40, "inline",     0.25)
    Set arr(6) = MakeEx("stichwort",    "ExtractStichwort",    "Stichwort (Callout und SDS, eigene Regex je kind) -- noch Platzhalter",      50, "inline",     0.10)
    Set arr(7) = MakeEx("einsatzmittel","ExtractEinsatzmittel","Einsatzmittel (nur Callout, mehrere durch ' | ' getrennt)",                  90, "extra_line", 0.0)

    ' Insertion-Sort nach (SortOrder, Name)
    Dim i, j, key
    For i = 1 To UBound(arr)
        Set key = arr(i)
        j = i - 1
        Do While j >= 0
            If ExGreater(arr(j), key) Then
                Set arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        Set arr(j + 1) = key
    Next
    BuildExtractors = arr
End Function

' True, wenn a in der Sortierung hinter b stehen soll.
Function ExGreater(a, b)
    If a.SortOrder > b.SortOrder Then
        ExGreater = True
    ElseIf a.SortOrder < b.SortOrder Then
        ExGreater = False
    Else
        ExGreater = (a.Name > b.Name)
    End If
End Function


' =====================================================================
'  EXTRAKTOREN  (eine Funktion pro Suchmuster)
'  Jede liefert den Wert als String oder "" (= kein Treffer).
' =====================================================================

Function ExtractPlz(content, kind)
    Dim d : Set d = ParseAddress(content, kind)
    ExtractPlz = d("plz")
End Function

Function ExtractCity(content, kind)
    Dim d : Set d = ParseAddress(content, kind)
    ExtractCity = d("ort")
End Function

Function ExtractStreet(content, kind)
    Dim d : Set d = ParseAddress(content, kind)
    ExtractStreet = d("strasse")
End Function

Function ExtractHouseNumber(content, kind)
    Dim d : Set d = ParseAddress(content, kind)
    ExtractHouseNumber = d("hausnummer")
End Function

Function ExtractObjekt(content, kind)
    Dim d : Set d = ParseAddress(content, kind)
    ExtractObjekt = d("objekt")
End Function

Function ExtractSchlagwort(content, kind)
    Dim re
    If kind = "sds" Then Set re = g_reSchlagSds Else Set re = g_reSchlagCo
    Dim ms : Set ms = re.Execute(content)
    If ms.Count > 0 Then ExtractSchlagwort = ms(0).SubMatches(0) Else ExtractSchlagwort = ""
End Function

Function ExtractStichwort(content, kind)
    Dim re
    If kind = "sds" Then Set re = g_reStichSds Else Set re = g_reStichCo
    Dim ms : Set ms = re.Execute(content)
    If ms.Count > 0 Then ExtractStichwort = ms(0).SubMatches(0) Else ExtractStichwort = ""
End Function

Function ExtractEinsatzmittel(content, kind)
    ExtractEinsatzmittel = ""
    If kind <> "callout" Then Exit Function

    Dim addr : Set addr = ParseAddress(content, kind)
    Dim plz, strasse, objekt
    plz     = addr("plz")
    strasse = addr("strasse")
    objekt  = addr("objekt")

    Dim ms : Set ms = g_rePipeBlocks.Execute(content)
    Dim parts() : ReDim parts(ms.Count) : Dim cnt : cnt = 0
    Dim m, token
    For Each m In ms
        token = Trim(m.SubMatches(0))
        If token = "" Then
            ' überspringen
        ElseIf plz <> "" And Left(token, Len(plz)) = plz Then
            ' PLZ-Ort-Block überspringen
        ElseIf strasse <> "" And Left(token, Len(strasse)) = strasse Then
            ' Straßen-Block überspringen
        ElseIf objekt <> "" And token = objekt Then
            ' Objekt-Block überspringen
        ElseIf g_reSchlagToken.Test(token) Then
            ' Schlagwort-Block überspringen
        Else
            parts(cnt) = token
            cnt = cnt + 1
        End If
    Next
    If cnt = 0 Then Exit Function
    ReDim Preserve parts(cnt - 1)
    ExtractEinsatzmittel = Join(parts, " | ")
End Function


' =====================================================================
'  ADRESS-PARSER  (entspricht patterns/_address.py)
'  Gecacht: pro (kind, content) läuft das Parsen nur einmal.
' =====================================================================

Function ParseAddress(content, kind)
    Dim key : key = kind & Chr(1) & content
    If g_addrCache.Exists(key) Then
        Set ParseAddress = g_addrCache(key)
        Exit Function
    End If

    Dim d : Set d = NewAddrDict()
    Dim ms, m

    If kind = "sds" Then
        Set ms = g_reSdsAddr.Execute(content)
        If ms.Count > 0 Then
            Set m = ms(0)
            d("plz")        = Trim(m.SubMatches(0))
            d("ort")        = Trim(m.SubMatches(1))
            d("strasse")    = Trim(m.SubMatches(2))
            d("hausnummer") = Trim(m.SubMatches(3))
        End If
        ' ||OT: ...|| hat Vorrang für den Ort
        Set ms = g_reSdsOt.Execute(content)
        If ms.Count > 0 Then d("ort") = Trim(ms(0).SubMatches(0))
        ' optionales ||OBJ: ...||
        Set ms = g_reSdsObj.Execute(content)
        If ms.Count > 0 Then d("objekt") = Trim(ms(0).SubMatches(0))

    ElseIf kind = "callout" Then
        Set ms = g_reCoFull.Execute(content)
        If ms.Count > 0 Then
            Set m = ms(0)
            d("plz")        = m.SubMatches(0)
            d("ort")        = CoOrtFromRaw(m.SubMatches(1))
            d("strasse")    = Trim(m.SubMatches(2))
            d("hausnummer") = g_reWs.Replace(m.SubMatches(3), "")   ' alle Leerzeichen raus
            ' direkt nach dem Straßen-Block nach einem ||Objekt||-Block suchen
            Dim rest : rest = Mid(content, m.FirstIndex + m.Length + 1)
            Dim obm : Set obm = g_reCoObjekt.Execute(rest)
            If obm.Count > 0 Then d("objekt") = Trim(obm(0).SubMatches(0))
        Else
            Set ms = g_reCoPlzOrt.Execute(content)
            If ms.Count > 0 Then
                d("plz") = ms(0).SubMatches(0)
                d("ort") = CoOrtFromRaw(ms(0).SubMatches(1))
            End If
        End If
    End If

    g_addrCache.Add key, d
    Set ParseAddress = d
End Function

Function NewAddrDict()
    Dim d : Set d = CreateObject("Scripting.Dictionary")
    d.Add "plz", ""
    d.Add "ort", ""
    d.Add "strasse", ""
    d.Add "hausnummer", ""
    d.Add "objekt", ""
    Set NewAddrDict = d
End Function

' "Ortsname - Gemeinde" -> "Ortsname"
Function CoOrtFromRaw(raw)
    Dim p : p = InStr(raw, " - ")
    If p > 0 Then
        CoOrtFromRaw = Trim(Left(raw, p - 1))
    Else
        CoOrtFromRaw = Trim(raw)
    End If
End Function


' =====================================================================
'  KLASSIFIKATION + NORMALISIERUNG
' =====================================================================

Function Classify(content, prefix)
    If Left(LTrim(content), Len(prefix)) = prefix Then
        Classify = "callout"
    Else
        Classify = "sds"
    End If
End Function

' strip + lower + Whitespace zusammenfassen + "str."/"strasse" -> "straße"
Function NormalizeValue(value)
    Dim v : v = LCase(Trim(value))
    v = g_reStr.Replace(v, g_STRASSE)
    v = g_reStrasse.Replace(v, g_STRASSE)
    v = g_reWs.Replace(v, " ")
    NormalizeValue = v
End Function


' =====================================================================
'  EXCEL EINLESEN
' =====================================================================

Function LoadRecords(path, sheetArg, active)
    Dim xl, wb, ws
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False

    On Error Resume Next
    Set wb = xl.Workbooks.Open(AbsPath(path), False, True)   ' UpdateLinks=False, ReadOnly=True
    If wb Is Nothing Then
        On Error GoTo 0
        xl.Quit
        Die "Excel-Datei konnte nicht geöffnet werden: " & path
    End If
    On Error GoTo 0

    If sheetArg = "" Then
        Set ws = wb.Worksheets(1)
    ElseIf IsNumeric(sheetArg) Then
        Set ws = wb.Worksheets(CLng(sheetArg) + 1)   ' Index 0 -> erstes Blatt
    Else
        Set ws = wb.Worksheets(sheetArg)
    End If

    Dim lastRow : lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    Dim data
    If lastRow >= 1 Then data = ws.Range("A1:C" & lastRow).Value

    Dim recs() : ReDim recs(15) : Dim cnt : cnt = 0
    Dim r, content, rec, ex, fn
    For r = 1 To lastRow
        content = CellToStr(data(r, 3))
        If Trim(content) <> "" Then
            Set rec = New Record
            rec.Timestamp = ParseTimestamp(data(r, 1), data(r, 2))
            rec.Kind = Classify(content, CALLOUT_PREFIX)
            rec.Content = content
            Set rec.Fields = CreateObject("Scripting.Dictionary")
            For Each ex In active
                Set fn = ex.Func
                rec.Fields(ex.Name) = fn(content, rec.Kind)
            Next
            If cnt > UBound(recs) Then ReDim Preserve recs(UBound(recs) * 2 + 1)
            Set recs(cnt) = rec
            cnt = cnt + 1
        End If
    Next

    wb.Close False
    xl.Quit
    Set ws = Nothing : Set wb = Nothing : Set xl = Nothing

    If cnt = 0 Then
        LoadRecords = Array()
    Else
        ReDim Preserve recs(cnt - 1)
        LoadRecords = recs
    End If
End Function

' Wandelt einen Zellwert sicher in einen String um.
Function CellToStr(v)
    If IsNull(v) Or IsEmpty(v) Then
        CellToStr = ""
    Else
        CellToStr = CStr(v)
    End If
End Function

' Datum + Uhrzeit zu einem Date-Wert. Hinweis: VBScript-Date kennt keine
' Millisekunden -- die werden verworfen (für 2-Stunden-Fenster irrelevant).
Function ParseTimestamp(dv, tv)
    ParseTimestamp = Empty
    If IsNull(dv) Or IsEmpty(dv) Or IsNull(tv) Or IsEmpty(tv) Then Exit Function

    Dim d, t
    ' --- Datum ---
    If IsNumeric(dv) Then
        d = CDate(CDbl(dv))
    ElseIf IsDate(dv) Then
        d = CDate(dv)
    Else
        d = ParseDateStr(CleanCell(CStr(dv)))
        If IsEmpty(d) Then Exit Function
    End If
    ' --- Uhrzeit ---
    If IsNumeric(tv) Then
        t = CDate(CDbl(tv))
    ElseIf IsDate(tv) Then
        t = CDate(tv)
    Else
        t = ParseTimeStr(Replace(CleanCell(CStr(tv)), ",", "."))
        If IsEmpty(t) Then Exit Function
    End If

    ParseTimestamp = DateSerial(Year(d), Month(d), Day(d)) + _
                     TimeSerial(Hour(t), Minute(t), Second(t))
End Function

Function ParseDateStr(s)        ' erwartet "TT.MM.JJJJ"
    ParseDateStr = Empty
    Dim p : p = Split(s, ".")
    If UBound(p) < 2 Then Exit Function
    If Not (IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2))) Then Exit Function
    ParseDateStr = DateSerial(CLng(p(2)), CLng(p(1)), CLng(p(0)))
End Function

Function ParseTimeStr(s)        ' erwartet "HH:MM:SS" oder "HH:MM:SS.fff"
    ParseTimeStr = Empty
    Dim p : p = Split(s, ":")
    If UBound(p) < 2 Then Exit Function
    Dim sec : sec = p(2)
    Dim dp : dp = InStr(sec, ".")
    If dp > 0 Then sec = Left(sec, dp - 1)
    If Not (IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(sec)) Then Exit Function
    ParseTimeStr = TimeSerial(CLng(p(0)), CLng(p(1)), CLng(sec))
End Function

Function CleanCell(s)
    s = Trim(CStr(s))
    If Left(s, 1) = "'" Then s = Mid(s, 2)
    CleanCell = s
End Function


' =====================================================================
'  PAARWEISER SCORE + FALL-BILDUNG  (entspricht matcher.py)
' =====================================================================

' Summe der Gewichte über Felder, in denen beide Records denselben
' normalisierten Wert haben. Nur "inline"-Muster zählen.
Function PairScore(a, b, weights, byName)
    Dim total : total = 0
    Dim nm, ex, va, vb
    For Each nm In weights.Keys
        If byName.Exists(nm) Then
            Set ex = byName(nm)
            If ex.DisplayMode = "inline" Then
                va = FieldVal(a, nm)
                vb = FieldVal(b, nm)
                If va <> "" And vb <> "" Then
                    If NormalizeValue(va) = NormalizeValue(vb) Then
                        total = total + weights(nm)
                    End If
                End If
            End If
        End If
    Next
    PairScore = total
End Function

Function FieldVal(rec, nm)
    If rec.Fields.Exists(nm) Then FieldVal = rec.Fields(nm) Else FieldVal = ""
End Function

Function FindCases(recs, active, weights, windowMin, threshold)
    Dim n : n = NumElems(recs)
    If n = 0 Then
        FindCases = Array()
        Exit Function
    End If

    ' Name -> Extractor
    Dim byName : Set byName = CreateObject("Scripting.Dictionary")
    Dim ex
    For Each ex In active
        If Not byName.Exists(ex.Name) Then byName.Add ex.Name, ex
    Next

    ' Reihenfolge: nach Zeit aufsteigend, Records ohne Zeit ans Ende
    Dim order() : ReDim order(n - 1)
    Dim i : For i = 0 To n - 1 : order(i) = i : Next
    Dim a, key, jj
    For a = 1 To n - 1
        key = order(a)
        jj = a - 1
        Do While jj >= 0
            If CmpRec(recs(order(jj)), recs(key)) > 0 Then
                order(jj + 1) = order(jj)
                jj = jj - 1
            Else
                Exit Do
            End If
        Loop
        order(jj + 1) = key
    Next

    Dim uf : Set uf = New UnionFind : uf.Init n
    Dim pairScores : Set pairScores = CreateObject("Scripting.Dictionary")
    Dim windowSec : windowSec = windowMin * 60

    Dim pos, k, ii, jx, ri, rj, s, lo, hi
    For pos = 0 To n - 1
        ii = order(pos)
        Set ri = recs(ii)
        For k = pos + 1 To n - 1
            jx = order(k)
            Set rj = recs(jx)
            ' Zeitfenster nur prüfen, wenn beide eine Zeit haben
            If Not IsEmpty(ri.Timestamp) And Not IsEmpty(rj.Timestamp) Then
                If DateDiff("s", ri.Timestamp, rj.Timestamp) > windowSec Then Exit For
            End If
            s = PairScore(ri, rj, weights, byName)
            If s >= threshold Then
                uf.Union ii, jx
                If ii < jx Then lo = ii : hi = jx Else lo = jx : hi = ii
                pairScores(lo & "_" & hi) = s
            End If
        Next
    Next

    ' Cluster sammeln: Wurzel -> kommaseparierte Indexliste
    Dim clusters : Set clusters = CreateObject("Scripting.Dictionary")
    Dim root
    For i = 0 To n - 1
        root = uf.Find(i)
        If clusters.Exists(root) Then
            clusters(root) = clusters(root) & "," & i
        Else
            clusters.Add root, CStr(i)
        End If
    Next

    ' Fälle bauen
    Dim cases() : ReDim cases(clusters.Count - 1) : Dim ci : ci = 0
    Dim rkey, members, x, y, mi, mj, lo2, hi2, pk, sv, sumS, maxS, cntS, c, crecs()
    For Each rkey In clusters.Keys
        members = Split(clusters(rkey), ",")
        sumS = 0 : maxS = 0 : cntS = 0
        For x = 0 To UBound(members)
            For y = x + 1 To UBound(members)
                mi = CLng(members(x)) : mj = CLng(members(y))
                If mi < mj Then lo2 = mi : hi2 = mj Else lo2 = mj : hi2 = mi
                pk = lo2 & "_" & hi2
                If pairScores.Exists(pk) Then
                    sv = pairScores(pk)
                    sumS = sumS + sv
                    cntS = cntS + 1
                    If sv > maxS Then maxS = sv
                End If
            Next
        Next
        Set c = New CaseRec
        ReDim crecs(UBound(members))
        For x = 0 To UBound(members)
            Set crecs(x) = recs(CLng(members(x)))
        Next
        c.Records = crecs
        c.ScoreMax = maxS
        If cntS > 0 Then c.ScoreAvg = sumS / cntS Else c.ScoreAvg = 0
        Set cases(ci) = c
        ci = ci + 1
    Next

    ' Fälle sortieren: nach frühestem Zeitstempel, Fälle ohne Zeit ans Ende
    Dim ckey
    For a = 1 To UBound(cases)
        Set ckey = cases(a)
        jj = a - 1
        Do While jj >= 0
            If CmpCase(cases(jj), ckey) > 0 Then
                Set cases(jj + 1) = cases(jj)
                jj = jj - 1
            Else
                Exit Do
            End If
        Loop
        Set cases(jj + 1) = ckey
    Next

    FindCases = cases
End Function

' Vergleich zweier Records nach Zeit (Empty = ganz nach hinten).
Function CmpRec(ra, rb)
    Dim ea, eb : ea = IsEmpty(ra.Timestamp) : eb = IsEmpty(rb.Timestamp)
    If ea And eb Then
        CmpRec = 0
    ElseIf ea Then
        CmpRec = 1
    ElseIf eb Then
        CmpRec = -1
    ElseIf ra.Timestamp < rb.Timestamp Then
        CmpRec = -1
    ElseIf ra.Timestamp > rb.Timestamp Then
        CmpRec = 1
    Else
        CmpRec = 0
    End If
End Function

Function CmpCase(ca, cb)
    Dim ma, mb : ma = CaseMinTs(ca) : mb = CaseMinTs(cb)
    Dim ea, eb : ea = IsEmpty(ma) : eb = IsEmpty(mb)
    If ea And eb Then
        CmpCase = 0
    ElseIf ea Then
        CmpCase = 1
    ElseIf eb Then
        CmpCase = -1
    ElseIf ma < mb Then
        CmpCase = -1
    ElseIf ma > mb Then
        CmpCase = 1
    Else
        CmpCase = 0
    End If
End Function

Function CaseMinTs(c)
    Dim mn : mn = Empty
    Dim arr : arr = c.Records
    Dim r
    For Each r In arr
        If Not IsEmpty(r.Timestamp) Then
            If IsEmpty(mn) Then
                mn = r.Timestamp
            ElseIf r.Timestamp < mn Then
                mn = r.Timestamp
            End If
        End If
    Next
    CaseMinTs = mn
End Function


' =====================================================================
'  AUSGABE
' =====================================================================

Sub PrintList(extractors)
    WScript.Echo "Verfügbare Suchmuster:"
    Dim e
    For Each e In extractors
        WScript.Echo "  " & PadRight(e.Name, 12) & "  " & e.Description
    Next
End Sub

Sub PrintCases(cases, active, showContent, onlyMulti)
    ' inline- und extra_line-Muster trennen
    Dim inlineArr : inlineArr = FilterByMode(active, "inline")
    Dim extraArr  : extraArr  = FilterByMode(active, "extra_line")

    Dim shown : shown = 0
    Dim caseNo : caseNo = 0
    Dim c, recsArr, nrec, r, e
    For Each c In cases
        recsArr = c.Records
        nrec = NumElems(recsArr)
        If Not (onlyMulti And nrec < 2) Then
            caseNo = caseNo + 1
            shown = shown + 1

            Dim n_sds, n_co : n_sds = 0 : n_co = 0
            For Each r In recsArr
                If r.Kind = "sds" Then n_sds = n_sds + 1 Else n_co = n_co + 1
            Next

            Dim head : head = BuildHead(n_co, n_sds)
            Dim suffix : suffix = "" : If nrec <= 1 Then suffix = ", ohne Gegenstück"

            WScript.Echo ""
            WScript.Echo "Fall #" & caseNo & "  " & nrec & " Record(s) (" & head & suffix & ")  " & _
                         "score_max=" & Fmt2(c.ScoreMax) & "  score_avg=" & Fmt2(c.ScoreAvg)

            Dim ic : ic = NumElems(inlineArr)
            For Each r In recsArr
                Dim fieldStr : fieldStr = ""
                If ic > 0 Then
                    Dim fieldParts() : ReDim fieldParts(ic - 1)
                    Dim fi : fi = 0
                    For Each e In inlineArr
                        fieldParts(fi) = e.Name & "=" & ReprVal(FieldVal(r, e.Name))
                        fi = fi + 1
                    Next
                    fieldStr = Join(fieldParts, "  ")
                End If
                WScript.Echo "  [" & PadRight(r.Kind, 7) & "] " & FmtTs(r.Timestamp) & "  " & fieldStr

                For Each e In extraArr
                    Dim val : val = FieldVal(r, e.Name)
                    If val <> "" Then WScript.Echo Space(11) & e.Name & ": " & val
                Next
                If showContent Then WScript.Echo Space(11) & r.Content
            Next
        End If
    Next

    If onlyMulti And shown = 0 Then
        WScript.Echo ""
        WScript.Echo "(Keine Fälle mit >= 2 Records gefunden.)"
    End If
End Sub

Sub PrintStats(recs, active)
    Dim n : n = NumElems(recs)
    Dim n_co, n_sds : n_co = 0 : n_sds = 0
    Dim r
    For Each r In recs
        If r.Kind = "callout" Then n_co = n_co + 1 Else n_sds = n_sds + 1
    Next

    WScript.Echo ""
    WScript.Echo "Match-Statistik (pro Suchmuster):"
    WScript.Echo "  " & PadRight("Muster", 12) & "  " & PadLeft("Callout", 14) & "  " & _
                 PadLeft("SDS", 14) & "  " & PadLeft("Gesamt", 14)

    Dim e, co_hit, sds_hit, total_hit, co_str, sds_str, tot_str
    For Each e In active
        co_hit = 0 : sds_hit = 0
        For Each r In recs
            If FieldVal(r, e.Name) <> "" Then
                If r.Kind = "callout" Then co_hit = co_hit + 1 Else sds_hit = sds_hit + 1
            End If
        Next
        total_hit = co_hit + sds_hit
        If n_co  > 0 Then co_str  = co_hit  & "/" & n_co  Else co_str  = "-"
        If n_sds > 0 Then sds_str = sds_hit & "/" & n_sds Else sds_str = "-"
        tot_str = total_hit & "/" & n
        WScript.Echo "  " & PadRight(e.Name, 12) & "  " & PadLeft(co_str, 14) & "  " & _
                     PadLeft(sds_str, 14) & "  " & PadLeft(tot_str, 14)
    Next
End Sub

Function BuildHead(n_co, n_sds)
    Dim parts() : ReDim parts(1) : Dim c : c = 0
    If n_co > 0 Then
        parts(c) = n_co & " Callout"
        If n_co <> 1 Then parts(c) = parts(c) & "s"
        c = c + 1
    End If
    If n_sds > 0 Then
        parts(c) = n_sds & " SDS"
        c = c + 1
    End If
    If c = 0 Then
        BuildHead = "0"
    Else
        ReDim Preserve parts(c - 1)
        BuildHead = Join(parts, ", ")
    End If
End Function

' Wie Python repr(): Wert in einfache Anführungszeichen, leer -> '-'
Function ReprVal(v)
    If v = "" Then v = "-"
    ReprVal = "'" & v & "'"
End Function

Function FmtTs(ts)
    If IsEmpty(ts) Then
        FmtTs = "-"
    Else
        FmtTs = Right("000" & Year(ts), 4) & "-" & Pad2(Month(ts)) & "-" & Pad2(Day(ts)) & _
                " " & Pad2(Hour(ts)) & ":" & Pad2(Minute(ts)) & ":" & Pad2(Second(ts))
    End If
End Function


' =====================================================================
'  CLI-HILFEN
' =====================================================================

Sub ConfigureExtractors(extractors, onlyArg, skipArg)
    Dim names : Set names = CreateObject("Scripting.Dictionary")
    Dim e
    For Each e In extractors
        names.Add e.Name, True
    Next

    If onlyArg <> "" Then
        Dim wanted : Set wanted = SplitSet(onlyArg)
        CheckUnknown wanted, names
        For Each e In extractors
            e.Enabled = wanted.Exists(e.Name)
        Next
    End If

    If skipArg <> "" Then
        Dim unwanted : Set unwanted = SplitSet(skipArg)
        CheckUnknown unwanted, names
        For Each e In extractors
            If unwanted.Exists(e.Name) Then e.Enabled = False
        Next
    End If
End Sub

Sub CheckUnknown(wanted, names)
    Dim k
    For Each k In wanted.Keys
        If Not names.Exists(k) Then
            Die "Unbekanntes Suchmuster: " & k & ". Verfügbar: " & Join(names.Keys, ", ")
        End If
    Next
End Sub

Function SplitSet(s)
    Dim d : Set d = CreateObject("Scripting.Dictionary")
    Dim parts : parts = Split(s, ",")
    Dim p, t
    For Each p In parts
        t = Trim(p)
        If t <> "" And Not d.Exists(t) Then d.Add t, True
    Next
    Set SplitSet = d
End Function

Function FilterActive(extractors)
    Dim tmp() : ReDim tmp(UBound(extractors)) : Dim c : c = 0
    Dim e
    For Each e In extractors
        If e.Enabled Then
            Set tmp(c) = e
            c = c + 1
        End If
    Next
    If c = 0 Then
        FilterActive = Array()
    Else
        ReDim Preserve tmp(c - 1)
        FilterActive = tmp
    End If
End Function

Function FilterByMode(active, mode)
    Dim tmp() : ReDim tmp(UBound(active)) : Dim c : c = 0
    Dim e
    For Each e In active
        If e.DisplayMode = mode Then
            Set tmp(c) = e
            c = c + 1
        End If
    Next
    If c = 0 Then
        FilterByMode = Array()
    Else
        ReDim Preserve tmp(c - 1)
        FilterByMode = tmp
    End If
End Function

Function JoinNames(active)
    Dim tmp() : ReDim tmp(UBound(active)) : Dim i : i = 0
    Dim e
    For Each e In active
        tmp(i) = e.Name
        i = i + 1
    Next
    JoinNames = Join(tmp, ", ")
End Function


' =====================================================================
'  KLEINE HELFER
' =====================================================================

Function NumElems(arr)
    Dim u : u = -1
    On Error Resume Next
    u = UBound(arr)
    On Error GoTo 0
    NumElems = u + 1
End Function

Function PadRight(s, n)
    s = CStr(s)
    If Len(s) < n Then PadRight = s & Space(n - Len(s)) Else PadRight = s
End Function

Function PadLeft(s, n)
    s = CStr(s)
    If Len(s) < n Then PadLeft = Space(n - Len(s)) & s Else PadLeft = s
End Function

Function Pad2(x)
    Pad2 = Right("0" & x, 2)
End Function

' Formatiert mit 2 Nachkommastellen und IMMER Punkt als Dezimaltrennzeichen.
Function Fmt2(x)
    Dim s : s = FormatNumber(x, 2, -1, 0, 0)
    Fmt2 = Replace(s, ",", ".")
End Function

' Liest eine Dezimalzahl unabhängig vom Gebietsschema (Punkt ODER Komma).
Function ToDouble(s)
    s = Replace(Trim(s), ",", ".")
    Dim neg : neg = False
    If Left(s, 1) = "-" Then neg = True : s = Mid(s, 2)
    Dim p : p = InStr(s, ".")
    Dim val
    If p = 0 Then
        val = CDbl(s)
    Else
        Dim ip, fp
        ip = Left(s, p - 1) : fp = Mid(s, p + 1)
        If ip = "" Then ip = "0"
        If fp = "" Then fp = "0"
        val = CDbl(ip) + CDbl(fp) / (10 ^ Len(fp))
    End If
    If neg Then val = -val
    ToDouble = val
End Function

Function AbsPath(p)
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    AbsPath = fso.GetAbsolutePathName(p)
End Function

Function FileExists(p)
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(p)
End Function

Sub Die(msg)
    WScript.Echo msg
    WScript.Quit 1
End Sub
