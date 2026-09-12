Option Explicit

'===================================================================================
'  modBOTImport
'-----------------------------------------------------------------------------------
'  Purpose : Read the "BOT Promo Collation" workbook, route each row by country
'            into the CZ / SK / HU sheets of THIS workbook (Food - Promo Link
'            EXIT Check), and copy rows where CAG/CAYG is not blank into the
'            matching "<Country> EXIT" sheet.
'
'  Run     : Open both workbooks in the same Excel session (or just this one -
'            the macro will ask you to browse for the BOT file if it isn't
'            already open), then run ImportBOTPromoData.
'
'  IMPORTANT ASSUMPTIONS - please review before relying on the output:
'  --------------------------------------------------------------------
'  1) "CAG" in the request is assumed to be the column headed "CAYG" in the
'     BOT workbook (there is no column literally called "CAG"). If that is
'     wrong, change BOT_HDR_CAG below to the correct header text.
'  2) The two country columns in BOT ("country" = full name, "Country" = 2
'     letter code SK/HU/CZ) always agreed in the sample file, so routing uses
'     the 2-letter code column ("Country").
'  3) "Start date" / "End date" are taken from BOT's "startdate"/"enddate"
'     columns (not the second pair "Start date"/"End date", which differed
'     in the sample file and looked like a reconciled/STRAPP date).
'  4) "Minus week until promo starts" uses a Tesco-style retail calendar
'     where fiscal Week 1 begins on the Monday on/or-after 1 March each
'     year (this matches every clean "yyyy-ww" example found in the data).
'     The value written is (current fiscal week) - (promo's fiscal week),
'     so a promo 3 weeks away shows -3, and 0/positive once it has started.
'     If your real fiscal calendar starts on a different date, fix the
'     single function FiscalAnchorForYear().
'  5) "Promo Week" text is very inconsistent (e.g. "2026-32 1WK",
'     "36 26 2WK GRNF", "W32 Harry Potter", "2026 Winter bird food" with no
'     week at all). ParsePromoWeek() tries three patterns in order and, if
'     none match, leaves the Minus-Week cell blank rather than guessing.
'  6) Columns that are driven by formulas in the destination sheets
'     ("Format" on the main sheets, and "Department"/"Format"/"Division" on
'     the EXIT sheets) are NEVER overwritten with values - the existing
'     formula from the row directly above is copied down instead, so all
'     existing formulas in the workbook are fully preserved.
'  7) Fields with no obvious source in BOT ("Promo num", "Active Open Links
'     in Stores") are left blank. "Subject" is stamped with a constant tag
'     and "Received" with today's date, purely so the row is traceable back
'     to this import - change the two constants below if you'd rather have
'     something else.
'  8) The macro is safe to re-run: it builds a duplicate-check on
'     TPN+Store+StartDate (main sheets) and TPN+Store (EXIT sheets) from the
'     rows already in the sheet, so re-importing the same BOT export will
'     not create duplicate rows.
'
'  MISSING HEADERS: if any BOT or destination header referenced below is not
'  found on its sheet, ImportBOTPromoData aborts up front and lists exactly
'  which ones are missing, rather than silently importing blank columns.
'===================================================================================

' ---- Configuration you may want to tweak -----------------------------------------
Private Const BOT_WORKBOOK_HINT As String = "Promo_Collation" ' partial file name to find/open
                                                                ' (matches "BOT_Promo_Collation_..." and
                                                                '  "Promo_Collation_Wednesday" etc. - if it
                                                                '  can't find a match among open workbooks it
                                                                '  just prompts you to browse for the file)
Private Const BOT_SHEET_NAME    As String = "Sheet1"

Private Const BOT_HDR_TPN        As String = "tpn"
Private Const BOT_HDR_STORE      As String = "storenumber"
Private Const BOT_HDR_BUYER      As String = "buyername"
Private Const BOT_HDR_PROMONAME  As String = "promonames"
Private Const BOT_HDR_STARTDATE  As String = "startdate"
Private Const BOT_HDR_ENDDATE    As String = "enddate"
Private Const BOT_HDR_DEPTDESC   As String = "Department description"
Private Const BOT_HDR_DIVDESC    As String = "Division description"
Private Const BOT_HDR_DRG        As String = "DRG"
Private Const BOT_HDR_DRGDESC    As String = "DRG description"
Private Const BOT_HDR_TYPE       As String = "OOCP/CP"
Private Const BOT_HDR_REASON     As String = "Reason"
Private Const BOT_HDR_ITEMDESC   As String = "Item description"
Private Const BOT_HDR_COUNTRYCD  As String = "Country"       ' 2-letter code column (SK/HU/CZ)
Private Const BOT_HDR_CAG        As String = "CAYG"          ' <-- assumption, see header notes

Private Const DEST_HDR_FROM      As String = "From"
Private Const DEST_HDR_ITEM      As String = "Item"
Private Const DEST_HDR_LOCATION  As String = "Location"
Private Const DEST_HDR_ITEMDESC  As String = "Item Description"   ' only exists on HU
Private Const DEST_HDR_PROMOWEEK As String = "Promo Week"
Private Const DEST_HDR_STARTDATE As String = "Start date"
Private Const DEST_HDR_ENDDATE   As String = "End date"
Private Const DEST_HDR_DEPT      As String = "Department"
Private Const DEST_HDR_DIV       As String = "Division"
Private Const DEST_HDR_REASON    As String = "Reason"
Private Const DEST_HDR_DRG       As String = "DRG"
Private Const DEST_HDR_DRGNAME   As String = "DRG Name"
Private Const DEST_HDR_TYPE      As String = "Type"
Private Const DEST_HDR_SUBJECT   As String = "Subject"
Private Const DEST_HDR_RECEIVED  As String = "Received"
Private Const DEST_HDR_MINUSWK   As String = "Minus week until promo starts"

Private Const EXIT_HDR_TPN       As String = "TPN"            ' sheet has a leading space, trimmed on read
Private Const EXIT_HDR_STORE     As String = "Store number"

Private Const IMPORT_SUBJECT_TAG As String = "BOT Import"

' Headers that MUST exist on the BOT sheet for the import to make sense.
' (This list intentionally excludes DEST_HDR_ITEMDESC, which is HU-only.)
Private Const REQUIRED_BOT_HEADERS As String = "tpn|storenumber|buyername|promonames|startdate|enddate|" & _
    "Department description|Division description|DRG|DRG description|OOCP/CP|Reason|Item description|Country|CAYG"

' ---- Small helper type to bundle everything we need for one destination sheet ----
Private Type TargetSheet
    ws          As Worksheet
    headers     As Object      ' Scripting.Dictionary: trimmed header text -> column index
    formulaCols As Object      ' Scripting.Dictionary: column index -> True, for formula-driven cols
    anchorCol   As Long        ' column used to find the current last row (a plain-value column)
    lastRow     As Long
    existing    As Object      ' Scripting.Dictionary: dedup key -> True
End Type


'===================================================================================
'  MAIN ENTRY POINT
'===================================================================================
Sub ImportBOTPromoData()

    Dim wbDest As Workbook, wbBOT As Workbook
    Dim wsBOT As Worksheet
    Dim hdrBOT As Object
    Dim lastRowBOT As Long, r As Long

    Dim tCZ As TargetSheet, tSK As TargetSheet, tHU As TargetSheet
    Dim eCZ As TargetSheet, eSK As TargetSheet, eHU As TargetSheet
    Dim tMain As TargetSheet, tExit As TargetSheet

    Dim countryCode As String
    Dim tpnVal As Variant, storeVal As Variant, startDateVal As Variant
    Dim mainKey As String, exitKey As String
    Dim cagVal As Variant
    Dim addedMain As Long, dupMain As Long, addedExit As Long, dupExit As Long, unmapped As Long
    Dim unparsedWeeks As Long
    Dim missingHeaders As String

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set wbDest = ThisWorkbook

    ' ---- 1. Locate the BOT source workbook -----------------------------------
    Set wbBOT = GetOrOpenWorkbook(BOT_WORKBOOK_HINT)
    If wbBOT Is Nothing Then
        MsgBox "Could not open the BOT Promo Collation workbook. Import cancelled.", vbExclamation
        GoTo CleanExit
    End If
    Set wsBOT = wbBOT.Sheets(BOT_SHEET_NAME)

    ' ---- 2. Read BOT headers and find last row --------------------------------
    Set hdrBOT = GetHeaderMap(wsBOT, 1)

    missingHeaders = FindMissingHeaders(hdrBOT, Split(REQUIRED_BOT_HEADERS, "|"))
    If Len(missingHeaders) > 0 Then
        MsgBox "The BOT workbook is missing expected header(s):" & vbCrLf & vbCrLf & _
               missingHeaders & vbCrLf & vbCrLf & _
               "Check the sheet layout hasn't changed (header text/case must match exactly), " & _
               "or update the BOT_HDR_* constants at the top of this module.", vbCritical
        GoTo CleanExit
    End If

    lastRowBOT = wsBOT.Cells(wsBOT.Rows.Count, hdrBOT(BOT_HDR_TPN)).End(xlUp).Row

    ' ---- 3. Prepare the six destination sheets --------------------------------
    ' NOTE: TargetSheet is a user-defined Type, not an object, so plain "="
    ' is used to assign it (never "Set").
    tCZ = PrepareTargetSheet(wbDest, "CZ", False)
    tSK = PrepareTargetSheet(wbDest, "SK", False)
    tHU = PrepareTargetSheet(wbDest, "HU", False)
    eCZ = PrepareTargetSheet(wbDest, "CZ EXIT", True)
    eSK = PrepareTargetSheet(wbDest, "SK EXIT", True)
    eHU = PrepareTargetSheet(wbDest, "HU EXIT", True)

    ' ---- 4. Walk every BOT row -------------------------------------------------
    For r = 2 To lastRowBOT

        tpnVal = GetVal(wsBOT, r, hdrBOT, BOT_HDR_TPN)
        If Len(Trim$(CStr(tpnVal))) = 0 Then GoTo NextRow    ' skip completely empty rows

        countryCode = Trim$(CStr(GetVal(wsBOT, r, hdrBOT, BOT_HDR_COUNTRYCD)))

        Select Case UCase$(countryCode)
            Case "CZ": tMain = tCZ: tExit = eCZ
            Case "SK": tMain = tSK: tExit = eSK
            Case "HU": tMain = tHU: tExit = eHU
            Case Else
                unmapped = unmapped + 1
                GoTo NextRow
        End Select

        storeVal = GetVal(wsBOT, r, hdrBOT, BOT_HDR_STORE)
        startDateVal = GetVal(wsBOT, r, hdrBOT, BOT_HDR_STARTDATE)

        ' ---- 4a. Main country sheet -------------------------------------------
        mainKey = MakeKey(tpnVal, storeVal, startDateVal)

        If Not tMain.existing.Exists(mainKey) Then
            AddMainRow tMain, wsBOT, r, hdrBOT, unparsedWeeks
            tMain.existing.Add mainKey, True
            addedMain = addedMain + 1
        Else
            dupMain = dupMain + 1
        End If

        ' ---- 4b. EXIT sheet - only when CAG/CAYG is not blank ------------------
        cagVal = GetVal(wsBOT, r, hdrBOT, BOT_HDR_CAG)
        If Len(Trim$(CStr(cagVal))) > 0 Then
            exitKey = MakeKey(tpnVal, storeVal)
            If Not tExit.existing.Exists(exitKey) Then
                AddExitRow tExit, tpnVal, storeVal
                tExit.existing.Add exitKey, True
                addedExit = addedExit + 1
            Else
                dupExit = dupExit + 1
            End If
        End If

        ' write back the (possibly updated) structures - needed because TargetSheet
        ' is passed by value into the Select Case assignment above
        Select Case UCase$(countryCode)
            Case "CZ": tCZ = tMain: eCZ = tExit
            Case "SK": tSK = tMain: eSK = tExit
            Case "HU": tHU = tMain: eHU = tExit
        End Select

NextRow:
    Next r

    MsgBox "BOT import finished." & vbCrLf & vbCrLf & _
           "Main sheets  - added: " & addedMain & "   duplicates skipped: " & dupMain & vbCrLf & _
           "EXIT sheets  - added: " & addedExit & "   duplicates skipped: " & dupExit & vbCrLf & _
           "Rows with unrecognised country: " & unmapped & vbCrLf & _
           "Promo weeks that could not be parsed (Minus-Week left blank): " & unparsedWeeks, _
           vbInformation, "Import complete"

CleanExit:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    MsgBox "Import stopped because of an error:" & vbCrLf & Err.Description, vbCritical
    Resume CleanExit

End Sub


'===================================================================================
'  Return a CrLf-joined list of any headers in requiredHeaders not present in hdr,
'  or "" if all are present. Used to fail fast and loudly instead of silently
'  importing blank columns (GetVal/WriteIfHeaderExists both fail silently by
'  design, so this check has to happen up front).
'===================================================================================
Private Function FindMissingHeaders(hdr As Object, requiredHeaders As Variant) As String
    Dim i As Long
    Dim missing As String
    For i = LBound(requiredHeaders) To UBound(requiredHeaders)
        If Not hdr.Exists(requiredHeaders(i)) Then
            missing = missing & "  - " & requiredHeaders(i) & vbCrLf
        End If
    Next i
    FindMissingHeaders = missing
End Function


'===================================================================================
'  Build a TargetSheet structure: header map, formula-column map, dedup keys
'===================================================================================
Private Function PrepareTargetSheet(wb As Workbook, sheetName As String, isExit As Boolean) As TargetSheet

    Dim result As TargetSheet
    Dim ws As Worksheet
    Dim hdr As Object
    Dim col As Long, lastCol As Long
    Dim anchorHeader As String
    Dim r As Long, lastRow As Long
    Dim tpnCol As Long, storeCol As Long, startCol As Long
    Dim k As String
    Dim requiredHeaders As Variant
    Dim missingHeaders As String

    Set ws = wb.Sheets(sheetName)
    Set hdr = GetHeaderMap(ws, 1)

    If isExit Then
        requiredHeaders = Array(EXIT_HDR_TPN, EXIT_HDR_STORE)
    Else
        requiredHeaders = Array(DEST_HDR_ITEM, DEST_HDR_LOCATION, DEST_HDR_STARTDATE)
    End If
    missingHeaders = FindMissingHeaders(hdr, requiredHeaders)
    If Len(missingHeaders) > 0 Then
        Err.Raise vbObjectError + 1, "PrepareTargetSheet", _
            "Sheet '" & sheetName & "' is missing expected header(s):" & vbCrLf & missingHeaders
    End If

    Set result.ws = ws
    Set result.headers = hdr

    ' Which column anchors "how many rows already exist" - use TPN/Item, a plain value column.
    ' A single-column End(xlUp) can understate lastRow if that one column has a stray blank
    ' in an otherwise-populated row, so corroborate against the sheet's true last used row.
    If isExit Then
        anchorHeader = EXIT_HDR_TPN
    Else
        anchorHeader = DEST_HDR_ITEM
    End If
    result.anchorCol = hdr(anchorHeader)
    lastRow = ws.Cells(ws.Rows.Count, result.anchorCol).End(xlUp).Row

    Dim usedLastRow As Long
    On Error Resume Next
    usedLastRow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
    On Error GoTo 0
    If usedLastRow > lastRow Then lastRow = usedLastRow

    If lastRow < 1 Then lastRow = 1
    result.lastRow = lastRow

    ' Detect which columns are formula-driven by scanning every data row (not just row 2),
    ' since a column's formula may not start until a later row (or row 2 may be blank/typed).
    Set result.formulaCols = CreateObject("Scripting.Dictionary")
    If lastRow >= 2 Then
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        For col = 1 To lastCol
            For r = 2 To lastRow
                If ws.Cells(r, col).HasFormula Then
                    If Not result.formulaCols.Exists(col) Then result.formulaCols.Add col, True
                    Exit For
                End If
            Next r
        Next col
    End If

    ' Build the dedup key set from whatever rows already exist
    Set result.existing = CreateObject("Scripting.Dictionary")
    If isExit Then
        tpnCol = hdr(EXIT_HDR_TPN)
        storeCol = hdr(EXIT_HDR_STORE)
        For r = 2 To lastRow
            k = MakeKey(ws.Cells(r, tpnCol).Value, ws.Cells(r, storeCol).Value)
            If Not result.existing.Exists(k) Then result.existing.Add k, True
        Next r
    Else
        tpnCol = hdr(DEST_HDR_ITEM)
        storeCol = hdr(DEST_HDR_LOCATION)
        startCol = hdr(DEST_HDR_STARTDATE)
        For r = 2 To lastRow
            k = MakeKey(ws.Cells(r, tpnCol).Value, ws.Cells(r, storeCol).Value, ws.Cells(r, startCol).Value)
            If Not result.existing.Exists(k) Then result.existing.Add k, True
        Next r
    End If

    PrepareTargetSheet = result

End Function


'===================================================================================
'  Append one row to a main country sheet (CZ / SK / HU)
'===================================================================================
Private Sub AddMainRow(ByRef t As TargetSheet, wsBOT As Worksheet, srcRow As Long, hdrBOT As Object, _
                        ByRef unparsedWeeks As Long)

    Dim newRow As Long
    Dim ws As Worksheet
    Dim h As Object
    Dim promoWeekText As String
    Dim minusWeek As Variant

    Set ws = t.ws
    Set h = t.headers
    newRow = t.lastRow + 1

    WriteIfHeaderExists t, newRow, DEST_HDR_FROM, GetVal(wsBOT, srcRow, hdrBOT, BOT_HDR_BUYER)
    WriteIfHeaderExists t, newRow, DEST_HDR_ITEM, GetVal(wsBOT, srcRow, hdrBOT, BOT_HDR_TPN)
    WriteIfHeaderExists t, newRow, DEST_HDR_LOCATION, GetVal(wsBOT, srcRow, hdrBOT, BOT_HDR_STORE)
    WriteIfHeaderExists t, newRow, DEST_HDR_ITEMDESC, GetVal(wsBOT, srcRow, hdrBOT, BOT_HDR_ITEMDESC) ' HU only
    promoWeekText = CStr(GetVal(wsBOT, srcRow, hdrBOT, BOT_HDR_PROMONAME))
    WriteIfHeaderExists t, newRow, DEST_HDR_PROMOWEEK, promoWeekText
    WriteIfHeaderExists t, newRow, DEST_HDR_STARTDATE, GetVal(wsBOT, srcRow, hdrBOT, BOT_HDR_STARTDATE)
    WriteIfHeaderExists t, newRow, DEST_HDR_ENDDATE, GetVal(wsBOT, srcRow, hdrBOT, BOT_HDR_ENDDATE)
    WriteIfHeaderExists t, newRow, DEST_HDR_DEPT, GetVal(wsBOT, srcRow, hdrBOT, BOT_HDR_DEPTDESC)
    WriteIfHeaderExists t, newRow, DEST_HDR_DIV, GetVal(wsBOT, srcRow, hdrBOT, BOT_HDR_DIVDESC)
    WriteIfHeaderExists t, newRow, DEST_HDR_REASON, GetVal(wsBOT, srcRow, hdrBOT, BOT_HDR_REASON)
    WriteIfHeaderExists t, newRow, DEST_HDR_DRG, GetVal(wsBOT, srcRow, hdrBOT, BOT_HDR_DRG)
    WriteIfHeaderExists t, newRow, DEST_HDR_DRGNAME, GetVal(wsBOT, srcRow, hdrBOT, BOT_HDR_DRGDESC)
    WriteIfHeaderExists t, newRow, DEST_HDR_TYPE, GetVal(wsBOT, srcRow, hdrBOT, BOT_HDR_TYPE)
    WriteIfHeaderExists t, newRow, DEST_HDR_SUBJECT, IMPORT_SUBJECT_TAG
    WriteIfHeaderExists t, newRow, DEST_HDR_RECEIVED, Date
    ' "Promo num" and "Active Open Links in Stores" have no reliable source in BOT - left blank

    ' Minus week until promo starts
    minusWeek = ParsePromoWeekToMinusWeeks(promoWeekText)
    If IsNull(minusWeek) Then
        unparsedWeeks = unparsedWeeks + 1
    Else
        WriteIfHeaderExists t, newRow, DEST_HDR_MINUSWK, minusWeek
    End If

    ' Preserve formulas: copy every formula-driven column down from the row above
    CopyFormulaColumnsDown t, newRow

    t.lastRow = newRow

End Sub


'===================================================================================
'  Append one row to an EXIT sheet (just TPN + Store number - the lookup formulas
'  for Department / Format / Division fill themselves in)
'===================================================================================
Private Sub AddExitRow(ByRef t As TargetSheet, tpnVal As Variant, storeVal As Variant)

    Dim newRow As Long
    Dim ws As Worksheet
    Dim h As Object

    Set ws = t.ws
    Set h = t.headers
    newRow = t.lastRow + 1

    WriteIfHeaderExists t, newRow, EXIT_HDR_TPN, tpnVal
    WriteIfHeaderExists t, newRow, EXIT_HDR_STORE, storeVal

    CopyFormulaColumnsDown t, newRow

    t.lastRow = newRow

End Sub


'===================================================================================
'  Copy every detected formula column from (newRow - 1) down into newRow, so that
'  VLOOKUP / IFS formulas already in the sheet are extended rather than overwritten.
'===================================================================================
Private Sub CopyFormulaColumnsDown(ByRef t As TargetSheet, newRow As Long)

    Dim col As Variant
    If newRow <= 2 Then Exit Sub   ' nothing above to copy from

    For Each col In t.formulaCols.Keys
        t.ws.Cells(newRow, CLng(col)).FormulaR1C1 = t.ws.Cells(newRow - 1, CLng(col)).FormulaR1C1
    Next col

End Sub


'===================================================================================
'  Write a value into a destination column only if that header exists on the
'  sheet (CZ/SK don't have "Item Description"; several sheets don't use every
'  header) - and never write into a column that is formula-driven, since those
'  are always populated via CopyFormulaColumnsDown instead.
'===================================================================================
Private Sub WriteIfHeaderExists(ByRef t As TargetSheet, rowNum As Long, headerName As String, val As Variant)
    Dim col As Long
    If t.headers.Exists(headerName) Then
        col = t.headers(headerName)
        If t.formulaCols.Exists(col) Then Exit Sub
        t.ws.Cells(rowNum, col).Value = val
    End If
End Sub


'===================================================================================
'  Safe read of a BOT cell by header name; returns "" if the header doesn't exist
'===================================================================================
Private Function GetVal(ws As Worksheet, rowNum As Long, hdr As Object, headerName As String) As Variant
    If hdr.Exists(headerName) Then
        GetVal = ws.Cells(rowNum, hdr(headerName)).Value
    Else
        GetVal = ""
    End If
End Function


'===================================================================================
'  Build a dictionary of Trim(header text) -> column number for a header row.
'  Blank headers are skipped. Comparison is case-sensitive on purpose, because
'  the BOT sheet uses two headers that differ only by case ("country"/"Country").
'===================================================================================
Private Function GetHeaderMap(ws As Worksheet, headerRow As Long) As Object
    Dim dict As Object
    Dim lastCol As Long, c As Long
    Dim h As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbBinaryCompare
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        h = Trim$(CStr(ws.Cells(headerRow, c).Value))
        If Len(h) > 0 Then
            If Not dict.Exists(h) Then dict.Add h, c
        End If
    Next c

    Set GetHeaderMap = dict
End Function


'===================================================================================
'  Build a "|" separated dedup key from 2 or 3 values. Dates are normalised to
'  yyyymmdd so formatting differences don't create false "new" rows.
'===================================================================================
Private Function MakeKey(v1 As Variant, v2 As Variant, Optional v3 As Variant) As String
    Dim s3 As String
    If Not IsMissing(v3) Then
        If IsDate(v3) Then
            s3 = Format$(CDate(v3), "yyyymmdd")
        Else
            s3 = Trim$(CStr(v3))
        End If
        MakeKey = Trim$(CStr(v1)) & "|" & Trim$(CStr(v2)) & "|" & s3
    Else
        MakeKey = Trim$(CStr(v1)) & "|" & Trim$(CStr(v2))
    End If
End Function


'===================================================================================
'  Find an already-open workbook whose name contains partialName, otherwise ask
'  the user to browse for it and open it.
'===================================================================================
Private Function GetOrOpenWorkbook(partialName As String) As Workbook
    Dim wb As Workbook
    Dim fPath As Variant

    For Each wb In Application.Workbooks
        If InStr(1, wb.Name, partialName, vbTextCompare) > 0 Then
            Set GetOrOpenWorkbook = wb
            Exit Function
        End If
    Next wb

    MsgBox "Please locate the BOT Promo Collation workbook.", vbInformation
    fPath = Application.GetOpenFilename( _
        FileFilter:="Excel Workbooks (*.xlsx;*.xlsm),*.xlsx;*.xlsm", _
        Title:="Select the BOT Promo Collation workbook")

    If fPath = False Then
        Set GetOrOpenWorkbook = Nothing
    Else
        Set GetOrOpenWorkbook = Workbooks.Open(CStr(fPath))
    End If
End Function


'===================================================================================
'  FISCAL CALENDAR HELPERS
'  Tesco-style retail calendar: fiscal Week 1, Day 1 = the Monday on/after
'  1 March of the fiscal year. Validated against the sample data:
'    "2026-29" -> 14-Sep-2026 (Mon)   "2026-32" -> 05-Oct-2026 (Mon)
'  Both come out exactly right with this anchor rule.
'===================================================================================
Private Function FiscalAnchorForYear(calYear As Long) As Date
    Dim march1 As Date
    Dim addDays As Long
    march1 = DateSerial(calYear, 3, 1)
    addDays = (1 - Weekday(march1, vbMonday) + 7) Mod 7
    FiscalAnchorForYear = march1 + addDays
End Function

' Returns fiscal year & fiscal week for any real date, as a 2-element array
Private Function FiscalYearWeek(d As Date) As Variant
    Dim fy As Long, anchor As Date
    fy = Year(d)
    anchor = FiscalAnchorForYear(fy)
    If d < anchor Then
        fy = fy - 1
        anchor = FiscalAnchorForYear(fy)
    End If
    FiscalYearWeek = Array(fy, Int((d - anchor) / 7) + 1)
End Function

' Converts (fiscalYear, fiscalWeek) into one comparable absolute number of weeks
' since a fixed epoch. Uses the ACTUAL number of fiscal weeks in each intervening
' year (52 or 53 - a 53rd week occurs whenever the gap between that year's anchor
' and the next year's anchor is 371 days instead of 364) rather than assuming every
' fiscal year has exactly 52 weeks, so week counts stay correct across a 53-week year.
Private Function AbsoluteWeekNumber(fiscalYear As Long, fiscalWeek As Long) As Long
    Const EPOCH_YEAR As Long = 2000
    Dim y As Long, total As Long

    total = 0
    If fiscalYear >= EPOCH_YEAR Then
        For y = EPOCH_YEAR To fiscalYear - 1
            total = total + WeeksInFiscalYear(y)
        Next y
    Else
        For y = EPOCH_YEAR - 1 To fiscalYear Step -1
            total = total - WeeksInFiscalYear(y)
        Next y
    End If

    AbsoluteWeekNumber = total + fiscalWeek
End Function

' Number of fiscal weeks in a given fiscal year (52 in most years, 53 whenever the
' anchor-to-anchor gap is 371 days rather than 364).
Private Function WeeksInFiscalYear(fiscalYear As Long) As Long
    Dim daysBetweenAnchors As Long
    daysBetweenAnchors = CLng(FiscalAnchorForYear(fiscalYear + 1) - FiscalAnchorForYear(fiscalYear))
    WeeksInFiscalYear = daysBetweenAnchors \ 7
End Function


'===================================================================================
'  Extract a (year, week) pair out of very inconsistent "Promo Week" free text,
'  e.g. "2026-32 1WK", "36 26 2WK GRNF", "W32 Harry Potter", or text with no week
'  at all ("2026 Winter bird food"). Returns Null if nothing usable is found.
'===================================================================================
Private Function ParsePromoWeek(promoText As String) As Variant
    Dim re As Object, mc As Object
    Dim yr As Long, wk As Long

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True

    ' Pattern 1: "yyyy-ww" e.g. "2026-32"
    re.Pattern = "(20\d{2})-(\d{1,2})"
    If re.Test(promoText) Then
        Set mc = re.Execute(promoText)
        yr = CLng(mc(0).SubMatches(0))
        wk = CLng(mc(0).SubMatches(1))
        ParsePromoWeek = Array(yr, wk)
        Exit Function
    End If

    ' Pattern 2: "ww yy" e.g. "36 26" (week then 2-digit year), at the start of the text
    re.Pattern = "^(\d{1,2})\s+(\d{2})\b"
    If re.Test(promoText) Then
        Set mc = re.Execute(promoText)
        wk = CLng(mc(0).SubMatches(0))
        yr = CLng(mc(0).SubMatches(1))
        If yr < 100 Then yr = 2000 + yr
        ParsePromoWeek = Array(yr, wk)
        Exit Function
    End If

    ' Pattern 3: "Www" e.g. "W32" - no year given, assume the current fiscal year
    re.Pattern = "\bW(\d{1,2})\b"
    If re.Test(promoText) Then
        Set mc = re.Execute(promoText)
        wk = CLng(mc(0).SubMatches(0))
        yr = FiscalYearWeek(Date)(0)
        ParsePromoWeek = Array(yr, wk)
        Exit Function
    End If

    ParsePromoWeek = Null

End Function


'===================================================================================
'  Public-facing calc: "Minus week until promo starts" = current fiscal week
'  minus the promo's fiscal week (negative = still in the future, 0/positive =
'  live or past). Returns Null if the promo week text couldn't be parsed.
'===================================================================================
Private Function ParsePromoWeekToMinusWeeks(promoText As String) As Variant
    Dim parsed As Variant
    Dim todayFY As Variant
    Dim promoAbs As Long, currentAbs As Long

    parsed = ParsePromoWeek(promoText)
    If IsNull(parsed) Then
        ParsePromoWeekToMinusWeeks = Null
        Exit Function
    End If

    todayFY = FiscalYearWeek(Date)
    currentAbs = AbsoluteWeekNumber(CLng(todayFY(0)), CLng(todayFY(1)))
    promoAbs = AbsoluteWeekNumber(CLng(parsed(0)), CLng(parsed(1)))

    ParsePromoWeekToMinusWeeks = currentAbs - promoAbs
End Function
