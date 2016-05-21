Attribute VB_Name = "ColumnWidth"
Option Explicit
Option Base 1
' =============================================================================
'
'       COLUMN WIDTH OPTIMIZATION
'
' =============================================================================
' by Erica Warren, erica.warren@macmillan.com
'
' This *should* optimize column width for printing (i.e., adjust column widths
' AND page setup to print to as few pages wide as possible while still being
' legible) for most spreadsheets. Assumes continuous range, row 1 header.
'
' NOTE: currently coerces all point measurements to Long data type, i.e. whole
' numbers. If this causes rounding errors, can change to Single.
'
' Right now only set up for US Letter, Legal, and Tabloid paper sizes, though
' adding international sizes would be simple.


' ===== FormatToPrint =========================================================
' This is the sub to call. Programmed into a Quick Access Toolbar button and a
' keyboard shortcut (PC: Ctrl+Shift+F).

Public Sub FormatToPrint()
Attribute FormatToPrint.VB_Description = "Optimize column width and other page settings to print."
Attribute FormatToPrint.VB_ProcData.VB_Invoke_Func = "F\n14"
    ' ----- STARTUP -----------------------------------------------------------
    Application.ScreenUpdating = False
    ' Save current selection, to return to at end
    Dim rngStartingSelection As Range
    Set rngStartingSelection = Selection

    ' ----- SET UP OBJECTS, VARIABLES -----------------------------------------
    Dim blnDone As Boolean: blnDone = False
    Dim thisSheet As Worksheet
    Dim lngColumns As Long
    Dim lngRows As Long
    Dim rngFitColumns As Range

    ' Set range of cells to fit, but not the header row, because often the only
    ' cell in a column with long text is the header

    Set thisSheet = ActiveSheet
    thisSheet.Cells(1, 1).Activate      ' do count header rows, though
    lngColumns = ActiveCell.End(xlToRight).Column
    lngRows = ActiveCell.End(xlDown).Row
    Set rngFitColumns = thisSheet.Range(Cells(2, 1), Cells(lngRows, lngColumns))

    ' ----- TRY JUST COLUMN-FIT FUNCTION W/O FORMATTING CHANGES ---------------
    If FitColumns(rngFitColumns, 1) = False Then

        ' ----- BASIC FORMATTING ----------------------------------------------
        ' Verdana is a common font that looks good small
        ' Should add error handling in case it's missing, though
        rngFitColumns.Font.Name = "Verdana"
        
        ' Landscape almost always better fit, right?
        thisSheet.PageSetup.Orientation = xlLandscape
    
        ' ----- WHAT PAGE SIZES CAN WE USE? -----------------------------------
        Dim pageSizes() As Variant
        ' Add error handling in case no paper sizes returned
        pageSizes = GetPaperSizes(thisSheet)
    
        ' ----- MAKE ADJUSTMENTS, TRY TO FIT COLUMNS --------------------------
        Dim lngPagesWide As Long: lngPagesWide = 0  ' counter for first Do loop
        Dim lngMargins As Long ' TOTAL L + R margins
        Dim A As Long, B As Long, C As Long

        Do
            lngPagesWide = lngPagesWide + 1
            For A = 10 To 6 Step -1
                rngFitColumns.Font.Size = A
                For B = LBound(pageSizes) To UBound(pageSizes)
                    thisSheet.PageSetup.PaperSize = pageSizes(B)
                    For C = 1 To 2
                        lngMargins = Application.InchesToPoints(1 / C)
                        thisSheet.PageSetup.LeftMargin = lngMargins / 2
                        thisSheet.PageSetup.RightMargin = lngMargins / 2
                        blnDone = FitColumns(rngFitColumns, lngPagesWide)
                    Next C
                Next B
            Next A
        Loop Until blnDone = True Or lngPagesWide = 4
        ' 4 pages wide seems like as good a place to give up as any

        ' ----- OTHER PAGE SETUP ----------------------------------------------
        ' Even if FitColumsn ultimately failed, these settings are still good
        With thisSheet.PageSetup
            .BottomMargin = Application.InchesToPoints(0.5)
            .TopMargin = Application.InchesToPoints(0.5)
            .CenterHorizontally = True
            .Zoom = False
            .FitToPagesTall = False
            .PrintGridlines = True
            .PrintTitleRows = thisSheet.Rows(1).Address
            .CenterFooter = "&C &P"         ' &C for center, &P for page number
            .FitToPagesWide = lngPagesWide
        End With
    End If
    Dim strMessage As String
    If blnDone = True Then
        strMessage = "It worked!" & vbNewLine
    Else
        strMessage = "Sad face :(" & vbNewLine
    End If
    strMessage = strMessage & _
        "Paper size: " & thisSheet.PageSetup.PaperSize & vbNewLine & _
        "Margins: " & thisSheet.PageSetup.LeftMargin & vbNewLine & _
        "Font size: " & rngFitColumns.Font.Size & vbNewLine & _
        "Pages wide: " & lngPagesWide
    MsgBox strMessage
        
    ' ----- FINISH ------------------------------------------------------------
CleanUp:
    ' Save current selection to return to at end
    rngStartingSelection.Select
    Application.ScreenUpdating = True

End Sub

' ===== FitColumns ============================================================
' Tries to fit the columns based on the current PageSetup settings. Returns
' True if it was successful, False if not. Minimum column size is 1.5 inches.
' The range passed to this function should NOT include headers. If successful,
' will also autofit rows and wrap text.

Private Function FitColumns(fitRange As Range, pagesWide As Long) As Boolean
    ' Set up variables
    Dim objPageSetup As PageSetup
    Dim lngPageWidth As Long  ' Full width of print-page in points
    Dim lngSideMargins As Long ' TOTAL side margin, i.e. L + R
    Dim lngAvailWidth As Long  ' Width we have avail for our columns
    Dim lngAvgColumnW As Long  ' Average width of columns to change
    Dim lngStartAvgColW As Long ' Start avg width to compare result to
    Dim ColumnCollect As Collection
    Dim blnSuccess As Boolean: blnSuccess = False
    Dim blnStop As Boolean
    Dim rngColumn As Range
    Dim colItem As Variant
    Dim lngColWidth As Long
    Dim lngMinColumnW As Long: lngMinColumnW = Application.InchesToPoints(1.5)
    Dim D As Long, E As Long, F As Long
    Dim dCount As Long: dCount = 0
    Dim rngCheckColumn As Range
    Dim strColumn As String

    Set objPageSetup = fitRange.Parent.PageSetup
    lngPageWidth = GetPageWidth(objPageSetup) * pagesWide
    lngSideMargins = objPageSetup.LeftMargin + objPageSetup.RightMargin
    lngAvailWidth = lngPageWidth - lngSideMargins

    ' ----- START COLUMN FITTING ---------------------------------------------
    ' NOTE: Range.Width property returns points, Range.ColumnWidth returns
    ' number of characters (counted as width of "0")

    ' AutoFit columns and check - maybe that's fine
    fitRange.Columns.AutoFit
    If fitRange.Width > lngAvailWidth Then ' Not fine - try to fit

    ' ----- BUILD COLLECTION OF COLUMN RANGES ---------------------------------

        Set ColumnCollect = New Collection
        For D = 1 To fitRange.Columns.Count
            Set rngColumn = fitRange.Range(Cells(1, D), _
                Cells(fitRange.Rows.Count, D))
            strColumn = rngColumn.Column
            ColumnCollect.Add rngColumn, strColumn
        Next D

    ' ----- LOOP - CALCULATE BEST COLUMN WIDTH --------------------------------
        Do
            dCount = dCount + 1  ' Counter to prevent infinite loop
            ' reset tests
            blnSuccess = False
            blnStop = True
            
            ' May need to add error handling for 0 columns? Or is just fail OK?
            If ColumnCollect.Count > 0 Then
                ' Record current average column width
                lngStartAvgColW = lngAvailWidth / ColumnCollect.Count
                
                ' Remove any columns already smaller than average
                For Each rngCheckColumn In ColumnCollect
                    If rngCheckColumn.Width < lngStartAvgColW Then
                        lngAvailWidth = lngAvailWidth - rngCheckColumn.Width
                        strColumn = rngCheckColumn.Column
                        ColumnCollect.Remove (strColumn)
                    End If
                Next


                ' Re-calculate average col width (rm cols already < average)
                lngAvgColumnW = lngAvailWidth / ColumnCollect.Count
                Debug.Print "Loop " & dCount & ": " & lngAvgColumnW
                If lngAvgColumnW > 108 Then
                    Stop
                End If
                ' Check if average is at least as large as the minimum we set
                If lngAvgColumnW < lngMinColumnW Then
                ' If not, function is a failure
                        blnStop = True
                        blnSuccess = False
                Else ' avg column width is OK
                    ' If nothing was removed, we're good!
                    If lngStartAvgColW = lngAvgColumnW Then
                        blnStop = True
                        blnSuccess = True
                    Else ' removed some columns last pass, need to recalculate
                        blnStop = False
                        blnSuccess = False
                    End If
                End If
            End If
        Loop Until blnStop = True
    Else
        blnSuccess = True
    End If

'   ----- TEST IF BEST COLUMN WIDTH WAS A SUCCESS -----------------------------
    ' If autofit along was ok, dCount = 0 and we don't need to change anything
    If blnSuccess = True And dCount > 0 Then
    ' Set col width of all remaining too-wide columns to average
        If ColumnCollect.Count > 0 Then
            Dim rngFinalColumn As Range
            For Each rngFinalColumn In ColumnCollect
                rngFinalColumn.Width = lngAvgColumnW
            Next rngFinalColumn
        End If
        Set rngColumn = Nothing
    
        ' Wrap , for rows we just shortened
        fitRange.WrapText = True
    
        ' Auto-fit rows so nothing is cut off
        fitRange.Rows.AutoFit
    
    End If
    Debug.Print "--- END FitColumns: " & blnSuccess
    Set ColumnCollect = Nothing
    
    ' ----- RETURN IF THIS WAS A SUCCESS --------------------------------------
    FitColumns = blnSuccess

End Function

' ===== GetPageWidth ==========================================================
' Returns page width in points (72 points = 1 inch). PageSetup.PaperSize will
' give you the xlPaperSize enum, but not the actual page size. Right now just
' deals with Letter, Legal, and Tabloid. Could easily add international sizes
' in the future: http://www.printernational.org/iso-paper-sizes.php
' In which case, note that Application.CentimetersToPoints() exists but does
' not return round values, though Long will coerce to integer values.

Private Function GetPageWidth(objOrigPageSetup As PageSetup) As Long
    ' Get orientation of paper (i.e., which dimension is "wide"
    Dim currentOrientation As XlPageOrientation
    currentOrientation = objOrigPageSetup.Orientation

    Dim currentWidth As Long
    Select Case objOrigPageSetup.PaperSize

        Case xlPaperLetter
            If currentOrientation = xlLandscape Then
                currentWidth = Application.InchesToPoints(11)
            Else
                currentWidth = Application.InchesToPoints(8.5)
            End If
            
        Case xlPaperLegal
            If currentOrientation = xlLandscape Then
                currentWidth = Application.InchesToPoints(14)
            Else
                currentWidth = Application.InchesToPoints(8.5)
            End If

        Case xlPaperTabloid
            If currentOrientation = xlLandscape Then
                currentWidth = Application.InchesToPoints(17)
            Else
                currentWidth = Application.InchesToPoints(11)
            End If

        Case xlPaper11x17
            If currentOrientation = xlLandscape Then
                currentWidth = Application.InchesToPoints(17)
            Else
                currentWidth = Application.InchesToPoints(11)
            End If
            
        Case Else ' some other paper size, just quit everything
            currentWidth = 0
    
    End Select

    GetPageWidth = currentWidth

End Function


' ===== GetPaperSizes =========================================================
' You can only set the paper size of the current worksheet to sizes you have
' available on your current default printer (I think?). This checks all sizes
' and returns an array of just the possible ones. NOTE that it returns sizes
' as xlPaperSize enumerations, which are in fact Long. Actual enumeration is:
' https://msdn.microsoft.com/en-us/library/office/ff839964.aspx

' NOTE: xlPaperTabloid and xlPaper11x17 are the same size, but some printers
' support one while some support the other.

Private Function GetPaperSizes(getPaperSheet As Worksheet) As Variant

    ' Record original page size of sheet
    Dim origSize As XlPaperSize: origSize = getPaperSheet.PageSetup.PaperSize
    Dim availSizes() As Variant
    Dim I As Long
    Dim J As Long: J = 0    ' for availSizes index; will be base 1
    Dim Sizes(1 To 4) As XlPaperSize

    ' Only trying these paper sizes. If wanted to try ALL paper sizes, loop
    ' through numbers 1 to 41. But just because you can set it doesn't mean
    ' that the paper is available in your printer right now, just that your
    ' printer COULD handle it
    Sizes(1) = xlPaperLetter
    Sizes(2) = xlPaperLegal
    Sizes(3) = xlPaperTabloid
    Sizes(4) = xlPaper11x17
    
    ' If not available, throws Error 1004: "Unable to set the PaperSize
    ' property of the PageSetup class"
    On Error Resume Next

    For I = LBound(Sizes) To UBound(Sizes)

        getPaperSheet.PageSetup.PaperSize = I

        Select Case Err.Number

            Case 0  ' no error, the paper size is available
                J = J + 1
                ReDim Preserve availSizes(1 To J)
                availSizes(J) = I
                Err.Clear

            Case 1004 ' Unable to set the PaperSize property of the PageSetup class
                Err.Clear  ' So same error doesn't trip this next loop

            Case Else  ' some other untrapped error
                ' Tell user something went wrong and quit
                MsgBox Err.Number & ": " & Err.Description
                Exit For

        End Select
    Next I

    ' Reset error instructions
    Err.Clear
    On Error GoTo 0

    ' Reset original pagesize
    getPaperSheet.PageSetup.PaperSize = origSize

    ' Return array of available sizes
    GetPaperSizes = availSizes

End Function

