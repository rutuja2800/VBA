Sub ExtractSpecificSheetsToFiles()
    Dim selectedSheets As Variant
    Dim customFileNames As Variant
    Dim i As Integer
    Dim ws As Worksheet
    Dim newWB As Workbook
    Dim newWS As Worksheet
    Dim cellData As Range
    Dim fileName As String
    Dim savePath As String
    Dim sampleFormatPath As String ' Path to the sample format file
    
    ' Path to the sample format file
    sampleFormatPath = "D:\c\Price book uploads\Sample files round.xlsx"
    
    ' Select folder path to save the files
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder to Save Extracted Files"
        .Show
        If .SelectedItems.Count > 0 Then
            savePath = .SelectedItems(1) & "\"
        Else
            MsgBox "No folder selected. Exiting process.", vbExclamation
            Exit Sub
        End If
    End With
    
    ' Define the sheet names you want to extract and corresponding custom file names
    selectedSheets = Array("RD_RAP", "WH EX (NON)", "WH VG (NON)", "WH GD (NON)", _
                           "INCLUSION", "ROUND_INCLUSION", "LOOSE (WHT) (OLD) ", "LOOSE (LB) (OLD)", _
                           "SNG SMC BACK", "MFG")
    
    customFileNames = Array("RBC_RapaportSampleFile.xlsx", "WH-EX-NON.xlsx", "WH-VG-NON.xlsx", _
                            "WH-GD-NON.xlsx", "LooseInclusionSampleFile_new.xlsx", "RoundInclusionSampleFile_new", "LooseWH-LBSampleFile_WH.xlsx", _
                            "LooseWH-LBSampleFile_LB.xlsx", "smccut_sample", "MFGSampleFile.xlsx")
    
    ' Loop through the selected sheets and extract data
    For i = LBound(selectedSheets) To UBound(selectedSheets)
        ' Find the sheet by name
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(selectedSheets(i))
        On Error GoTo 0
        
        ' If sheet exists, extract data and create a new file
        If Not ws Is Nothing Then
            ' Define the file name for the new workbook
            fileName = customFileNames(i)
            
            ' Define the range for each sheet differently (modify these ranges as needed)
            Dim sourceRange As Range
            Select Case selectedSheets(i)
                Case "INCLUSION"
                    Set sourceRange = ws.Range("N1:Y1575")
                Case "ROUND_INCLUSION"
                    Set sourceRange = ws.Range("A1:EY1575")
                Case "LOOSE (LB) (OLD)"
                    Set sourceRange = ws.Range("N1:Y1575")
                ' Add more cases for other sheets as needed...
                Case "SNG SMC BACK"
                    Set sourceRange = ws.Range("A1:SY1576")
                Case "WH VG (NON)"
                    Set sourceRange = ws.Range("A1:L1575")
                Case "WH GD (NON)"
                    Set sourceRange = ws.Range("A1:L1575")
                Case "LOOSE (WHT) (OLD) "
                    Set sourceRange = ws.Range("A1:L1575")
                Case "MFG"
                    Set sourceRange = ws.Range("A1:JL1579")
                Case Else
                    ' Default range if not specified
                    Set sourceRange = ws.Range("A1:L1590")
            End Select
            
            ' Open the sample format file
            Dim formatWB As Workbook
            Set formatWB = Workbooks.Open(sampleFormatPath)
            
            ' Define the target range in the sample format sheet (modify these ranges as needed)
            Dim targetRange As Range
            Select Case selectedSheets(i)
                Case "INCLUSION"
                    Set targetRange = formatWB.Sheets("INCLUSION").Range("N1") ' Modify the range in the sample format sheet where you want to paste
                Case "ROUND_INCLUSION"
                    Set targetRange = formatWB.Sheets("ROUND_INCLUSION").Range("FA1") ' Modify the range in the sample format sheet where you want to paste
                ' Add more cases for other sheets as needed...
                Case "SNG SMC BACK"
                    Set targetRange = formatWB.Sheets("SNG SMC BACK").Range("TA1") ' Modify the range in the sample format sheet where you want to paste
                Case "MFG"
                    Set targetRange = formatWB.Sheets("MFG").Range("JN1") ' Modify the range in the sample format sheet where you want to paste
                Case "LOOSE (WHT) (OLD) "
                    Set targetRange = formatWB.Sheets("LOOSE (WHT) (OLD) ").Range("N1") ' Modify the range in the sample format sheet where you want to paste
                Case "LOOSE (LB) (OLD)"
                    Set targetRange = formatWB.Sheets("LOOSE (LB) (OLD)").Range("N1")
                Case "WH EX (NON)"
                    Set targetRange = formatWB.Sheets("WH EX (NON)").Range("N1")
                Case "WH VG (NON)"
                    Set targetRange = formatWB.Sheets("WH VG (NON)").Range("N1")
                Case "WH GD (NON)"
                    Set targetRange = formatWB.Sheets("WH GD (NON)").Range("N1")
                Case Else
                    ' Default range if not specified
                    Set targetRange = formatWB.Sheets("RD_RAP").Range("N1") ' Modify the default range in the sample format sheet where you want to paste
            End Select
            
            ' Copy data from pricebook sheet to corresponding sheet in the sample format file
            sourceRange.Copy
            targetRange.PasteSpecial Paste:=xlPasteValues ' Paste values only
            Application.CutCopyMode = False ' Clear the clipboard
            
            ' Create a new workbook and copy values and formats from the modified sample format sheet
            Set newWB = Workbooks.Add
            Set newWS = newWB.Sheets(1)
            formatWB.Sheets(selectedSheets(i)).UsedRange.Copy
            newWS.Range("A1").PasteSpecial Paste:=xlPasteValues
            newWS.Range("A1").PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False ' Clear the clipboard
            
            ' Define the columns to clear in the new worksheet
            Dim columnsToClear As Range

            Select Case selectedSheets(i)
                Case "MFG"
                    Set columnsToClear = newWS.Range("JN:AZZ")
                Case "ROUND_INCLUSION"
                    Set columnsToClear = newWS.Range("FA:AZZ")
                Case "SNG SMC BACK"
                    Set columnsToClear = newWS.Range("TA:AZZ")
                Case Else
                    Set columnsToClear = newWS.Range("N:AZZ")
            End Select

            ' Clear contents in the specified columns
            If Not columnsToClear Is Nothing Then
                columnsToClear.ClearContents
            End If
            
            ' Auto-fit columns in the new worksheet
            newWS.Cells.EntireColumn.AutoFit
            
            ' Save the new workbook and close it
            newWB.SaveAs savePath & fileName
            newWB.Close SaveChanges:=False
            
            ' Close the sample format file without saving changes
            formatWB.Close SaveChanges:=False
            Set formatWB = Nothing ' Release object reference
            
        End If
    Next i
End Sub
