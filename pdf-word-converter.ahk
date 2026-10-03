#Requires AutoHotkey v2.0

; PDF - Word Converter - Word ↔ PDF with Drag & Drop Support
; Uses Microsoft Word COM automation for conversions

; Global variables
global fileList := []
global mainGui := ""
global fileListView := ""
global progressBar := ""
global statusText := ""
global outputFolder := ""
global outputFolderEdit := ""
global wordFormat := "docx"   ; "docx" or "doc"

; Create the main GUI
CreateGUI()

CreateGUI() {
    global mainGui, fileListView, progressBar, statusText

    ; Create main window
    mainGui := Gui("+Resize -MaximizeBox", "PDF - Word Converter")
    mainGui.SetFont("s10", "Segoe UI")
    mainGui.BackColor := "0xF0F0F0"

    ; Title
    mainGui.SetFont("s14 Bold", "Segoe UI")
    mainGui.Add("Text", "x20 y20 w560 Center", "PDF - Word Converter")
    mainGui.SetFont("s10", "Segoe UI")
    mainGui.Add("Text", "x20 y50 w560 Center c0x666666", "Convert between Word and PDF formats")

    ; Drag & Drop Zone
    mainGui.SetFont("s11", "Segoe UI")
    dropZone := mainGui.Add("Text", "x20 y90 w560 h100 Border Center 0x200", "")
    dropZone.SetFont("s11", "Segoe UI")
    mainGui.Add("Text", "x20 y120 w560 h40 Center BackgroundTrans c0x0078D4", "📁 Drag & Drop Files Here`nor click Browse to select files")

    ; Browse button
    browseBtn := mainGui.Add("Button", "x230 y200 w140 h35", "📂 Browse Files")
    browseBtn.OnEvent("Click", BrowseFiles)

    ; File List
    mainGui.Add("Text", "x20 y250 w560", "Selected Files:")
    fileListView := mainGui.Add("ListView", "x20 y275 w560 h150", ["Filename", "Type", "Status"])
    fileListView.ModifyCol(1, 300)
    fileListView.ModifyCol(2, 100)
    fileListView.ModifyCol(3, 140)

    ; Remove and Clear buttons
    removeBtn := mainGui.Add("Button", "x20 y435 w120 h30", "❌ Remove")
    removeBtn.OnEvent("Click", RemoveSelected)

    clearBtn := mainGui.Add("Button", "x150 y435 w120 h30", "🗑️ Clear All")
    clearBtn.OnEvent("Click", ClearAll)

    ; Word output format radio buttons
    mainGui.SetFont("s10", "Segoe UI")
    mainGui.Add("Text", "x20 y472 w170 h20 +0x200", "Word Output Format:")
    wordFormatDocx := mainGui.Add("Radio", "x195 y470 w90 h20 Checked", "DOCX")
    wordFormatDocx.OnEvent("Click", (*) => (wordFormat := "docx"))
    wordFormatDoc := mainGui.Add("Radio", "x295 y470 w80 h20", "DOC")
    wordFormatDoc.OnEvent("Click", (*) => (wordFormat := "doc"))

    ; Output folder row
    mainGui.SetFont("s10", "Segoe UI")
    mainGui.Add("Text", "x20 y503 w105 h25 +0x200", "Output Folder:")
    outputFolderEdit := mainGui.Add("Edit", "x130 y500 w390 h25 ReadOnly", "(same as source)")
    folderBrowseBtn := mainGui.Add("Button", "x525 y499 w55 h27", "📁")
    folderBrowseBtn.OnEvent("Click", SelectOutputFolder)

    ; Single convert button
    mainGui.SetFont("s11 Bold", "Segoe UI")
    convertBtn := mainGui.Add("Button", "x20 y535 w560 h45", "⚡ Convert All")
    convertBtn.OnEvent("Click", ConvertFiles)

    ; Progress bar
    mainGui.SetFont("s10", "Segoe UI")
    progressBar := mainGui.Add("Progress", "x20 y595 w560 h20", 0)

    ; Status text
    statusText := mainGui.Add("Text", "x20 y623 w560 h30 c0x006600", "Ready. Add files to begin.")

    ; Enable drag and drop
    mainGui.OnEvent("DropFiles", HandleDrop)

    ; Show the GUI
    mainGui.Show("w600 h670")
}

; Select output folder
SelectOutputFolder(*) {
    global outputFolder, outputFolderEdit

    chosen := DirSelect("", 3, "Select Output Folder")
    if (chosen != "") {
        outputFolder := chosen
        outputFolderEdit.Value := chosen
    }
}

; Browse for files
BrowseFiles(*) {
    global fileList, fileListView

    selectedFiles := FileSelect("M", "", "Select Word or PDF files", "Word/PDF Files (*.docx; *.doc; *.pdf)")

    if (selectedFiles.Length = 0)
        return

    ; First element is the directory, rest are filenames
    directory := selectedFiles[1]

    ; If only one file selected, it's the full path
    if (selectedFiles.Length = 1) {
        AddFileToList(directory)
    } else {
        ; Multiple files - combine directory with each filename
        loop selectedFiles.Length - 1 {
            filePath := directory "\" selectedFiles[A_Index + 1]
            AddFileToList(filePath)
        }
    }

    UpdateStatus("Files added. Select conversion type.")
}

; Handle drag and drop
HandleDrop(GuiObj, GuiCtrlObj, FileArray, X, Y) {
    for filePath in FileArray {
        AddFileToList(filePath)
    }
    UpdateStatus(FileArray.Length " file(s) added via drag & drop.")
}

; Add file to list
AddFileToList(filePath) {
    global fileList, fileListView

    ; Check if file already in list
    for file in fileList {
        if (file = filePath)
            return
    }

    ; Get file extension
    SplitPath(filePath, &fileName, , &fileExt)

    ; Validate file type
    if (!RegExMatch(fileExt, "i)^(doc|docx|pdf)$")) {
        UpdateStatus("Skipped: " fileName " (unsupported format)")
        return
    }

    ; Add to array and listview
    fileList.Push(filePath)

    fileType := (fileExt = "pdf") ? "PDF" : "Word"
    fileListView.Add("", fileName, fileType, "Pending")
}

; Remove selected file
RemoveSelected(*) {
    global fileList, fileListView

    rowNum := fileListView.GetNext()
    if (rowNum = 0) {
        UpdateStatus("No file selected to remove.")
        return
    }

    fileList.RemoveAt(rowNum)
    fileListView.Delete(rowNum)
    UpdateStatus("File removed.")
}

; Clear all files
ClearAll(*) {
    global fileList, fileListView

    fileList := []
    fileListView.Delete()
    progressBar.Value := 0
    UpdateStatus("All files cleared.")
}

; Convert files (autodetects type: Word→PDF, PDF→Word)
ConvertFiles(*) {
    global fileList, fileListView, progressBar, outputFolder, wordFormat

    if (fileList.Length = 0) {
        MsgBox("Please add files first!", "No Files", "Icon!")
        return
    }

    ; Initialize progress
    progressBar.Value := 0
    UpdateStatus("Starting conversion...")

    ; Try to create Word application
    try {
        wordApp := ComObject("Word.Application")
        wordApp.Visible := false
        wordApp.DisplayAlerts := 0
    } catch {
        MsgBox("Microsoft Word is not installed or unavailable!`n`nThis tool requires MS Word to perform conversions.", "Error", "IconX")
        return
    }

    ; Process each file
    successCount := 0
    failCount := 0
    total := fileList.Length

    loop total {
        filePath := fileList[A_Index]
        rowNum := A_Index

        SplitPath(filePath, &fileName, &fileDir, &fileExt, &fileNameNoExt)

        ; Determine output directory and path
        outDir := (outputFolder != "") ? outputFolder : fileDir
        isWordFile := RegExMatch(fileExt, "i)^(doc|docx)$")
        wordExt := (wordFormat = "doc") ? ".doc" : ".docx"
        outputPath := outDir "\" fileNameNoExt (isWordFile ? ".pdf" : wordExt)

        ; Update status
        fileListView.Modify(rowNum, , , , "Converting...")
        UpdateStatus("Converting " A_Index "/" total ": " fileName)

        ; Perform conversion based on autodetected type
        try {
            if (isWordFile)
                ConvertWordToPDF(wordApp, filePath, outputPath)
            else
                ConvertPDFToWord(wordApp, filePath, outputPath, wordFormat)

            fileListView.Modify(rowNum, , , , "✓ Completed")
            successCount++
        } catch {
            fileListView.Modify(rowNum, , , , "✗ Failed")
            failCount++
        }

        ; Update progress
        progressBar.Value := (A_Index / total) * 100
    }

    ; Close Word
    try {
        wordApp.Quit()
    }

    ; Final status
    UpdateStatus("Conversion complete! Success: " successCount ", Failed: " failCount)

    if (successCount > 0) {
        outMsg := (outputFolder != "") ? outputFolder : "same folder as source"
        MsgBox("Conversion complete!`n`nSuccess: " successCount "`nFailed: " failCount "`n`nOutput folder: " outMsg, "Complete", "Iconi")
    }
}

; Convert Word to PDF
ConvertWordToPDF(wordApp, inputPath, outputPath) {
    doc := wordApp.Documents.Open(inputPath)

    ; wdFormatPDF = 17
    doc.SaveAs2(outputPath, 17)
    doc.Close(0)
}

; Convert PDF to Word
ConvertPDFToWord(wordApp, inputPath, outputPath, fmt := "docx") {
    doc := wordApp.Documents.Open(inputPath)

    ; wdFormatDocument = 0 (.doc), wdFormatDocumentDefault = 16 (.docx)
    fmtCode := (fmt = "doc") ? 0 : 16
    doc.SaveAs2(outputPath, fmtCode)
    doc.Close(0)
}

; Update status text
UpdateStatus(message) {
    global statusText
    statusText.Value := message
}
