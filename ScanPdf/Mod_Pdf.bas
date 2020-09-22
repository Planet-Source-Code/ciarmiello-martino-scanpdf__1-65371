Attribute VB_Name = "Mod_Pdf"
Public strFileName As OPENFILENAME
Public Immagini(100) As String
Public ContImg As Integer

Public Declare Function lstrcpy Lib "Kernel" (lpszString1 As Any, lpszString2 As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)
Public Declare Function VeryOpen Lib "verywrite.dll" (ByVal lpFileName As String) As Long
Public Declare Function VeryCreate Lib "verywrite.dll" (ByVal lpFileName As String) As Long
Public Declare Sub VeryClose Lib "verywrite.dll" (ByVal id As Long)
Public Declare Function VeryAddImage Lib "verywrite.dll" (ByVal id As Long, ByVal lpFileName As String) As Long
Public Declare Function VeryDecryptPDF Lib "verywrite.dll" (ByVal inFileName As String, ByVal outFileName As String, ByVal OwnerPassword As String, ByVal UserPassword As String) As Long
Public Declare Function VeryIsPDFEncrypted Lib "verywrite.dll" (ByVal inFileName As String) As Long
Public Declare Function VeryEncryptPDF Lib "verywrite.dll" (ByVal inFileName As String, ByVal outFileName As String, ByVal EnctyptLen As Long, ByVal permission As Long, ByVal OwnerPassword As String, ByVal UserPassword As String) As Long
Public Declare Function VeryAddInfo Lib "verywrite.dll" (ByVal id As Long, ByVal Title As String, ByVal Subject As String, ByVal Author As String, ByVal Keywords As String, ByVal CREATOR As String) As Long
Public Declare Function VerySetFunction Lib "verywrite.dll" (ByVal id As Long, ByVal func_code As Long, ByVal Para1 As Long, ByVal Para2 As Long, ByVal szPara3 As String, ByVal szPara4 As String) As Long
Public Declare Function VeryStampOpen Lib "verywrite.dll" (ByVal sIn As String, ByVal sOut As String) As Long
Public Declare Sub VeryStampClose Lib "verywrite.dll" (ByVal id As Long)
Public Declare Function VeryStampAddText Lib "verywrite.dll" (ByVal id As Long, ByVal position As Long, ByVal sstring As String, ByVal color As Long, ByVal alignment As Long, ByVal shift_lr As Long, ByVal shift_tb As Long, ByVal Rotate As Long, ByVal layer As Long, ByVal hollow As Long, ByVal fontcode As Long, ByVal FontName As String, ByVal Fontsize As Long, ByVal Action As Long, ByVal link As String, ByVal pageno As Long) As Long
Public Declare Function VeryAddImageData Lib "verywrite.dll" (ByVal id As Long, ByVal MemData As Long, ByVal width As Long, ByVal height As Long, ByVal color As Long) As Long
Public Declare Function VeryAddTextEx Lib "verywrite.dll" (ByVal id As Long, ByVal x As Long, ByVal y As Long, ByVal str As String, ByVal color As Long) As Long
Public Declare Function VeryAddText1 Lib "verywrite.dll" Alias "VeryAddText" (ByVal id As Long, ByVal x As Long, ByVal y As Long, ByRef width As Long, ByRef height As Long, ByVal str As String, ByVal color As Long, ByVal bkcolor As Long, ByVal lformat As Long) As Long
Public Declare Function VeryAddText2 Lib "verywrite.dll" Alias "VeryAddText" (ByVal id As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal str As String, ByVal color As Long, ByVal bkcolor As Long, ByVal lformat As Long) As Long
Public Declare Function VeryAddTxt Lib "verywrite.dll" (ByVal id As Long, ByVal lpFileName As String, ByVal color As Long, ByVal pagewidth As Long, ByVal pageheight As Long, ByVal AutoNewLine As Long, ByVal AutoWidthAdjust As Long, ByVal LeftMargin As Long, ByVal RightMargin As Long, ByVal TopMargin As Long, ByVal BottomMargin As Long, ByVal TabSize As Long) As Long
Public Declare Function VeryAddLine Lib "verywrite.dll" (ByVal id As Long, ByVal sx As Long, ByVal sy As Long, ByVal ex As Long, ByVal ey As Long, ByVal side_width As Long, ByVal side_color As Long) As Long
Public Declare Function VeryAddRect Lib "verywrite.dll" (ByVal id As Long, ByVal sx As Long, ByVal sy As Long, ByVal ex As Long, ByVal ey As Long, ByVal side_width As Long, ByVal side_color As Long, ByVal flagFill As Long, ByVal fill_color As Long) As Long
Public Declare Function VeryGetFunction Lib "verywrite.dll" (ByVal id As Long, ByVal func_code As Long, ByVal Para1 As Long, ByVal Para2 As Long, ByVal szPara3 As String, ByVal szPara4 As String) As Long
Public Declare Function VerySplitMergePDF Lib "verywrite.dll" (ByVal szCommand As String) As Long
Public Declare Function VeryStampSetFunction Lib "verywrite.dll" (ByVal id As Long, ByVal func_code As Long, ByVal Para1 As Long, ByVal Para2 As Long, ByVal szPara3 As String, ByVal szPara4 As String) As Long
Public Declare Function VeryStampAddImage Lib "verywrite.dll" (ByVal id As Long, ByVal position As Long, ByVal FileName As String, ByVal shift_lr As Long, ByVal shift_tb As Long, ByVal Rotate As Long, ByVal layer As Long, ByVal zoomW As Long, ByVal zoomH As Long, ByVal Action As Long, ByVal link As String, ByVal pageno As Long) As Long
Public Declare Function VeryStampAddLine Lib "verywrite.dll" (ByVal id As Long, ByVal position As Long, ByVal line_width As Long, ByVal color As Long, ByVal shift_lr As Long, ByVal shift_tb As Long, ByVal Rotate As Long, ByVal layer As Long, ByVal zoomW As Long, ByVal zoomH As Long) As Long

Public Declare Function TWAIN_AcquireToClipboard Lib "EZTW32.DLL" (ByVal hwndApp&, ByVal wPixTypes&) As Long
Public Declare Function TWAIN_SelectImageSource Lib "EZTW32.DLL" (ByVal hwndApp&) As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Public Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type


Public Function File_CommonDialog_Open(Optional FileFilter As String = "All Files|*.*", Optional DefaultExtention As String = "*.*", Optional CommonDialog_Title As String = "Apri") As String
    On Error Resume Next
    Dim lngReturnValue As Long
    Dim intRest As Integer
    strFileName.lpstrTitle = CommonDialog_Title
    strFileName.lpstrDefExt = DefaultExtention
    DialogFilter (FileFilter)
    strFileName.hInstance = App.hInstance
    strFileName.lpstrFile = Space(259)
    strFileName.nMaxFile = 260
    strFileName.Flags = &H4
    strFileName.lStructSize = Len(strFileName)
    lngReturnValue = GetOpenFileName(strFileName)
    File_CommonDialog_Open = strFileName.lpstrFile
    DoEvents
End Function


Public Sub DialogFilter(WantedFilter As String)
    Dim intLoopCount As Integer
    strFileName.lpstrFilter = ""
    For intLoopCount = 1 To Len(WantedFilter)
        If Mid$(WantedFilter, intLoopCount, 1) = "|" Then strFileName.lpstrFilter = _
        strFileName.lpstrFilter Else strFileName.lpstrFilter = _
        strFileName.lpstrFilter + Mid$(WantedFilter, intLoopCount, 1)
    Next intLoopCount
    strFileName.lpstrFilter = strFileName.lpstrFilter
End Sub


Public Function File_CommonDialog_Save(Optional FileFilter As String = "All Files|*.*", Optional DefaultExtention As String = "*.*", Optional CommonDialog_Title As String = "Salva con nome") As String
    On Error Resume Next
    Dim lngReturnValue As Long
    Dim intRest As Integer
    strFileName.lpstrTitle = CommonDialog_Title
    strFileName.lpstrDefExt = DefaultExtention
    DialogFilter (FileFilter)
    strFileName.hInstance = App.hInstance
    strFileName.lpstrFile = Chr(0) & Space(259)
    strFileName.nMaxFile = 260
    strFileName.Flags = &H80000 Or &H4
    strFileName.lStructSize = Len(strFileName)
    lngReturnValue = GetSaveFileName(strFileName)
    File_CommonDialog_Save = strFileName.lpstrFile
    DoEvents
End Function


Public Function GestioneImmagini(bmpPath As String, jpgPath As String, pictureName As Object)
    Dim oBmp As New FREEIMAGE_ACTIVEXLib.Image
    Dim ergBmp As Integer
    
    SavePicture pictureName, bmpPath
    
    On Error Resume Next
    ergBmp = oBmp.Open(bmpPath)
    If Err.Number <> 0 Or ergBmp <> 1 Then
        MsgBox "File BMP not Valid!"
    Else
        ergBmp = oBmp.Save(jpgPath)
        If Err.Number <> 0 Or ergBmp <> 1 Then
            MsgBox "Conversion BMP To JPG Failed!", vbOKOnly + vbCritical
            Exit Function
        End If
    End If
    oBmp.Close
    Set oBmp = Nothing
    Kill bmpPath
    DoEvents
End Function


Public Function CreaPDF(NomeFile As String, Stampa_Pdf As Long, Copia_Pdf As Long, Modifica_Pdf As Long, Password As String, Pagine As Integer, Optional CopiaConforme As Boolean)
    Dim FileName As String
    Dim id As Long
    Dim totale As Long
    Dim ContPag As Integer
    FileName = NomeFile & ".pdf"
    id = VeryCreate(FileName)
        
    If (id > 0) Then
        ret = VerySetFunction(id, 101, 1, Pagine, "", "")
        For ContPag = 1 To Pagine
            ret = VeryAddImage(id, Immagini(ContPag))
        Next ContPag
        VeryClose id
    End If

    For ContPag = 1 To Pagine
        If Left(Immagini(ContPag), Len(Immagini(ContPag)) - 4 - Len(Trim(ContPag))) = NomeFile Then Kill Immagini(ContPag)
    Next ContPag

    If CopiaConforme = True Then
        id = VeryStampOpen(FileName, FileName)
        If (id > 0) Then
            Code = VeryStampAddText(id, 9, "COPY IN COMPLIANCE WHIT ORIGINATES", 12632256, 17, 10, 0, 45, 0, 1, 0, "", 50, 0, "", 0)
            Code = VeryStampAddText(id, 5, "", 0, 0, -50, -30, 0, 0, 0, 0, NullString, 2, 0, NullString, 0)
            VeryStampClose (id)
        End If
    End If
End Function


Public Sub VisualizzaPdf(PDFFileName As String)
    On Error Resume Next
    If Trim(PDFFileName) = "" Then Exit Sub
    If Dir(PDFFileName) <> "" Then
        If Dir("C:\Programmi\Adobe\Acrobat 5.0\Reader\AcroRd32.exe") <> "" Then AppPercorso = "C:\Programmi\Adobe\Acrobat 5.0\Reader\AcroRd32.exe"
        If Dir("D:\Programmi\Adobe\Acrobat 5.0\Reader\AcroRd32.exe") <> "" Then AppPercorso = "D:\Programmi\Adobe\Acrobat 5.0\Reader\AcroRd32.exe"
        If Dir("C:\Programmi\Adobe\Acrobat 5.0\Acrobat\Acrobat.exe") <> "" Then AppPercorso = "C:\Programmi\Adobe\Acrobat 5.0\Acrobat\Acrobat.exe"
        If Dir("D:\Programmi\Adobe\Acrobat 5.0\Acrobat\Acrobat.exe") <> "" Then AppPercorso = "D:\Programmi\Adobe\Acrobat 5.0\Acrobat\Acrobat.exe"
        If Dir("C:\Programmi\Adobe\Acrobat 6.0\Acrobat\Acrobat.exe") <> "" Then AppPercorso = "C:\Programmi\Adobe\Acrobat 6.0\Acrobat\Acrobat.exe"
        If Dir("D:\Programmi\Adobe\Acrobat 6.0\Acrobat\Acrobat.exe") <> "" Then AppPercorso = "D:\Programmi\Adobe\Acrobat 6.0\Acrobat\Acrobat.exe"
        If AppPercorso <> "" Then Call Shell(AppPercorso & " " & PDFFileName) Else MsgBox "It has not been installed Reader of Acrobat", vbOKOnly, "Operation cancelled"
    End If
End Sub


Public Sub AcquisisciDoc(PdfPercorso As String, PDFFileName As String, Optional CopiaConforme As Boolean = False)
    Dim ContP As Integer
    Dim Pagine As Integer
    Dim ResPagine As String
    If Right(PdfPercorso, 1) = "\" Then PdfPercorso = Left(PdfPercorso, Len(PdfPercorso) - 1)
    If Dir(PdfPercorso & "\" & PDFFileName & ".pdf") <> "" Then
        MsgBox "Already existing document!", vbOKOnly + vbCritical, "Operation cancelled"
    Else
        frm_SelezioneDocumento.mnuModifica.Enabled = False
        ResPagine = InputBox("Digitare il numero di pagine del documento da acquisire.", "Document Acquisition", 1)
        If Val(ResPagine) > 0 Then Pagine = Val(ResPagine) Else Exit Sub
        ContImg = 0
        For ContP = 1 To Pagine
            MsgBox "Insert the page " & ContP & " of " & Pagine & " of the document to acquire," & vbCrLf & "then press OK in order to start the next scansion.", vbOKOnly, "Document Acquisition"
            Call ZScanPdf.Acquisizione(PdfPercorso & "\" & PDFFileName & ContP)
            ContImg = ContImg + 1
            Immagini(ContImg) = PdfPercorso & "\" & PDFFileName & ContP & ".jpg"
        Next ContP
        Call CreaPDF(PdfPercorso & "\" & PDFFileName, 1, 0, 0, "", Pagine, CopiaConforme)
        DoEvents
        intMsg = MsgBox("The document " & PDFFileName & ".pdf it has been saved correctly!" & vbCrLf & vbCrLf & "Want you open the new file ?", vbYesNo + vbDefaultButton2 + vbInformation, "ScanPdf")
        If intMsg = vbYes Then Call VisualizzaPdf(PdfPercorso & "\" & PDFFileName & ".pdf")
    End If
End Sub


Public Sub Allega(PdfPercorso As String, PDFFileName As String)
    Call ZScanPdf.Acquisizione(PdfPercorso & "\" & PDFFileName & ContImg + 1)
    If Dir(PdfPercorso & "\" & PDFFileName & ContImg + 1 & ".jpg") <> "" Then
        ContImg = ContImg + 1
        Immagini(ContImg) = PdfPercorso & "\" & PDFFileName & ContImg & ".jpg"
        frm_SelezioneDocumento.List1.AddItem Immagini(ContImg)
        frm_SelezioneDocumento.mnuScan.Enabled = False
        frm_SelezioneDocumento.mnuConverti.Enabled = False
    End If
End Sub


Public Sub Salva(PdfPercorso As String, PDFFileName As String, Optional CopiaConforme As Boolean = False)
    If ContImg > 0 Then
        Call CreaPDF(PdfPercorso & "\" & PDFFileName, 1, 0, 0, "", ContImg, CopiaConforme)
        DoEvents
        intMsg = MsgBox("The document " & PDFFileName & ".pdf it has been saved correctly!" & vbCrLf & vbCrLf & "Want you open the new file ?", vbYesNo + vbDefaultButton2 + vbInformation, "ScanPdf")
        If intMsg = vbYes Then Call VisualizzaPdf(PdfPercorso & "\" & PDFFileName & ".pdf")
    End If
End Sub


Public Sub CentraForm(myForm As Form)
    On Error Resume Next
    myForm.WindowState = 0
    myForm.Left = (Screen.width - myForm.width) / 2
    myForm.Top = ((Screen.height - myForm.height) / 2) - 670
End Sub


Public Sub SvuotaVar()
    Dim cont As Integer
    For cont = 1 To 100
        Immagini(cont) = ""
    Next cont
    ContImg = 0
End Sub


Public Sub AcquisisciWord(FileWord As String, WPercorso As String, WFileName As String, Optional CopiaConforme As Boolean = False)
    On Error Resume Next
    Dim wApp As Word.Application
    Dim wdoc As Word.Document
    Dim DocumentoWord As String
    Dim ContP As Integer
    Dim PageCount As Integer
    Dim CountSu As Integer
    Screen.MousePointer = vbHourglass
    frm_SelezioneDocumento.Visible = False
    DoEvents
    DocumentoWord = Left(FileWord, Len(FileWord) - 4) & "2.doc"
    Copia = CopyFile(FileWord, DocumentoWord, True)
    DoEvents
    Set wApp = New Word.Application
    Set wdoc = wApp.Documents.Open(DocumentoWord)
    DoEvents
    wApp.Visible = True
    PageCount = wdoc.ActiveWindow.Selection.Information(wdNumberOfPagesInDocument)
    wdoc.Close (False)
    wApp.Quit (False)
    DoEvents
    For ContP = 1 To PageCount
        Set wApp = New Word.Application
        Set wdoc = wApp.Documents.Open(DocumentoWord)
        DoEvents
        wApp.Visible = True
        DoEvents
        Sleep 500
        Clipboard.Clear
        wdoc.ActiveWindow.Selection.WholeStory
        wdoc.ActiveWindow.Selection.CopyAsPicture
        DoEvents
        If wdoc.ActiveWindow.Selection.Information(wdNumberOfPagesInDocument) > 1 Then
            wdoc.ActiveWindow.Selection.GoTo wdGoToPage, wdGoToAbsolute, 2
            For CountSu = 1 To 100
                SendKeys "+{UP}", True
            Next CountSu
            Sleep 300
            SendKeys "{DELETE}", True
            Sleep 300
            DoEvents
        End If
        wdoc.Close (True)
        wApp.Quit (True)
        DoEvents
        ContImg = ContP
        Immagini(ContP) = WPercorso & "\" & WFileName & ContP & ".jpg"
        Copia = CopyFile(App.Path & "\Temp.JPG", WPercorso & "\" & WFileName & ContP & ".jpg", True)
        DoEvents
        Sleep 300
        ReturnValue = Shell("mspaint " & DosPath(WPercorso & "\" & WFileName & ContP & ".jpg"), vbMaximizedFocus)
        DoEvents
        Sleep 1000
        AppActivate ReturnValue
        Sleep 200
        SendKeys "%f", True
        Sleep 100
        SendKeys "{RIGHT}", True
        Sleep 100
        SendKeys "{DOWN}", True
        Sleep 100
        SendKeys "{DOWN}", True
        Sleep 100
        SendKeys "{DOWN}", True
        Sleep 100
        SendKeys "{DOWN}", True
        Sleep 100
        SendKeys "{ENTER}", True
        Sleep 100
        SendKeys "%{F4}", True
        Sleep 300
        SendKeys "{ENTER}", True
        Sleep 300
        DoEvents
    Next ContP
    DoEvents
    frm_SelezioneDocumento.Visible = True
    Call Salva(WPercorso, WFileName)
    Screen.MousePointer = vbDefault
End Sub


Public Function DosPath(Percorso As String) As String
    Dim ContC As Double
    Dim TempPath As String
    Dim TempStr As String
    Dim ContTS As Integer
    Dim TempPos As Double
    Dim Estens As Integer
    TempPos = 1
    For ContC = 1 To Len(Percorso)
        Estens = 0
        If Mid(Percorso, ContC, 1) = "\" Then Estens = 1
        If Mid(Percorso, ContC, 4) = ".jpg" Then Estens = 3
        If Estens > 0 Then
            If InStr(1, Mid(Percorso, TempPos + 1, ContC - TempPos - 1), " ") > 0 Then
                ContTS = TempPos
                TempStr = ""
                If Len(Mid(Percorso, TempPos + 1, ContC - TempPos - 1)) > 6 Then
                    While ContTS < ContC - 1
                        ContTS = ContTS + 1
                        If Mid(Percorso, ContTS, 1) <> " " Then TempStr = TempStr & Mid(Percorso, ContTS, 1)
                        If Len(TempStr) = 6 Then GoTo Fine
                    Wend
                Else
                    TempStr = Replace(Mid(Percorso, TempPos + 1, ContC - TempPos - 1), " ", "")
                End If
Fine:
                TempStr = TempStr & "~1"
                TempPath = TempPath & "\" & TempStr
            Else
                If TempPath <> "" Then TempPath = TempPath & "\"
                If TempPath <> "" Then TempPath = TempPath & Mid(Percorso, TempPos + 1, ContC - TempPos - 1) Else TempPath = TempPath & Mid(Percorso, TempPos, ContC - TempPos)
            End If
            TempPos = ContC
        End If
    Next ContC
    DosPath = TempPath & ".jpg"
End Function



'Public Sub WordToPdf(strDOC As String, strPDF As String)
'    On Error Resume Next
'    Dim WordToPdf As PDFmaker.CreatePDF
'    Call WordToPdf.CloseAcrobat
'    WordToPdf.CreatePDFfromPowerPoint strPDF, strDOC, blnApplySecurity:=chkSecurity.Value
'    WordToPdf.OpenPDF strPDF, , 1, 1, ""
'End Sub
