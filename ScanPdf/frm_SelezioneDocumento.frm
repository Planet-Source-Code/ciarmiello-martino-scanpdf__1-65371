VERSION 5.00
Begin VB.Form frm_SelezioneDocumento 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ScanPdf"
   ClientHeight    =   5685
   ClientLeft      =   1590
   ClientTop       =   2190
   ClientWidth     =   4950
   ClipControls    =   0   'False
   FillColor       =   &H80000000&
   Icon            =   "frm_SelezioneDocumento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5685
   ScaleWidth      =   4950
   Begin VB.Frame Frame1 
      Caption         =   "Destination selected for the Pdf file :"
      Height          =   975
      Left            =   60
      TabIndex        =   8
      Top             =   2760
      Width           =   4815
      Begin VB.TextBox TxtFileName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   300
         Width           =   4695
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Confirm Path"
      Height          =   315
      Left            =   3120
      TabIndex        =   4
      Top             =   3840
      Width           =   1635
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove Enclosed"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3120
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      Height          =   1425
      Left            =   60
      TabIndex        =   6
      Top             =   4200
      Width           =   4815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Copy in compliance with originates"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   3900
      Width           =   2835
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4815
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   2340
      Width           =   3795
   End
   Begin VB.Label Label2 
      Caption         =   "Name of File :"
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   2400
      Width           =   1035
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuApri 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuNuovo 
         Caption         =   "New"
      End
      Begin VB.Menu mnuEsci 
         Caption         =   "End"
      End
   End
   Begin VB.Menu mnuModifica 
      Caption         =   "&Edit"
      Begin VB.Menu mnuAllega 
         Caption         =   "Enclose from File"
      End
      Begin VB.Menu mnuAcquisisci 
         Caption         =   "Enclose from Scanner"
      End
      Begin VB.Menu mnuClipboard 
         Caption         =   "Enclose from Clipboard"
      End
      Begin VB.Menu mnuSalva 
         Caption         =   "Save"
      End
   End
   Begin VB.Menu mnuScan 
      Caption         =   "&Acquisition"
      Begin VB.Menu mnuAcquisisciDoc 
         Caption         =   "Acquire Document from Scanner"
      End
   End
   Begin VB.Menu mnuConverti 
      Caption         =   "&Convert"
      Begin VB.Menu mnuWord 
         Caption         =   "Convert Document from Word"
      End
   End
End
Attribute VB_Name = "frm_SelezioneDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    On Error Resume Next
    'Dim WordToPdf As PDFmaker.CreatePDF
    Drive1.Drive = "C:\"
    Dir1.Path = "C:\"
    Me.Tag = "File"
    CentraForm Me
    Call SvuotaVar
    frm_SelezioneDocumento.List1.Clear
    DoEvents
    Call Abilita(False)
    TxtFileName.Text = ""
End Sub


Private Sub Dir1_Change()
    If Right(Dir1.Path, 1) <> "\" Then
        If Text1.Text <> "" Then TxtFileName.Text = Dir1.Path & "\" & Text1.Text & ".pdf" Else TxtFileName.Text = Dir1.Path & "\Document.pdf"
    Else
        If Text1.Text <> "" Then TxtFileName.Text = Dir1.Path & Text1.Text & ".pdf" Else TxtFileName.Text = Dir1.Path & "Document.pdf"
    End If
End Sub


Private Sub Drive1_Change()
    On Error GoTo errore
    Dir1.Path = Drive1.Drive
    If Right$(Dir1.Path, 1) <> "\" Then
        If Text1.Text <> "" Then TxtFileName.Text = Dir1.Path & "\" & Text1.Text & ".pdf" Else TxtFileName.Text = Dir1.Path & "\Document.pdf"
    Else
        If Text1.Text <> "" Then TxtFileName.Text = Dir1.Path & Text1.Text & ".pdf" Else TxtFileName.Text = Dir1.Path & "Document.pdf"
    End If
    Exit Sub
    
errore:
    Drive1.Drive = "C:"
    Resume
End Sub


Private Sub Text1_Change()
    If Right(Dir1.Path, 1) <> "\" Then
        If Text1.Text <> "" Then TxtFileName.Text = Dir1.Path & "\" & Text1.Text & ".pdf" Else TxtFileName.Text = Dir1.Path & "\Document.pdf"
    Else
        If Text1.Text <> "" Then TxtFileName.Text = Dir1.Path & Text1.Text & ".pdf" Else TxtFileName.Text = Dir1.Path & "Document.pdf"
    End If
End Sub


Private Sub Command2_Click()
    If List1.ListCount > 0 Then
        List1.RemoveItem (List1.ListIndex)
        Call SvuotaVar
        For cont = 0 To List1.ListCount - 1
            List1.ListIndex = cont
            ContImg = ContImg + 1
            Immagini(ContImg) = List1.List(cont)
        Next cont
    End If
End Sub


Private Sub Command1_Click()
    If TxtFileName <> "" Then Call Abilita(True)
End Sub


Private Sub mnuAcquisisciDoc_Click()
    On Error Resume Next
    Dim Nome As String
    Dim CCAO As Boolean
    If Check1.Value = vbChecked Then CCAO = True Else CCAO = False
    If Text1.Text <> "" Then Nome = Text1.Text Else Nome = "Document"
    If TxtFileName <> "" Then
        Call AcquisisciDoc(Dir1.Path, Nome, CCAO)
        Call SvuotaVar
        frm_SelezioneDocumento.List1.Clear
        DoEvents
        Call Abilita(False)
        TxtFileName.Text = ""
    Else
        MsgBox "The destination selected for the file to acquire is not valid !", vbOKOnly + vbCritical, "Operation cancelled"
    End If
End Sub
Private Sub mnuWord_Click()
    On Error Resume Next
    Dim Nome As String
    Dim CCAO As Boolean
    Dim FileDoc As String
    If Check1.Value = vbChecked Then CCAO = True Else CCAO = False
    If Text1.Text <> "" Then Nome = Text1.Text Else Nome = "Document"
    If TxtFileName <> "" Then
        If Right(PdfPercorso, 1) = "\" Then PdfPercorso = Left(PdfPercorso, Len(PdfPercorso) - 1)
        If Dir(PdfPercorso & "\" & PDFFileName & ".pdf") <> "" Then
            MsgBox "Already existing document!", vbOKOnly + vbCritical, "Operation cancelled"
        Else
            FileDoc = Trim(File_CommonDialog_Open("*.doc", "*.doc", "Select Word document"))
            FileDoc = Left(FileDoc, Len(FileDoc) - 1)
            If FileDoc <> "" And Dir(FileDoc) <> "" And UCase(Right(FileDoc, 3)) = "DOC" Then
                Call AcquisisciWord(FileDoc, Dir1.Path, Nome, CCAO)
                Call Abilita(False)
            Else
                MsgBox "The selected file is not a valid Word document !", vbOKOnly + vbCritical, "Operation cancelled"
            End If
        End If
        Call SvuotaVar
        frm_SelezioneDocumento.List1.Clear
        DoEvents
        Call Abilita(False)
        TxtFileName.Text = ""
    Else
        MsgBox "The destination selected for the new file is not valid  !", vbOKOnly + vbCritical, "Operation cancelled"
    End If
End Sub
Private Sub mnuAcquisisci_Click()
    Dim Nome As String
    If Text1.Text <> "" Then Nome = Text1.Text Else Nome = "Document"
    If TxtFileName <> "" Then Call Allega(Dir1.Path, Nome)
End Sub
Private Sub mnuClipboard_Click()
    On Error Resume Next
    Dim Nome As String
    Dim Percorso As String
    If Text1.Text <> "" Then Nome = Text1.Text Else Nome = "Document"
    If Right(Dir1.Path, 1) <> "\" Then Percorso = Dir1.Path Else Percorso = Left(Dir1.Path, Len(Dir1.Path) - 1)
    If TxtFileName <> "" Then
        Copia = CopyFile(App.Path & "\Temp.JPG", Percorso & "\" & Nome & ContImg + 1 & ".jpg", True)
        ReturnValue = Shell("mspaint " & DosPath(Percorso & "\" & Nome & ContImg + 1 & ".jpg"), vbMaximizedFocus)
        Sleep 1000
        AppActivate ReturnValue
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
        Sleep 100
        DoEvents
        If Dir(Percorso & "\" & Nome & ContImg + 1 & ".jpg") <> "" Then
            ContImg = ContImg + 1
            Immagini(ContImg) = Percorso & "\" & Nome & ContImg & ".jpg"
            List1.AddItem Immagini(ContImg)
            mnuConverti.Enabled = False
            mnuScan.Enabled = False
        Else
            MsgBox "In the clipboard there are no valid selection !", vbOKOnly + vbCritical, "Operation cancelled"
        End If
    Else
        MsgBox "The destination selected for the new file is not valid  !", vbOKOnly + vbCritical, "Operation cancelled"
    End If
End Sub
Private Sub mnuAllega_Click()
    Dim FileImg As String
    FileImg = Trim(File_CommonDialog_Open("*.jpg", "*.jpg", "Allega immagine JPG"))
    FileImg = Left(FileImg, Len(FileImg) - 1)
    If FileImg <> "" And Dir(FileImg) <> "" And UCase(Right(FileImg, 3)) = "JPG" Then
        ContImg = ContImg + 1
        FileImg = Left(FileImg, Len(FileImg) - 1)
        Immagini(ContImg) = FileImg
        frm_SelezioneDocumento.List1.AddItem Immagini(ContImg)
        mnuScan.Enabled = False
        mnuConverti.Enabled = False
    Else
        MsgBox "The selected file is not valid !", vbOKOnly + vbCritical, "Operation cancelled"
    End If
End Sub
Private Sub mnuApri_Click()
    Dim Filepdf As String
    Call SvuotaVar
    frm_SelezioneDocumento.List1.Clear
    DoEvents
    Call Abilita(False)
    TxtFileName.Text = ""
    Filepdf = File_CommonDialog_Open("*.pdf", "*.pdf", "Apri documento PDF")
    Filepdf = Left(Filepdf, Len(Filepdf) - 1)
    If Filepdf <> "" And Dir(Filepdf) <> "" And UCase(Right(Filepdf, 3)) = "PDF" Then
        Call VisualizzaPdf(Filepdf)
    Else
        MsgBox "The selected file is not a valid document !", vbOKOnly + vbCritical, "Operation cancelled"
    End If
End Sub
Private Sub mnuEsci_Click()
    Unload Me
End Sub
Private Sub mnuNuovo_Click()
    Call SvuotaVar
    frm_SelezioneDocumento.List1.Clear
    DoEvents
    Call Abilita(False)
    TxtFileName.Text = ""
End Sub
Private Sub mnuSalva_Click()
    On Error Resume Next
    Dim Nome As String
    Dim CCAO As Boolean
    If Check1.Value = vbChecked Then CCAO = True Else CCAO = False
    If Text1.Text <> "" Then Nome = Text1.Text Else Nome = "Document"
    If TxtFileName <> "" Then Call Salva(Dir1.Path, Nome, CCAO) Else MsgBox "The destination selected for the new file is not valid  !", vbOKOnly + vbCritical, "Operation cancelled"
    Call SvuotaVar
    frm_SelezioneDocumento.List1.Clear
    DoEvents
    Call Abilita(False)
    TxtFileName.Text = ""
End Sub


Public Sub Abilita(AbilitaPdf As Boolean)
    mnuModifica.Enabled = AbilitaPdf
    mnuScan.Enabled = AbilitaPdf
    mnuConverti.Enabled = AbilitaPdf
    Command2.Enabled = AbilitaPdf
    List1.Enabled = AbilitaPdf
    Drive1.Enabled = Not AbilitaPdf
    Dir1.Enabled = Not AbilitaPdf
    Text1.Enabled = Not AbilitaPdf
    Command2.Visible = AbilitaPdf
    Command1.Visible = Not AbilitaPdf
End Sub
