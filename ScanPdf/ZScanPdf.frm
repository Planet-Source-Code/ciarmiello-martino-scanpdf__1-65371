VERSION 5.00
Begin VB.Form ZScanPdf 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Acquisizione Documento"
   ClientHeight    =   1530
   ClientLeft      =   1590
   ClientTop       =   1890
   ClientWidth     =   3090
   FillColor       =   &H80000000&
   Icon            =   "ZScanPdf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1530
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   -6105
      ScaleHeight     =   1470
      ScaleWidth      =   9135
      TabIndex        =   0
      Top             =   0
      Width           =   9195
      Begin VB.PictureBox picScreen 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   5964
         Left            =   0
         ScaleHeight     =   5970
         ScaleWidth      =   9660
         TabIndex        =   1
         Top             =   0
         Width           =   9660
         Begin VB.PictureBox picHolder 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   636
            Left            =   2424
            ScaleHeight     =   630
            ScaleWidth      =   1455
            TabIndex        =   2
            Top             =   1488
            Visible         =   0   'False
            Width           =   1452
         End
         Begin VB.Line linTool 
            Visible         =   0   'False
            X1              =   36
            X2              =   36
            Y1              =   98
            Y2              =   114
         End
      End
   End
End
Attribute VB_Name = "ZScanPdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Acquisizione(NomeFile As String)
    On Error Resume Next
    Me.Visible = False
    R = TWAIN_AcquireToClipboard(Me.hWnd, t%)
    Picture1.Picture = Clipboard.GetData(vbCFDIB)
    DoEvents
    If Picture1.Picture <> 0 Then Else Exit Sub
    Call GestioneImmagini(NomeFile & ".bmp", NomeFile & ".jpg", Picture1.Picture)
    DoEvents
    Unload Me
End Sub

