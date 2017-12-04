VERSION 5.00
Begin VB.Form Form9 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Movie"
   ClientHeight    =   9675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10440
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9675
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   9705
      Left            =   0
      ScaleHeight     =   9645
      ScaleWidth      =   10365
      TabIndex        =   0
      Top             =   -30
      Width           =   10425
      Begin VB.VScrollBar VScroll1 
         Height          =   6225
         Left            =   9630
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2760
         Width           =   435
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   780
         ScaleHeight     =   1545
         ScaleWidth      =   1425
         TabIndex        =   6
         Top             =   3330
         Width           =   1455
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   780
         ScaleHeight     =   1545
         ScaleWidth      =   1425
         TabIndex        =   5
         Top             =   5010
         Width           =   1455
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1545
         Left            =   780
         ScaleHeight     =   1515
         ScaleWidth      =   1425
         TabIndex        =   4
         Top             =   6720
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   720
         TabIndex        =   1
         Top             =   960
         Width           =   8985
      End
      Begin VB.Shape Shape1 
         Height          =   6225
         Left            =   300
         Top             =   2760
         Width           =   9765
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   2310
         TabIndex        =   13
         Top             =   3420
         Width           =   5955
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   2310
         TabIndex        =   12
         Top             =   4260
         Width           =   3855
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   2310
         TabIndex        =   11
         Top             =   5100
         Width           =   5955
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   2310
         TabIndex        =   10
         Top             =   5910
         Width           =   3945
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   2310
         TabIndex        =   9
         Top             =   6870
         Width           =   5955
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   2340
         TabIndex        =   8
         Top             =   7650
         Width           =   3975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SEARCH MOVIE TO DELETE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   750
         TabIndex        =   3
         Top             =   510
         Width           =   3885
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1890
         TabIndex        =   2
         Top             =   1680
         Width           =   4935
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim image_folder_path As String
    image_folder_path = App.Path & "\db\image_database"
    Picture1.AutoRedraw = True
    Set Pic = LoadPicture(image_folder_path + "\white.jpg")
    Picture1.PaintPicture Pic, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
    Set Picture1.Picture = Picture1.Image
    DeleteSearch
End Sub

Private Sub Picture2_Click()
    DeleteMovieFromDB (Label3.Caption)
End Sub

Private Sub Picture3_Click()
    DeleteMovieFromDB (Label5.Caption)
End Sub

Private Sub Picture4_Click()
    DeleteMovieFromDB (Label7.Caption)
End Sub

Private Sub Text1_Change()
    DeleteSearch
End Sub

Private Sub VScroll1_Change()
    DeleteSearchScrollChange
End Sub
