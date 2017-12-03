VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movie Rating System"
   ClientHeight    =   10215
   ClientLeft      =   3090
   ClientTop       =   1635
   ClientWidth     =   17580
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10215
   ScaleWidth      =   17580
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   15270
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   40
      Top             =   6900
      Width           =   495
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   15270
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   39
      Top             =   4770
      Width           =   495
   End
   Begin VB.PictureBox Picture5 
      Height          =   10215
      Left            =   0
      ScaleHeight     =   10155
      ScaleWidth      =   13035
      TabIndex        =   17
      Top             =   0
      Width           =   13095
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ratings and Reviews"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   8280
         Width           =   3375
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Movie Information"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   7560
         Width           =   3375
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ratings and Reviews"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   8280
         Width           =   3375
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Movie Information"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   7560
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ratings and Reviews"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   540
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   8280
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Movie Information"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   540
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   7560
         Width           =   3375
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   9270
         ScaleHeight     =   3825
         ScaleWidth      =   3075
         TabIndex        =   21
         Top             =   1620
         Width           =   3105
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   4950
         ScaleHeight     =   3825
         ScaleWidth      =   3075
         TabIndex        =   20
         Top             =   1590
         Width           =   3105
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   690
         ScaleHeight     =   3825
         ScaleWidth      =   3075
         TabIndex        =   19
         Top             =   1590
         Width           =   3105
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   465
         Left            =   0
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   9690
         Width           =   13065
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label20"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   9300
         TabIndex        =   37
         Top             =   6900
         Width           =   3105
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label19"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   4950
         TabIndex        =   36
         Top             =   6900
         Width           =   3105
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label18"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   630
         TabIndex        =   35
         Top             =   6900
         Width           =   3105
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label17"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   585
         Left            =   9300
         TabIndex        =   28
         Top             =   6330
         Width           =   3105
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label16"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   585
         Left            =   4950
         TabIndex        =   27
         Top             =   6330
         Width           =   3105
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label15"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   630
         TabIndex        =   26
         Top             =   6330
         Width           =   3105
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label14"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   795
         Left            =   9300
         TabIndex        =   25
         Top             =   5640
         Width           =   3105
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   795
         Left            =   4950
         TabIndex        =   24
         Top             =   5640
         Width           =   3105
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label12"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   795
         Left            =   630
         TabIndex        =   23
         Top             =   5640
         Width           =   3105
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   765
         Left            =   1830
         TabIndex        =   22
         Top             =   330
         Width           =   9765
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         Height          =   7875
         Left            =   8760
         Top             =   1200
         Width           =   4065
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   7875
         Left            =   4470
         Top             =   1200
         Width           =   4065
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   7875
         Left            =   210
         Top             =   1200
         Width           =   4065
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   1455
      Left            =   13410
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   8
      Top             =   8160
      Width           =   1455
   End
   Begin VB.PictureBox Picture3 
      Height          =   1455
      Left            =   13410
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   7
      Top             =   6090
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Height          =   1455
      Left            =   13410
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   6
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   10335
      Left            =   13110
      TabIndex        =   0
      Top             =   -90
      Width           =   4485
      Begin VB.PictureBox Picture16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3150
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   45
         Top             =   9060
         Width           =   495
      End
      Begin VB.PictureBox Picture15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3150
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   44
         Top             =   6960
         Width           =   495
      End
      Begin VB.PictureBox Picture14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3150
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   43
         Top             =   4860
         Width           =   495
      End
      Begin VB.PictureBox Picture13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3150
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   42
         Top             =   2820
         Width           =   495
      End
      Begin VB.PictureBox Picture12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   2160
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   41
         Top             =   9060
         Width           =   495
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   2160
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   38
         Top             =   2820
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Height          =   1455
         Left            =   330
         ScaleHeight     =   1395
         ScaleWidth      =   1395
         TabIndex        =   5
         Top             =   1950
         Width           =   1455
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   10215
         Left            =   4200
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   90
         Width           =   285
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   270
         TabIndex        =   2
         Top             =   720
         Width           =   3705
      End
      Begin VB.Line Line4 
         X1              =   1830
         X2              =   4110
         Y1              =   4770
         Y2              =   4770
      End
      Begin VB.Line Line3 
         X1              =   1830
         X2              =   4110
         Y1              =   6900
         Y2              =   6900
      End
      Begin VB.Line Line2 
         X1              =   1830
         X2              =   4110
         Y1              =   8970
         Y2              =   8970
      End
      Begin VB.Line Line1 
         X1              =   1830
         X2              =   4110
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label10"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   1830
         TabIndex        =   16
         Top             =   8730
         Width           =   2295
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   1830
         TabIndex        =   15
         Top             =   8190
         Width           =   2295
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   1830
         TabIndex        =   14
         Top             =   6660
         Width           =   2295
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   1830
         TabIndex        =   13
         Top             =   6120
         Width           =   2295
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   1830
         TabIndex        =   12
         Top             =   4530
         Width           =   2295
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   1830
         TabIndex        =   11
         Top             =   3990
         Width           =   2295
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   1830
         TabIndex        =   10
         Top             =   2490
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   1830
         TabIndex        =   9
         Top             =   1980
         Width           =   2295
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1050
         TabIndex        =   3
         Top             =   1410
         Width           =   2325
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Search Movies.."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   300
         TabIndex        =   1
         Top             =   390
         Width           =   1965
      End
   End
   Begin VB.Menu File 
      Caption         =   "File"
   End
   Begin VB.Menu SortMovies 
      Caption         =   "Sort Movies"
      Begin VB.Menu CurrentlyRunning 
         Caption         =   "Currently Running"
         Checked         =   -1  'True
      End
      Begin VB.Menu Ratings 
         Caption         =   "Ratings"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MovieDetailsDisplay (Label12.Caption)
End Sub

Private Sub Command2_Click()
    RatingsDisplay (Label12.Caption)
End Sub

Private Sub Command3_Click()
    MovieDetailsDisplay (Label13.Caption)
End Sub

Private Sub Command4_Click()
    RatingsDisplay (Label13.Caption)
End Sub

Private Sub Command5_Click()
    MovieDetailsDisplay (Label14.Caption)
End Sub

Private Sub Command6_Click()
    RatingsDisplay (Label14.Caption)
End Sub

Private Sub CurrentlyRunning_Click()
    CurrentlyRunning.Checked = True
    Ratings.Checked = False
    Sort_Current
End Sub

Private Sub Form_Load()
    SearchMovie
    Sort_Current
    Dim image_folder_path As String
    image_folder_path = App.Path & "\db\image_database"
    Picture5.AutoRedraw = True
    Set Pic = LoadPicture(image_folder_path + "\blue.jpg")
    Picture5.PaintPicture Pic, 0, 0, Picture5.ScaleWidth, Picture5.ScaleHeight
    Set Picture5.Picture = Picture5.Image
    
    Picture9.AutoRedraw = True
    Set Pic = LoadPicture(image_folder_path + "\info.bmp")
    Picture9.PaintPicture Pic, 0, 0, Picture9.ScaleWidth, Picture9.ScaleHeight
    Set Picture9.Picture = Picture9.Image
    
    Picture10.AutoRedraw = True
    Set Pic = LoadPicture(image_folder_path + "\info.bmp")
    Picture10.PaintPicture Pic, 0, 0, Picture10.ScaleWidth, Picture10.ScaleHeight
    Set Picture10.Picture = Picture10.Image
    
    Picture11.AutoRedraw = True
    Set Pic = LoadPicture(image_folder_path + "\info.bmp")
    Picture11.PaintPicture Pic, 0, 0, Picture11.ScaleWidth, Picture11.ScaleHeight
    Set Picture9.Picture = Picture9.Image
    
    Picture12.AutoRedraw = True
    Set Pic = LoadPicture(image_folder_path + "\info.bmp")
    Picture12.PaintPicture Pic, 0, 0, Picture12.ScaleWidth, Picture12.ScaleHeight
    Set Picture12.Picture = Picture12.Image
    
    Picture13.AutoRedraw = True
    Set Pic = LoadPicture(image_folder_path + "\star.bmp")
    Picture13.PaintPicture Pic, 0, 0, Picture13.ScaleWidth, Picture13.ScaleHeight
    Set Picture13.Picture = Picture13.Image
    
    Picture14.AutoRedraw = True
    Set Pic = LoadPicture(image_folder_path + "\star.bmp")
    Picture14.PaintPicture Pic, 0, 0, Picture14.ScaleWidth, Picture14.ScaleHeight
    Set Picture14.Picture = Picture14.Image
    
    Picture15.AutoRedraw = True
    Set Pic = LoadPicture(image_folder_path + "\star.bmp")
    Picture15.PaintPicture Pic, 0, 0, Picture15.ScaleWidth, Picture15.ScaleHeight
    Set Picture15.Picture = Picture15.Image
    
    Picture16.AutoRedraw = True
    Set Pic = LoadPicture(image_folder_path + "\star.bmp")
    Picture16.PaintPicture Pic, 0, 0, Picture16.ScaleWidth, Picture16.ScaleHeight
    Set Picture16.Picture = Picture16.Image
End Sub

Private Sub HScroll1_Change()
    HScrollChange
End Sub

Private Sub Picture1_Click()
    MovieDetailsDisplay (Label3.Caption)
End Sub

Private Sub Picture10_Click()
     MovieDetailsDisplay (Label5.Caption)
End Sub

Private Sub Picture11_Click()
     MovieDetailsDisplay (Label7.Caption)
End Sub

Private Sub Picture12_Click()
     MovieDetailsDisplay (Label9.Caption)
End Sub

Private Sub Picture13_Click()
    RatingsDisplay (Label3.Caption)
End Sub

Private Sub Picture14_Click()
    RatingsDisplay (Label5.Caption)
End Sub

Private Sub Picture15_Click()
    RatingsDisplay (Label7.Caption)
End Sub

Private Sub Picture16_Click()
    RatingsDisplay (Label9.Caption)
End Sub

Private Sub Picture2_Click()
    MovieDetailsDisplay (Label5.Caption)
End Sub

Private Sub Picture3_Click()
    MovieDetailsDisplay (Label7.Caption)
End Sub

Private Sub Picture4_Click()
    MovieDetailsDisplay (Label9.Caption)
End Sub

Private Sub Picture6_Click()
     MovieDetailsDisplay (Label12.Caption)
End Sub

Private Sub Picture7_Click()
    MovieDetailsDisplay (Label13.Caption)
End Sub

Private Sub Picture8_Click()
    MovieDetailsDisplay (Label14.Caption)
End Sub

Private Sub Picture9_Click()
     MovieDetailsDisplay (Label3.Caption)
End Sub

Private Sub Ratings_Click()
    Ratings.Checked = True
    CurrentlyRunning.Checked = False
End Sub

Private Sub Text1_Change()
    SearchMovie
End Sub

Private Sub VScroll1_Change()
    SearchScrollChange
End Sub
