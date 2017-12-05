VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin Panel"
   ClientHeight    =   6795
   ClientLeft      =   4575
   ClientTop       =   2070
   ClientWidth     =   10695
   FillColor       =   &H00FFFFFF&
   FillStyle       =   4  'Upward Diagonal
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   10695
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5760
      Width           =   4995
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "CHANGE MOVIE STATUS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4950
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4710
      Width           =   4995
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "DELETE A MOVIE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4950
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3630
      Width           =   4995
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "ADD A MOVIE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4950
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2610
      Width           =   4995
   End
   Begin VB.PictureBox Picture1 
      Height          =   6795
      Left            =   0
      ScaleHeight     =   6735
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CHOOSE ONE OF THE FOLLOWING OPTIONS TO BE PERFORMED :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   4800
      TabIndex        =   2
      Top             =   1530
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME ADMINISTRATOR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4800
      TabIndex        =   1
      Top             =   240
      Width           =   5505
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form7.Show
End Sub

Private Sub Command2_Click()
    Form9.Show
End Sub

Private Sub Command3_Click()
    Form10.Show
End Sub

Private Sub Command4_Click()
Unload Form1
Unload Form2
Unload Form3
Unload Form4
Unload Form5
Unload Form6
Unload Me
Unload frmSplash
Unload Form7
Unload Form8
Unload Form9
Unload Form10
End Sub

Private Sub Form_Load()
    Dim image_folder_path As String
    image_folder_path = App.Path & "\db\image_database"
    Picture1.AutoRedraw = True
    Set Pic = LoadPicture(image_folder_path + "\mrs.bmp")
    Picture1.PaintPicture Pic, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
    Set Picture1.Picture = Picture1.Image
End Sub

