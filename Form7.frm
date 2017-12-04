VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form7 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Movie"
   ClientHeight    =   10305
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12765
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10305
   ScaleWidth      =   12765
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   10305
      Left            =   0
      ScaleHeight     =   10245
      ScaleWidth      =   12705
      TabIndex        =   0
      Top             =   0
      Width           =   12765
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "YES"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4020
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   12
         Top             =   6570
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   585
         Left            =   3360
         TabIndex        =   11
         Top             =   690
         Width           =   4875
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   585
         Left            =   3360
         TabIndex        =   10
         Top             =   1620
         Width           =   3105
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   585
         Left            =   3360
         TabIndex        =   9
         Top             =   3240
         Width           =   4845
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   585
         Left            =   3360
         TabIndex        =   8
         Top             =   4050
         Width           =   4845
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   585
         Left            =   3360
         TabIndex        =   7
         Top             =   4890
         Width           =   4845
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   585
         Left            =   3360
         TabIndex        =   6
         Top             =   5700
         Width           =   4845
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   2415
         Left            =   420
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   7620
         Width           =   11685
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ADD  MOVIE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8670
         TabIndex        =   4
         Top             =   720
         Width           =   3075
      End
      Begin VB.CommandButton Command3 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8670
         TabIndex        =   3
         Top             =   1950
         Width           =   3075
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   6480
         TabIndex        =   2
         Top             =   1620
         Width           =   1755
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5970
         TabIndex        =   1
         Top             =   6570
         Width           =   1305
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   555
         Left            =   3360
         TabIndex        =   13
         Top             =   2490
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   979
         _Version        =   393216
         Format          =   106561537
         CurrentDate     =   43042
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   7710
         Top             =   1650
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NAME OF MOVIE : "
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
         Height          =   555
         Left            =   390
         TabIndex        =   22
         Top             =   810
         Width           =   3255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "RELEASE DATE : "
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
         Height          =   555
         Left            =   30
         TabIndex        =   21
         Top             =   2610
         Width           =   3915
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LANGUAGE : "
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
         Height          =   525
         Left            =   300
         TabIndex        =   20
         Top             =   3390
         Width           =   3285
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "GENRE : "
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
         Height          =   555
         Left            =   270
         TabIndex        =   19
         Top             =   4200
         Width           =   3135
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DIRECTOR : "
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
         Height          =   555
         Left            =   60
         TabIndex        =   18
         Top             =   5040
         Width           =   3705
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CAST : "
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
         Height          =   555
         Left            =   300
         TabIndex        =   17
         Top             =   5820
         Width           =   3105
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SYNOPSIS : "
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
         Height          =   585
         Left            =   330
         TabIndex        =   16
         Top             =   7230
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CURRENTLY IN THEATRES : "
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
         Height          =   555
         Left            =   330
         TabIndex        =   15
         Top             =   6600
         Width           =   3495
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "IMAGE : "
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
         Height          =   555
         Left            =   330
         TabIndex        =   14
         Top             =   1770
         Width           =   3255
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    CommonDialog1.Filter = "Apps (*.txt|*.txt|All files (*.*)|*.*"
    CommonDialog1.DefaultExt = "jpeg"
    CommonDialog1.DialogTitle = "Select File"
    CommonDialog1.ShowOpen
    Text2.Text = CommonDialog1.filename
End Sub

Private Sub Command2_Click()
    AddRecord
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim image_folder_path As String
    image_folder_path = App.Path & "\db\image_database"
    Picture1.AutoRedraw = True
    Set Pic = LoadPicture(image_folder_path + "\white.jpg")
    Picture1.PaintPicture Pic, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
    Set Picture1.Picture = Picture1.Image
End Sub
