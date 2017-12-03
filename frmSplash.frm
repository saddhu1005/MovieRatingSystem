VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8220
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   8235
      Left            =   0
      ScaleHeight     =   8175
      ScaleWidth      =   6375
      TabIndex        =   0
      Top             =   0
      Width           =   6435
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   1650
         Top             =   7290
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   435
         Left            =   270
         TabIndex        =   1
         Top             =   6810
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   767
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2430
         TabIndex        =   2
         Top             =   7410
         Width           =   3555
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim image_folder_path As String
    image_folder_path = App.Path & "\db\image_database"
    Picture1.AutoRedraw = True
    Set Pic = LoadPicture(image_folder_path + "\mrs.bmp")
    Picture1.PaintPicture Pic, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
    Set Picture1.Picture = Picture1.Image
    CalculateOverallRatings
End Sub



Private Sub Timer1_Timer()
    Timer1.Interval = Rnd * 300 + 10
    ProgressBar1.Value = ProgressBar1.Value + 2
    Label1.Caption = "Making System Environment Ready...."
    If ProgressBar1.Value >= 30 And ProgressBar1.Value <= 50 Then
        Label1.Caption = "Getting Environment Variables Ready...."
    ElseIf ProgressBar1.Value >= 50 And ProgressBar1.Value <= 80 Then
        Label1.Caption = "Loading Database...."
    ElseIf ProgressBar1.Value >= 80 Then
        Label1.Caption = "Starting Movie Rating System..."
    End If
    If ProgressBar1.Value = 100 Then
        Unload Me
        Form1.Show
        Timer1.Enabled = False
    End If
End Sub
