VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   6795
   ClientLeft      =   4335
   ClientTop       =   1830
   ClientWidth     =   10695
   FillColor       =   &H00FFFFFF&
   FillStyle       =   4  'Upward Diagonal
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   10695
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "LOGIN"
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
      Left            =   7980
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3750
      Width           =   2325
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   4800
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2940
      Width           =   5505
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4800
      TabIndex        =   2
      Top             =   1710
      Width           =   5505
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "User Login"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   7620
      TabIndex        =   8
      Top             =   5010
      Width           =   1635
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Not a Admin ? "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   5010
      Width           =   1485
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4800
      TabIndex        =   4
      Top             =   2370
      Width           =   2625
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   1140
      Width           =   2625
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER ADMIN CREDENTIALS"
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
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim username, password, db_path, sql As String
    username = Text1.Text
    password = Text2.Text
    If (username = "" Or password = "") Then
        MsgBox ("Please Enter Your Credentials")
    Else
        sql = "SELECT * FROM UserDetails WHERE UserName = '" & username & "'"
        db_path = App.Path + "\db\" + "MovieRatingSystem.mdb"
        Dim cn As ADODB.Connection
        Dim rs As ADODB.Recordset
        Set cn = New ADODB.Connection
            cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
        Set rs = New ADODB.Recordset
            rs.Open sql, cn, adOpenStatic, adLockOptimistic
        If (rs.Fields("PASSWORD").Value = password And rs.Fields("IS ADMIN").Value = 1) Then
            Unload Me
            Form8.Show
        Else
            Text2.Text = ""
            If (rs.Fields("IS ADMIN").Value = 1) Then
                MsgBox ("Invalid Credentials! Please Re-enter Your credentials.")
            Else
                MsgBox ("Sorry! You do not have administrative rights.")
                Unload Me
                Form1.Show
            End If
        End If
        cn.Close
    End If
End Sub

Private Sub Form_Load()
    Dim image_folder_path As String
    image_folder_path = App.Path & "\db\image_database"
    Picture1.AutoRedraw = True
    Set Pic = LoadPicture(image_folder_path + "\mrs.bmp")
    Picture1.PaintPicture Pic, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
    Set Picture1.Picture = Picture1.Image
End Sub

Private Sub Label5_Click()
    Unload Me
    Form1.Show
End Sub
