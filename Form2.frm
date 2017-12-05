VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register"
   ClientHeight    =   6795
   ClientLeft      =   4575
   ClientTop       =   1830
   ClientWidth     =   10695
   FillColor       =   &H00FFFFFF&
   FillStyle       =   4  'Upward Diagonal
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   10695
   Begin VB.TextBox Text5 
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
      Left            =   7110
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   3990
      Width           =   3285
   End
   Begin VB.TextBox Text4 
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
      Left            =   7110
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   3270
      Width           =   3285
   End
   Begin VB.TextBox Text3 
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
      Left            =   7110
      TabIndex        =   10
      Top             =   2520
      Width           =   3285
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "REGISTER"
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
      Left            =   6360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
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
      Left            =   7110
      TabIndex        =   5
      Top             =   1800
      Width           =   3285
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
      Left            =   7110
      TabIndex        =   2
      Top             =   1110
      Width           =   3285
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
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Login Here"
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
      Left            =   7590
      TabIndex        =   14
      Top             =   5880
      Width           =   1485
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Already a User ? "
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
      Left            =   5790
      TabIndex        =   13
      Top             =   5910
      Width           =   1875
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "LAST NAME :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4770
      TabIndex        =   9
      Top             =   2610
      Width           =   1635
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "FIRST NAME :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4770
      TabIndex        =   8
      Top             =   1890
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CONFIRM PASSWORD :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4770
      TabIndex        =   7
      Top             =   4110
      Width           =   2085
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4770
      TabIndex        =   4
      Top             =   3360
      Width           =   1515
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4770
      TabIndex        =   3
      Top             =   1200
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTER"
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
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim first_name, last_name, username, password1, password2, final_password As String
    
    'assigning values
    first_name = Text2.Text
    last_name = Text3.Text
    username = Text1.Text
    password1 = Text4.Text
    password2 = Text5.Text
    
    If (first_name = "" Or last_name = "" Or username = "" Or password1 = "" Or password2 = "") Then
        MsgBox ("All Fields Are Mandotory.Please Fill the Required Details")
        Exit Sub
    Else

        'password checking
         If (password1 = password2) Then
             final_password = password1
         Else
             Text4.Text = ""
             Text5.Text = ""
             MsgBox ("Passwords Do Not Match.Please Renter the Passwords !")
             Exit Sub
         End If
         
         Dim cn As ADODB.Connection
         Dim rs As ADODB.Recordset
         Dim db_path As String
        
         db_path = App.Path + "\db\" + "MovieRatingSystem.mdb"
         Set cn = New ADODB.Connection
             cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
         Set rs = New ADODB.Recordset
             rs.Open "Select * from UserDetails", cn, adOpenStatic, adLockOptimistic
             
         With rs
             .AddNew
             .Fields("FIRST NAME").Value = first_name
             .Fields("LAST NAME").Value = last_name
             .Fields("USERNAME").Value = username
             .Fields("PASSWORD").Value = final_password
             .Fields("IS ADMIN").Value = 0
             .Update
         End With
         
         Register_to_MovieRatings (username)
         Register_to_MovieComments (username)
         
         MsgBox ("User Successfully Registered !")
         Unload Form2
         Form1.Show
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

Private Sub Label8_Click()
    Unload Me
    Form1.Show
End Sub

Private Sub Register_to_MovieRatings(username As String)
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim db_path As String
    db_path = App.Path + "\db\" + "MovieRatingSystem.mdb"
    Set cn = New ADODB.Connection
        cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
    Set rs = New ADODB.Recordset
        rs.Open "Select * from MovieRatings", cn, adOpenStatic, adLockOptimistic
    With rs
        .AddNew
        .Fields("USER").Value = username
        .Update
    End With
    cn.Close
End Sub

Private Sub Register_to_MovieComments(username As String)
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim db_path As String
    db_path = App.Path + "\db\" + "MovieRatingSystem.mdb"
    Set cn = New ADODB.Connection
        cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
    Set rs = New ADODB.Recordset
        rs.Open "Select * from MovieComments", cn, adOpenStatic, adLockOptimistic
    With rs
        .AddNew
        .Fields("USER").Value = username
        .Update
    End With
    cn.Close
End Sub
