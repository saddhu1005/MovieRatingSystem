Attribute VB_Name = "RRDisplay"
Public current_user As String
Public Sub RatingsDisplay(movie_name As String)
    Form6.Show
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim db_path, query As String
    
    db_path = App.Path + "\db\MovieRatingSystem.mdb"
    Set cn = New ADODB.Connection
        cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
    query = "select * from MovieRatings where USER = '" & current_user & "'"
    Set rs = New ADODB.Recordset
        rs.Open query, cn, adOpenStatic, adLockOptimistic
    
    If (rs.Fields(movie_name).Value) Then
        Form6.Label7.Caption = rs.Fields(movie_name).Value
        Form6.Combo1.Visible = False
        Form6.Command2.Visible = True
        Form6.Command3.Visible = False
    Else
        Form6.Label7.Caption = "NOT YET RATED"
        Form6.Combo1.Visible = True
        Form6.Command2.Visible = False
        Form6.Command3.Visible = True
    End If
    AddmovieDetails (movie_name)
    ReviewDisplay (movie_name)
End Sub

Private Sub AddmovieDetails(movie_name)
    Dim db_path, query As String
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim no_of_records As Integer
    Dim search_results As Variant
    db_path = App.Path + "\db\MovieRatingSystem.mdb"
    Set cn = New ADODB.Connection
        cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
    query = "select * from MovieDetails where NAME = '" & UCase(movie_name) & "'"
    Set rs = New ADODB.Recordset
        rs.Open query, cn, adOpenStatic, adLockOptimistic
    rs.MoveLast
    no_of_records = rs.RecordCount
    rs.MoveFirst
    search_results = rs.GetRows(no_of_records)
    cn.Close
    
    Form6.Picture2.AutoRedraw = True
    Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 0))
    Form6.Picture2.PaintPicture Pic, 0, 0, Form6.Picture2.ScaleWidth, Form6.Picture2.ScaleHeight
    Set Form6.Picture2.Picture = Form6.Picture2.Image
    
    Form6.Label1.Caption = search_results(0, 0)
    
     If (Not (search_results(2, search_index) = "")) Then
        Form6.Label5.Caption = CStr(search_results(2, 0))
    Else
        Form6.Label5.Caption = "-"
    End If
End Sub
Private Sub ReviewDisplay(movie_name)
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim db_path, query As String
    
    db_path = App.Path + "\db\MovieRatingSystem.mdb"
    Set cn = New ADODB.Connection
        cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
    query = "select * from MovieComments where USER = '" & current_user & "'"
    Set rs = New ADODB.Recordset
        rs.Open query, cn, adOpenStatic, adLockOptimistic
    If (rs.Fields(movie_name).Value <> "") Then
        Form6.Text1.Text = rs.Fields(movie_name).Value
        Form6.Text1.Enabled = False
        Form6.Command1.Visible = False
        Form6.Command4.Visible = True
    Else
        Form6.Text1.Enabled = True
        Form6.Command1.Visible = True
        Form6.Command4.Visible = False
    End If
    CollectComments (movie_name)
End Sub
