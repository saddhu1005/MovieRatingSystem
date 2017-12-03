Attribute VB_Name = "MovieDisplay"
Public Sub MovieDetailsDisplay(movie_name As String)
    Form5.Show
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
     
     
     
    Form5.Picture2.AutoRedraw = True
    Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 0))
    Form5.Picture2.PaintPicture Pic, 0, 0, Form5.Picture2.ScaleWidth, Form5.Picture2.ScaleHeight
    Set Form5.Picture2.Picture = Form5.Picture2.Image
    
    Form5.Label1.Caption = search_results(0, 0)
        
    If (Not (search_results(2, search_index) = "")) Then
        Form5.Label8.Caption = CStr(search_results(2, 0))
    Else
        Form5.Label8.Caption = "Yet to be Rated"
    End If
    
    Form5.Label3.Caption = "LANGUAGE: " + search_results(4, 0)
    
    Form5.Label2.Caption = "RELEASE DATE: " + CStr(search_results(3, 0))
    
    Dim Synopsis As String
    For k = 8 To 14
        If (Not (search_results(k, 0) = "")) Then
            Synopsis = Synopsis + search_results(k, 0)
        End If
    Next
    Form5.Label10.Caption = Synopsis
    
    Form1.Label3.Caption = "LANGUAGE: " + search_results(4, 0)
    
    Form5.Label4.Caption = "DIRECTOR: " + search_results(5, 0)
    
    Form5.Label5.Caption = "GENRE: " + search_results(6, 0)
    
    Form5.Label6.Caption = "CAST: " + search_results(7, 0)
End Sub
