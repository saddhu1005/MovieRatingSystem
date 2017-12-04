Attribute VB_Name = "DeleteMovie"
Private delete_search_index As Integer
Private search_results As Variant

Public Sub DeleteMovieFromDB(movie_name As String)
    Dim confrim As String
    confirm = MsgBox("Confirm Deletion!", vbYesNo, "Delete Movie")
    If confirm = vbYes Then
        Dim db As Database
        Dim query As String
        Set db = OpenDatabase(App.Path + "/db/MovieRatingSystem.mdb")
        DeleteFromMovieRatings (movie_name)
        DeleteFromMovieComments (movie_name)
        query = "DELETE FROM MovieDetails WHERE NAME='" & movie_name & "';"
        db.Execute (query)
        db.Close
        MsgBox ("Movie Deleted Successfully")
        DeleteSearch
    End If
End Sub
Private Sub DeleteFromMovieRatings(movie_name As String)
    Dim cn As ADODB.Connection
    Dim query, db_path As String
    db_path = App.Path + "\db\MovieRatingSystem.mdb"
    Set cn = New ADODB.Connection
        cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
    query = "ALTER TABLE MovieRatings  DROP [" & movie_name & "];"
    cn.Execute (query)
    cn.Close
End Sub
Private Sub DeleteFromMovieComments(movie_name As String)
    Dim db As Database
    Dim query As String
    Set db = OpenDatabase(App.Path + "/db/MovieRatingSystem.mdb")
    query = "ALTER TABLE MovieComments  DROP [" & movie_name & "];"
    db.Execute (query)
    db.Close
End Sub

Public Sub DeleteSearch()
    Dim search_string As String
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim no_of_records As Integer
    Dim Pic As Picture
    Dim db_path As String
    
    search_string = UCase(Form9.Text1.Text)
    If search_string = "" Then
        Form9.Label2.Visible = False
        Form9.Label3.Visible = False
        Form9.Label4.Visible = False
        Form9.Label5.Visible = False
        Form9.Label6.Visible = False
        Form9.Label7.Visible = False
        Form9.Label8.Visible = False
        Form9.Shape1.Visible = False
        Form9.Picture2.Visible = False
        Form9.Picture3.Visible = False
        Form9.Picture4.Visible = False
        Form9.VScroll1.Visible = False
    Else
        db_path = App.Path + "\db\MovieRatingSystem.mdb"
        Set cn = New ADODB.Connection
            cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
        query = "select * from MovieDetails where NAME like '" & search_string & "%'"
        Set rs = New ADODB.Recordset
            rs.Open query, cn, adOpenStatic, adLockOptimistic
        
        If rs.EOF And rs.BOF Then
            Form9.Label2.Visible = True
            Form9.Label3.Visible = False
            Form9.Label4.Visible = False
            Form9.Label5.Visible = False
            Form9.Label6.Visible = False
            Form9.Label7.Visible = False
            Form9.Label8.Visible = False
            Form9.Shape1.Visible = False
            Form9.Picture2.Visible = False
            Form9.Picture3.Visible = False
            Form9.Picture4.Visible = False
            Form9.VScroll1.Visible = False
            Form9.Label2.Caption = "No Movies Found"
        Else
            rs.MoveLast
            no_of_records = rs.RecordCount
            rs.MoveFirst
            search_results = rs.GetRows(no_of_records)
            cn.Close
            If no_of_records = 1 Then
                Form9.Label2.Visible = True
                Form9.Label3.Visible = True
                Form9.Label4.Visible = True
                Form9.Label5.Visible = False
                Form9.Label6.Visible = False
                Form9.Label7.Visible = False
                Form9.Label8.Visible = False
                Form9.Shape1.Visible = True
                Form9.Picture2.Visible = True
                Form9.Picture3.Visible = False
                Form9.Picture4.Visible = False
                Form9.VScroll1.Visible = False
                Form9.Label2.Caption = "Search Results"
                Form9.Picture2.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 0))
                Form9.Picture2.PaintPicture Pic, 0, 0, Form9.Picture2.ScaleWidth, Form9.Picture2.ScaleHeight
                Set Form9.Picture2.Picture = Form9.Picture2.Image
                Form9.Label3.Caption = search_results(0, 0)
                Form9.Label4.Caption = search_results(5, 0)
            ElseIf no_of_records = 2 Then
                Form9.Label2.Visible = True
                Form9.Label3.Visible = True
                Form9.Label4.Visible = True
                Form9.Label5.Visible = True
                Form9.Label6.Visible = True
                Form9.Label7.Visible = False
                Form9.Label8.Visible = False
                Form9.Shape1.Visible = True
                Form9.Picture2.Visible = True
                Form9.Picture3.Visible = True
                Form9.Picture4.Visible = False
                Form9.VScroll1.Visible = False
                Form9.Label2.Caption = "Search Results"
                Form9.Picture2.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 0))
                Form9.Picture2.PaintPicture Pic, 0, 0, Form9.Picture2.ScaleWidth, Form9.Picture2.ScaleHeight
                Set Form9.Picture2.Picture = Form9.Picture2.Image
                Form9.Label3.Caption = search_results(0, 0)
                Form9.Label4.Caption = search_results(5, 0)
                Form9.Picture3.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 1))
                Form9.Picture3.PaintPicture Pic, 0, 0, Form9.Picture3.ScaleWidth, Form9.Picture3.ScaleHeight
                Set Form9.Picture3.Picture = Form9.Picture3.Image
                Form9.Label5.Caption = search_results(0, 1)
                Form9.Label6.Caption = search_results(5, 1)
            ElseIf no_of_records = 3 Then
                Form9.Label2.Visible = True
                Form9.Label3.Visible = True
                Form9.Label4.Visible = True
                Form9.Label5.Visible = True
                Form9.Label6.Visible = True
                Form9.Label7.Visible = True
                Form9.Label8.Visible = True
                Form9.Shape1.Visible = True
                Form9.Picture2.Visible = True
                Form9.Picture3.Visible = True
                Form9.Picture4.Visible = True
                Form9.VScroll1.Visible = False
                Form9.Label2.Caption = "Search Results"
                Form9.Picture2.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 0))
                Form9.Picture2.PaintPicture Pic, 0, 0, Form9.Picture2.ScaleWidth, Form9.Picture2.ScaleHeight
                Set Form9.Picture2.Picture = Form9.Picture2.Image
                Form9.Label3.Caption = search_results(0, 0)
                Form9.Label4.Caption = search_results(5, 0)
                Form9.Picture3.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 1))
                Form9.Picture3.PaintPicture Pic, 0, 0, Form9.Picture3.ScaleWidth, Form9.Picture3.ScaleHeight
                Set Form9.Picture3.Picture = Form9.Picture3.Image
                Form9.Label5.Caption = search_results(0, 1)
                Form9.Label6.Caption = search_results(5, 1)
                Form9.Picture4.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 2))
                Form9.Picture4.PaintPicture Pic, 0, 0, Form9.Picture4.ScaleWidth, Form9.Picture4.ScaleHeight
                Set Form9.Picture4.Picture = Form9.Picture4.Image
                Form9.Label7.Caption = search_results(0, 2)
                Form9.Label8.Caption = search_results(5, 2)
            Else
                Form9.Label2.Visible = True
                Form9.Label3.Visible = True
                Form9.Label4.Visible = True
                Form9.Label5.Visible = True
                Form9.Label6.Visible = True
                Form9.Label7.Visible = True
                Form9.Label8.Visible = True
                Form9.Shape1.Visible = True
                Form9.Picture2.Visible = True
                Form9.Picture3.Visible = True
                Form9.Picture4.Visible = True
                Form9.VScroll1.Visible = True
                Form9.Label2.Caption = "Search Results"
                delete_search_index = 0
                Form9.Picture2.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, delete_search_index))
                Form9.Picture2.PaintPicture Pic, 0, 0, Form9.Picture2.ScaleWidth, Form9.Picture2.ScaleHeight
                Set Form9.Picture2.Picture = Form9.Picture2.Image
                Form9.Label3.Caption = search_results(0, delete_search_index)
                Form9.Label4.Caption = search_results(5, delete_search_index)
                Form9.Picture3.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, delete_search_index + 1))
                Form9.Picture3.PaintPicture Pic, 0, 0, Form9.Picture3.ScaleWidth, Form9.Picture3.ScaleHeight
                Set Form9.Picture3.Picture = Form9.Picture3.Image
                Form9.Label5.Caption = search_results(0, delete_search_index + 1)
                Form9.Label6.Caption = search_results(5, delete_search_index + 1)
                Form9.Picture4.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, delete_search_index + 2))
                Form9.Picture4.PaintPicture Pic, 0, 0, Form9.Picture4.ScaleWidth, Form9.Picture4.ScaleHeight
                Set Form9.Picture4.Picture = Form9.Picture4.Image
                Form9.Label7.Caption = search_results(0, delete_search_index + 2)
                Form9.Label8.Caption = search_results(5, delete_search_index + 2)
                
                Form9.VScroll1.Value = 0
                Form9.VScroll1.Max = no_of_records - 3
                Form9.VScroll1.Min = 0
            End If
        End If
    End If
End Sub
Public Sub DeleteSearchScrollChange()
    Dim j As Integer
    Dim Pic As Picture
    j = Form9.VScroll1.Value
    If j > delete_search_index Then
        Form9.Picture2.Picture = Form9.Picture3.Picture
        Form9.Label3.Caption = Form9.Label5.Caption
        Form9.Label4.Caption = Form9.Label6.Caption
        
        Form9.Picture3.Picture = Form9.Picture4.Picture
        Form9.Label5.Caption = Form9.Label7.Caption
        Form9.Label6.Caption = Form9.Label8.Caption
        
        Form9.Picture4.AutoRedraw = True
        Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, j + 2))
        Form9.Picture4.PaintPicture Pic, 0, 0, Form9.Picture4.ScaleWidth, Form9.Picture4.ScaleHeight
        Set Form9.Picture4.Picture = Form9.Picture4.Image
        Form9.Label7.Caption = search_results(0, j + 2)
        Form9.Label8.Caption = search_results(5, j + 2)
        delete_search_index = j
    End If
    If j < delete_search_index Then
        Form9.Picture4.Picture = Form9.Picture3.Picture
        Form9.Label7.Caption = Form9.Label5.Caption
        Form9.Label8.Caption = Form9.Label6.Caption
        
        Form9.Picture3.Picture = Form9.Picture2.Picture
        Form9.Label5.Caption = Form9.Label3.Caption
        Form9.Label6.Caption = Form9.Label4.Caption
    
        Form9.Picture2.AutoRedraw = True
        Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, j))
        Form9.Picture2.PaintPicture Pic, 0, 0, Form9.Picture2.ScaleWidth, Form9.Picture2.ScaleHeight
        Set Form9.Picture2.Picture = Form9.Picture2.Image
        Form9.Label3.Caption = search_results(0, j)
        Form9.Label4.Caption = search_results(5, j)
        delete_search_index = j
    End If
End Sub
