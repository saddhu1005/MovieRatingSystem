Attribute VB_Name = "SortMovies"
Private current_movies As Variant
Private Index As Integer

Public Sub Sort_Current()
    Dim db_path As String
    Dim db As Database
    Dim rs As Recordset
    Dim total_current_movies As Integer
    Form4.Label11.Caption = "MOVIES : CURRENTLY RUNNING"
    Form4.Label11.FontUnderline = True
    db_path = App.Path + "\db\MovieRatingSystem.mdb"
    Set db = OpenDatabase(db_path)
    Set rs = db.OpenRecordset("select * from MovieDetails where CURRENT=1")
    rs.MoveLast
    total_current_records = rs.RecordCount
    rs.MoveFirst
    current_movies = rs.GetRows(total_current_records)
    
    If total_current_records > 3 Then
        Index = 0
        Form4.Picture6.AutoRedraw = True
        Set Pic = LoadPicture(App.Path + "\db\" + current_movies(1, Index))
        Form4.Picture6.PaintPicture Pic, 0, 0, Form4.Picture6.ScaleWidth, Form4.Picture6.ScaleHeight
        Set Form4.Picture6.Picture = Form4.Picture6.Image
        Form4.Label12.Caption = current_movies(0, Index)
        Form4.Label18.Caption = "DIRECTOR: " + current_movies(5, Index)
        If current_movies(2, Index) Then
            Form4.Label15.Caption = "RATINGS : " + CStr(current_movies(2, Index))
        Else
            Form4.Label15.Caption = "UNRATED YET"
        End If
            
        
        Form4.Picture7.AutoRedraw = True
        Set Pic = LoadPicture(App.Path + "\db\" + current_movies(1, Index + 1))
        Form4.Picture7.PaintPicture Pic, 0, 0, Form4.Picture7.ScaleWidth, Form4.Picture7.ScaleHeight
        Set Form4.Picture7.Picture = Form4.Picture7.Image
        Form4.Label13.Caption = current_movies(0, Index + 1)
        Form4.Label19.Caption = "DIRECTOR: " + current_movies(5, Index + 1)
        If current_movies(2, Index + 1) Then
            Form4.Label16.Caption = "RATINGS : " + CStr(current_movies(2, Index + 1))
        Else
            Form4.Label16.Caption = "UNRATED YET"
        End If
        
        Form4.Picture8.AutoRedraw = True
        Set Pic = LoadPicture(App.Path + "\db\" + current_movies(1, Index + 2))
        Form4.Picture8.PaintPicture Pic, 0, 0, Form4.Picture8.ScaleWidth, Form4.Picture8.ScaleHeight
        Set Form4.Picture8.Picture = Form4.Picture8.Image
        Form4.Label14.Caption = current_movies(0, Index + 2)
        Form4.Label20.Caption = "DIRECTOR: " + current_movies(5, Index + 2)
        If current_movies(2, Index + 2) Then
            Form4.Label17.Caption = "RATINGS : " + CStr(current_movies(2, Index + 2))
        Else
            Form4.Label17.Caption = "UNRATED YET"
        End If
        
        Form4.HScroll1.Value = 0                                                                 'Setting Values For Vertical scroll bar
        Form4.HScroll1.Max = total_current_records - 3
        Form4.HScroll1.Min = 0
    End If
End Sub

Public Sub HScrollChange()
    Dim j As Integer
    Dim Pic As Picture
    j = Form4.HScroll1.Value
    If j > Index Then
        Form4.Picture6.Picture = Form4.Picture7.Picture
        Form4.Label12.Caption = Form4.Label13.Caption
        Form4.Label15.Caption = Form4.Label16.Caption
        Form4.Label18.Caption = Form4.Label19.Caption
        
        Form4.Picture7.Picture = Form4.Picture8.Picture
        Form4.Label13.Caption = Form4.Label14.Caption
        Form4.Label16.Caption = Form4.Label17.Caption
        Form4.Label19.Caption = Form4.Label20.Caption
        
        Form4.Picture8.AutoRedraw = True
        Set Pic = LoadPicture(App.Path + "\db\" + current_movies(1, j + 2))
        Form4.Picture8.PaintPicture Pic, 0, 0, Form4.Picture8.ScaleWidth, Form4.Picture8.ScaleHeight
        Set Form4.Picture8.Picture = Form4.Picture8.Image
        Form4.Label14.Caption = current_movies(0, j + 2)
        If current_movies(2, j + 2) Then
            Form4.Label17.Caption = "RATINGS : " + CStr(current_movies(2, j + 2))
        Else
            Form4.Label17.Caption = "UNRATED YET"
        End If
        Form4.Label20.Caption = "DIRECTOR: " + current_movies(5, j + 2)
        Index = j
    End If
    If j < Index Then
        Form4.Picture8.Picture = Form4.Picture7.Picture
        Form4.Label14.Caption = Form4.Label13.Caption
        Form4.Label17.Caption = Form4.Label16.Caption
        Form4.Label20.Caption = Form4.Label19.Caption
        
        
        Form4.Picture7.Picture = Form4.Picture6.Picture
        Form4.Label13.Caption = Form4.Label12.Caption
        Form4.Label16.Caption = Form4.Label15.Caption
        Form4.Label19.Caption = Form4.Label18.Caption
        
    
        Form4.Picture6.AutoRedraw = True
        Set Pic = LoadPicture(App.Path + "\db\" + current_movies(1, j))
        Form4.Picture6.PaintPicture Pic, 0, 0, Form4.Picture6.ScaleWidth, Form4.Picture6.ScaleHeight
        Set Form4.Picture6.Picture = Form4.Picture6.Image
        Form4.Label12.Caption = current_movies(0, j)
        If current_movies(2, j) Then
            Form4.Label15.Caption = "RATINGS : " + CStr(current_movies(2, j))
        Else
            Form4.Label15.Caption = "UNRATED YET"
        End If
        Form4.Label18.Caption = "DIRECTOR: " + current_movies(5, j)
        Index = j
    End If

End Sub
