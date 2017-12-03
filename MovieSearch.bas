Attribute VB_Name = "MovieSearch"
Private search_index As Integer
Private search_results As Variant

Public Sub SearchMovie()
    Dim search_string, db_path As String
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim no_of_records As Integer
    Dim Pic As Picture
    
    search_string = UCase(Form4.Text1.Text)
    
    If search_string = "" Then
        Form4.Label2.Visible = False
        Form4.Label3.Visible = False
        Form4.Label4.Visible = False
        Form4.Label5.Visible = False
        Form4.Label6.Visible = False
        Form4.Label7.Visible = False
        Form4.Label8.Visible = False
        Form4.Label9.Visible = False
        Form4.Label10.Visible = False
        Form4.Picture1.Visible = False
        Form4.Picture2.Visible = False
        Form4.Picture3.Visible = False
        Form4.Picture4.Visible = False
        Form4.Picture9.Visible = False
        Form4.Picture10.Visible = False
        Form4.Picture11.Visible = False
        Form4.Picture12.Visible = False
        Form4.Picture13.Visible = False
        Form4.Picture14.Visible = False
        Form4.Picture15.Visible = False
        Form4.Picture16.Visible = False
        Form4.VScroll1.Visible = False
        Form4.Line1.Visible = False
        Form4.Line2.Visible = False
        Form4.Line3.Visible = False
        Form4.Line4.Visible = False
    Else
        db_path = App.Path + "\db\MovieRatingSystem.mdb"
        Set cn = New ADODB.Connection
            cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
        query = "select * from MovieDetails where NAME like '" & search_string & "%'"
        Set rs = New ADODB.Recordset
            rs.Open query, cn, adOpenStatic, adLockOptimistic
        If rs.EOF And rs.BOF Then
            Form4.Label2.Visible = True
            Form4.Label2.Caption = "No Such Movies Found"
            Form4.Label3.Visible = False
            Form4.Label4.Visible = False
            Form4.Label5.Visible = False
            Form4.Label6.Visible = False
            Form4.Label7.Visible = False
            Form4.Label8.Visible = False
            Form4.Label9.Visible = False
            Form4.Label10.Visible = False
            Form4.Picture1.Visible = False
            Form4.Picture2.Visible = False
            Form4.Picture3.Visible = False
            Form4.Picture4.Visible = False
            Form4.Picture9.Visible = False
            Form4.Picture10.Visible = False
            Form4.Picture11.Visible = False
            Form4.Picture12.Visible = False
            Form4.Picture13.Visible = False
            Form4.Picture14.Visible = False
            Form4.Picture15.Visible = False
            Form4.Picture16.Visible = False
            Form4.VScroll1.Visible = False
            Form4.Line1.Visible = False
            Form4.Line2.Visible = False
            Form4.Line3.Visible = False
            Form4.Line4.Visible = False
        Else
            rs.MoveLast
            no_of_records = rs.RecordCount
            rs.MoveFirst
            search_results = rs.GetRows(no_of_records)
            cn.Close
            
            If no_of_records = 1 Then
                Form4.Label2.Visible = True
                Form4.Label2.Caption = "Search Results"
                Form4.Label3.Visible = True
                Form4.Label4.Visible = True
                Form4.Label5.Visible = False
                Form4.Label6.Visible = False
                Form4.Label7.Visible = False
                Form4.Label8.Visible = False
                Form4.Label9.Visible = False
                Form4.Label10.Visible = False
                Form4.Picture1.Visible = True
                Form4.Picture2.Visible = False
                Form4.Picture3.Visible = False
                Form4.Picture4.Visible = False
                Form4.Picture9.Visible = True
                Form4.Picture10.Visible = False
                Form4.Picture11.Visible = False
                Form4.Picture12.Visible = False
                Form4.Picture13.Visible = True
                Form4.Picture14.Visible = False
                Form4.Picture15.Visible = False
                Form4.Picture16.Visible = False
                Form4.VScroll1.Visible = False
                Form4.Line1.Visible = True
                Form4.Line2.Visible = False
                Form4.Line3.Visible = False
                Form4.Line4.Visible = False
                Form4.Picture1.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 0))
                Form4.Picture1.PaintPicture Pic, 0, 0, Form4.Picture1.ScaleWidth, Form4.Picture1.ScaleHeight
                Set Form4.Picture1.Picture = Form4.Picture1.Image
                Form4.Label3.Caption = search_results(0, 0)
                Form4.Label4.Caption = search_results(5, 0)
            ElseIf no_of_records = 2 Then
                Form4.Label2.Visible = True
                Form4.Label2.Caption = "Search Results"
                Form4.Label3.Visible = True
                Form4.Label4.Visible = True
                Form4.Label5.Visible = True
                Form4.Label6.Visible = True
                Form4.Label7.Visible = False
                Form4.Label8.Visible = False
                Form4.Label9.Visible = False
                Form4.Label10.Visible = False
                Form4.Picture1.Visible = True
                Form4.Picture2.Visible = True
                Form4.Picture3.Visible = False
                Form4.Picture4.Visible = False
                Form4.Picture9.Visible = True
                Form4.Picture10.Visible = True
                Form4.Picture11.Visible = False
                Form4.Picture12.Visible = False
                Form4.Picture13.Visible = True
                Form4.Picture14.Visible = True
                Form4.Picture15.Visible = False
                Form4.Picture16.Visible = False
                Form4.VScroll1.Visible = False
                Form4.Line1.Visible = True
                Form4.Line2.Visible = True
                Form4.Line3.Visible = False
                Form4.Line4.Visible = False
                Form4.Picture1.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 0))
                Form4.Picture1.PaintPicture Pic, 0, 0, Form4.Picture1.ScaleWidth, Form4.Picture1.ScaleHeight
                Set Form4.Picture1.Picture = Form4.Picture1.Image
                Form4.Label3.Caption = search_results(0, 0)
                Form4.Label4.Caption = search_results(5, 0)
                Form4.Picture2.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 1))
                Form4.Picture2.PaintPicture Pic, 0, 0, Form4.Picture2.ScaleWidth, Form4.Picture2.ScaleHeight
                Set Form4.Picture2.Picture = Form4.Picture2.Image
                Form4.Label5.Caption = search_results(0, 1)
                Form4.Label6.Caption = search_results(5, 1)
            ElseIf no_of_records = 3 Then
                Form4.Label2.Visible = True
                Form4.Label2.Caption = "Search Results"
                Form4.Label3.Visible = True
                Form4.Label4.Visible = True
                Form4.Label5.Visible = True
                Form4.Label6.Visible = True
                Form4.Label7.Visible = True
                Form4.Label8.Visible = True
                Form4.Label9.Visible = False
                Form4.Label10.Visible = False
                Form4.Picture1.Visible = True
                Form4.Picture2.Visible = True
                Form4.Picture3.Visible = True
                Form4.Picture4.Visible = False
                Form4.Picture9.Visible = True
                Form4.Picture10.Visible = True
                Form4.Picture11.Visible = True
                Form4.Picture12.Visible = False
                Form4.Picture13.Visible = True
                Form4.Picture14.Visible = True
                Form4.Picture15.Visible = True
                Form4.Picture16.Visible = False
                Form4.VScroll1.Visible = False
                Form4.Line1.Visible = True
                Form4.Line2.Visible = True
                Form4.Line3.Visible = True
                Form4.Line4.Visible = False
                Form4.Picture1.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 0))
                Form4.Picture1.PaintPicture Pic, 0, 0, Form4.Picture1.ScaleWidth, Form4.Picture1.ScaleHeight
                Set Form4.Picture1.Picture = Form4.Picture1.Image
                Form4.Label3.Caption = search_results(0, 0)
                Form4.Label4.Caption = search_results(5, 0)
                Form4.Picture2.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 1))
                Form4.Picture2.PaintPicture Pic, 0, 0, Form4.Picture2.ScaleWidth, Form4.Picture2.ScaleHeight
                Set Form4.Picture2.Picture = Form4.Picture2.Image
                Form4.Label5.Caption = search_results(0, 1)
                Form4.Label6.Caption = search_results(5, 1)
                Form4.Picture3.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 2))
                Form4.Picture3.PaintPicture Pic, 0, 0, Form4.Picture3.ScaleWidth, Form4.Picture3.ScaleHeight
                Set Form4.Picture3.Picture = Form4.Picture3.Image
                Form4.Label7.Caption = search_results(0, 2)
                Form4.Label8.Caption = search_results(5, 2)
            ElseIf no_of_records = 4 Then
                Form4.Label2.Visible = True
                Form4.Label2.Caption = "Search Results"
                Form4.Label3.Visible = True
                Form4.Label4.Visible = True
                Form4.Label5.Visible = True
                Form4.Label6.Visible = True
                Form4.Label7.Visible = True
                Form4.Label8.Visible = True
                Form4.Label9.Visible = True
                Form4.Label10.Visible = True
                Form4.Picture1.Visible = True
                Form4.Picture2.Visible = True
                Form4.Picture3.Visible = True
                Form4.Picture4.Visible = True
                Form4.Picture9.Visible = True
                Form4.Picture10.Visible = True
                Form4.Picture11.Visible = True
                Form4.Picture12.Visible = True
                Form4.Picture13.Visible = True
                Form4.Picture14.Visible = True
                Form4.Picture15.Visible = True
                Form4.Picture16.Visible = True
                Form4.VScroll1.Visible = False
                Form4.Line1.Visible = True
                Form4.Line2.Visible = True
                Form4.Line3.Visible = True
                Form4.Line4.Visible = True
                Form4.Picture1.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 0))
                Form4.Picture1.PaintPicture Pic, 0, 0, Form4.Picture1.ScaleWidth, Form4.Picture1.ScaleHeight
                Set Form4.Picture1.Picture = Form4.Picture1.Image
                Form4.Label3.Caption = search_results(0, 0)
                Form4.Label4.Caption = search_results(5, 0)
                Form4.Picture2.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 1))
                Form4.Picture2.PaintPicture Pic, 0, 0, Form4.Picture2.ScaleWidth, Form4.Picture2.ScaleHeight
                Set Form4.Picture2.Picture = Form4.Picture2.Image
                Form4.Label5.Caption = search_results(0, 1)
                Form4.Label6.Caption = search_results(5, 1)
                Form4.Picture3.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 2))
                Form4.Picture3.PaintPicture Pic, 0, 0, Form4.Picture3.ScaleWidth, Form4.Picture3.ScaleHeight
                Set Form4.Picture3.Picture = Form4.Picture3.Image
                Form4.Label7.Caption = search_results(0, 2)
                Form4.Label8.Caption = search_results(5, 2)
                Form4.Picture4.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 3))
                Form4.Picture4.PaintPicture Pic, 0, 0, Form4.Picture4.ScaleWidth, Form4.Picture4.ScaleHeight
                Set Form4.Picture4.Picture = Form4.Picture4.Image
                Form4.Label9.Caption = search_results(0, 3)
                Form4.Label10.Caption = search_results(5, 3)
            Else
                Form4.Label2.Visible = True
                Form4.Label2.Caption = "Search Results"
                Form4.Label3.Visible = True
                Form4.Label4.Visible = True
                Form4.Label5.Visible = True
                Form4.Label6.Visible = True
                Form4.Label7.Visible = True
                Form4.Label8.Visible = True
                Form4.Label9.Visible = True
                Form4.Label10.Visible = True
                Form4.Picture1.Visible = True
                Form4.Picture2.Visible = True
                Form4.Picture3.Visible = True
                Form4.Picture4.Visible = True
                Form4.Picture9.Visible = True
                Form4.Picture10.Visible = True
                Form4.Picture11.Visible = True
                Form4.Picture12.Visible = True
                Form4.Picture13.Visible = True
                Form4.Picture14.Visible = True
                Form4.Picture15.Visible = True
                Form4.Picture16.Visible = True
                Form4.VScroll1.Visible = True
                Form4.Line1.Visible = True
                Form4.Line2.Visible = True
                Form4.Line3.Visible = True
                Form4.Line4.Visible = True
                search_index = 0
                Form4.Picture1.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, search_index))
                Form4.Picture1.PaintPicture Pic, 0, 0, Form4.Picture1.ScaleWidth, Form4.Picture1.ScaleHeight
                Set Form4.Picture1.Picture = Form4.Picture1.Image
                Form4.Label3.Caption = search_results(0, search_index)
                Form4.Label4.Caption = search_results(5, search_index)
                Form4.Picture2.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, search_index + 1))
                Form4.Picture2.PaintPicture Pic, 0, 0, Form4.Picture2.ScaleWidth, Form4.Picture2.ScaleHeight
                Set Form4.Picture2.Picture = Form4.Picture2.Image
                Form4.Label5.Caption = search_results(0, search_index + 1)
                Form4.Label6.Caption = search_results(5, search_index + 1)
                Form4.Picture3.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, search_index + 2))
                Form4.Picture3.PaintPicture Pic, 0, 0, Form4.Picture3.ScaleWidth, Form4.Picture3.ScaleHeight
                Set Form4.Picture3.Picture = Form4.Picture3.Image
                Form4.Label7.Caption = search_results(0, search_index + 2)
                Form4.Label8.Caption = search_results(5, search_index + 2)
                Form4.Picture4.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, search_index + 3))
                Form4.Picture4.PaintPicture Pic, 0, 0, Form4.Picture4.ScaleWidth, Form4.Picture4.ScaleHeight
                Set Form4.Picture4.Picture = Form4.Picture4.Image
                Form4.Label9.Caption = search_results(0, search_index + 3)
                Form4.Label10.Caption = search_results(5, search_index + 3)
                
                Form4.VScroll1.Value = 0
                Form4.VScroll1.Max = no_of_records - 4
                Form4.VScroll1.Min = 0
            End If
        End If
    End If
End Sub

Public Sub SearchScrollChange()
    Dim j As Integer
    Dim Pic As Picture
    j = Form4.VScroll1.Value
    If j > search_index Then
        Form4.Picture1.Picture = Form4.Picture2.Picture
        Form4.Label3.Caption = Form4.Label5.Caption
        Form4.Label4.Caption = Form4.Label6.Caption
        
        Form4.Picture2.Picture = Form4.Picture3.Picture
        Form4.Label5.Caption = Form4.Label7.Caption
        Form4.Label6.Caption = Form4.Label8.Caption
        
        Form4.Picture3.Picture = Form4.Picture4.Picture
        Form4.Label7.Caption = Form4.Label9.Caption
        Form4.Label8.Caption = Form4.Label10.Caption
        
        Form4.Picture4.AutoRedraw = True
        Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, j + 3))
        Form4.Picture4.PaintPicture Pic, 0, 0, Form4.Picture4.ScaleWidth, Form4.Picture4.ScaleHeight
        Set Form4.Picture4.Picture = Form4.Picture4.Image
        Form4.Label9.Caption = search_results(0, j + 3)
        Form4.Label10.Caption = search_results(5, j + 3)
        search_index = j
    End If
    If j < search_index Then
        Form4.Picture4.Picture = Form4.Picture3.Picture
        Form4.Label9.Caption = Form4.Label7.Caption
        Form4.Label10.Caption = Form4.Label8.Caption
        
        Form4.Picture3.Picture = Form4.Picture2.Picture
        Form4.Label7.Caption = Form4.Label5.Caption
        Form4.Label8.Caption = Form4.Label6.Caption
        
        Form4.Picture2.Picture = Form4.Picture1.Picture
        Form4.Label5.Caption = Form4.Label3.Caption
        Form4.Label6.Caption = Form4.Label4.Caption
    
        Form4.Picture1.AutoRedraw = True
        Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, j))
        Form4.Picture1.PaintPicture Pic, 0, 0, Form4.Picture1.ScaleWidth, Form4.Picture1.ScaleHeight
        Set Form4.Picture1.Picture = Form4.Picture1.Image
        Form4.Label3.Caption = search_results(0, j)
        Form4.Label4.Caption = search_results(5, j)
        search_index = j
    End If
End Sub
