Attribute VB_Name = "MovieStatus"
Private status_search_index As Integer
Private search_results As Variant
Public Sub StatusSearch()
    Dim search_string As String
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim no_of_records As Integer
    Dim Pic As Picture
    Dim db_path As String
    
    search_string = UCase(Form10.Text1.Text)
    If search_string = "" Then
        Form10.Label2.Visible = False
        Form10.Label3.Visible = False
        Form10.Label4.Visible = False
        Form10.Label5.Visible = False
        Form10.Label6.Visible = False
        Form10.Label7.Visible = False
        Form10.Label8.Visible = False
        Form10.Shape1.Visible = False
        Form10.Picture2.Visible = False
        Form10.Picture3.Visible = False
        Form10.Picture4.Visible = False
        Form10.VScroll1.Visible = False
    Else
        db_path = App.Path + "\db\MovieRatingSystem.mdb"
        Set cn = New ADODB.Connection
            cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
        query = "select * from MovieDetails where NAME like '" & search_string & "%'"
        Set rs = New ADODB.Recordset
            rs.Open query, cn, adOpenStatic, adLockOptimistic
        
        If rs.EOF And rs.BOF Then
            Form10.Label2.Visible = True
            Form10.Label3.Visible = False
            Form10.Label4.Visible = False
            Form10.Label5.Visible = False
            Form10.Label6.Visible = False
            Form10.Label7.Visible = False
            Form10.Label8.Visible = False
            Form10.Shape1.Visible = False
            Form10.Picture2.Visible = False
            Form10.Picture3.Visible = False
            Form10.Picture4.Visible = False
            Form10.VScroll1.Visible = False
            Form10.Label2.Caption = "No Movies Found"
        Else
            rs.MoveLast
            no_of_records = rs.RecordCount
            rs.MoveFirst
            search_results = rs.GetRows(no_of_records)
            cn.Close
            If no_of_records = 1 Then
                Form10.Label2.Visible = True
                Form10.Label3.Visible = True
                Form10.Label4.Visible = True
                Form10.Label5.Visible = False
                Form10.Label6.Visible = False
                Form10.Label7.Visible = False
                Form10.Label8.Visible = False
                Form10.Shape1.Visible = True
                Form10.Picture2.Visible = True
                Form10.Picture3.Visible = False
                Form10.Picture4.Visible = False
                Form10.VScroll1.Visible = False
                Form10.Label2.Caption = "Search Results"
                Form10.Picture2.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 0))
                Form10.Picture2.PaintPicture Pic, 0, 0, Form10.Picture2.ScaleWidth, Form10.Picture2.ScaleHeight
                Set Form10.Picture2.Picture = Form10.Picture2.Image
                Form10.Label3.Caption = search_results(0, 0)
                Form10.Label4.Caption = search_results(5, 0)
            ElseIf no_of_records = 2 Then
                Form10.Label2.Visible = True
                Form10.Label3.Visible = True
                Form10.Label4.Visible = True
                Form10.Label5.Visible = True
                Form10.Label6.Visible = True
                Form10.Label7.Visible = False
                Form10.Label8.Visible = False
                Form10.Shape1.Visible = True
                Form10.Picture2.Visible = True
                Form10.Picture3.Visible = True
                Form10.Picture4.Visible = False
                Form10.VScroll1.Visible = False
                Form10.Label2.Caption = "Search Results"
                Form10.Picture2.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 0))
                Form10.Picture2.PaintPicture Pic, 0, 0, Form10.Picture2.ScaleWidth, Form10.Picture2.ScaleHeight
                Set Form9.Picture2.Picture = Form9.Picture2.Image
                Form10.Label3.Caption = search_results(0, 0)
                Form10.Label4.Caption = search_results(5, 0)
                Form10.Picture3.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 1))
                Form10.Picture3.PaintPicture Pic, 0, 0, Form10.Picture3.ScaleWidth, Form10.Picture3.ScaleHeight
                Set Form9.Picture3.Picture = Form9.Picture3.Image
                Form10.Label5.Caption = search_results(0, 1)
                Form10.Label6.Caption = search_results(5, 1)
            ElseIf no_of_records = 3 Then
                Form10.Label2.Visible = True
                Form10.Label3.Visible = True
                Form10.Label4.Visible = True
                Form10.Label5.Visible = True
                Form10.Label6.Visible = True
                Form10.Label7.Visible = True
                Form10.Label8.Visible = True
                Form10.Shape1.Visible = True
                Form10.Picture2.Visible = True
                Form10.Picture3.Visible = True
                Form10.Picture4.Visible = True
                Form10.VScroll1.Visible = False
                Form10.Label2.Caption = "Search Results"
                Form10.Picture2.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 0))
                Form10.Picture2.PaintPicture Pic, 0, 0, Form10.Picture2.ScaleWidth, Form10.Picture2.ScaleHeight
                Set Form10.Picture2.Picture = Form10.Picture2.Image
                Form10.Label3.Caption = search_results(0, 0)
                Form10.Label4.Caption = search_results(5, 0)
                Form10.Picture3.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 1))
                Form10.Picture3.PaintPicture Pic, 0, 0, Form10.Picture3.ScaleWidth, Form10.Picture3.ScaleHeight
                Set Form9.Picture3.Picture = Form9.Picture3.Image
                Form10.Label5.Caption = search_results(0, 1)
                Form10.Label6.Caption = search_results(5, 1)
                Form10.Picture4.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, 2))
                Form10.Picture4.PaintPicture Pic, 0, 0, Form10.Picture4.ScaleWidth, Form10.Picture4.ScaleHeight
                Set Form9.Picture4.Picture = Form9.Picture4.Image
                Form10.Label7.Caption = search_results(0, 2)
                Form10.Label8.Caption = search_results(5, 2)
            Else
                Form10.Label2.Visible = True
                Form10.Label3.Visible = True
                Form10.Label4.Visible = True
                Form10.Label5.Visible = True
                Form10.Label6.Visible = True
                Form10.Label7.Visible = True
                Form10.Label8.Visible = True
                Form10.Shape1.Visible = True
                Form10.Picture2.Visible = True
                Form10.Picture3.Visible = True
                Form10.Picture4.Visible = True
                Form10.VScroll1.Visible = True
                Form10.Label2.Caption = "Search Results"
                Form10.Picture2.AutoRedraw = True
                status_search_index = 0
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, status_search_index))
                Form10.Picture2.PaintPicture Pic, 0, 0, Form10.Picture2.ScaleWidth, Form10.Picture2.ScaleHeight
                Set Form10.Picture2.Picture = Form10.Picture2.Image
                Form10.Label3.Caption = search_results(0, status_search_index)
                Form10.Label4.Caption = search_results(5, status_search_index)
                Form10.Picture3.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, status_search_index + 1))
                Form10.Picture3.PaintPicture Pic, 0, 0, Form10.Picture3.ScaleWidth, Form10.Picture3.ScaleHeight
                Set Form9.Picture3.Picture = Form9.Picture3.Image
                Form10.Label5.Caption = search_results(0, status_search_index + 1)
                Form10.Label6.Caption = search_results(5, status_search_index + 1)
                Form10.Picture4.AutoRedraw = True
                Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, status_search_index + 2))
                Form10.Picture4.PaintPicture Pic, 0, 0, Form10.Picture4.ScaleWidth, Form10.Picture4.ScaleHeight
                Set Form9.Picture4.Picture = Form9.Picture4.Image
                Form10.Label7.Caption = search_results(0, status_search_index + 2)
                Form10.Label8.Caption = search_results(5, status_search_index + 2)
                
                Form10.VScroll1.Value = 0
                Form10.VScroll1.Max = no_of_records - 3
                Form10.VScroll1.Min = 0
            End If
        End If
    End If
End Sub
Public Sub StatusSearchScrollChange()
    Dim j As Integer
    Dim Pic As Picture
    j = Form10.VScroll1.Value
    If j > status_search_index Then
        Form10.Picture2.Picture = Form10.Picture3.Picture
        Form10.Label3.Caption = Form10.Label5.Caption
        Form10.Label4.Caption = Form10.Label6.Caption
        
        Form10.Picture3.Picture = Form10.Picture4.Picture
        Form10.Label5.Caption = Form10.Label7.Caption
        Form10.Label6.Caption = Form10.Label8.Caption
        
        Form10.Picture4.AutoRedraw = True
        Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, j + 2))
        Form10.Picture4.PaintPicture Pic, 0, 0, Form10.Picture4.ScaleWidth, Form10.Picture4.ScaleHeight
        Set Form10.Picture4.Picture = Form10.Picture4.Image
        Form10.Label7.Caption = search_results(0, j + 2)
        Form10.Label8.Caption = search_results(5, j + 2)
        status_search_index = j
    End If
    If j < status_search_index Then
        Form10.Picture4.Picture = Form10.Picture3.Picture
        Form10.Label7.Caption = Form10.Label5.Caption
        Form10.Label8.Caption = Form10.Label6.Caption
        
        Form10.Picture3.Picture = Form10.Picture2.Picture
        Form10.Label5.Caption = Form10.Label3.Caption
        Form10.Label6.Caption = Form10.Label4.Caption
    
        Form10.Picture2.AutoRedraw = True
        Set Pic = LoadPicture(App.Path + "\db\" + search_results(1, j))
        Form10.Picture2.PaintPicture Pic, 0, 0, Form10.Picture2.ScaleWidth, Form10.Picture2.ScaleHeight
        Set Form10.Picture2.Picture = Form10.Picture2.Image
        Form10.Label3.Caption = search_results(0, j)
        Form10.Label4.Caption = search_results(5, j)
        status_search_index = j
    End If
End Sub
Public Sub UpdatemovieStatus(movie_name As String)
    Dim choice As String
    Dim current As Integer
    choice = MsgBox("Yes = Current Active, No = Currently Inactive", vbYesNo, "Choose Status")
    If choice = vbYes Then
        current = 1
    Else
        current = 0
    End If
    Dim cn As ADODB.Connection
    Dim db_path, query As String
    db_path = App.Path + "\db\MovieRatingSystem.mdb"
    Set cn = New ADODB.Connection
        cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
    query = "Update MovieDetails Set CURRENT=" & Val(current) & " where NAME=['" & movie_name & "'];"
    cn.Execute (query)
    MsgBox ("Status Updated SuccessFully!")
End Sub


