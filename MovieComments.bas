Attribute VB_Name = "MovieComments"
Private comment_index As Integer
Private Comment_array As Variant
Public Sub SubmitMovieComments(movie_name As String)
    If Len(Form6.Text1.Text) > 254 Then
        MsgBox ("Review Should have length greater than 255 characters!")
    Else
        If Form6.Text1.Text = "" Or Form6.Text1.Text = "Write Your Review...." Then
            MsgBox ("Please Enter your review and then press Submit")
        Else
            Dim cn As ADODB.Connection
            Dim db_path, query As String
            db_path = App.Path + "\db\MovieRatingSystem.mdb"
            Set cn = New ADODB.Connection
                cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
            query = "Update MovieComments Set [" & movie_name & "]='" & CStr(Form6.Text1.Text) & "' where USER = '" & current_user & "'"
            cn.Execute (query)
            MsgBox ("Review Submitted SuccessFully!")
            RatingsDisplay (movie_name)
        End If
    End If
End Sub

Public Sub EditMovieComments(movie_name As String)
    Dim cn As ADODB.Connection
    Dim db_path, query As String
    db_path = App.Path + "\db\MovieRatingSystem.mdb"
    Set cn = New ADODB.Connection
        cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
    query = "Update MovieComments Set [" & movie_name & "]='" & "" & "' where USER = '" & current_user & "'"
    cn.Execute (query)
    RatingsDisplay (movie_name)
End Sub

Public Sub CollectComments(movie_name As String)
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim no_of_records As Integer
    Dim db_path, query As String
    db_path = App.Path + "\db\MovieRatingSystem.mdb"
    Set cn = New ADODB.Connection
        cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
    query = "select USER,[" & movie_name & "] from Moviecomments where [" & movie_name & "] IS NOT NULL AND [" & movie_name & "] <> ''"
    Set rs = New ADODB.Recordset
        rs.Open query, cn, adOpenStatic, adLockOptimistic
    If rs.EOF And rs.BOF Then
        Form6.Label8.Visible = False
        Form6.Label9.Visible = False
        Form6.Label10.Visible = False
        Form6.Label11.Visible = False
        Form6.Label12.Visible = False
        Form6.Label13.Visible = False
        Form6.Line1.Visible = False
        Form6.Line2.Visible = False
        Form6.VScroll1.Visible = False
        Form6.Shape1.Visible = False
    Else
        rs.MoveLast
        no_of_records = rs.RecordCount
        rs.MoveFirst
        Comment_array = rs.GetRows(no_of_records)
        If no_of_records = 1 Then
            Form6.Label8.Visible = True
            Form6.Label9.Visible = True
            Form6.Label10.Visible = False
            Form6.Label11.Visible = False
            Form6.Label12.Visible = False
            Form6.Label13.Visible = False
            Form6.Line1.Visible = True
            Form6.Line2.Visible = False
            Form6.VScroll1.Visible = False
            Form6.Shape1.Visible = True
            Form6.Label8.Caption = Comment_array(0, 0) + ":"
            Form6.Label8.FontUnderline = True
            Form6.Label9.Caption = Comment_array(1, 0)
         ElseIf no_of_records = 2 Then
            Form6.Label8.Visible = True
            Form6.Label9.Visible = True
            Form6.Label10.Visible = True
            Form6.Label11.Visible = True
            Form6.Label12.Visible = False
            Form6.Label13.Visible = False
            Form6.Line1.Visible = True
            Form6.Line2.Visible = True
            Form6.VScroll1.Visible = False
            Form6.Shape1.Visible = True
            Form6.Label8.Caption = Comment_array(0, 0) + ":"
            Form6.Label8.FontUnderline = True
            Form6.Label9.Caption = Comment_array(1, 0)
            Form6.Label10.Caption = Comment_array(0, 1) + ":"
            Form6.Label10.FontUnderline = True
            Form6.Label11.Caption = Comment_array(1, 1)
         ElseIf no_of_records = 3 Then
            Form6.Label8.Visible = True
            Form6.Label9.Visible = True
            Form6.Label10.Visible = True
            Form6.Label11.Visible = True
            Form6.Label12.Visible = True
            Form6.Label13.Visible = True
            Form6.Line1.Visible = True
            Form6.Line2.Visible = True
            Form6.VScroll1.Visible = False
            Form6.Shape1.Visible = True
            Form6.Label8.Caption = Comment_array(0, 0) + ":"
            Form6.Label8.FontUnderline = True
            Form6.Label9.Caption = Comment_array(1, 0)
            Form6.Label10.Caption = Comment_array(0, 1) + ":"
            Form6.Label10.FontUnderline = True
            Form6.Label11.Caption = Comment_array(1, 1)
            Form6.Label12.Caption = Comment_array(0, 2) + ":"
            Form6.Label12.FontUnderline = True
            Form6.Label13.Caption = Comment_array(1, 2)
        Else
            comment_index = 0
            Form6.Label8.Visible = True
            Form6.Label9.Visible = True
            Form6.Label10.Visible = True
            Form6.Label11.Visible = True
            Form6.Label12.Visible = True
            Form6.Label13.Visible = True
            Form6.Line1.Visible = True
            Form6.Line2.Visible = True
            Form6.VScroll1.Visible = True
            Form6.Shape1.Visible = True
            Form6.Label8.Caption = Comment_array(0, comment_index) + ":"
            Form6.Label8.FontUnderline = True
            Form6.Label9.Caption = Comment_array(1, comment_index)
            Form6.Label10.Caption = Comment_array(0, comment_index + 1) + ":"
            Form6.Label10.FontUnderline = True
            Form6.Label11.Caption = Comment_array(1, comment_index + 1)
            Form6.Label12.Caption = Comment_array(0, comment_index + 2) + ":"
            Form6.Label12.FontUnderline = True
            Form6.Label13.Caption = Comment_array(1, comment_index + 2)
            
            Form6.VScroll1.Value = 0
            Form6.VScroll1.Max = no_of_records - 3
            Form6.VScroll1.Min = 0
        End If
    End If
End Sub

Public Sub CommentScrollChange()
    Dim j As Integer
    j = Form6.VScroll1.Value
    If j > comment_index Then
        Form6.Label8.Caption = Form6.Label10.Caption
        Form6.Label8.FontUnderline = True
        Form6.Label9.Caption = Form6.Label11.Caption
        
        Form6.Label10.Caption = Form6.Label12.Caption
        Form6.Label10.FontUnderline = True
        Form6.Label11.Caption = Form6.Label13.Caption
        

        
        Form6.Label12.Caption = Comment_array(0, j + 2) + ":"
        Form6.Label12.FontUnderline = True
        Form6.Label13.Caption = Comment_array(1, j + 2)
        comment_index = j
    End If
    If j < comment_index Then
        Form6.Label12.Caption = Form6.Label10.Caption
        Form6.Label12.FontUnderline = True
        Form6.Label13.Caption = Form6.Label11.Caption

        
        Form6.Label10.Caption = Form6.Label8.Caption
        Form6.Label10.FontUnderline = True
        Form6.Label11.Caption = Form6.Label9.Caption
        
        Form6.Label8.Caption = Comment_array(0, j) + ":"
        Form6.Label8.FontUnderline = True
        Form6.Label9.Caption = Comment_array(1, j)
    
        comment_index = j
    End If
End Sub
