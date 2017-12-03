Attribute VB_Name = "MovieRatings"
Public Sub SubmitMovieRatings(movie_name As String)
    If Form6.Combo1.Text = "RATE THE MOVIE" Then
        MsgBox ("Please Select Your rating Then press Submit")
    Else
        Dim cn As ADODB.Connection
        Dim db_path, query As String
        db_path = App.Path + "\db\MovieRatingSystem.mdb"
        Set cn = New ADODB.Connection
            cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
        query = "Update MovieRatings Set [" & movie_name & "]=" & Val(Form6.Combo1.Text) & " where USER = '" & current_user & "'"
        cn.Execute (query)
        MsgBox ("Movie Rated SuccessFully!")
        CalculateOverallRatings
        RatingsDisplay (movie_name)
    End If
End Sub
Public Sub EditMovieRatings(movie_name As String)
        Dim cn As ADODB.Connection
        Dim db_path, query As String
        db_path = App.Path + "\db\MovieRatingSystem.mdb"
        Set cn = New ADODB.Connection
            cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
        query = "select * from MovieRatings where USER = '" & current_user & "'"
        query = "Update MovieRatings Set [" & movie_name & "]=" & "NULL" & " where USER = '" & current_user & "'"
        cn.Execute (query)
        RatingsDisplay (movie_name)
End Sub
Public Sub CalculateOverallRatings()
    Dim cn As ADODB.Connection
    Dim rs1, rs2 As ADODB.Recordset
    Dim db_path, query1, query2, query As String
    Dim Ratings_array As Variant
    Dim no_of_records, no_of_fields, sum, count As Integer
    Dim rating As Double
    db_path = App.Path + "\db\MovieRatingSystem.mdb"
    Set cn = New ADODB.Connection
        cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
    query1 = "select * from MovieDetails"
    query2 = "select * from MovieRatings"
    Set rs1 = New ADODB.Recordset
        rs1.Open query1, cn, adOpenStatic, adLockOptimistic
    Set rs2 = New ADODB.Recordset
        rs2.Open query2, cn, adOpenStatic, adLockOptimistic
    rs2.MoveLast
    no_of_records = rs2.RecordCount
    rs2.MoveFirst
    no_of_fields = rs2.Fields.count - 1
    Ratings_array = rs2.GetRows(no_of_records)
    
    For i = 1 To no_of_fields
        sum = 0
        count = 0
        For j = 0 To (no_of_records - 1)
            If Ratings_array(i, j) Then
                sum = sum + Val(Ratings_array(i, j))
                count = count + 1
            End If
        Next
        If (count = 0) Then
            query = "Update MovieDetails Set RATING=" & "NULL" & " where NAME = '" & rs2.Fields(i).Name & "'"
        Else
            rating = sum / count
            query = "Update MovieDetails Set RATING=" & CStr(rating) & " where NAME = '" & rs2.Fields(i).Name & "'"
        End If
        cn.Execute (query)
    Next
    cn.Close
End Sub

