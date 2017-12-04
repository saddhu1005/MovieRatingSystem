Attribute VB_Name = "AddMovie"
Public Sub AddRecord()

    If (Form7.Text1.Text = "" Or Form7.Text2.Text = "" Or Form7.Text4.Text = "" Or Form7.Text5.Text = "" Or Form7.Text6.Text = "" Or Form7.Text7.Text = "" Or Form7.Text8.Text = "") Then
        MsgBox ("Please Enter All The Necessary details")
    Else
        If Form7.Option1.Value = False And Form7.Option2.Value = False Then
            MsgBox ("Current Status Not Checked")
        Else
            Dim cn As ADODB.Connection
            Dim rs As ADODB.Recordset
            Dim name, director, genre, language, cast, image_path, relative_image_path, filename, ds As String
            Dim synopsis1, synopsis2, synopsis3, synopsis4, synopsis5, synopsis6, synopsis7, original_synopsis As String
            Dim is_current, no_of_synopsis, extra_length As Integer
            Dim release_date As Date
            
            
            ds = App.Path + "/db/MovieRatingSystem.mdb"
            
            Set cn = New ADODB.Connection
                cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & ds
            Set rs = New ADODB.Recordset
                rs.Open "Select * from MovieDetails", cn, adOpenStatic, adLockOptimistic
            
            
            name = UCase(Form7.Text1.Text)                                                      'Movie Name
            
            director = Form7.Text6.Text                                                         'Movie Director
            
            cast = Form7.Text7.Text                                                             'Movie Cast
            
            genre = Form7.Text5.Text                                                            'Movie Genre
            
            language = Form7.Text4.Text                                                         'Movie Language
            
            If Form7.Option1.Value = True Then                                                  'Movie is current
                is_current = 1
            End If
            If Form7.Option2.Value = True Then
                is_current = 0
            End If
            
            release_date = Form7.DTPicker1.Value                                               'Movie Release Date
            
            no_of_synopsis = (Len(Form7.Text8.Text) \ 255) + 1                                 'Breaking the Original Synopsis
            original_synopsis = Form7.Text8.Text
            If no_of_synopsis = 1 Then
                synopsis1 = original_synopsis
            ElseIf no_of_synopsis = 2 Then
                synopsis1 = Left(original_synopsis, 255)
                original_synopsis = Replace(original_synopsis, synopsis1, "")
                synopsis2 = original_synopsis
            ElseIf no_of_synopsis = 3 Then
                synopsis1 = Left(original_synopsis, 255)
                original_synopsis = Replace(original_synopsis, synopsis1, "")
                synopsis2 = Left(original_synopsis, 255)
                original_synopsis = Replace(original_synopsis, synopsis2, "")
                synopsis3 = original_synopsis
            ElseIf no_of_synopsis = 4 Then
                synopsis1 = Left(original_synopsis, 255)
                original_synopsis = Replace(original_synopsis, synopsis1, "")
                synopsis2 = Left(original_synopsis, 255)
                original_synopsis = Replace(original_synopsis, synopsis2, "")
                synopsis3 = Left(original_synopsis, 255)
                original_synopsis = Replace(original_synopsis, synopsis3, "")
                synopsis4 = original_synopsis
            ElseIf no_of_synopsis = 5 Then
                synopsis1 = Left(original_synopsis, 255)
                original_synopsis = Replace(original_synopsis, synopsis1, "")
                synopsis2 = Left(original_synopsis, 255)
                original_synopsis = Replace(original_synopsis, synopsis2, "")
                synopsis3 = Left(original_synopsis, 255)
                original_synopsis = Replace(original_synopsis, synopsis3, "")
                synopsis4 = Left(original_synopsis, 255)
                original_synopsis = Replace(original_synopsis, synopsis4, "")
                synopsis5 = original_synopsis
            ElseIf no_of_synopsis = 6 Then
                synopsis1 = Left(original_synopsis, 255)
                original_synopsis = Replace(original_synopsis, synopsis1, "")
                synopsis2 = Left(original_synopsis, 255)
                original_synopsis = Replace(original_synopsis, synopsis2, "")
                synopsis3 = Left(original_synopsis, 255)
                original_synopsis = Replace(original_synopsis, synopsis3, "")
                synopsis4 = Left(original_synopsis, 255)
                original_synopsis = Replace(original_synopsis, synopsis4, "")
                synopsis5 = Left(original_synopsis, 255)
                original_synopsis = Replace(original_synopsis, synopsis5, "")
                synopsis6 = original_synopsis
            ElseIf no_of_synopsis = 7 Then
                synopsis1 = Left(original_synopsis, 255)
                original_synopsis = Replace(original_synopsis, synopsis1, "")
                synopsis2 = Left(original_synopsis, 255)
                original_synopsis = Replace(original_synopsis, synopsis2, "")
                synopsis3 = Left(original_synopsis, 255)
                original_synopsis = Replace(original_synopsis, synopsis3, "")
                synopsis4 = Left(original_synopsis, 255)
                original_synopsis = Replace(original_synopsis, synopsis4, "")
                synopsis5 = Left(original_synopsis, 255)
                original_synopsis = Replace(original_synopsis, synopsis5, "")
                synopsis6 = Left(original_synopsis, 255)
                original_synopsis = Replace(original_synopsis, synopsis6, "")
                synopsis7 = original_synopsis
            Else
                MsgBox ("Synosis Is Too Long")
            End If
            
            image_folder_path = App.Path + "\db\" + relative_image_path
            image_path = Form7.CommonDialog1.filename                                                                 'Getting the Image path Ready
            filename = Mid(image_path, InStrRev(image_path, "\") + 1, Len(image_path))
            relative_image_path = "image_database\" + filename
            image_folder_path = App.Path + "\db\" + relative_image_path
            FileCopy image_path, image_folder_path
            
            addMovieToMovieRatings name
            addMovieToMovieComments name
            
            With rs
                .AddNew
                .Fields("NAME").Value = name
                .Fields("IMAGE").Value = relative_image_path
                .Fields("RELEASE DATE").Value = release_date
                .Fields("LANGUAGE").Value = language
                .Fields("DIRECTOR").Value = director
                .Fields("GENRE").Value = genre
                .Fields("CAST").Value = cast
                .Fields("CURRENT").Value = is_current
                If (Not (synopsis1 = "")) Then
                    .Fields("SYNOPSIS1") = synopsis1
                End If
                If (Not (synopsis2 = "")) Then
                    .Fields("SYNOPSIS2") = synopsis2
                End If
                If (Not (synopsis3 = "")) Then
                    .Fields("SYNOPSIS3") = synopsis3
                End If
                If (Not (synopsis4 = "")) Then
                    .Fields("SYNOPSIS4") = synopsis4
                End If
                If (Not (synopsis5 = "")) Then
                    .Fields("SYNOPSIS5") = synopsis5
                End If
                If (Not (synopsis6 = "")) Then
                    .Fields("SYNOPSIS1") = synopsis6
                End If
                If (Not (synopsis7 = "")) Then
                    .Fields("SYNOPSIS7") = synopsis7
                End If
                .Update
            End With
            
            
            MsgBox ("Movie Successfully Added")
            cn.Close
            Form7.Text1.Text = ""
            Form7.Text2.Text = ""
            Form7.Text4.Text = ""
            Form7.Text5.Text = ""
            Form7.Text6.Text = ""
            Form7.Text7.Text = ""
            Form7.Text8.Text = ""
        End If
    End If
End Sub

Private Sub addMovieToMovieRatings(ByVal movie_name As String)
    Dim db As Database
    Dim query As String
    Set db = OpenDatabase(App.Path + "/db/MovieRatingSystem.mdb")
    query = "ALTER TABLE MovieRatings ADD COLUMN [" & movie_name & "] INT;"
    db.Execute (query)
    db.Close
End Sub
Private Sub addMovieToMovieComments(ByVal movie_name As String)
    Dim db As Database
    Dim query As String
    Set db = OpenDatabase(App.Path + "/db/MovieRatingSystem.mdb")
    query = "ALTER TABLE MovieComments ADD COLUMN [" & movie_name & "] CHAR(255);"
    db.Execute (query)
    db.Close
End Sub
