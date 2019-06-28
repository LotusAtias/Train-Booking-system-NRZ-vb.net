Imports System.Data.OleDb

Imports System.Math
Imports System.IO
Public Class FrmAdmin
    Dim provider As String
    Dim datafile As String
    Dim constring As String
    Dim myconnection As OleDbConnection = New OleDbConnection


    Dim choice As String
    Dim mnthchoice As String
    Dim yrchoice As String
    Dim a As String = "Chinhoyi"
    Dim b As String = "Bulawayo"
    Dim c As String = "Mutare"
    Dim d As String = "Shamva"


    Private Sub FrmAdmin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        provider = "provider = microsoft.ace.oledb.12.0; data source="
        datafile = " C:\Users\CHEKAZ\Documents\railwayticketbooking.accdb"
        constring = provider & datafile
        myconnection.ConnectionString = constring

        Dim sql1 As String
        sql1 = "SELECT * FROM cashier"
        Dim adapter1 As New OleDbDataAdapter(sql1, myconnection)
        Dim cmd1 As New OleDbCommand(sql1, myconnection)
        Dim dt1 As New DataTable("cashier")
        adapter1.Fill(dt1)
        DataGridView1.DataSource = dt1

        Dim sql2 As String
        sql2 = "SELECT * FROM officerincharge"
        Dim adapter2 As New OleDbDataAdapter(sql2, myconnection)
        Dim cmd2 As New OleDbCommand(sql2, myconnection)
        Dim dt2 As New DataTable("officerincharge")
        adapter2.Fill(dt2)
        DataGridView2.DataSource = dt2
        Try
            Dim adp As OleDbDataAdapter = New OleDbDataAdapter
            Dim dtt As New DataTable("studentss")
            adp.SelectCommand = New OleDbCommand("Select * FROM dailybookings ", myconnection)
            adp.Fill(dtt)

            Dim rpt As New CrystalReport2
            rpt.SetDataSource(dtt)
            CrystalReportViewer1.ReportSource = rpt

        Catch ex As Exception

        End Try

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            myconnection.Open()
            Dim sqlinsert As String
            sqlinsert = "INSERT INTO cashier([id], [name], [surname], [dateofbirth], [age], [username], [pwsword], [PicturePath] )" _
              & "VALUES(@id, @name, @surname, @dateofbirth, @age, @username, @pwsword, '" & Replace$(TextBoxPictureFilePath.Text, "'", "''") & "')"
            Dim cmd As New OleDbCommand(sqlinsert, myconnection)
            cmd.Parameters.Add(New OleDbParameter("@id", TextBox11.Text))
            cmd.Parameters.Add(New OleDbParameter("@name", TextBox1.Text))
            cmd.Parameters.Add(New OleDbParameter("@surname", TextBox2.Text))
            cmd.Parameters.Add(New OleDbParameter("@dateofbirth", DateTimePicker1.Value))
            cmd.Parameters.Add(New OleDbParameter("@age", TextBox3.Text))
            cmd.Parameters.Add(New OleDbParameter("@username", TextBox5.Text))
            cmd.Parameters.Add(New OleDbParameter("@pwsword", TextBox14.Text))
            cmd.Parameters.Add(New OleDbParameter("@PicturePath", TextBox3.Text))

            cmd.ExecuteNonQuery()
            MsgBox("You Have Successfully saved")
            myconnection.Close()

            refreshdgvAdmin()
            TextBox11.Clear()
            TextBox1.Clear()
            TextBox2.Clear()
            DateTimePicker1.Value = DateTime.Today
            TextBox3.Clear()
            TextBox14.Clear()
            TextBox5.Clear()
            PictureBox1.Image = Nothing
            TextBoxPictureFilePath.Text = "C:\Users\CHEKAZ\Downloads\images.png"
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            myconnection.Close()
        End Try


    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Dim iDagdag As Integer
        iDagdag = CInt(DateDiff(DateInterval.Year, DateTimePicker1.Value, Now) / 4)
        TextBox3.Text = Floor((DateDiff(DateInterval.Day, DateTimePicker1.Value, Now) - iDagdag) / 365)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim img As String

        Dim myStream As Stream = Nothing
        Dim openFileDialog1 As New OpenFileDialog()

        openFileDialog1.InitialDirectory = "c:\"
        openFileDialog1.Filter = Nothing
        openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = True
        openFileDialog1.FileName = ""

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                myStream = openFileDialog1.OpenFile()
                If (myStream IsNot Nothing) Then

                    TextBoxPictureFilePath.Text = ""

                    img = openFileDialog1.FileName
                    PictureBox1.Image = System.Drawing.Bitmap.FromFile(img)

                    TextBoxPictureFilePath.Text = openFileDialog1.FileName
                End If
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            Finally
                If (myStream IsNot Nothing) Then
                    myStream.Close()
                End If
            End Try
        End If
    End Sub
    Private Sub saveimagepath()
        Try
            Dim myConnection As OleDbConnection
            Dim myCommand As OleDbCommand
            Dim mySQLString As String
            myConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Users\CHEKAZ\Documents\railwayticketbooking.accdb")
            myConnection.Open()
            mySQLString = "INSERT INTO cashier (PicturePath) VALUES('" & Replace$(TextBoxPictureFilePath.Text, "'", "''") & "')"
            myCommand = New OleDbCommand(mySQLString, myConnection)
            myCommand.ExecuteNonQuery()
            Dim sql As String
            sql = "SELECT * FROM cashier"
            Dim adapter As New OleDbDataAdapter(sql, myConnection)
            Dim dt As New DataTable("cashier")
            adapter.Fill(dt)
            DataGridView1.DataSource = dt

        Catch ex As Exception
            MessageBox.Show(ex.Message & " - " & ex.Source)
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            ' This is our DELETE Statement. To be sure we delete the correct record and not all of 
            ' them.
            ' We use the WHERE to be sure only that record that the user has selected is deleted.
            Dim sqldelete As String
            sqldelete = "DELETE * FROM officerincharge WHERE id='" & DataGridView2.CurrentRow.Cells(0).Value.ToString & "'"

            ' This is our DataAdapter. This executes our SQL Statement above against the Database
            ' we defined in the Connection String
            Dim adapter As New OleDbDataAdapter(sqldelete, myconnection)
            ' Gets the records from the table and fills our adapter with those.
            Dim dt As New DataTable("oficerincharge")
            adapter.Fill(dt)
            ' Assigns the edited DataSource on the DataGridView and the refreshes the 
            ' view to ensure everything is up to date in real time.
            DataGridView2.DataSource = dt
            ' This is a Sub in Module 1 to refresh the DataGridView when information is added,
            '  updated, or deleted.
            refreshdgvAdmin()
        Catch ex As Exception
            MsgBox(ex.Message & "or There is nothing to Delete")

        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            myconnection.Open()
            Dim sqlinsert As String
            sqlinsert = "INSERT INTO officerincharge([id], [name], [surname], [dateofbirth], [age], [username], [pwsword], [PicturePath] )" _
              & "VALUES(@id, @name, @surname, @dateofbirth, @age, @username, @pwsword, '" & Replace$(TextBox13.Text, "'", "''") & "')"
            Dim cmd As New OleDbCommand(sqlinsert, myconnection)
            cmd.Parameters.Add(New OleDbParameter("@id", TextBox12.Text))
            cmd.Parameters.Add(New OleDbParameter("@name", TextBox10.Text))
            cmd.Parameters.Add(New OleDbParameter("@surname", TextBox9.Text))
            cmd.Parameters.Add(New OleDbParameter("@dateofbirth", DateTimePicker2.Value))
            cmd.Parameters.Add(New OleDbParameter("@age", TextBox8.Text))
            cmd.Parameters.Add(New OleDbParameter("@username", TextBox7.Text))
            cmd.Parameters.Add(New OleDbParameter("@pwsword", TextBox4.Text))
            cmd.Parameters.Add(New OleDbParameter("@PicturePath", TextBox13.Text))

            cmd.ExecuteNonQuery()
            MsgBox("You Have Successfully saved")
            refreshdgvAdmin()
            myconnection.Close()
            TextBox12.Clear()
            TextBox10.Clear()
            TextBox9.Clear()
            DateTimePicker2.Value = Date.Now
            TextBox8.Clear()
            TextBox7.Clear()
            TextBox4.Clear()
            PictureBox2.Image = Nothing
            TextBox13.Text = "C:\Users\CHEKAZ\Downloads\images.png"
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            myconnection.Close()
        End Try
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Try
            TextBox11.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString
            TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString
            TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString
            DateTimePicker1.Value = DataGridView1.Rows(e.RowIndex).Cells(3).Value
            TextBox3.Text = DataGridView1.Rows(e.RowIndex).Cells(4).Value.ToString
            TextBox5.Text = DataGridView1.Rows(e.RowIndex).Cells(5).Value.ToString
            TextBox14.Text = DataGridView1.Rows(e.RowIndex).Cells(6).Value.ToString
            PictureBox1.ImageLocation = DataGridView1.Rows(e.RowIndex).Cells(7).Value.ToString
            TextBoxPictureFilePath.Text = DataGridView1.Rows(e.RowIndex).Cells(7).Value.ToString
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Dim sqlupdate As String
        ' Here we use the UPDATE Statement to update the information. To be sure we are 
        ' updating the right record we also use the WHERE clause to be sureno information
        ' is added or changed in the other records
        sqlupdate = "UPDATE cashier SET name=@name, surname=@surname, dateofbirth=@dateofbirth," _
            & "age=@age, username=@username, pwsword=@pwsword, PicturePath=@PicturePath WHERE [id]='" & TextBox11.Text & "'"

        Dim cmd As New OleDbCommand(sqlupdate, myconnection)
        'This assigns the values for our columns in the DataBase. 
        ' To ensure the correct values are written to the correct column
        cmd.Parameters.Add(New OleDbParameter("@name", TextBox1.Text))
        cmd.Parameters.Add(New OleDbParameter("@surname", TextBox2.Text))
        cmd.Parameters.Add(New OleDbParameter("@dateofbirth", DateTimePicker1.Value))
        cmd.Parameters.Add(New OleDbParameter("@age", TextBox3.Text))
        cmd.Parameters.Add(New OleDbParameter("@username", TextBox5.Text))
        cmd.Parameters.Add(New OleDbParameter("@pwsword", TextBox14.Text))
        cmd.Parameters.Add(New OleDbParameter("@PicturePath", TextBoxPictureFilePath.Text))
        ' This is what actually writes our changes to the DataBase.
        ' You have to open the connection, execute the commands and
        ' then close connection.
        Try
            myconnection.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try
        ' This are subs in Module1, to clear all the TextBoxes on the form
        ' and refresh the DataGridView on the MainForm to show our new records.
        refreshdgvAdmin()
        TextBox11.Clear()
        TextBox1.Clear()
        TextBox2.Clear()
        DateTimePicker1.Value = Date.Now
        TextBox3.Clear()
        TextBox14.Clear()
        TextBox5.Clear()
        PictureBox1.Image = Nothing
        TextBoxPictureFilePath.Text = "C:\Users\CHEKAZ\Downloads\images.png"
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            ' This is our DELETE Statement. To be sure we delete the correct record and not all of 
            ' them.
            ' We use the WHERE to be sure only that record that the user has selected is deleted.
            Dim sqldelete As String
            sqldelete = "DELETE * FROM cashier WHERE id='" & DataGridView1.CurrentRow.Cells(0).Value.ToString & "'"

            ' This is our DataAdapter. This executes our SQL Statement above against the Database
            ' we defined in the Connection String
            Dim adapter As New OleDbDataAdapter(sqldelete, myconnection)
            ' Gets the records from the table and fills our adapter with those.
            Dim dt As New DataTable("cashier")
            adapter.Fill(dt)
            ' Assigns the edited DataSource on the DataGridView and the refreshes the 
            ' view to ensure everything is up to date in real time.
            DataGridView1.DataSource = dt
            ' This is a Sub in Module 1 to refresh the DataGridView when information is added,
            '  updated, or deleted.
            refreshdgvAdmin()
        Catch ex As Exception
            MsgBox(ex.Message & "or There is nothing to Delete")

        End Try
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        Dim iDagdag As Integer
        iDagdag = CInt(DateDiff(DateInterval.Year, DateTimePicker2.Value, Now) / 4)
        TextBox8.Text = Floor((DateDiff(DateInterval.Day, DateTimePicker2.Value, Now) - iDagdag) / 365)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim img As String

        Dim myStream As Stream = Nothing
        Dim openFileDialog1 As New OpenFileDialog()

        openFileDialog1.InitialDirectory = "c:\"
        openFileDialog1.Filter = Nothing
        openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = True
        openFileDialog1.FileName = ""

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                myStream = openFileDialog1.OpenFile()
                If (myStream IsNot Nothing) Then

                    TextBox13.Text = ""

                    img = openFileDialog1.FileName
                    PictureBox2.Image = System.Drawing.Bitmap.FromFile(img)

                    TextBox13.Text = openFileDialog1.FileName
                End If
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            Finally
                If (myStream IsNot Nothing) Then
                    myStream.Close()
                End If
            End Try
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim sqlupdate As String
        ' Here we use the UPDATE Statement to update the information. To be sure we are 
        ' updating the right record we also use the WHERE clause to be sureno information
        ' is added or changed in the other records
        sqlupdate = "UPDATE officerincharge SET name=@name, surname=@surname, dateofbirth=@dateofbirth," _
            & "age=@age, username=@username, pwsword=@pwsword, PicturePath=@PicturePath WHERE [id]='" & TextBox12.Text & "'"

        Dim cmd As New OleDbCommand(sqlupdate, myconnection)
        'This assigns the values for our columns in the DataBase. 
        ' To ensure the correct values are written to the correct column
        cmd.Parameters.Add(New OleDbParameter("@name", TextBox10.Text))
        cmd.Parameters.Add(New OleDbParameter("@surname", TextBox9.Text))
        cmd.Parameters.Add(New OleDbParameter("@dateofbirth", DateTimePicker2.Value))
        cmd.Parameters.Add(New OleDbParameter("@age", TextBox8.Text))
        cmd.Parameters.Add(New OleDbParameter("@username", TextBox7.Text))
        cmd.Parameters.Add(New OleDbParameter("@pwsword", TextBox4.Text))
        cmd.Parameters.Add(New OleDbParameter("@PicturePath", TextBox13.Text))
        ' This is what actually writes our changes to the DataBase.
        ' You have to open the connection, execute the commands and
        ' then close connection.
        Try
            myconnection.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try
        ' This are subs in Module1, to clear all the TextBoxes on the form
        ' and refresh the DataGridView on the MainForm to show our new records.
        refreshdgvAdmin()
        TextBox12.Clear()
        TextBox10.Clear()
        TextBox9.Clear()
        DateTimePicker2.Value = Date.Now
        TextBox8.Clear()
        TextBox7.Clear()
        TextBox4.Clear()
        PictureBox2.Image = Nothing
        TextBox13.Text = "C:\Users\CHEKAZ\Downloads\images.png"
    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Try
            TextBox12.Text = DataGridView2.Rows(e.RowIndex).Cells(0).Value.ToString
            TextBox10.Text = DataGridView2.Rows(e.RowIndex).Cells(1).Value.ToString
            TextBox9.Text = DataGridView2.Rows(e.RowIndex).Cells(2).Value.ToString
            DateTimePicker2.Value = DataGridView2.Rows(e.RowIndex).Cells(3).Value
            TextBox8.Text = DataGridView2.Rows(e.RowIndex).Cells(4).Value.ToString
            TextBox7.Text = DataGridView2.Rows(e.RowIndex).Cells(5).Value.ToString
            TextBox4.Text = DataGridView2.Rows(e.RowIndex).Cells(6).Value.ToString
            PictureBox2.ImageLocation = DataGridView2.Rows(e.RowIndex).Cells(7).Value.ToString
            TextBox13.Text = DataGridView2.Rows(e.RowIndex).Cells(7).Value.ToString
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim count, count1, count2, count3 As Integer
        Dim fnldate As String = mnthchoice + yrchoice
        Chart1.Series(0).Points.Clear()
        myconnection.Open()
        Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM [dailybookings] WHERE [to_train] ='" & a & "' AND  date_train LIKE '%" & fnldate & "%' ", myconnection)
        Dim dr As OleDbDataReader = cmd.ExecuteReader
        While dr.Read
            count = count + 1
        End While
        Chart1.Series("Number of People").Points.AddXY(a, count)
        dr.Close()
        cmd.Dispose()
        myconnection.Close()

        myconnection.Open()
        Dim cmd1 As OleDbCommand = New OleDbCommand("SELECT * FROM [dailybookings] WHERE [to_train] ='" & b & "' AND  date_train LIKE '%" & fnldate & "%' ", myconnection)
        Dim dr1 As OleDbDataReader = cmd1.ExecuteReader
        While dr1.Read
            count1 = count1 + 1
        End While
        Chart1.Series("Number of People").Points.AddXY(b, count1)
        dr1.Close()
        cmd1.Dispose()
        myconnection.Close()

        myconnection.Open()
        Dim cmd2 As OleDbCommand = New OleDbCommand("SELECT * FROM [dailybookings] WHERE [to_train] ='" & c & "' AND  date_train LIKE '%" & fnldate & "%' ", myconnection)
        Dim dr2 As OleDbDataReader = cmd2.ExecuteReader
        While dr2.Read
            count2 = count2 + 1
        End While
        Chart1.Series("Number of People").Points.AddXY(c, count2)
        dr2.Close()
        cmd2.Dispose()
        myconnection.Close()

        myconnection.Open()
        Dim cmd3 As OleDbCommand = New OleDbCommand("SELECT * FROM [dailybookings] WHERE [to_train] ='" & d & "' AND  date_train LIKE '%" & fnldate & "%' ", myconnection)
        Dim dr3 As OleDbDataReader = cmd3.ExecuteReader
        While dr3.Read
            count3 = count3 + 1
        End While
        Chart1.Series("Number of People").Points.AddXY(d, count3)
        dr3.Close()
        cmd3.Dispose()
        myconnection.Close()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedItem = "January" Then
            mnthchoice = "1/"
        ElseIf ComboBox1.SelectedItem = "February"
            mnthchoice = "2/"
        ElseIf ComboBox1.SelectedItem = "March"
            mnthchoice = "3/"
        ElseIf ComboBox1.SelectedItem = "April"
            mnthchoice = "4/"
        ElseIf ComboBox1.SelectedItem = "May"
            mnthchoice = "5/"
        ElseIf ComboBox1.SelectedItem = "June"
            mnthchoice = "6/"
        ElseIf ComboBox1.SelectedItem = "July"
            mnthchoice = "7/"
        ElseIf ComboBox1.SelectedItem = "August"
            mnthchoice = "8/"
        ElseIf ComboBox1.SelectedItem = "September"
            mnthchoice = "9/"
        ElseIf ComboBox1.SelectedItem = "October"
            mnthchoice = "10/"
        ElseIf ComboBox1.SelectedItem = "November"
            mnthchoice = "11/"
        ElseIf ComboBox1.SelectedItem = "December"
            mnthchoice = "12/"
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedItem = "2018" Then
            yrchoice = "2018"
        ElseIf ComboBox2.SelectedItem = "2019"
            yrchoice = "2019"
        ElseIf ComboBox2.SelectedItem = "2020"
            yrchoice = "2020"
        ElseIf ComboBox2.SelectedItem = "2021"
            yrchoice = "2021"
        ElseIf ComboBox2.SelectedItem = "2022"
            yrchoice = "2022"
        ElseIf ComboBox2.SelectedItem = "2023"
            yrchoice = "2023"
        ElseIf ComboBox2.SelectedItem = "2024"
            yrchoice = "2024"
        ElseIf ComboBox2.SelectedItem = "2025"
            yrchoice = "2025"
        End If
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        FrmLogin.Show()
        Me.Hide()
        myconnection.Close()
    End Sub
End Class