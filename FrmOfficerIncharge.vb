Imports System.Data.OleDb
Public Class FrmOfficerIncharge
    Dim provider As String
    Dim datafile As String
    Dim constring As String
    Public imagevariablee As String
    Dim myconnection As OleDbConnection = New OleDbConnection

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If CheckBox1.Checked = True Then
            Dim a As String = "Verified"
            myconnection.ConnectionString = constring
            Dim sqlupdate As String
            sqlupdate = "UPDATE dailybookings SET flag=@flag WHERE [ticket_no]='" & TextBox1.Text & "'"
            Dim cmd As New OleDbCommand(sqlupdate, myconnection)

            cmd.Parameters.Add(New OleDbParameter("@flag", a))
            Try
                myconnection.Open()
                cmd.ExecuteNonQuery()
                TextBox1.Clear()
                TextBox2.Clear()
                TextBox4.Clear()
                CheckBox1.Checked = Nothing
                refreshofficer()
            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                myconnection.Close()
            End Try
        End If
    End Sub

    Private Sub FrmOfficerIncharge_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        provider = "provider = microsoft.ace.oledb.12.0; data source="
        datafile = " C:\Users\CHEKAZ\Documents\railwayticketbooking.accdb"
        constring = provider & datafile
        myconnection.ConnectionString = constring
        Label6.Text = Format(Now, "d/M/yyyy")
        insertimage()
        Dim sql As String
        sql = "SELECT ticket_no, name, surname, flag FROM dailybookings"
        Dim adapter As New OleDbDataAdapter(sql, myconnection)
        Dim cmd As New OleDbCommand(sql, myconnection)
        Dim dt As New DataTable("ticket")
        adapter.Fill(dt)
        DataGridView2.DataSource = dt

        Dim sql1 As String
        sql1 = "SELECT * FROM dailybookings"
        Dim adapter1 As New OleDbDataAdapter(sql1, myconnection)
        Dim cmd1 As New OleDbCommand(sql1, myconnection)
        Dim dt1 As New DataTable("ticket")
        adapter1.Fill(dt1)
        DataGridView1.DataSource = dt1

    End Sub
    Private Sub insertimage()
        Dim cnn As New OleDb.OleDbConnection
        cnn = New OleDb.OleDbConnection
        cnn.ConnectionString = "Provider=Microsoft.ace.OLEDB.12.0;Data Source=C:\Users\CHEKAZ\Documents\railwayticketbooking.accdb;"
        If cnn.State = ConnectionState.Open Then
            cnn.Open()
        End If
        Dim da As New OleDb.OleDbDataAdapter("SELECT * FROM officerincharge WHERE username='" & imagevariablee & "'", cnn)
        Dim dt As New DataTable
        da.Fill(dt)
        Try
            Dim ad As String = CStr(dt.Rows(0).Item("PicturePath"))
            cnn.Close()
            PictureBox1.Image = System.Drawing.Bitmap.FromFile(ad)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub MonthCalendar1_DateChanged(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar1.DateChanged
        myconnection.ConnectionString = constring
        Label6.Text = MonthCalendar1.SelectionStart.ToShortDateString
        Dim sqlsearch As String
        sqlsearch = "SELECT ticket_no, name, surname, flag FROM dailybookings WHERE date_train LIKE '%" & Label6.Text & "%'"
        ' Once again we execute the SQL statements against our DataBase
        Dim adapter As New OleDbDataAdapter(sqlsearch, myconnection)
        ' Shows the records and updates the DataGridView
        Dim dt As New DataTable("dailybookings")
        adapter.Fill(dt)
        DataGridView2.DataSource = dt

        Dim sqlsearch1 As String
        sqlsearch1 = "SELECT * FROM dailybookings WHERE date_train LIKE '%" & Label6.Text & "%'"
        ' Once again we execute the SQL statements against our DataBase
        Dim adapter1 As New OleDbDataAdapter(sqlsearch1, myconnection)
        ' Shows the records and updates the DataGridView
        Dim dt1 As New DataTable("dailybookings")
        adapter1.Fill(dt1)
        DataGridView1.DataSource = dt1
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Try
            TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString
            TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value.ToString
            TextBox4.Text = DataGridView1.Rows(e.RowIndex).Cells(4).Value.ToString
        Catch ex As Exception

        End Try
    End Sub
    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        myconnection.ConnectionString = constring
        Label6.Text = MonthCalendar1.SelectionStart.ToShortDateString
        Dim sqlsearch As String
        sqlsearch = "SELECT * FROM dailybookings WHERE ticket_no LIKE '%" & TextBox3.Text & "%' or name LIKE '%" & TextBox3.Text & "%' or surname LIKE '%" & TextBox3.Text & "%'"
        ' Once again we execute the SQL statements against our DataBase
        Dim adapter As New OleDbDataAdapter(sqlsearch, myconnection)
        ' Shows the records and updates the DataGridView
        Dim dt As New DataTable("dailybookings")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
    End Sub
    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        FrmLogin.Show()
        Me.Hide()
        myconnection.Close()
    End Sub

End Class