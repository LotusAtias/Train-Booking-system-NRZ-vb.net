Imports System.Data.OleDb

Module Module1
    Dim provider As String
    Dim source As String
    Dim constring As String
    Dim myconnection As OleDbConnection = New OleDbConnection
    Public d_equiry As String

    Sub checkpassenger()
        provider = "provider = microsoft.ace.oledb.12.0; data source="
        source = "C:\Users\CHEKAZ\Documents\railwayticketbooking.accdb"
        constring = provider & source
        myconnection.ConnectionString = constring
        Dim a As String = "BNRZ20"
        Dim b As String = "CNRZ20"
        Dim c As String = "MNRZ20"
        Dim d As String = "SNRZ20"
#Region "one"
        myconnection.ConnectionString = constring
        Dim READER As OleDbDataReader
        Dim tltpass As Integer = 2
        Try
            myconnection.Open()
            Dim Query As String

            Query = "Select * from dailybookings WHERE train_name='" & a & "' AND date_train='" & d_equiry & "' "
            Dim Commandd = New OleDbCommand(Query, myconnection)
            READER = Commandd.ExecuteReader
            Dim count As Integer
            count = 0
            While READER.Read
                count = count + 1
            End While
            frmCashier.Label42.Text = tltpass
            frmCashier.Label37.Text = count
            frmCashier.Label39.Text = tltpass - count
            myconnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            myconnection.Close()
        Finally
            myconnection.Dispose()
        End Try
#End Region
#Region "sec"
        myconnection.ConnectionString = constring
        Dim READER1 As OleDbDataReader
        Dim tltpass1 As Integer = 2
        Try
            myconnection.Open()
            Dim Query1 As String

            Query1 = "Select * from dailybookings WHERE train_name='" & b & "' AND date_train='" & d_equiry & "' "
            Dim Commandd1 = New OleDbCommand(Query1, myconnection)
            READER1 = Commandd1.ExecuteReader
            Dim count1 As Integer
            count1 = 0
            While READER1.Read
                count1 = count1 + 1
            End While
            frmCashier.Label44.Text = tltpass1
            frmCashier.Label48.Text = count1
            frmCashier.Label46.Text = tltpass1 - count1
            myconnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            myconnection.Close()
        Finally
            myconnection.Dispose()
        End Try
#End Region
#Region "thir"
        myconnection.ConnectionString = constring
        Dim READER2 As OleDbDataReader
        Dim tltpass2 As Integer = 2
        Try
            myconnection.Open()
            Dim Query2 As String

            Query2 = "Select * from dailybookings WHERE train_name='" & c & "' AND date_train='" & d_equiry & "' "
            Dim Commandd2 = New OleDbCommand(Query2, myconnection)
            READER2 = Commandd2.ExecuteReader
            Dim count2 As Integer
            count2 = 0
            While READER2.Read
                count2 = count2 + 1
            End While
            frmCashier.Label56.Text = tltpass2
            frmCashier.Label60.Text = count2
            frmCashier.Label58.Text = tltpass2 - count2
            myconnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            myconnection.Close()
        Finally
            myconnection.Dispose()
        End Try
#End Region
#Region "thur"
        myconnection.ConnectionString = constring
        Dim READER3 As OleDbDataReader
        Dim tltpass3 As Integer = 2
        Try
            myconnection.Open()
            Dim Query3 As String

            Query3 = "Select * from dailybookings WHERE train_name='" & d & "' AND date_train='" & d_equiry & "' "
            Dim Commandd3 = New OleDbCommand(Query3, myconnection)
            READER3 = Commandd3.ExecuteReader
            Dim count3 As Integer
            count3 = 0
            While READER3.Read
                count3 = count3 + 1
            End While
            frmCashier.Label50.Text = tltpass3
            frmCashier.Label54.Text = count3
            frmCashier.Label52.Text = tltpass3 - count3
            myconnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            myconnection.Close()
        Finally
            myconnection.Dispose()
        End Try
#End Region
    End Sub
    Sub checkpassenger1()
        provider = "provider = microsoft.ace.oledb.12.0; data source="
        source = "C:\Users\CHEKAZ\Documents\railwayticketbooking.accdb"
        constring = provider & source
        myconnection.ConnectionString = constring
        Dim a As String = "MNRZ"
        Dim b As String = "WNRZ"
#Region "one"
        myconnection.ConnectionString = constring
        Dim READER As OleDbDataReader
        Dim tltpass As Integer = 2
        Try
            myconnection.Open()
            Dim Query As String
            Query = "Select * from dailybookings WHERE train_name='" & a & "' AND date_train='" & Format(Now, "d/M/yyyy") & "' "
            Dim Commandd = New OleDbCommand(Query, myconnection)
            READER = Commandd.ExecuteReader
            Dim count As Integer
            count = 0
            While READER.Read
                count = count + 1
            End While
            frmCashier.Label65.Text = tltpass
            frmCashier.Label69.Text = count
            frmCashier.Label67.Text = tltpass - count
            myconnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            myconnection.Close()
        Finally
            myconnection.Dispose()
        End Try
#End Region
#Region "two"
        myconnection.ConnectionString = constring
        Dim READER1 As OleDbDataReader
        Dim tltpass1 As Integer = 2
        Try
            myconnection.Open()
            Dim Query1 As String

            Query1 = "Select * from dailybookings WHERE train_name='" & b & "' AND date_train='" & Format(Now, "d/M/yyyy") & "' "
            Dim Commandd1 = New OleDbCommand(Query1, myconnection)
            READER1 = Commandd1.ExecuteReader
            Dim count1 As Integer
            count1 = 0
            While READER1.Read
                count1 = count1 + 1
            End While
            frmCashier.Label85.Text = tltpass1
            frmCashier.Label89.Text = count1
            frmCashier.Label87.Text = tltpass1 - count1
            myconnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            myconnection.Close()
        Finally
            myconnection.Dispose()
        End Try
#End Region
    End Sub
    Sub clearcontrols()
        frmCashier.ComboBox7.SelectedItem = Nothing
        frmCashier.TextBox7.Clear()
        frmCashier.TextBox1.Clear()
        frmCashier.TextBox2.Clear()
        frmCashier.ComboBox1.SelectedItem = Nothing
        frmCashier.TextBox3.Clear()
        frmCashier.ComboBox3.SelectedItem = Nothing
        frmCashier.ComboBox4.SelectedItem = Nothing
        frmCashier.TextBox9.Clear()
        frmCashier.Label23.Text = "- - - - - - - - - - - - "
    End Sub
    Sub clearcontrols1()
        frmCashier.TextBox11.Text = "Standard"
        frmCashier.TextBox10.Clear()
        frmCashier.TextBox6.Clear()
        frmCashier.TextBox5.Clear()
        frmCashier.ComboBox6.SelectedItem = Nothing
        frmCashier.TextBox4.Clear()
        frmCashier.TextBox8.Clear()
        frmCashier.RadioButton1.Checked = Nothing
        frmCashier.RadioButton2.Checked = Nothing
        frmCashier.Label62.Text = "- - - - - - - - - - - - "
    End Sub
    Sub checktraincomplete()
        provider = "provider = microsoft.ace.oledb.12.0; data source="
        source = "C:\Users\CHEKAZ\Documents\railwayticketbooking.accdb"
        constring = provider & source
        myconnection.ConnectionString = constring

        Dim READER As OleDbDataReader
        Dim tltpass As Integer = 5
        Try
            myconnection.Open()
            Dim Query As String

            Query = "Select * from dailybookings WHERE date_train='" & Format(Now, "d/MM/yyyy") & "' "
            Dim Commandd = New OleDbCommand(Query, myconnection)
            READER = Commandd.ExecuteReader
            Dim count As Integer
            count = 0
            While READER.Read
                count = count + 1
            End While
            If count <= tltpass Then

            Else


            End If
            myconnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            myconnection.Close()
        Finally
            myconnection.Dispose()

        End Try
    End Sub
    Sub RefreshDGVCashierRecords()
        provider = "provider = microsoft.ace.oledb.12.0; data source="
        source = "C:\Users\CHEKAZ\Documents\railwayticketbooking.accdb"
        constring = provider & source
        myconnection.ConnectionString = constring

        Dim sql As String
        sql = "SELECT * FROM dailybookings"
        Dim adapter As New OleDbDataAdapter(sql, myconnection)
        Dim dt As New DataTable("dailybookings")
        adapter.Fill(dt)
        frmCashier.DataGridView3.DataSource = dt

        Dim sql1 As String
        sql1 = "SELECT * FROM dailybookings where tlabel='" & 1 & "'"
        Dim adapter1 As New OleDbDataAdapter(sql1, myconnection)
        Dim dt1 As New DataTable("dailybookings")
        adapter1.Fill(dt1)
        frmCashier.DataGridView1.DataSource = dt1

        Dim sql2 As String
        sql2 = "SELECT * FROM dailybookings where tlabel='" & 2 & "'"
        Dim adapter2 As New OleDbDataAdapter(sql2, myconnection)
        Dim dt2 As New DataTable("dailybookings")
        adapter2.Fill(dt2)
        frmCashier.DataGridView2.DataSource = dt2
    End Sub
    Sub refreshdgvAdmin()
        provider = "provider = microsoft.ace.oledb.12.0; data source="
        source = "C:\Users\CHEKAZ\Documents\railwayticketbooking.accdb"
        constring = provider & source
        myconnection.ConnectionString = constring

        Dim sql As String
        sql = "SELECT * FROM cashier"
        Dim adapter As New OleDbDataAdapter(sql, myconnection)
        Dim dt As New DataTable("cashier")
        adapter.Fill(dt)
        FrmAdmin.DataGridView1.DataSource = dt

        Dim sql1 As String
        sql1 = "SELECT * FROM officerincharge"
        Dim adapter1 As New OleDbDataAdapter(sql1, myconnection)
        Dim dt1 As New DataTable("officerincharge")
        adapter1.Fill(dt1)
        FrmAdmin.DataGridView2.DataSource = dt1
    End Sub
    Sub refreshofficer()
        provider = "provider = microsoft.ace.oledb.12.0; data source="
        source = "C:\Users\CHEKAZ\Documents\railwayticketbooking.accdb"
        constring = provider & source
        myconnection.ConnectionString = constring

        Dim sql As String
        sql = "SELECT * FROM dailybookings"
        Dim adapter As New OleDbDataAdapter(sql, myconnection)
        Dim dt As New DataTable("cashier")
        adapter.Fill(dt)
        FrmOfficerIncharge.DataGridView1.DataSource = dt

        Dim sql1 As String
        sql1 = "SELECT ticket_no, name, surname, flag FROM dailybookings"
        Dim adapter1 As New OleDbDataAdapter(sql1, myconnection)
        Dim dt1 As New DataTable("officerincharge")
        adapter1.Fill(dt1)
        FrmOfficerIncharge.DataGridView2.DataSource = dt1
    End Sub
End Module
