Imports System.Data.OleDb
Public Class frmCashier
    Dim provider As String
    Dim datafile As String
    Dim constring As String
    Dim myconnection As OleDbConnection = New OleDbConnection

    Dim TextToPrint As String = ""
    Dim cmode As String
    Dim cmode1 As String
    Dim bmod As String
    Dim mode As String
    Dim mode1 As String
    Dim price As Double
    Dim choice As Integer
    Dim finalprice As Double
    Dim u As String
    Dim a As Double
    Dim daily_from As String
    Dim daily_to As String
    Public imagevariable As String
    Dim WithEvents PDB As New ToolStripButton("Confirm Payment")
    Dim WithEvents PDBB As New ToolStripButton("Confirm Payment")
    Private Sub frmCashier_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        provider = "provider = microsoft.ace.oledb.12.0; data source="
        datafile = " C:\Users\CHEKAZ\Documents\railwayticketbooking.accdb"
        constring = provider & datafile
        myconnection.ConnectionString = constring

        insertimage()
        checkpassenger()
        checkpassenger1()
        ticketnum()
        ticketnum1()
        CType(PrintPreviewDialog1.Controls(1), ToolStrip).Items.Add(PDB)
        CType(PrintPreviewDialog2.Controls(1), ToolStrip).Items.Add(PDBB)
        Label8.Text = Format(Now, "d/M/yyyy")
        Dim sql1 As String
        sql1 = "SELECT * FROM dailybookings"
        Dim adapter1 As New OleDbDataAdapter(sql1, myconnection)
        Dim cmd1 As New OleDbCommand(sql1, myconnection)
        Dim dt1 As New DataTable("ticket")
        adapter1.Fill(dt1)
        DataGridView3.DataSource = dt1

        Dim sql As String
        sql = "SELECT * FROM dailybookings where tlabel= '" & 1 & "'"
        Dim adapter As New OleDbDataAdapter(sql, myconnection)
        Dim cmd As New OleDbCommand(sql, myconnection)
        Dim dt As New DataTable("ticket")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt

        Dim sql2 As String
        sql2 = "SELECT * FROM dailybookings where tlabel= '" & 2 & "'"
        Dim adapter2 As New OleDbDataAdapter(sql2, myconnection)
        Dim cmd2 As New OleDbCommand(sql2, myconnection)
        Dim dt2 As New DataTable("ticket")
        adapter2.Fill(dt2)
        DataGridView2.DataSource = dt2
    End Sub
    Private Sub insertimage()
        Dim cnn As New OleDb.OleDbConnection
        cnn = New OleDb.OleDbConnection
        cnn.ConnectionString = "Provider=Microsoft.ace.OLEDB.12.0;Data Source=C:\Users\CHEKAZ\Documents\railwayticketbooking.accdb;"
        If cnn.State = ConnectionState.Open Then
            cnn.Open()
        End If
        Dim da As New OleDb.OleDbDataAdapter("SELECT * FROM cashier WHERE username='" & imagevariable & "'", cnn)
        Dim dt As New DataTable
        da.Fill(dt)
        Try
            Dim ad As String = CStr(dt.Rows(0).Item("PicturePath"))
            cnn.Close()
            PictureBox1.Image = System.Drawing.Bitmap.FromFile(ad)
        Catch ex As Exception

        End Try
    End Sub
    Private Sub ticketnum()
        provider = "provider = microsoft.ace.oledb.12.0; data source="
        datafile = " C:\Users\CHEKAZ\Documents\railwayticketbooking.accdb"
        constring = provider & datafile
        myconnection.ConnectionString = constring

        myconnection.Open()
        Dim READER
        Dim Query As String
        Dim trainname As Integer = 0
        Query = "Select * from dailybookings WHERE ticket_no='" & trainname & "' and tlabel='" & 1 & "'"
        Dim Commandd = New OleDbCommand(Query, myconnection)
        READER = Commandd.ExecuteReader

        Dim found As Boolean = False
        While READER.Read
            found = True
        End While

        If found = True Then
            Dim cmd As New OleDbCommand
            cmd.CommandType = CommandType.Text
            cmd.Connection = myconnection
            cmd.CommandText = "Select Max (ticket_no) from [dailybookings] where tlabel= '" & 1 & "'"
            Dim aa As Integer = cmd.ExecuteScalar
            myconnection.Close()
            Dim temp As Integer = aa + 1
            TextBox7.Text = temp
            If temp = Nothing Then

            End If
        Else
            Try
                Dim sqlinsert As String
                sqlinsert = "INSERT INTO dailybookings([train_name],[train_type], [ticket_no], [name], [surname], [age_range], [phone_number], [from_train], [to_train], [price], [date_train], [time], [tlabel])" _
              & "VALUES(@train_name, @train_type, @ticket_no, @name, @surname, @age_range, @phone_number, @from_train, @to_train, @price, @date_train, @time, @tlabel)"

                Dim cmd As New OleDbCommand(sqlinsert, myconnection)
                cmd.Parameters.Add(New OleDbParameter("@train_name", "NNRZ"))
                cmd.Parameters.Add(New OleDbParameter("@train_type", "Economic"))
                cmd.Parameters.Add(New OleDbParameter("@ticket_no", 0))
                cmd.Parameters.Add(New OleDbParameter("@name", "lotus"))
                cmd.Parameters.Add(New OleDbParameter("@surname", "lotus"))
                cmd.Parameters.Add(New OleDbParameter("@age_range", "Adult"))
                cmd.Parameters.Add(New OleDbParameter("@phone_number", "0777147424"))
                cmd.Parameters.Add(New OleDbParameter("@from_train", "Harare"))
                cmd.Parameters.Add(New OleDbParameter("@to_train", "Harare"))
                cmd.Parameters.Add(New OleDbParameter("@price", 00))
                cmd.Parameters.Add(New OleDbParameter("@date_train", "05/08/18"))
                cmd.Parameters.Add(New OleDbParameter("@time", Format(Now, "hh:mm:ss tt")))
                cmd.Parameters.Add(New OleDbParameter("@tlabel", 1))
                cmd.ExecuteNonQuery()
                myconnection.Close()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                myconnection.Close()
            End Try
        End If
    End Sub
    Private Sub ticketnum1()
        provider = "provider = microsoft.ace.oledb.12.0; data source="
        datafile = " C:\Users\CHEKAZ\Documents\railwayticketbooking.accdb"
        constring = provider & datafile
        myconnection.ConnectionString = constring

        myconnection.Open()
        Dim READER
        Dim Query As String
        Dim trainname As Integer = 200000
        Query = "Select * from dailybookings WHERE ticket_no='" & trainname & "' and tlabel='" & 2 & "'"
        Dim Commandd = New OleDbCommand(Query, myconnection)
        READER = Commandd.ExecuteReader

        Dim found As Boolean = False
        While READER.Read
            found = True
        End While

        If found = True Then
            Dim cmd As New OleDbCommand
            cmd.CommandType = CommandType.Text
            cmd.Connection = myconnection
            cmd.CommandText = "Select Max (ticket_no) from [dailybookings] where tlabel= '" & 2 & "'"
            Dim aa As Integer = cmd.ExecuteScalar
            myconnection.Close()
            Dim temp As Integer = aa + 1
            TextBox10.Text = temp
            If temp = Nothing Then

            End If
        Else
            Try
                Dim sqlinsert As String
                sqlinsert = "INSERT INTO dailybookings([train_name],[train_type], [ticket_no], [name], [surname], [age_range], [phone_number], [from_train], [to_train], [price], [date_train], [time], [tlabel])" _
              & "VALUES(@train_name, @train_type, @ticket_no, @name, @surname, @age_range, @phone_number, @from_train, @to_train, @price, @date_train, @time, @tlabel)"

                Dim cmd As New OleDbCommand(sqlinsert, myconnection)
                cmd.Parameters.Add(New OleDbParameter("@train_name", "NbRZ"))
                cmd.Parameters.Add(New OleDbParameter("@train_type", "standard"))
                cmd.Parameters.Add(New OleDbParameter("@ticket_no", 200000))
                cmd.Parameters.Add(New OleDbParameter("@name", "atias"))
                cmd.Parameters.Add(New OleDbParameter("@surname", "atias"))
                cmd.Parameters.Add(New OleDbParameter("@age_range", "kids"))
                cmd.Parameters.Add(New OleDbParameter("@phone_number", "0773825486"))
                cmd.Parameters.Add(New OleDbParameter("@from_train", "Harare"))
                cmd.Parameters.Add(New OleDbParameter("@to_train", "mufakose"))
                cmd.Parameters.Add(New OleDbParameter("@price", 01))
                cmd.Parameters.Add(New OleDbParameter("@date_train", "05/08/18"))
                cmd.Parameters.Add(New OleDbParameter("@time", Format(Now, "hh:mm:ss tt")))
                cmd.Parameters.Add(New OleDbParameter("@tlabel", 2))
                cmd.ExecuteNonQuery()
                myconnection.Close()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                myconnection.Close()
            End Try
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        mode = ComboBox1.SelectedItem
        Label10.Text = ComboBox1.SelectedItem
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Label5.Text = ComboBox3.SelectedItem
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        mode1 = ComboBox4.SelectedItem
        If mode1 = "Bulawayo" Then
            a = 8
            TextBox9.Text = "BNRZ20"
        ElseIf mode1 = "Chinhoyi" Then
            a = 4
            TextBox9.Text = "CNRZ20"
        ElseIf mode1 = "Mutare" Then
            a = 7
            TextBox9.Text = "MNRZ20"
        ElseIf mode1 = "Shamva" Then
            a = 11
            TextBox9.Text = "SNRZ20"
        End If
        Label6.Text = ComboBox4.SelectedItem
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        Label9.Text = TextBox7.Text
    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        Label32.Text = ComboBox7.SelectedItem
    End Sub

    Private Sub Label23_Click(sender As Object, e As EventArgs) Handles Label23.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If mode = "Adults (16 & above)" Then
            If mode1 = "Bulawayo" Then
                Label23.Text = a
            ElseIf mode1 = "Chinhoyi" Then
                Label23.Text = a
            ElseIf mode1 = "Mutare" Then
                Label23.Text = a
            ElseIf mode1 = "Shamva" Then
                Label23.Text = a
            End If
        ElseIf mode = "Kids (15 & below)" Then
            If mode1 = "Bulawayo" Then
                Label23.Text = a * 0.5
            ElseIf mode1 = "Chinhoyi" Then
                Label23.Text = a * 0.5
            ElseIf mode1 = "Mutare" Then
                Label23.Text = a * 0.5
            ElseIf mode1 = "Shamva" Then
                Label23.Text = a * 0.5
            End If
        End If
    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub
#Region "Code for print preview"
    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Static currentChar As Integer
        Dim textfont As Font = New Font("Courier New", 10, FontStyle.Bold)

        Dim h, w As Integer
        Dim left, top As Integer
        With PrintDocument1.DefaultPageSettings
            h = 0
            w = 0
            left = 0
            top = 0
        End With

        Dim lines As Integer = CInt(Math.Round(h / 1))
        Dim b As New Rectangle(left, top, w, h)
        Dim format As StringFormat
        format = New StringFormat(StringFormatFlags.LineLimit)
        Dim line, chars As Integer

        e.Graphics.MeasureString(Mid(TextToPrint, currentChar + 1), textfont, New SizeF(w, h), format, chars, line)
        e.Graphics.DrawString(TextToPrint.Substring(currentChar, chars), New Font("Courier New", 10, FontStyle.Bold), Brushes.Black, b, format)

        currentChar = currentChar + chars
        If currentChar < TextToPrint.Length Then
            e.HasMorePages = True
        Else
            e.HasMorePages = False
            currentChar = 0
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        PrintHeader()
        ItemsToBePrinted()
        printFooter()
        Dim printControl = New Printing.StandardPrintController
        PrintDocument1.PrintController = printControl
        Try
            PrintPreviewDialog1.Document = PrintDocument1
            PrintPreviewDialog1.ShowDialog()
            'PrintDocument1.Print()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub PrintHeader()

        TextToPrint = ""
        'send Business Name
        Dim StringToPrint As String = "Business Name"
        Dim LineLen As Integer = StringToPrint.Length
        Dim spcLen1 As New String(" "c, Math.Round((33 - LineLen) / 2)) 'This line is used to center text in the middle of the receipt
        TextToPrint &= spcLen1 & StringToPrint & Environment.NewLine

        'send address name
        StringToPrint = "12345 Street Avenue"
        LineLen = StringToPrint.Length
        Dim spcLen2 As New String(" "c, Math.Round((33 - LineLen) / 2))
        TextToPrint &= spcLen2 & StringToPrint & Environment.NewLine

        ' send city, state, zip
        StringToPrint = "City, State, Zip code"
        LineLen = StringToPrint.Length
        Dim spcLen3 As New String(" "c, Math.Round((33 - LineLen) / 2))
        TextToPrint &= spcLen3 & StringToPrint & Environment.NewLine

        ' send phone number
        StringToPrint = "999-999-9999"
        LineLen = StringToPrint.Length
        Dim spcLen4 As New String(" "c, Math.Round((33 - LineLen) / 2))
        TextToPrint &= spcLen4 & StringToPrint & Environment.NewLine

        'send website
        StringToPrint = "website.com"
        LineLen = StringToPrint.Length
        Dim spcLen4b As New String(" "c, Math.Round((33 - LineLen) / 2))
        TextToPrint &= spcLen4b & StringToPrint & Environment.NewLine

    End Sub

    Public Sub ItemsToBePrinted()
        Dim l As String = TextBox9.Text
        Dim a As String = ComboBox7.SelectedItem
        Dim b As String = TextBox7.Text
        Dim k As String = TextBox1.Text
        Dim d As String = TextBox2.Text
        Dim e As String = ComboBox1.SelectedItem
        Dim f As String = TextBox3.Text
        Dim g As String = ComboBox3.SelectedItem
        Dim h As String = ComboBox4.SelectedItem
        Dim j As String = Label23.Text
        '  cmd.Parameters.Add(New OleDbParameter("@date", Label8.Text))
        '  cmd.Parameters.Add(New OleDbParameter("@time", Format(Now, "hh:mm:ss tt")))

        Dim StringToPrint As String = "Train Type: " & a
        Dim LineLen As String = StringToPrint.Length
        Dim spcLen5 As New String(" "c, Math.Round((30 - LineLen)))
        TextToPrint &= "Description" & Environment.NewLine
        TextToPrint &= "Train Name: " & l & Environment.NewLine
        TextToPrint &= StringToPrint & Environment.NewLine
        TextToPrint &= "ticket No#: " & b & Environment.NewLine
        TextToPrint &= "Name      : " & k & Environment.NewLine
        TextToPrint &= "Surname   : " & d & Environment.NewLine
        TextToPrint &= "Age Range : " & e & Environment.NewLine
        TextToPrint &= "Phone No# : " & f & Environment.NewLine
        TextToPrint &= "From      : " & g & Environment.NewLine
        TextToPrint &= "To        : " & h & Environment.NewLine
        TextToPrint &= "Price     : " & j & Environment.NewLine

    End Sub
    Public Sub ItemsToBePrinted1()
        Dim l As String = TextBox8.Text
        Dim a As String = TextBox11.Text
        Dim b As String = TextBox10.Text
        Dim k As String = TextBox6.Text
        Dim d As String = TextBox5.Text
        Dim e As String = ComboBox6.SelectedItem
        Dim f As String = TextBox4.Text
        Dim g As String = daily_from
        Dim h As String = daily_to
        Dim j As String = Label62.Text

        Dim StringToPrint As String = "Train Type: " & a
        Dim LineLen As String = StringToPrint.Length
        Dim spcLen5 As New String(" "c, Math.Round((30 - LineLen)))
        TextToPrint &= "Description" & Environment.NewLine
        TextToPrint &= "Train Name: " & l & Environment.NewLine
        TextToPrint &= StringToPrint & Environment.NewLine
        TextToPrint &= "ticket No#: " & b & Environment.NewLine
        TextToPrint &= "Name      : " & k & Environment.NewLine
        TextToPrint &= "Surname   : " & d & Environment.NewLine
        TextToPrint &= "Age Range : " & e & Environment.NewLine
        TextToPrint &= "Phone No# : " & f & Environment.NewLine
        TextToPrint &= "From      : " & g & Environment.NewLine
        TextToPrint &= "To        : " & h & Environment.NewLine
        TextToPrint &= "Price     : " & j & Environment.NewLine

    End Sub

    Public Sub printFooter()
        TextToPrint &= Environment.NewLine & Environment.NewLine
        Dim globalLengt As Integer = 0

        'SubTotal Amount
        Dim StringToPrint As String = "Sub Total   " & FormatCurrency("3.99", , , TriState.True, TriState.True)  'Change here to subtotal
        Dim LineLen As String = StringToPrint.Length
        globalLengt = StringToPrint.Length
        Dim spcLen5 As New String(" "c, Math.Round((26 - LineLen)))
        TextToPrint &= Environment.NewLine & spcLen5 & StringToPrint & Environment.NewLine

        'Tax Amount
        StringToPrint = "Tax         " & FormatCurrency("0.05", , , TriState.True, TriState.True) 'Change to tax amount
        LineLen = globalLengt
        Dim spcLen6 As New String(" "c, Math.Round((26 - LineLen)))
        If Not StringToPrint = "Tax         $0.00" Then
            TextToPrint &= spcLen6 & StringToPrint & Environment.NewLine
        End If

        'Total Amount
        StringToPrint = "Total       " & "$4.04"
        LineLen = globalLengt
        Dim spcLen8 As New String(" "c, Math.Round((26 - LineLen)))
        TextToPrint &= spcLen8 & StringToPrint & Environment.NewLine & Environment.NewLine

        'Cash Entered Amount
        StringToPrint = "Cash        " & FormatCurrency("5.00", , , TriState.True, TriState.True)
        LineLen = globalLengt
        Dim spcLen9 As New String(" "c, Math.Round((26 - LineLen)))
        If Not StringToPrint = "Cash        $0.00" Then
            TextToPrint &= spcLen9 & StringToPrint & Environment.NewLine
        End If

        'Change Amount
        StringToPrint = "Change      " & FormatCurrency("0.96", , , TriState.True, TriState.True)
        LineLen = globalLengt
        Dim spcLen10 As New String(" "c, Math.Round((26 - LineLen)))
        TextToPrint &= Environment.NewLine & spcLen10 & StringToPrint & Environment.NewLine
    End Sub
    Private Sub PDB_Click1(ByVal sender As Object, ByVal e As EventArgs) Handles PDB.Click
        myconnection.ConnectionString = constring
        Dim READER As OleDbDataReader
        Dim tltpass As Integer = 2
        Try
            myconnection.Open()
            Dim Query As String
            Dim trainname As String = TextBox9.Text
            Dim datt As String = d_equiry
            Query = "Select * from dailybookings WHERE train_name='" & trainname & "' AND date_train='" & datt & "'"
            Dim Commandd = New OleDbCommand(Query, myconnection)
            READER = Commandd.ExecuteReader
            Dim count As Integer
            count = 0
            While READER.Read
                count = count + 1
            End While
            If count < tltpass Then
                Try
                    Dim sqlinsert As String
                    sqlinsert = "INSERT INTO dailybookings([train_name],[train_type], [ticket_no], [name], [surname], [age_range], [phone_number], [from_train], [to_train], [price], [date_train], [time], [tlabel])" _
                      & "VALUES(@train_name, @train_type, @ticket_no, @name, @surname, @age_range, @phone_number, @from_train, @to_train, @price, @date_train, @time, @tlabel)"

                    Dim cmd As New OleDbCommand(sqlinsert, myconnection)
                    cmd.Parameters.Add(New OleDbParameter("@train_name", TextBox9.Text))
                    cmd.Parameters.Add(New OleDbParameter("@train_type", ComboBox7.SelectedItem))
                    cmd.Parameters.Add(New OleDbParameter("@ticket_no", TextBox7.Text))
                    cmd.Parameters.Add(New OleDbParameter("@name", TextBox1.Text))
                    cmd.Parameters.Add(New OleDbParameter("@surname", TextBox2.Text))
                    cmd.Parameters.Add(New OleDbParameter("@age_range", ComboBox1.SelectedItem))
                    cmd.Parameters.Add(New OleDbParameter("@phone_number", TextBox3.Text))
                    cmd.Parameters.Add(New OleDbParameter("@from_train", ComboBox3.SelectedItem))
                    cmd.Parameters.Add(New OleDbParameter("@to_train", ComboBox4.SelectedItem))
                    cmd.Parameters.Add(New OleDbParameter("@price", Label23.Text))
                    cmd.Parameters.Add(New OleDbParameter("@date_train", Label8.Text))
                    cmd.Parameters.Add(New OleDbParameter("@time", Format(Now, "hh:mm:ss tt")))
                    cmd.Parameters.Add(New OleDbParameter("@tlabel", 1))
                    cmd.ExecuteNonQuery()
                    MsgBox("You Have Successfully Paid")
                    RefreshDGVCashierRecords()
                    myconnection.Close()
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                    myconnection.Close()
                End Try
                clearcontrols()
                d_equiry = Label8.Text
                checkpassenger()
                ticketnum()
            Else
                MsgBox("Please the train is now full you cant add more customers")
                myconnection.Close()
                clearcontrols()
                ticketnum()
            End If
            myconnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            myconnection.Close()
        Finally
            myconnection.Dispose()
        End Try
        PrintPreviewDialog1.Close()
    End Sub
    Private Sub PDBB_Click1(ByVal sender As Object, ByVal e As EventArgs) Handles PDBB.Click
        myconnection.ConnectionString = constring
        Dim READER As OleDbDataReader
        Dim tltpass As Integer = 2
        Try
            myconnection.Open()
            Dim Query As String
            Dim trainname As String = TextBox8.Text
            Dim datt As String = Label8.Text
            Query = "Select * from dailybookings WHERE train_name='" & trainname & "' AND date_train='" & datt & "'"
            Dim Commandd = New OleDbCommand(Query, myconnection)
            READER = Commandd.ExecuteReader
            Dim count As Integer
            count = 0
            While READER.Read
                count = count + 1
            End While
            If count < tltpass Then
                Try
                    Dim sqlinsert As String
                    sqlinsert = "INSERT INTO dailybookings([train_name],[train_type], [ticket_no], [name], [surname], [age_range], [phone_number], [from_train], [to_train], [price], [date_train], [time], [tlabel])" _
                      & "VALUES(@train_name, @train_type, @ticket_no, @name, @surname, @age_range, @phone_number, @from_train, @to_train, @price, @date_train, @time, @tlabel)"

                    Dim cmd As New OleDbCommand(sqlinsert, myconnection)
                    cmd.Parameters.Add(New OleDbParameter("@train_name", TextBox8.Text))
                    cmd.Parameters.Add(New OleDbParameter("@train_type", TextBox11.Text))
                    cmd.Parameters.Add(New OleDbParameter("@ticket_no", TextBox10.Text))
                    cmd.Parameters.Add(New OleDbParameter("@name", TextBox6.Text))
                    cmd.Parameters.Add(New OleDbParameter("@surname", TextBox5.Text))
                    cmd.Parameters.Add(New OleDbParameter("@age_range", ComboBox6.SelectedItem))
                    cmd.Parameters.Add(New OleDbParameter("@phone_number", TextBox4.Text))
                    cmd.Parameters.Add(New OleDbParameter("@from_train", daily_from))
                    cmd.Parameters.Add(New OleDbParameter("@to_train", daily_to))
                    cmd.Parameters.Add(New OleDbParameter("@price", Label62.Text))
                    cmd.Parameters.Add(New OleDbParameter("@date_train", Format(Now, "d/M/yyyy")))
                    cmd.Parameters.Add(New OleDbParameter("@time", Format(Now, "hh:mm:ss tt")))
                    cmd.Parameters.Add(New OleDbParameter("@tlabel", 2))
                    cmd.ExecuteNonQuery()
                    MsgBox("You Have Successfully Paid")
                    RefreshDGVCashierRecords()
                    myconnection.Close()
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                    myconnection.Close()
                End Try
                clearcontrols1()
                checkpassenger1()
                ticketnum1()
            Else
                MsgBox("Please the train is now full you cant add more customers")
                myconnection.Close()
                clearcontrols1()
                ticketnum1()
            End If
            myconnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            myconnection.Close()
        Finally
            myconnection.Dispose()
        End Try
        PrintPreviewDialog2.Close()
    End Sub
#End Region
    Private Sub MonthCalendar1_DateChanged(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar1.DateChanged
        myconnection.ConnectionString = constring

        Label8.Text = MonthCalendar1.SelectionStart.ToShortDateString
        Dim sqlsearch As String
        sqlsearch = "SELECT * FROM dailybookings WHERE date_train LIKE '%" & Label8.Text & "%' and tlabel= '" & 1 & "'"
        ' Once again we execute the SQL statements against our DataBase
        Dim adapter As New OleDbDataAdapter(sqlsearch, myconnection)
        ' Shows the records and updates the DataGridView
        Dim dt As New DataTable("dailybookings")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        d_equiry = Label8.Text
        checkpassenger()

    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged
        myconnection.ConnectionString = constring
        Label8.Text = MonthCalendar1.SelectionStart.ToShortDateString
        Dim sqlsearch As String
        sqlsearch = "SELECT * FROM dailybookings WHERE ticket_no LIKE '%" & TextBox12.Text & "%' or name LIKE '%" & TextBox12.Text & "%' or surname LIKE '%" & TextBox12.Text & "%'"
        ' Once again we execute the SQL statements against our DataBase
        Dim adapter As New OleDbDataAdapter(sqlsearch, myconnection)
        ' Shows the records and updates the DataGridView
        Dim dt As New DataTable("dailybookings")
        adapter.Fill(dt)
        DataGridView3.DataSource = dt
    End Sub

    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        Try
            TextBox17.Text = DataGridView3.Rows(e.RowIndex).Cells(0).Value.ToString
            TextBox20.Text = DataGridView3.Rows(e.RowIndex).Cells(1).Value.ToString
            TextBox15.Text = DataGridView3.Rows(e.RowIndex).Cells(2).Value.ToString
            TextBox16.Text = DataGridView3.Rows(e.RowIndex).Cells(3).Value.ToString
            TextBox14.Text = DataGridView3.Rows(e.RowIndex).Cells(4).Value.ToString
            ComboBox2.Text = DataGridView3.Rows(e.RowIndex).Cells(5).Value.ToString
            TextBox13.Text = DataGridView3.Rows(e.RowIndex).Cells(6).Value.ToString
            TextBox18.Text = DataGridView3.Rows(e.RowIndex).Cells(7).Value.ToString
            ComboBox5.Text = DataGridView3.Rows(e.RowIndex).Cells(8).Value.ToString
            TextBox22.Text = DataGridView3.Rows(e.RowIndex).Cells(9).Value.ToString
            TextBox23.Text = DataGridView3.Rows(e.RowIndex).Cells(10).Value.ToString
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Try
            ' This is our DELETE Statement. To be sure we delete the correct record and not all of 
            ' them.
            ' We use the WHERE to be sure only that record that the user has selected is deleted.
            Dim sqldelete As String
            sqldelete = "DELETE * FROM dailybookings WHERE ticket_no='" & DataGridView3.CurrentRow.Cells(2).Value.ToString & "'"

            ' This is our DataAdapter. This executes our SQL Statement above against the Database
            ' we defined in the Connection String
            Dim adapter As New OleDbDataAdapter(sqldelete, myconnection)
            ' Gets the records from the table and fills our adapter with those.
            Dim dt As New DataTable("dailybookings")
            adapter.Fill(dt)
            ' Assigns the edited DataSource on the DataGridView and the refreshes the 
            ' view to ensure everything is up to date in real time.
            DataGridView3.DataSource = dt
            ' This is a Sub in Module 1 to refresh the DataGridView when information is added,
            '  updated, or deleted.
            RefreshDGVCashierRecords()
        Catch ex As Exception
            MsgBox(ex.Message & "or There is nothing to Delete")

        End Try
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim sqlupdate As String
        ' Here we use the UPDATE Statement to update the information. To be sure we are 
        ' updating the right record we also use the WHERE clause to be sureno information
        ' is added or changed in the other records
        'sqlupdate = "UPDATE dailybookings SET train_name=@train_name, train_type=@train_type, name=@name," _
        '    & "surname=@surname, age_range=@age_range, phone_number=@phone_number, from=@from, to_train=@to_train, price=@price, date_train=@date_train WHERE [ticket_no]='" & TextBox15.Text & "'"
        sqlupdate = "UPDATE dailybookings SET train_name=@train_name, train_type=@train_type, name=@name," _
            & "surname=@surname, age_range=@age_range, phone_number=@phone_number, to_train=@to_train, price=@price, date_train=@date_train WHERE [ticket_no]='" & TextBox15.Text & "'"
        '  sqlupdate = "UPDATE dailybookings SET from_train=@from_train WHERE [ticket_no]='" & TextBox15.Text & "'"

        Dim cmd As New OleDbCommand(sqlupdate, myconnection)
        'This assigns the values for our columns in the DataBase. 
        ' To ensure the correct values are written to the correct column
        cmd.Parameters.Add(New OleDbParameter("@train_name", TextBox17.Text))
        cmd.Parameters.Add(New OleDbParameter("@train_type", TextBox20.Text))
        cmd.Parameters.Add(New OleDbParameter("@name", TextBox16.Text))
        cmd.Parameters.Add(New OleDbParameter("@surname", TextBox14.Text))
        cmd.Parameters.Add(New OleDbParameter("@age_range", ComboBox2.SelectedItem))
        cmd.Parameters.Add(New OleDbParameter("@phone_number", TextBox13.Text))
        cmd.Parameters.Add(New OleDbParameter("@to_train", ComboBox5.SelectedItem))
        cmd.Parameters.Add(New OleDbParameter("@price", TextBox22.Text))
        cmd.Parameters.Add(New OleDbParameter("@date_train", TextBox23.Text))
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
        Button8.Enabled = False
        RefreshDGVCashierRecords()
        TextBox17.Clear()
        TextBox20.Clear()
        TextBox15.Clear()
        TextBox16.Clear()
        TextBox14.Clear()
        ComboBox2.SelectedItem = Nothing
        TextBox13.Clear()
        TextBox18.Clear()
        ComboBox5.SelectedItem = Nothing
        TextBox22.Clear()
        TextBox23.Clear()
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        cmode1 = ComboBox5.SelectedItem
        If cmode1 = "Bulawayo" Then
            a = 8
            TextBox17.Text = "BNRZ20"
        ElseIf cmode1 = "Chinhoyi" Then
            a = 4
            TextBox17.Text = "CNRZ20"
        ElseIf cmode1 = "Mutare" Then
            a = 7
            TextBox17.Text = "MNRZ20"
        ElseIf cmode1 = "Shamva" Then
            a = 11
            TextBox17.Text = "SNRZ20"
        End If
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        cmode = ComboBox2.SelectedItem

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        FrmLogin.Show()
        Me.Hide()
        myconnection.Close()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If cmode = "Adults (16 & above)" Then
            If cmode1 = "Bulawayo" Then
                TextBox22.Text = a
            ElseIf cmode1 = "Chinhoyi" Then
                TextBox22.Text = a
            ElseIf cmode1 = "Mutare" Then
                TextBox22.Text = a
            ElseIf cmode1 = "Shamva" Then
                TextBox22.Text = a
            End If
        ElseIf cmode = "Kids (15 & below)" Then
            If cmode1 = "Bulawayo" Then
                TextBox22.Text = a * 0.5
            ElseIf cmode1 = "Chinhoyi" Then
                TextBox22.Text = a * 0.5
            ElseIf cmode1 = "Mutare" Then
                TextBox22.Text = a * 0.5
            ElseIf cmode1 = "Shamva" Then
                TextBox22.Text = a * 0.5
            End If
        End If
        Button8.Enabled = True
    End Sub
    Private Sub TabControl1_TabIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.TabIndexChanged
        ticketnum()
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        bmod = ComboBox6.SelectedItem
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            u = 1
        End If
        daily_from = "CBD"
        daily_to = "Mufakose"
        TextBox8.Text = "MNRZ"
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            u = 1
        End If
        daily_from = "CBD"
        daily_to = "Wadzana"
        TextBox8.Text = "WNRZ"
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If bmod = "Adults (16 & above)" Then
            Label62.Text = u
        ElseIf bmod = "Kids (15 & below)" Then
            Label62.Text = u * 0.5
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        PrintHeader()
        ItemsToBePrinted1()
        printFooter()
        Dim printControl = New Printing.StandardPrintController
        PrintDocument1.PrintController = printControl
        Try
            PrintPreviewDialog2.Document = PrintDocument1
            PrintPreviewDialog2.ShowDialog()
            'PrintDocument1.Print()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

End Class