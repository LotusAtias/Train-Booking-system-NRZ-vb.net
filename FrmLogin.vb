Imports System.Data.OleDb
Public Class FrmLogin
    Dim provider As String
    Dim datafile As String
    Dim constring As String
    Dim myconnection As OleDbConnection = New OleDbConnection
    Private Sub FrmLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        provider = "provider = microsoft.ace.oledb.12.0; data source="
        datafile = " C:\Users\CHEKAZ\Documents\railwayticketbooking.accdb"
        constring = provider & datafile
        myconnection.ConnectionString = constring
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        myconnection.Open()
        Dim cmd As OleDbCommand = New OleDbCommand(" select * from [cashier] where [username] = '" & TextBox1.Text & "' and [pwsword]= '" & TextBox2.Text & "'", myconnection)
        Dim dr As OleDbDataReader = cmd.ExecuteReader
        Dim userfound As Boolean = False
        While dr.Read
            userfound = True
            frmCashier.imagevariable = dr("username")
            frmCashier.Label34.Text = dr("name") & "  " & dr("surname")
        End While
        If userfound = True Then
            Me.Hide()
            frmCashier.Show()
            TextBox1.Clear()
            TextBox2.Clear()
            Label4.Text = ""
            myconnection.Close()
        Else
            Label4.Text = "username or password incorrect"
            TextBox1.Clear()
            TextBox2.Clear()
            myconnection.Close()
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        myconnection.Open()
        Dim cmd As OleDbCommand = New OleDbCommand(" select * from [officerincharge] where [username] = '" & TextBox1.Text & "' and [pwsword]= '" & TextBox2.Text & "'", myconnection)
        Dim dr As OleDbDataReader = cmd.ExecuteReader
        Dim userfound As Boolean = False
        While dr.Read
            userfound = True
            FrmOfficerIncharge.imagevariablee = dr("username")
            FrmOfficerIncharge.Label34.Text = dr("name") & "  " & dr("surname")
        End While

        If userfound = True Then
            Me.Hide()
            FrmOfficerIncharge.Show()
            TextBox1.Clear()
            TextBox2.Clear()
            Label4.Text = ""
            myconnection.Close()
        Else
            Label4.Text = "username or password incorrect "
            TextBox1.Clear()
            TextBox2.Clear()
            myconnection.Close()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        myconnection.Open()
        Dim cmd As OleDbCommand = New OleDbCommand(" select * from [admin] where [username] = '" & TextBox1.Text & "' and [password]= '" & TextBox2.Text & "'", myconnection)
        Dim dr As OleDbDataReader = cmd.ExecuteReader
        Dim userfound As Boolean = False
        While dr.Read
            userfound = True
        End While
        If userfound = True Then
            Me.Hide()
            FrmAdmin.Show()
            TextBox1.Clear()
            TextBox2.Clear()
            Label4.Text = ""
            myconnection.Close()
        Else
            Label4.Text = "username or password incorrect "
            TextBox1.Clear()
            TextBox2.Clear()
            myconnection.Close()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox1.Clear()
        TextBox2.Clear()
        Label4.Text = ""
    End Sub
End Class
