Imports System.Data.OleDb
Public Class Form1


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sno As Integer
        Dim na As String
        Dim da As String
        sno = TextBox1.Text
        na = TextBox2.Text
        da = TextBox3.Text

        Dim connetionString As String
        Dim connection As OleDbConnection
        Dim oledbAdapter As New OleDbDataAdapter
        Dim sql As String
        connetionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Database.accdb;"
        connection = New OleDbConnection(connetionString)
        sql = "insert into user emp('& sno & ',' & na &',' & DateTimePicker1.value & ',' & da & ')"
        Try
            connection.Open()
            oledbAdapter.InsertCommand = New OleDbCommand(sql, connection)
            oledbAdapter.InsertCommand.ExecuteNonQuery()
            MsgBox("Row(s) Inserted !! ")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
End Class
