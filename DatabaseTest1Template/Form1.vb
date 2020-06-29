Imports System.Data.OleDb

Public Class frmTestBed
    Private Sub frmTestBed_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub

    Sub Testlog(ByVal tid As Integer, ByVal testdate As Date,
                ByVal reqmt As String, ByVal Tester As String,
                ByVal result As Boolean)
        Dim cnTestResults As New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;Data Source=TestResultsDB.mdb")
        Dim iRet As Integer
        cnTestResults.Open()
        Dim cmd As New OleDbCommand("Insert into TestResults (TestID, TestDate, Reqmt, Tester, Result) " &
                                    "Values (" & tid & ", '" & testdate & "', '" & reqmt & "', '" & Tester _
                                    & "', " & result & ")", cnTestResults)
        iRet = cmd.ExecuteNonQuery
        cnTestResults.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim cn As New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;Data Source=TestResultsDB.mdb")
        Dim iRet1, iret2 As Integer
        cn.Open()
        Dim cmd As New OleDbCommand("select count(*) from contactList", cn)
        iRet1 = cmd.ExecuteScalar
        cmd.CommandText = "select count(*) from customers"
        iret2 = cmd.ExecuteScalar

        If iRet1 = iret2 Then
            Testlog(100, DateTime.Now, "XYZ122", "Sweeny, M.", True)
            MessageBox.Show("Test pass reported to database")
        Else
            Testlog(100, DateTime.Now, "XYZ122", "Sweeny, M.", False)
            MessageBox.Show("Test failure reported to database")
        End If
        cn.Close()
    End Sub
End Class
