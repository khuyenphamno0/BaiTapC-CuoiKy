Imports System.Data.OleDb

Public Class Form1
    Private con As OleDbConnection
    Private WithEvents qly_BHg, qly_Hg As BindingManagerBase
    Public lenh As String

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim constring As String = "Provider= Microsoft.Jet.OLEDB.4.0;" _
    & "Data Source = " & Application.StartupPath & "\QLSV.mdb;"
        con = New OleDbConnection(constring)
        Xuat_HOADONBH()
        Xuat_HANG()
    End Sub
    Private Sub Xuat_HOADONBH()
        Dim lenh As String
        lenh = "select Shd, HD.mahg, tenhg, slg, giaban,slg*giaban As TTien from HOADONBH as HD, HANG as Hg Where HD.mahg=Hg.mahg"
        Dim cmd As New OleDbCommand(lenh, con)
        con.Open()
        Dim bang_doc As OleDbDataReader = cmd.ExecuteReader
        Dim dttable As New DataTable("HOADON")
        dttable.Load(bang_doc, LoadOption.OverwriteChanges)
        con.Close()
        DataGrid1.DataSource = dttable
        qly_BHg = Me.BindingContext(dttable)
    End Sub
    Private Sub Xuat_HANG()
        Dim lenh As String
        lenh = "select * From Hang"
        Dim cmd As New OleDbCommand(lenh, con)
        con.Open()
        Dim bang_doc As OleDbDataReader = cmd.ExecuteReader
        Dim dttable As New DataTable("HANG")
        dttable.Load(bang_doc, LoadOption.OverwriteChanges)
        con.Close()
        qly_Hg = Me.BindingContext(dttable)
        CB.Text = qly_Hg.Current("mahg") & " | " & qly_Hg.Current("tenhg")
        'Lấy dữ liệu từ bảng đưa vào Combobox
        For I = 0 To qly_Hg.Count - 1
            qly_Hg.Position = I
            CB.Items.Add(qly_Hg.Current("mahg") & " | " & qly_Hg.Current("tenhg"))
        Next

    End Sub
    Private Sub CB_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles CB.SelectedIndexChanged
        Dim tach As String
        tach = CB.Text.Substring(0, InStr(CB.Text, " |")).Trim
        lenh = "select Shd, HD.mahg, tenhg, slg, giaban,slg*giaban As TTien from HOADONBH as HD, HANG as Hg Where HD.mahg=Hg.mahg and HD.mahg = '" & tach & "'"
        Dim cmd As New OleDbCommand(lenh, con)
        con.Open()
        Dim bang_doc As OleDbDataReader = cmd.ExecuteReader
        Dim dttable As New DataTable("hang")
        dttable.Load(bang_doc, LoadOption.OverwriteChanges)
        con.Close()
        qly_BHg = Me.BindingContext(dttable)
        DataGrid1.DataSource = dttable

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Xuat_HOADONBH()
    End Sub
End Class
