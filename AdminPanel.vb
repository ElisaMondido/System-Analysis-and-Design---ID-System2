Public Class AdminPanel

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        Me.Close()
        admin.Enabled = False

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If txtUsername.Text = "" Or txtPassword.Text = "" Or txtFullname.Text = "" Then
            MsgBox("Please Fill in All Fields")
            Exit Sub

        End If
        rs = New ADODB.Recordset
        rs.Open("SELECT * from tbl_userAdmin where username='" & txtUsername.Text & "'", con, 1, 2)
        If rs.EOF Then
            rs.AddNew()
            rs.Fields("username").Value = txtUsername.Text
            rs.Fields("password").Value = txtPassword.Text
            rs.Fields("fullname").Value = txtFullname.Text
            rs.Update()
            MsgBox("Successfully Added!")
        Else
            MsgBox("Already Existing")
        End If
        rs = Nothing
    End Sub

    Private Sub AdminPanel_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Left = 300
        Me.Top = 150
        TabControl1.Enabled = False
        txtUsername.Enabled = False
        txtPassword.Enabled = False
        txtFullname.Enabled = False

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        txtUsername.Enabled = True
        txtPassword.Enabled = True
        txtFullname.Enabled = True
    End Sub

    Private Sub TextBox3_LostFocus(sender As Object, e As EventArgs) Handles TextBox3.LostFocus
        rs = New ADODB.Recordset
        rs.Open("SELECT * from tblLogin where username='" & TextBox3.Text & "'", con, 1, 2)
        If rs.EOF Then
            TextBox3.Text = ""
            TextBox2.Text = ""
            TextBox1.Text = ""

        Else
            TextBox3.Text = rs.Fields("username").Value
            TextBox2.Text = rs.Fields("password").Value
            TextBox1.Text = rs.Fields("fullname").Value

        End If
        rs = Nothing
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        rs = New ADODB.Recordset
        rs.Open("SELECT *  from tblLogin where username='" & TextBox3.Text & "'", con, 1, 2)
        If CheckBox1.Checked = True Then
            rs.Fields("stat").Value = 0
            rs.Update()

        Else
            rs.Fields("stat").Value = 1
            rs.Update()
        End If
        rs = Nothing
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        rs = New ADODB.Recordset
        rs.Open("SELECT * from tbl_Useradmin where password='" & TextBox4.Text & "'", con, 1, 2)
        If rs.EOF Then
            TabControl1.Enabled = False
        Else
            TabControl1.Enabled = True
        End If
        rs = Nothing
    End Sub
End Class