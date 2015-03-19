Public Class User

    Private Sub User_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call mycon()
        Me.Left = 500
        Me.Top = 250

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        End
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        rs = New ADODB.Recordset
        rs.Open("SELECT * from tblLogIn where username='" & TextBox1.Text & "' and password='" & TextBox2.Text & "' AND STAT=1 ", con, 1, 2)
        If rs.EOF Then
            MsgBox("Access Denied")
        Else
            Dim FName As String
            FName = rs.Fields("FullName").Value
            NameUser = rs.Fields("username").Value
            Dim MaxID As Integer
            rs = New ADODB.Recordset
            rs.Open("SELECT MAX(logInID) as MaxID from tblHistory", con, 1, 2)
            MaxIDUser = rs.Fields("MaxID").Value + 1
            rs = Nothing


            rs = New ADODB.Recordset
            rs.Open("SELECT * from tblHistory", con, 1, 2)
            rs.AddNew()

            rs.Fields("LogInID").Value = MaxIDUser
            rs.Fields("Fullname").Value = FName
            rs.Fields("DateIn").Value = DateValue(Now)
            rs.Fields("TimeIn").Value = TimeValue(Now)
            rs.Fields("TimeOut").Value = "--"
            rs.Update()
            rs = Nothing


            admin.Show()
            Me.Hide()
        End If

        rs = Nothing

    End Sub
End Class