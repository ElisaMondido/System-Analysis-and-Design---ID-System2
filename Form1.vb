Public Class Form1
    Public Sub item_fill()
        Dim y As Integer
        rs = New ADODB.Recordset
        rs.Open("SELECT * from tblEmpInfo order by lastname asc", con, 1, 2)
        If rs.EOF Then
            ListView1.Items.Clear()
            y = 0
        Else
            ListView1.Items.Clear()
            Do While Not rs.EOF
                ListView1.Items.Add(rs.Fields("empnum").Value)
                ListView1.Items(y).SubItems.Add(rs.Fields("lastname").Value)
                ListView1.Items(y).SubItems.Add(rs.Fields("firstname").Value)
                ListView1.Items(y).SubItems.Add(rs.Fields("midname").Value)
                rs.MoveNext()
                y = y + 1

            Loop
        End If
        rs = Nothing
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call mycon()
        Call item_fill()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        Me.Close()
    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged
        '        admin.txtEmpNum.Text = ListView1.FocusedItem.Text
        '       admin.Label26.Text = ListView1.FocusedItem.SubItems(1).Text + ",  " + ListView1.FocusedItem.SubItems(2).Text + ",  " + ListView1.FocusedItem.SubItems(3).Text
        '
    End Sub
End Class