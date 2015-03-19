Public Class schedule

    Public Sub item_fill()
        Dim u As Integer
        rs = New ADODB.Recordset
        rs.Open("SELECT * from tblSchedule order by days asc", con, 1, 2)
        If rs.EOF Then
            ListView1.Items.Clear()
            u = 0
        Else
            ListView1.Items.Clear()
            Do While Not rs.EOF
                ListView1.Items.Add(rs.Fields("empnum").Value)
                ListView1.Items(u).SubItems.Add(rs.Fields("subjcode").Value)
                ListView1.Items(u).SubItems.Add(rs.Fields("subjdesc").Value)
                ListView1.Items(u).SubItems.Add(rs.Fields("unit").Value)
                ListView1.Items(u).SubItems.Add(rs.Fields("days").Value)
                ListView1.Items(u).SubItems.Add(rs.Fields("sem").Value)
                ListView1.Items(u).SubItems.Add(rs.Fields("syear").Value)
                ListView1.Items(u).SubItems.Add(rs.Fields("timefrom").Value)
                ListView1.Items(u).SubItems.Add(rs.Fields("timeto").Value)
                rs.MoveNext()
                u = u + 1

            Loop
        End If
        rs = Nothing

    End Sub

    Private Sub schedule_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call mycon()
        Call item_fill()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        admin.Show()
        Me.Hide()
    End Sub
End Class