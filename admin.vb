Public Class admin
    Public Sub dis()
        cmdSave.Enabled = False
        cmdCancel.Enabled = False
        cmdAdd.Enabled = True
        cmdUpdate.Enabled = True
        cmdDelete.Enabled = True
    End Sub
    Public Sub ena()
        cmdSave.Enabled = True
        cmdCancel.Enabled = True
        cmdAdd.Enabled = False
        cmdUpdate.Enabled = False
        cmdDelete.Enabled = False
    End Sub

    Public Sub texdis()
        txtEmp.Enabled = False
        txtLast.Enabled = False
        txtFirst.Enabled = False
        txtMiddle.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False

    End Sub

    Public Sub texena()
        txtEmp.Enabled = True
        txtLast.Enabled = True
        txtFirst.Enabled = True
        txtMiddle.Enabled = True
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        ComboBox3.Enabled = True
        ComboBox4.Enabled = True
    End Sub
    Public Sub disCourse()
        Button6.Enabled = True
        Button5.Enabled = True
        Button4.Enabled = True
        Button3.Enabled = False
        Button2.Enabled = False
    End Sub
    Public Sub EnCourse()
        Button6.Enabled = False
        Button5.Enabled = False
        Button4.Enabled = False
        Button3.Enabled = True
        Button2.Enabled = True
    End Sub
    Public Sub disDept()
        Button12.Enabled = True
        Button11.Enabled = True
        Button10.Enabled = True
        Button9.Enabled = False
        Button8.Enabled = False
    End Sub
    Public Sub EnDept()
        Button12.Enabled = False
        Button11.Enabled = False
        Button10.Enabled = False
        Button9.Enabled = True
        Button8.Enabled = True
    End Sub
    Public Sub disAnnounce()
        Button24.Enabled = True
        Button23.Enabled = True
        Button22.Enabled = True
        Button21.Enabled = False
        Button20.Enabled = False
    End Sub
    Public Sub EnAnnounce()
        Button24.Enabled = False
        Button23.Enabled = False
        Button22.Enabled = False
        Button21.Enabled = True
        Button20.Enabled = True
    End Sub
    Public Sub item_fill()
        Dim i As Integer
        rs = New ADODB.Recordset
        rs.Open("SELECT * from tblEmpInfo order by lastname asc", con, 1, 2)
        If rs.EOF Then
            ListView1.Items.Clear()
            i = 0
        Else
            ListView1.Items.Clear()
            Do While Not rs.EOF
                ListView1.Items.Add(rs.Fields("empnum").Value)
                ListView1.Items(i).SubItems.Add(rs.Fields("lastname").Value)
                ListView1.Items(i).SubItems.Add(rs.Fields("firstname").Value)
                ListView1.Items(i).SubItems.Add(rs.Fields("midname").Value)
                ListView1.Items(i).SubItems.Add(rs.Fields("position").Value)
                ListView1.Items(i).SubItems.Add(rs.Fields("dept").Value)
                ListView1.Items(i).SubItems.Add(rs.Fields("syear").Value)
                ListView1.Items(i).SubItems.Add(rs.Fields("sem").Value)
                rs.MoveNext()
                i = i + 1

            Loop
        End If
        rs = Nothing
    End Sub
    Public Sub item_fill_Subject()
        Dim i As Integer
        rs = New ADODB.Recordset
        rs.Open("SELECT * from tblAnnounce", con, 1, 2)
        If rs.EOF Then
            ListView5.Items.Clear()
            i = 0
        Else
            ListView5.Items.Clear()
            Do While Not rs.EOF
                ListView5.Items.Add(rs.Fields("Studnum").Value)
                ListView5.Items(i).SubItems.Add(rs.Fields("Message").Value)
                ListView5.Items(i).SubItems.Add(rs.Fields("sender").Value)
                ListView5.Items(i).SubItems.Add(rs.Fields("DateNeeded").Value)

                rs.MoveNext()
                i = i + 1

            Loop
        End If
        rs = Nothing
    End Sub
    Public Sub item_fill_Pos()
        Dim i As Integer
        rs = New ADODB.Recordset
        rs.Open("SELECT * from tblPosition order by posCode asc", con, 1, 2)
        If rs.EOF Then
            ListView2.Items.Clear()
            i = 0
        Else
            ListView2.Items.Clear()
            Do While Not rs.EOF
                ListView2.Items.Add(rs.Fields("posCode").Value)
                ListView2.Items(i).SubItems.Add(rs.Fields("PosDesc").Value)
                rs.MoveNext()
                i = i + 1

            Loop
        End If
        rs = Nothing
    End Sub
    Public Sub item_fill_Dept()
        Dim i As Integer
        rs = New ADODB.Recordset
        rs.Open("SELECT * from tblDepartment order by deptCode asc", con, 1, 2)
        If rs.EOF Then
            ListView3.Items.Clear()
            i = 0
        Else
            ListView3.Items.Clear()
            Do While Not rs.EOF
                ListView3.Items.Add(rs.Fields("deptCode").Value)
                ListView3.Items(i).SubItems.Add(rs.Fields("deptDesc").Value)
                rs.MoveNext()
                i = i + 1
            Loop
        End If
        rs = Nothing
    End Sub
    Public Sub item_fill_User()
        Dim y As Integer
        rs = New ADODB.Recordset
        rs.Open("SELECT * from tblLogin", con, 1, 2)
        If rs.EOF Then
            ListView6.Items.Clear()
            y = 0
        Else
            ListView6.Items.Clear()
            Do While Not rs.EOF
                ListView6.Items.Add(rs.Fields("fullname").Value)
                ListView6.Items(y).SubItems.Add(rs.Fields("username").Value)
                ListView6.Items(y).SubItems.Add(rs.Fields("password").Value)
                rs.MoveNext()
                y = y + 1
            Loop
        End If
        rs = Nothing
    End Sub
    Public Sub item_fill_History()
        Dim y As Integer
        rs = New ADODB.Recordset
        rs.Open("SELECT * from tblHistory where LogInID <> 0", con, 1, 2)
        If rs.EOF Then
            ListView7.Items.Clear()
            y = 0
        Else
            ListView7.Items.Clear()
            Do While Not rs.EOF
                ListView7.Items.Add(rs.Fields("fullname").Value)
                ListView7.Items(y).SubItems.Add(rs.Fields("dateIn").Value)
                ListView7.Items(y).SubItems.Add(rs.Fields("TimeIN").Value)
                ListView7.Items(y).SubItems.Add(rs.Fields("TimeOut").Value)
                rs.MoveNext()
                y = y + 1
            Loop
        End If
        rs = Nothing
    End Sub
    Public Sub item_fill_Report()

        Dim y As Integer
        rs = New ADODB.Recordset
        rs.Open("SELECT * from tbl_log", con, 1, 2)
        If rs.EOF Then
            ListView4.Items.Clear()
            y = 0
        Else
            ListView4.Items.Clear()
            Do While Not rs.EOF
                ListView4.Items.Add(rs.Fields("empNum").Value)
                ListView4.Items(y).SubItems.Add(rs.Fields("fname").Value)
                ListView4.Items(y).SubItems.Add(rs.Fields("TimeIn").Value)
                ListView4.Items(y).SubItems.Add(rs.Fields("TimeOut").Value)
                ListView4.Items(y).SubItems.Add(rs.Fields("DateNow").Value)
                ListView4.Items(y).SubItems.Add(rs.Fields("day").Value)
                rs.MoveNext()
                y = y + 1
            Loop
        End If
        rs = Nothing
    End Sub
    Private Sub ListView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.Click
        txtEmp.Text = ListView1.FocusedItem.Text
        txtLast.Text = ListView1.FocusedItem.SubItems(1).Text
        txtFirst.Text = ListView1.FocusedItem.SubItems(2).Text
        txtMiddle.Text = ListView1.FocusedItem.SubItems(3).Text
    End Sub

    Private Sub TabPage1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage1.Click
        Call mycon()
        Call item_fill()
        Call dis()
        Call texdis()
        ListView1.Visible = True
    End Sub

    Private Sub cmdAdd_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        flag = "ADD"
        txtEmp.Clear()
        txtLast.Clear()
        txtFirst.Clear()
        txtMiddle.Clear()

        Call ena()
        Call texena()

    End Sub

    Private Sub cmdSave_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call mycon()

        If txtEmp.Text = "" Then
            MsgBox("PLease Fill in Student Number")
            Exit Sub
        End If

        If txtLast.Text = "" Then
            MsgBox("PLease Fill In Last Name")
            Exit Sub

        End If

        If txtFirst.Text = "" Then
            MsgBox("PLease Fill In Firstname")
            Exit Sub

        End If
        If txtMiddle.Text = "" Then
            MsgBox("PLease Fill In MIddleName")
            Exit Sub
        End If

        If ComboBox1.Text = "" Then
            MsgBox("Please select Course")
            Exit Sub
        End If

        If ComboBox3.Text = "" Then
            MsgBox("Please select School Year")
            Exit Sub
        End If

        If ComboBox4.Text = "" Then
            MsgBox("Please select Semester")
            Exit Sub
        End If

        If ComboBox4.Text = "" Then
            MsgBox("Please select Semester")
            Exit Sub
        End If

        If TextBox8.Text = "" Then
            MsgBox("Please Enter Student PIN")
            Exit Sub
        End If

        If flag = "ADD" Then
            rs = New ADODB.Recordset
            rs.Open("SELECT * from tblEmpInfo where empnum='" & txtEmp.Text & "'", con, 1, 2)
            If rs.EOF Then
                rs.AddNew()
                rs.Fields("empnum").Value = txtEmp.Text
                rs.Fields("lastname").Value = txtLast.Text
                rs.Fields("firstname").Value = txtFirst.Text
                rs.Fields("midname").Value = txtMiddle.Text
                rs.Fields("position").Value = ComboBox1.Text
                rs.Fields("dept").Value = ComboBox2.Text
                rs.Fields("syear").Value = ComboBox3.Text
                rs.Fields("sem").Value = ComboBox4.Text
                rs.Fields("pinCode").Value = TextBox8.Text

                rs.Fields("picpath").Value = lblstudpic.Text
                rs.Update()
                MsgBox("Successfully Saved")
                item_fill()
                Call dis()
                Call texdis()
            Else
                MsgBox("Already Existing")
            End If
            rs = Nothing
        ElseIf flag = "EDIT" Then
            rs = New ADODB.Recordset
            rs.Open("SELECT * from tblEmpInfo where empnum = '" & EmpNumber & "'", con, 1, 2)
            rs.Fields("empnum").Value = txtEmp.Text
            rs.Fields("lastname").Value = txtLast.Text
            rs.Fields("firstname").Value = txtFirst.Text
            rs.Fields("midname").Value = txtMiddle.Text

            rs.Update()
            MsgBox("Successfully Saved")
            item_fill()
            Call dis()
            Call texdis()
            rs = Nothing
        End If
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Call dis()
        Call texdis()
    End Sub

    Private Sub cmdUpdate_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
        flag = "EDIT"
        EmpNumber = txtEmp.Text
        Call ena()
        Call texena()
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        Dim x
        x = MsgBox("DELETE this record?", vbYesNo + vbCritical, "Deleting Record")
        If x = vbYes Then
            rs = New ADODB.Recordset
            rs.Open("DELETE FROM TBLEMPINFO WHERE EMPNUM='" & txtEmp.Text & "'", con, 1, 2)
            MsgBox("Successfully Removed")
            item_fill()
            Call dis()
            Call texdis()
            rs = Nothing
        Else
        End If
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub TabPage4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call mycon()
        Call item_fill()
        Call dis()
        Call texdis()
    End Sub
    Private Sub cmdShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        schedule.Show()
        Me.Hide()
    End Sub

    Private Sub cmdAdd3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        flag = "ADD"

        Call ena()
        Call texena()
    End Sub

    Private Sub cmdUpdate3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        flag = "EDIT"
        EmpNumber = txtEmp.Text
    End Sub

    Private Sub cmdCancel3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call dis()
        Call texdis()
    End Sub

    Private Sub cmdExit3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Please Fill Up Properly")
            Exit Sub
        End If
        If flag = "ADD" Then
            rs = New ADODB.Recordset
            rs.Open("SELECT * from tblPosition where poscode='" & TextBox1.Text & "'", con, 1, 2)
            If rs.EOF Then
                rs.AddNew()
                rs.Fields("poscode").Value = TextBox1.Text
                rs.Fields("posdesc").Value = TextBox2.Text
                rs.Update()
                MsgBox("Successfully Added")
                item_fill_Pos()
                Call disCourse()
            Else
                MsgBox("Already Existing")
            End If
            rs = Nothing
        ElseIf flag = "EDIT" Then
            rs = New ADODB.Recordset
            rs.Open("SELECT * from tblPosition where poscode='" & PosCodeNUm & "'", con, 1, 2)
           
                rs.Fields("poscode").Value = TextBox1.Text
                rs.Fields("posdesc").Value = TextBox2.Text
                rs.Update()
            MsgBox("Successfully Saved the Changes")
            disCourse()
                item_fill_Pos()
           
            rs = Nothing

        End If

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If TextBox5.Text = "" Or TextBox3.Text = "" Then
            MsgBox("Please Fill Up Properly")
            Exit Sub
        End If
        If flag = "ADD" Then
            rs = New ADODB.Recordset
            rs.Open("SELECT * from tblDepartment where deptcode='" & TextBox5.Text & "'", con, 1, 2)
            If rs.EOF Then
                rs.AddNew()
                rs.Fields("deptcode").Value = TextBox5.Text
                rs.Fields("deptdesc").Value = TextBox3.Text
                rs.Update()
                MsgBox("Successfully Added")
                item_fill_Dept()
                Call disDept()
            Else
                MsgBox("Already Existing")
            End If
            rs = Nothing
        ElseIf flag = "EDIT" Then
            rs = New ADODB.Recordset
            rs.Open("SELECT * from tblDepartment where deptcode='" & DeptCodeNUm & "'", con, 1, 2)
           
                rs.Fields("deptcode").Value = TextBox1.Text
                rs.Fields("Deptdesc").Value = TextBox2.Text
                rs.Update()
            MsgBox("Successfully Added")
            Call disDept()
            item_fill_Dept()
         
            rs = Nothing
        End If
    End Sub

    Private Sub admin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call mycon()
        item_fill_History()
        item_fill_Pos()
        item_fill()
        item_fill_Dept()
        item_fill_Subject()
        item_fill_Report()
        Call disDept()
        Call dis()

        rs = New ADODB.Recordset
        rs.Open("SELECT * from tblLogIn where username='" & NameUser & "'", con, 1, 2)
        Label34.Text = rs.Fields("Fullname").Value
        rs = Nothing
        texdis()
        Call disCourse()
        Call disAnnounce()
        TextBox1.Enabled = False
        TextBox2.Enabled = False

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        flag = "ADD"
        Call EnCourse()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        Call EnCourse()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        flag = "EDIT"
        PosCodeNUm = TextBox1.Text
        Call EnCourse()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        flag = "EDIT"
        DeptCodeNUm = TextBox5.Text
        Call EnDept()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim x
        x = MsgBox("Delete this Record", vbYesNo + vbCritical, "Deleting Records")
        If x = vbYes Then
            rs = New ADODB.Recordset
            rs.Open("DELETE FROM TBLPosition where posCode='" & TextBox1.Text & "'", con, 1, 2)
            MsgBox("Successfully Deleted!")
            Call item_fill_Pos()
            Call disCourse()
            rs = Nothing
        Else
        End If

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim x
        x = MsgBox("Delete this Record", vbYesNo + vbCritical, "Deleting Records")
        If x = vbYes Then
            rs = New ADODB.Recordset
            rs.Open("DELETE FROM TBLDepartment where DeptCode='" & TextBox5.Text & "'", con, 1, 2)
            MsgBox("Successfully Deleted!")

            Call disDept()
            rs = Nothing
        Else
        End If
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        flag = "ADD"
        Call EnDept()
    End Sub
    Private Sub Button13_Click(sender As Object, e As EventArgs)
        Form1.Show()
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        If flag = "ADD" Then
            rs = New ADODB.Recordset
            rs.Open("SELECT * from tblAnnounce", con, 1, 2)
            rs.AddNew()
            rs.Fields("studNUm").Value = txtSubjCode.Text
            rs.Fields("studentName").Value = TextBox4.Text
            rs.Fields("message").Value = txtSubjDesc.Text
            rs.Fields("sender").Value = txtUnits.Text
            rs.Fields("DatePublish").Value = DateValue(Now)
            rs.Fields("DateNeeded").Value = DateTimePicker1.Value
            rs.Update()
            MsgBox("Successfully Saved")
            disAnnounce()
            Call item_fill_Subject()
            rs = Nothing
        ElseIf flag = "EDIT" Then
        rs = New ADODB.Recordset
            rs.Open("SELECT * from tblAnnounce where studnum='" & SubjectCode & "'", con, 1, 2)
        rs.Fields("studNUm").Value = txtSubjCode.Text
        rs.Fields("studentName").Value = TextBox4.Text
        rs.Fields("message").Value = txtSubjDesc.Text
        rs.Fields("sender").Value = txtUnits.Text
        rs.Fields("DatePublish").Value = DateValue(Now)
        rs.Fields("DateNeeded").Value = DateTimePicker1.Value

        rs.Update()
        MsgBox("Successfully Saved")
        Call item_fill_Subject()

        rs = Nothing
        End If

    End Sub

    Private Sub ListView5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView5.SelectedIndexChanged
        txtSubjCode.Text = ListView5.FocusedItem.Text
        txtSubjDesc.Text = ListView5.FocusedItem.SubItems(1).Text
        txtUnits.Text = ListView5.FocusedItem.SubItems(2).Text

    End Sub
    Private Sub Label28_Click(sender As Object, e As EventArgs) Handles Label28.Click
        rs = New ADODB.Recordset
        rs.Open("SELECT * from tblHistory where LogInID='" & MaxIDUser & "'", con, 1, 2)
        rs.Fields("TimeOut").Value = TimeValue(Now)
        rs.Update()
        rs = Nothing

        Me.Close()

    End Sub

    

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        Dim i As Integer
        rs = New ADODB.Recordset
        rs.Open("SELECT * from tblEmpInfo where [lastname]like'" & TextBox6.Text & "%' order by lastname asc", con, 1, 2)
        If rs.EOF Then
            ListView1.Items.Clear()
            i = 0
        Else
            ListView1.Items.Clear()
            Do While Not rs.EOF
                ListView1.Items.Add(rs.Fields("empnum").Value)
                ListView1.Items(i).SubItems.Add(rs.Fields("lastname").Value)
                ListView1.Items(i).SubItems.Add(rs.Fields("firstname").Value)
                ListView1.Items(i).SubItems.Add(rs.Fields("midname").Value)
                ListView1.Items(i).SubItems.Add(rs.Fields("position").Value)
                ListView1.Items(i).SubItems.Add(rs.Fields("dept").Value)
                ListView1.Items(i).SubItems.Add(rs.Fields("syear").Value)
                ListView1.Items(i).SubItems.Add(rs.Fields("sem").Value)
                rs.MoveNext()
                i = i + 1

            Loop
        End If
        rs = Nothing
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        If flag = "ADD" Then
            rs = New ADODB.Recordset
            rs.Open("SELECT * from tblLogIn where username='" & txtUSer.Text & "'", con, 1, 2)
            rs.AddNew()
            rs.Fields("username").Value = txtUSer.Text
            rs.Fields("password").Value = txtPassword.Text
            rs.Fields("Fullname").Value = txtFullname.Text
            rs.Fields("stat").Value = 0

            rs.Update()
            MsgBox("Successfully Saved")
            item_fill_User()
            rs = Nothing
        ElseIf flag = "EDIT" Then
            rs = New ADODB.Recordset
            rs.Open("SELECT * from tblLogIn where username='" & UserCode & "'", con, 1, 2)

            rs.Fields("username").Value = txtUSer.Text
            rs.Fields("password").Value = txtPassword.Text
            rs.Fields("Fullname").Value = txtFullname.Text
            rs.Update()
            MsgBox("Successfully Updated")
            item_fill_User()
            rs = Nothing
        End If

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        flag = "ADD"
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        flag = "EDIT"
        UserCode = txtUSer.Text
    End Sub

    Private Sub ListView6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView6.SelectedIndexChanged
        txtFullname.Text = ListView6.FocusedItem.Text
        txtPassword.Text = ListView6.FocusedItem.SubItems(2).Text
        txtUSer.Text = ListView6.FocusedItem.SubItems(1).Text
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        rs = New ADODB.Recordset
        rs.Open("delete from tblLogin where username='" & txtUSer.Text & "'", con, 1, 2)
        MsgBox("Record has been Remove")
        rs = Nothing

    End Sub
   
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label33.Text = Format(DateValue(Now), "MMMM dd, yyyy")

    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        Dim x
        x = MsgBox("DELETE this Record?", vbYesNo + vbQuestion, "Deleting Subject")
        If x = vbYes Then
            rs = New ADODB.Recordset
            rs.Open("DELETE from tblAnnounce where studnum='" & txtSubjCode.Text & "'", con, 1, 2)
            MsgBox("Successfully Deleted")
            disAnnounce()
            Call item_fill_Subject()
            rs = Nothing

        Else

        End If
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        flag = "ADD"
        Call EnAnnounce()
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        flag = "EDIT"
        SubjectCode = txtSubjCode.Text
        Call EnAnnounce()
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Dim OpenFileDialog1 As New OpenFileDialog
        OpenFileDialog1.Filter = "Picture Files (*)|*.jpg;*"
        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            studpic.ImageLocation = OpenFileDialog1.FileName
            lblstudpic.Text = studpic.ImageLocation
        End If
    End Sub

    Private Sub txtEmp_LostFocus(sender As Object, e As EventArgs) Handles txtEmp.LostFocus
        rs = New ADODB.Recordset
        rs.Open("SELECT * from tblEmpInfo where empNum='" & txtEmp.Text & "'", con, 1, 2)
        If rs.EOF Then
            txtLast.Focus()
        Else
            MsgBox("Already Used...")
            txtEmp.Text = ""
            txtEmp.Focus()
        End If
        rs = Nothing

    End Sub

    Private Sub txtEmp_TextChanged(sender As Object, e As EventArgs) Handles txtEmp.TextChanged
       
    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged
        rs = New ADODB.Recordset
        rs.Open("select picpath from tblEmpInfo where empnum='" & ListView1.FocusedItem.Text & "'", con, 1, 2)
        If rs.EOF Then
            studpic.ImageLocation = "D:\PICTURES\New Pictures\Ms. Eartrh 2011 Coronation Night\1.jpg"
        Else
            lblstudpic.Text = rs.Fields("picpath").Value
            studpic.ImageLocation = lblstudpic.Text
        End If
        rs = Nothing
    End Sub

    Private Sub ComboBox1_GotFocus(sender As Object, e As EventArgs) Handles ComboBox1.GotFocus
        rs = New ADODB.Recordset
        rs.Open("SELECT * from tblPosition order by poscode asc", con, 1, 2)
        If rs.EOF Then
            ComboBox1.Items.Clear()
        Else
            ComboBox1.Items.Clear()
            Do While Not rs.EOF
                ComboBox1.Items.Add(rs.Fields("poscode").Value)
                rs.MoveNext()

            Loop
        End If
        rs = Nothing

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub ComboBox2_GotFocus(sender As Object, e As EventArgs) Handles ComboBox2.GotFocus
        rs = New ADODB.Recordset
        rs.Open("SELECT * from tblDepartment order by deptcode asc", con, 1, 2)
        If rs.EOF Then
            ComboBox2.Items.Clear()
        Else
            ComboBox2.Items.Clear()
            Do While Not rs.EOF
                ComboBox2.Items.Add(rs.Fields("deptcode").Value)
                rs.MoveNext()

            Loop
        End If
        rs = Nothing
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub

    Private Sub ComboBox3_GotFocus(sender As Object, e As EventArgs) Handles ComboBox3.GotFocus
        rs = New ADODB.Recordset
        rs.Open("SELECT * from  tbl_SchoolYear order by sy asc", con, 1, 2)
        If rs.EOF Then
            ComboBox3.Items.Clear()
        Else
            ComboBox3.Items.Clear()
            Do While Not rs.EOF
                ComboBox3.Items.Add(rs.Fields("sy").Value)
                rs.MoveNext()

            Loop
        End If
        rs = Nothing
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

    End Sub

    Private Sub txtSubjCode_LostFocus(sender As Object, e As EventArgs) Handles txtSubjCode.LostFocus
        rs = New ADODB.Recordset
        rs.Open("SELECT * from tblEmpInfo where empNum='" & txtSubjCode.Text & "'", con, 1, 2)
        If rs.EOF Then
            TextBox4.Text = ""
        Else
            TextBox4.Text = rs.Fields("lastname").Value + "," + rs.Fields("firstname").Value + "," + rs.Fields("midname").Value
        End If
        rs = Nothing


    End Sub

    Private Sub txtSubjCode_TextChanged(sender As Object, e As EventArgs) Handles txtSubjCode.TextChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call disCourse()
    End Sub

    Private Sub ListView2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView2.SelectedIndexChanged
        TextBox1.Text = ListView2.FocusedItem.Text
        TextBox2.Text = ListView2.FocusedItem.SubItems(1).Text
    End Sub

    Private Sub ComboBox5_GotFocus(sender As Object, e As EventArgs) Handles ComboBox5.GotFocus
        rs = New ADODB.Recordset
        rs.Open("SELECT distinct(poscode) from tblPosition order by poscode asc", con, 1, 2)
        If rs.EOF Then
            ComboBox5.Items.Clear()
        Else
            ComboBox5.Items.Clear()
            Do While Not rs.EOF
                ComboBox5.Items.Add(rs.Fields("poscode").Value)
                rs.MoveNext()

            Loop

        End If
        rs = Nothing

    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        Dim i As Integer
        rs = New ADODB.Recordset
        rs.Open("SELECT * from tblEmpInfo where position='" & ComboBox5.Text & "' order by lastname asc", con, 1, 2)
        If rs.EOF Then
            ListView1.Items.Clear()
            i = 0
        Else
            ListView1.Items.Clear()
            Do While Not rs.EOF
                ListView1.Items.Add(rs.Fields("empnum").Value)
                ListView1.Items(i).SubItems.Add(rs.Fields("lastname").Value)
                ListView1.Items(i).SubItems.Add(rs.Fields("firstname").Value)
                ListView1.Items(i).SubItems.Add(rs.Fields("midname").Value)
                ListView1.Items(i).SubItems.Add(rs.Fields("position").Value)
                ListView1.Items(i).SubItems.Add(rs.Fields("dept").Value)
                ListView1.Items(i).SubItems.Add(rs.Fields("syear").Value)
                ListView1.Items(i).SubItems.Add(rs.Fields("sem").Value)
                rs.MoveNext()
                i = i + 1

            Loop
        End If
        rs = Nothing
    End Sub

    Private Sub TabPage7_Click(sender As Object, e As EventArgs) Handles TabPage7.Click

    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        Dim y As Integer
        rs = New ADODB.Recordset
        rs.Open("SELECT * from tblHistory where [fullname]like'" & TextBox7.Text & "%'", con, 1, 2)
        If rs.EOF Then
            ListView7.Items.Clear()
            y = 0
        Else
            ListView7.Items.Clear()
            Do While Not rs.EOF
                ListView7.Items.Add(rs.Fields("fullname").Value)
                ListView7.Items(y).SubItems.Add(rs.Fields("dateIn").Value)
                ListView7.Items(y).SubItems.Add(rs.Fields("TimeIN").Value)
                ListView7.Items(y).SubItems.Add(rs.Fields("TimeOut").Value)
                rs.MoveNext()
                y = y + 1
            Loop
        End If
        rs = Nothing

    End Sub

    Private Sub Button13_Click_1(sender As Object, e As EventArgs) Handles Button13.Click
        rs = New ADODB.Recordset
        Dim x
        x = MsgBox("Delete History?", vbYesNo + vbCritical, "Deleting History")
        If x = vbYes Then
            rs.Open("DELETE from tblHistory where LogInID <> 0 and TimeOut<>'--'", con, 1, 2)
            item_fill_History()
        Else

        End If
        rs = Nothing

    End Sub

    Private Sub TabPage4_Click_1(sender As Object, e As EventArgs) Handles TabPage4.Click

    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged

        Dim y As Integer
        rs = New ADODB.Recordset
        rs.Open("SELECT * from tbl_log where [fname]like'" & TextBox9.Text & "%'", con, 1, 2)
        If rs.EOF Then
            ListView4.Items.Clear()
            y = 0
        Else
            ListView4.Items.Clear()
            Do While Not rs.EOF
                ListView4.Items.Add(rs.Fields("empNum").Value)
                ListView4.Items(y).SubItems.Add(rs.Fields("fname").Value)
                ListView4.Items(y).SubItems.Add(rs.Fields("TimeIn").Value)
                ListView4.Items(y).SubItems.Add(rs.Fields("TimeOut").Value)
                ListView4.Items(y).SubItems.Add(rs.Fields("DateNow").Value)
                ListView4.Items(y).SubItems.Add(rs.Fields("day").Value)
                rs.MoveNext()
                y = y + 1
            Loop
        End If
        rs = Nothing
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        AdminPanel.Show()
        Me.Enabled = False

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Call disDept()
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        disAnnounce()
    End Sub

    Private Sub ListView3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView3.SelectedIndexChanged
        TextBox5.Text = ListView3.FocusedItem.Text
        TextBox3.Text = ListView3.FocusedItem.SubItems(1).Text
    End Sub

    Private Sub TabPage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Click

    End Sub
End Class
