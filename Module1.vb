Module Module1
    Public rs As New ADODB.Recordset
    Public con As New ADODB.Connection
    Public com As New ADODB.Command
    Public flag As String
    Public EmpNumber As String
    Public PosCodeNUm As String
    Public DeptCodeNUm As String
    Public UserCode As String
    Public SubjectCode As String
    Public MaxIDUser As Integer
    Public NameUser As String


    Public Sub mycon()
        rs = New ADODB.Recordset
        con = New ADODB.Connection
        com = New ADODB.Command
        con.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=sa;Initial Catalog=timekeeping"
        con.Open()
        com.ActiveConnection = con
    End Sub
End Module
