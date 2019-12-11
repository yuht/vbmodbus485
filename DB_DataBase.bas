Attribute VB_Name = "DB_database"
Option Explicit

'��Ҫ���� Microsoft ADO Ext x.x for DDL and Security �����ADOX
'��Ҫ���� Microsoft ActiveX Data Objects XXX Library �����ADODB��� ��������ACCESS���ݿ����SQLSERVER���ݿ�

Dim Cat As New ADOX.Catalog '����cat������һ������Ҳ����
Dim Conn As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Public DBpath_name As String  '���ݿ����
'

Public Function DB_CreateDataBase(DataBase_Path_And_Name As String) As Boolean
    On Error Resume Next
    Debug.Print "create DB"
    DBpath_name = "provider=Microsoft.Jet.OLEDB.4.0;data source=" & DataBase_Path_And_Name & ";" '�������ݿ�
    'Debug.Print DBpath_name
    
    Cat.Create DBpath_name '    �������ݿ�
    
    If Err And Err <> &H80040E17 Then  '���ݿ� DataBase_Path_And_Name �Ѿ����ڡ�
        
        Debug.Print Err.Number, Hex(Err.Number)
        Debug.Print Err.Description
        Err.Clear
        DB_CreateDataBase = False
        Exit Function
    End If
    
    DB_CreateDataBase = True
End Function
'

Public Function DB_Create_Table(DB_TableName As Table) As Boolean
    On Error Resume Next
    Debug.Print "create table"
    Cat.ActiveConnection = DBpath_name
    Cat.Tables.Append DB_TableName '�������ݱ�
    
    If Err And Err <> &H80040E3F Then   '�� DB_TableName �Ѵ��ڡ�
        Debug.Print Err.Number, Hex(Err.Number)
        Debug.Print Err.Description
        Err.Clear
        DB_Create_Table = False
        Exit Function
    End If
    DB_Create_Table = True
End Function
'

Public Function DB_InsertRecord(Table_Name As String, RecordArray() As String) As Boolean
    On Error Resume Next
    Dim i As Integer
    Conn.Open DBpath_name
    Rs.CursorLocation = adUseClient
    Rs.Open Table_Name, Conn, adOpenKeyset, adLockPessimistic
    Rs.AddNew '����������¼�¼
    For i = 0 To UBound(RecordArray)
        Rs.Fields(i).Value = RecordArray(i)
    Next
    Debug.Print "rs.RecordCount ", Rs.RecordCount
    Rs.Update
    Rs.Close
    Conn.Close
    
    If Err Then
        Debug.Print Err.Number
        Debug.Print Err.Description
        Err.Clear
    End If
    
End Function
 
