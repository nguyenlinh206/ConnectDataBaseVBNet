Imports System.Data.OleDb
Imports System.Data.OleDb.OleDbException
Public Class Oledb
    Private conn As OleDbConnection
    Public Sub New(ByVal oledb As String)
        Dim StrConnect As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & oledb
        conn = New OleDbConnection(StrConnect)
    End Sub
    Public Sub Insert(ByVal Cmd As String)
        Dim _Excute As New OleDbCommand(Cmd, conn)
        If conn.State = ConnectionState.Closed Then conn.Open()
        _Excute.ExecuteNonQuery()
        If conn.State = ConnectionState.Open Then conn.Close()
    End Sub

    Public Function GetTable(ByVal Cmd As String, Optional _Table As String = "")
        Dim _Excute As New OleDbCommand(Cmd, conn)
        If conn.State = ConnectionState.Closed Then conn.Open()
        Dim Read_Table As OleDbDataReader = _Excute.ExecuteReader
        Dim Get_Table As New DataTable(_Table)
        Get_Table.Load(Read_Table, LoadOption.OverwriteChanges)
        If conn.State = ConnectionState.Open Then conn.Close()
        Return Get_Table
    End Function
    Public Sub Edit(ByVal Cmd As String)
        Dim _Excute As New OleDbCommand(Cmd, conn)
        If conn.State = ConnectionState.Closed Then conn.Open()
        _Excute.CommandType = CommandType.Text
        _Excute.ExecuteNonQuery()
        If conn.State = ConnectionState.Open Then conn.Close()
    End Sub
    Public Sub InsertPar(ByVal _Table As String, ByVal _Columns As String, ByVal _Content As ArrayList)
        'thempar("Title", "TieuDe,NoiDung", noidung)
        Dim ValueColumn As String = "@" & _Columns
        ValueColumn = Replace(ValueColumn, ",", ",@")
        Dim ListColumn As Array
        ListColumn = Split(ValueColumn, ",")
        If ListColumn.Length = _Content.Count Then
            If conn.State = ConnectionState.Closed Then conn.Open()
            Dim cmd As OleDbCommand = New OleDbCommand("Insert Into " & _Table & "(" & _Columns & ") Values(" & ValueColumn & ")", conn)
            ' MsgBox("Insert Into " & bang & "(" & caccot & ") Values(" & cotvalue & ")")
            For i = 0 To _Content.Count - 1
                cmd.Parameters.AddWithValue(Trim(ListColumn(i)), _Content(i))
            Next
            cmd.ExecuteNonQuery()
        Else
            MsgBox("Error!", vbInformation, "Error")
        End If
        If conn.State = ConnectionState.Open Then conn.Close()
    End Sub
    Public Sub EditPar(ByVal _Table As String, ByVal _Comlumns As String, ByVal _ListContent As ArrayList, ByVal _Cmd As String)
        'suapar("bang", "TieuDe,NoiDung", noidung,"Update bang set caccot = @caccot,...where ....")
        Dim ValueColumn As String = "@" & _Comlumns
        ValueColumn = Replace(ValueColumn, ",", ",@")
        Dim ListColumn As Array
        ListColumn = Split(ValueColumn, ",")
        If ListColumn.Length = _ListContent.Count Then
            If conn.State = ConnectionState.Closed Then conn.Open()
            Dim cmd As OleDbCommand = New OleDbCommand(_Cmd, conn)
            For i = 0 To _ListContent.Count - 1
                cmd.Parameters.AddWithValue(Trim(ListColumn(i)), _ListContent(i))
            Next
            cmd.ExecuteNonQuery()
        Else
            MsgBox("Error!", vbInformation, "Error")
        End If
        If conn.State = ConnectionState.Open Then conn.Close()
    End Sub
End Class
