﻿Imports System.Data
Imports System.Data.OleDb.OleDbConnection
Public Class Template
    Declare Function GetUserName Lib "advapi32.dll" Alias _
        "GetUserNameA" (ByVal lpBuffer As String, _
        ByRef nSize As Integer) As Integer
    Public CurrentDataSet As DataSet = Nothing
    Public CurrentDataAdapter As OleDb.OleDbDataAdapter = Nothing
    Public CurrentBindingSource As BindingSource = Nothing
    Private ConnectString As String = Nothing
    Public Login As Boolean = False
    Private UserTable As String = Nothing
    Private UserField As String = Nothing
    Private LockTable As String = Nothing
    Private ActiveUsersTable As String = Nothing
    Private Contact As String = Nothing
    Private con As OleDb.OleDbConnection

    Public Function QueryTest(SQLCode As String) As Long
        'Execute a SQL Command and return the number of records

        Dim Counter As Long
        Dim dt As New DataTable
        Dim da As New OleDb.OleDbDataAdapter(SQLCode, ConnectString)

        Try
            'Connect
            da.Fill(dt)
            'Assign
            Counter = dt.Rows.Count

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
            'Close Off & Clean up
            dt = Nothing
            da = Nothing

        End Try
        QueryTest = Counter

    End Function

    Public Sub ExecuteSQL(SQLCode As String)
        'Execute a SQL Command - No return

        Dim ErrorMessage As String = vbNullString

        'Create connection & Command
        Dim cmd As New OleDb.OleDbCommand(SQLCode, con)
        Dim trans As OleDb.OleDbTransaction
        Dim Attempts As Integer = 0

        'Open connection - assign a transaction
        OpenCon()
        UpdateActiveUsers(True)
        trans = con.BeginTransaction(IsolationLevel.ReadCommitted)
        cmd.Transaction = trans


        Try
            
            'Set the action as a transaction
            cmd.Transaction = trans
            'Execute SQL Command
            cmd.ExecuteNonQuery()
            'If OK. Commit changes
            Call TryCommit(trans)


        Catch ex As Exception
            ErrorMessage = ex.Message
            Call TryRollBack(trans)

        Finally
            'Close Off & Clean up
            UpdateActiveUsers(False)
            CloseCon()
            cmd = Nothing
            trans = Nothing
            If ErrorMessage <> vbNullString Then MsgBox(ErrorMessage)

        End Try

    End Sub

    Private Sub TryCommit(Trans As OleDb.OleDbTransaction)

        Dim Attempts As Integer = 0

        'Loop and try to commit changes with a delay
        Do While Attempts <= 3
            Try
                OpenCon()
                trans.Commit()
                Exit Sub

            Catch ex As OleDb.OleDbException
                Threading.Thread.Sleep(10000)
                Attempts = Attempts + 1
                MsgBox(Attempts)
                If Attempts = 4 Then
                    If MsgBox("A record appears to be locked by another user" _
                              & vbNewLine & vbNewLine & "Do you want to try again?", vbYes) = vbYes Then
                        'Back to start       
                        Attempts = 0
                    Else
                        Call TryRollBack(Trans)
                    End If

                End If

            Catch ex2 As Exception
                MsgBox("Failed to commit changes - " & ex2.Message)
                Call TryRollBack(Trans)

            End Try

        Loop

    End Sub

    Private Sub TryRollBack(Trans As OleDb.OleDbTransaction)

        'Loop and try to RollBack a transaction with delay
        Dim Attempts As Integer = 0


        Do While Attempts <= 3

            Try
                OpenCon()
                Trans.Rollback()
                Exit Sub

            Catch ex As Exception
                Threading.Thread.Sleep(10000)
                Attempts = Attempts + 1
                If Attempts = 4 Then
                    MsgBox("Failed to rollback changes - " & ex.Message)
                    Exit Sub
                End If

            End Try

        Loop

    End Sub

    Public Sub CreateDataSet(SQLCode As String, BindSource As BindingSource, ctl As Object)
        'Create a new dataset, set a bindining source and object to that binding source

        Dim ErrorMessage As String = vbNullString

        Try
            'Open connection
            OpenCon()

            'Create New Dataset & adapter
            CurrentDataAdapter = New OleDb.OleDbDataAdapter(SQLCode, con)
            CurrentDataSet = New DataSet()
            CurrentBindingSource = BindSource

            'Use adapter to fill dataset
            CurrentDataAdapter.Fill(CurrentDataSet)

            'Set bindsource datasource as dataset, set object datasource as bindsource
            BindSource.DataSource = CurrentDataSet.Tables(0)
            ctl.DataSource = BindSource

        Catch ex As Exception
            ErrorMessage = ex.Message

        Finally

            'Close off & Clean up
            CloseCon()
            If ErrorMessage <> vbNullString Then MsgBox(ErrorMessage)

        End Try

    End Sub

    Public Sub UpdateBackend(ctl As Object)
        'Saving function to update access backend

        Dim ErrorMessage As String = vbNullString

        'Is the data dirty / has errors that have auto-undone
        If CurrentDataSet.HasChanges() = False Then
            MsgBox("Errors present/No changes to upload")
            Exit Sub
        End If

        'Open Connection & Set transaction
        OpenCon()
        UpdateActiveUsers(True)
        Dim trans As OleDb.OleDbTransaction
        trans = con.BeginTransaction(IsolationLevel.ReadCommitted)

        If Not IsNothing(CurrentDataAdapter.UpdateCommand) Then CurrentDataAdapter.UpdateCommand.Transaction = trans
        If Not IsNothing(CurrentDataAdapter.InsertCommand) Then CurrentDataAdapter.InsertCommand.Transaction = trans
        If Not IsNothing(CurrentDataAdapter.DeleteCommand) Then CurrentDataAdapter.DeleteCommand.Transaction = trans

        Try

            CurrentBindingSource.EndEdit()

            'Use dataadapter to update the backend (Commands already set)
            CurrentDataAdapter.Update(CurrentDataSet)
            Call TryCommit(trans)
            MsgBox("Table Updated")
            'Remove any error messages & accept changes
            CurrentDataSet.AcceptChanges()
            Call Refresher(ctl)

        Catch ex As Exception
            ErrorMessage = ex.Message
            Call TryRollBack(trans)

        Finally
            'Close off & clean up
            UpdateActiveUsers(False)
            CloseCon()
            trans = Nothing
            If ErrorMessage <> vbNullString Then MsgBox(ErrorMessage)

        End Try

    End Sub

    Public Function UnloadData() As Boolean
        'Close down currnt dataset, dataadapter & bindingsource

        'Variable if user wants to save
        Dim Cancel As Boolean = False
        Dim ErrorMessage As String = vbNullString

        'Is there currently a dataset to close?
        If IsNothing(CurrentDataSet) Then
            UnloadData = False
            Exit Function
        End If

        Try

            'Is the dataset dirty?
            If CurrentDataSet.HasChanges() Then

                'Ask user if they want to proceed and lose data?
                If (MsgBox("Changes to data will be lost unless saved first. Do you wish to discard changes?", vbYesNo) = vbNo) Then Cancel = True

            End If


            'If want to continue, clear all current data items
            If Cancel = False Then
                CurrentDataSet = Nothing
                CurrentDataAdapter = Nothing
                CurrentBindingSource = Nothing
            End If

        Catch ex As Exception
            ErrorMessage = ex.Message
        Finally
            'Pass back whether clean up happened
            UnloadData = Cancel
            If ErrorMessage <> vbNullString Then MsgBox(ErrorMessage)
        End Try

    End Function

    Public Function TempDataSet(SQLCode As String) As DataSet
        'Create a temporary dataset for things such as combo box which arent based on the initial query

        Dim ErrorMessage As String = vbNullString

        Try
            'Open connection
            OpenCon()
            'New temporary data adapter and dataset
            Dim TempDataAdapter = New OleDb.OleDbDataAdapter(SQLCode, con)
            TempDataSet = New DataSet()
            'Use temp adapter to fill temp dataset
            TempDataAdapter.Fill(TempDataSet)

        Catch ex As Exception
            errormessage = ex.Message
            TempDataSet = Nothing
        Finally

            'Close off & Clean up
            CloseCon()
            If ErrorMessage <> vbNullString Then MsgBox(ErrorMessage)

        End Try

    End Function

    Public Function CreateCSVString(SQLCode As String) As String

        Dim da As New OleDb.OleDbDataAdapter(SQLCode, ConnectString)
        Dim dt As New DataTable
        Dim Output As String = vbNullString
        Dim ErrorMessage As String = vbNullString

        Try
            da.Fill(dt)

            For Each row As DataRow In dt.Rows

                If Not IsNothing(row.Item(0)) And Not IsDBNull(row.Item(0)) Then
                    Output = Output & row.Item(0).ToString & ","
                End If

            Next

            Output = Left(Output, Len(Output) - 1)

        Catch ex As Exception
            ErrorMessage = ex.Message

        Finally
            CreateCSVString = Output
            dt = Nothing
            da = Nothing
            If ErrorMessage <> vbNullString Then MsgBox(ErrorMessage)
        End Try

    End Function

    Public Sub Refresher(DataItem As Object)

        Try
            Call CreateDataSet(CurrentDataAdapter.SelectCommand.CommandText, CurrentBindingSource, DataItem)
            DataItem.Parent.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Public Sub LoginCheck()

        Dim SQLString As String = "SELECT * FROM " & UserTable & " WHERE " & UserField & "='" & GetUserName() & "'"
        Dim ErrorMessage As String = "You do not have permission to use this database. Please contact David Burnside or " & Contact

        If QueryTest(SQLString) = 0 Then
            MsgBox(ErrorMessage)
            Call Quitter(True)
        Else
            Login = True
        End If

    End Sub

    Public Sub LockCheck()

        Dim SQLString As String = "SELECT * FROM " & LockTable
        Dim ErrorMessage As String = "The database is currently locked. Please contact David Burnside"

        If QueryTest(SQLString) <> 0 Then
            If GetUserName <> "d.burnside" Then
                MsgBox(ErrorMessage)
                Call Quitter(True)
            Else
                MsgBox("Database is locked")

            End If
        End If


    End Sub

    Private Sub LockUnlock()

        Dim SQLTest As String = "SELECT * FROM " & LockTable
        Dim SQLInsert As String = "INSERT INTO " & LockTable & " Values ('" & GetUserName() & "')"
        Dim SQLDelete As String = "DELETE * FROM" & LockTable
        Dim Message As String

        If QueryTest(SQLTest) = 0 Then
            ExecuteSQL(SQLInsert)
            Message = "Locked"
        Else
            ExecuteSQL(SQLDelete)
            Message = "Unlocked"
        End If

        MsgBox(Message)

    End Sub

    Public Function GetUserName() As String
        Dim iReturn As Integer
        Dim userName As String
        userName = New String(CChar(" "), 50)
        iReturn = GetUserName(userName, 50)
        GetUserName = userName.Substring(0, userName.IndexOf(Chr(0)))
    End Function

    Public Sub Quitter(Optional CloseAnyway As Boolean = False)

        On Error Resume Next

        If Login = False Or CloseAnyway = True Then
            CloseCon()
            UpdateActiveUsers(False)
            Application.Exit()
        End If

    End Sub

    Public Sub SetPrivate(UserTbl As String, _
                          UserFld As String, _
                          LockTbl As String, _
                          ContactPerson As String, _
                          ConnectionString As String, _
                          ActiveUsersTbl As String)

        UserTable = UserTbl
        UserField = UserFld
        LockTable = LockTbl
        Contact = ContactPerson
        ConnectString = ConnectionString
        con = New OleDb.OleDbConnection(ConnectString)
        ActiveUsersTable = ActiveUsersTbl


    End Sub

    Public Sub SetCommandConnection(Optional Command As OleDb.OleDbCommand = Nothing)

        On Error Resume Next
        CurrentDataAdapter.InsertCommand.Connection = con
        CurrentDataAdapter.UpdateCommand.Connection = con
        CurrentDataAdapter.DeleteCommand.Connection = con
        Command.Connection = con

    End Sub

    Public Sub ErrorHandler(sender As Object, e As Object)

        Dim Obj As Object

        Try
            If TypeOf (sender) Is DataGridView Then
                Obj = CType(sender, DataGridView)
                Obj.Rows(e.RowIndex).Cells(e.ColumnIndex).ErrorText = e.exception.message
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub

    Public Function NewDataAdapter(SQLCode As String) As OleDb.OleDbDataAdapter
        NewDataAdapter = New OleDb.OleDbDataAdapter(SQLCode, ConnectString)
    End Function

    Public Sub OpenCon()
        If (con.State = ConnectionState.Closed) Then con.Open()
    End Sub

    Public Sub CloseCon()
        If (con.State = ConnectionState.Open) Then con.Close()
    End Sub

    Private Sub Auditter(DataAdapter As OleDb.OleDbDataAdapter)

        Dim Person As String = GetUserName()
        Dim Action As String = vbNullString

        If Not IsNothing(CurrentDataAdapter.UpdateCommand) Then
            Action = "UPDATE"
        End If
        If Not IsNothing(CurrentDataAdapter.InsertCommand) Then
            Action = "INSERT"
        End If
        If Not IsNothing(CurrentDataAdapter.DeleteCommand) Then
            Action = "DELETE"
        End If

    End Sub

    Private Sub UpdateActiveUsers(Insert As Boolean)

        Dim cmd As New OleDb.OleDbCommand("DELETE * FROM " & ActiveUsersTable & " WHERE User='" & GetUserName & "'", con)
        Dim cmd2 As New OleDb.OleDbCommand("INSERT INTO " & ActiveUsersTable & " VALUES ('" & GetUserName & "')", con)

        OpenCon()
        If Insert = True Then
            cmd.ExecuteNonQuery()
            cmd2.ExecuteNonQuery()
        Else
            cmd.ExecuteNonQuery()
        End If

    End Sub

    Public Sub SingleClick(sender As Object, e As DataGridViewCellEventArgs)
        Dim dgv As DataGridView = CType(sender, DataGridView)

        If dgv(e.ColumnIndex, e.RowIndex).EditType.ToString() = "System.Windows.Forms.DataGridViewComboBoxEditingControl" Then
            SendKeys.Send("{F4}")
        End If
    End Sub
End Class
