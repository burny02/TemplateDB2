Option Explicit On
Imports System.Data
Imports System.Data.OleDb.OleDbConnection

Public Class MyCmbColumn
    Inherits DataGridViewComboBoxColumn

    Public TempOverclass As TemplateDB.OverClass
    Public SQLString As String = ""
    Public ReliantComboBox As ComboBox = Nothing
    Public Parent As DataGridView = Nothing
    Public ClmName As String = ""

    Public Sub RefreshDataTable()

        Dim comboValue As String = "'"
        Try
            comboValue = comboValue & ReliantComboBox.SelectedValue.ToString
        Catch ex As Exception
        End Try
        comboValue = comboValue & "'"
        Dim Dt As DataTable = TempOverclass.TempDataTable(SQLString & comboValue)
        Dim clm As MyCmbColumn = Parent.Columns(ClmName)
        clm.DataSource = Dt

    End Sub

End Class

Public Class MyCombo
    Public ComboName As String = ""
    Public Blank As Boolean = True
    Public FullDataSQL As String = ""
    Public WhichCmb As ComboBox = Nothing
    Public TableOfData As DataTable = Nothing
    Public LiveData As Boolean = False
    Public ActiveFilter As String = ""
    Public SecondColumnSQL As String = ""
    Public OnlyFilterableBy As New Collection
End Class


Public Class CentralFunctions
    Declare Function GetUserName Lib "advapi32.dll" Alias _
        "GetUserNameA" (ByVal lpBuffer As String,
        ByRef nSize As Integer) As Integer
    Public CurrentDataSet As DataSet = Nothing
    Public CurrentDataAdapter As OleDb.OleDbDataAdapter = Nothing
    Public CurrentBindingSource As BindingSource = Nothing
    Private ConnectString As String = Nothing
    Private AuditTable As String = Nothing
    Private UserTable As String = Nothing
    Private UserField As String = Nothing
    Private LockTable As String = Nothing
    Public ReadOnlyUser As Boolean = True
    Private Contact As String = Nothing
    Private con As OleDb.OleDbConnection
    Public CmdList As New List(Of OleDb.OleDbCommand)
    Private CurrentTrans As OleDb.OleDbTransaction = Nothing
    Public DataItemCollection As New Collection
    Public ComboCollection As New Collection
    Public ComboColumnCollection As New Collection

    Public Function SELECTCount(SQLCode As String) As Long
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
            CloseCon()
            Throw

        Finally
            'Close Off & Clean up
            CloseCon()
            dt = Nothing
            da = Nothing

        End Try
        SELECTCount = Counter

    End Function

    Public Sub AddToMassSQL(SQLCodeOrCmd As Object, Optional Audit As Boolean = True)
        'Execute many SQL Commands - No return. Execute submits them all

        Dim Cmd As OleDb.OleDbCommand = Nothing

        'Create connection & Command
        If TypeOf (SQLCodeOrCmd) Is String Then
            Cmd = New OleDb.OleDbCommand(SQLCodeOrCmd, con)

        ElseIf TypeOf (SQLCodeOrCmd) Is OleDb.OleDbCommand Then
            Cmd = New OleDb.OleDbCommand(SQLCodeOrCmd.CommandText, con)
        End If
        CmdList.Add(Cmd)

        If Audit = True Then
            Dim AuditPerson As String = vbNullString
            Dim AuditValues As String = vbNullString
            Dim ActionTable As String = vbNullString
            Dim AuditAction As String = vbNullString

            GetSQLAudit(Cmd.CommandText, AuditAction, ActionTable, AuditPerson, AuditValues)
            Dim AuditSQLCode As String = "'" & AuditPerson & "','" & AuditAction &
                    "','" & ActionTable & "','" & AuditValues & "'"
            AuditSQLCode = "INSERT INTO " & AuditTable & " ([Person], [Action], [TName], [NValue]) VALUES (" & AuditSQLCode & ")"
            Dim AuditCmd = New OleDb.OleDbCommand(AuditSQLCode, con)
            CmdList.Add(AuditCmd)
        End If

    End Sub

    Public Sub ExecuteMassSQL(Optional AddTransactionOnly As Boolean = False)
        'Executes all commands in CmdList
        'Option to not commit - For simulatenous transactions

        If ReadOnlyUser = True Then
            MsgBox("Read only permissions have been granted")
            Exit Sub
        End If

        Dim Attempts As Integer = 0

        'Open connection - assign a transaction
        OpenCon()


        If CurrentTrans Is Nothing Then CurrentTrans = con.BeginTransaction(IsolationLevel.ReadCommitted)

        Dim i As Integer = 0

        Try
            'Add all to transaction & Execute
            Do While i < CmdList.Count
                CmdList(i).Transaction = CurrentTrans
                CmdList(i).ExecuteNonQuery()
                i = i + 1
            Loop

            'If wanted, commit changes
            If AddTransactionOnly = False Then TryCommit()

        Catch ex As Exception
            Call TryRollBack()
            CloseCon()
            Throw

        Finally
            'Close Off & Clean up
            If AddTransactionOnly = False Then
                CloseCon()
                CurrentTrans = Nothing
            End If
            CmdList.Clear()


        End Try

    End Sub

    Public Sub ExecuteSQL(SQLCodeOrCmd As Object) 'Execute a SQL Command - No return

        If ReadOnlyUser = True Then
            MsgBox("Read only permissions have been granted")
            Exit Sub
        End If

        Dim Cmd As OleDb.OleDbCommand = Nothing

        'Create connection & Command
        If TypeOf (SQLCodeOrCmd) Is String Then
            Cmd = New OleDb.OleDbCommand(SQLCodeOrCmd, con)
        ElseIf TypeOf (SQLCodeOrCmd) Is OleDb.OleDbCommand Then
            Cmd = New OleDb.OleDbCommand(SQLCodeOrCmd.CommandText, con)
        End If

        Dim Attempts As Integer = 0
        Dim AuditPerson As String = vbNullString
        Dim AuditValues As String = vbNullString
        Dim ActionTable As String = vbNullString
        Dim AuditAction As String = vbNullString

        'Open connection - assign a transaction
        OpenCon()
        If CurrentTrans Is Nothing Then CurrentTrans = con.BeginTransaction(IsolationLevel.ReadCommitted)
        Cmd.Transaction = CurrentTrans


        Try
            'Audit
            GetSQLAudit(Cmd.CommandText, AuditAction, ActionTable, AuditPerson, AuditValues)
            Dim AuditSQLCode As String = "'" & AuditPerson & "','" & AuditAction &
                    "','" & ActionTable & "','" & AuditValues & "'"
            AuditSQLCode = "INSERT INTO " & AuditTable & " ([Person], [Action], [TName], [NValue]) VALUES (" & AuditSQLCode & ")"
            Dim AuditCmd = New OleDb.OleDbCommand(AuditSQLCode, con)
            AuditCmd.Transaction = CurrentTrans
            AuditCmd.ExecuteNonQuery()

            'Set the action as a transaction
            Cmd.Transaction = CurrentTrans
            'Execute SQL Command
            Cmd.ExecuteNonQuery()
            'If OK. Commit changes
            Call TryCommit()

        Catch ex As Exception
            Call TryRollBack()
            CloseCon()
            Throw

        Finally
            'Close Off & Clean up
            CloseCon()
            Cmd = Nothing
            CurrentTrans = Nothing


        End Try

    End Sub

    Private Sub TryCommit()

        If ReadOnlyUser = True Then
            MsgBox("Read only permissions have been granted")
            Exit Sub
        End If

        Dim Attempts As Integer = 0

        'Loop and try to commit changes with a delay
        Do While Attempts <= 3
            Try
                OpenCon()
                CurrentTrans.Commit()
                Exit Sub

            Catch ex As OleDb.OleDbException
                Threading.Thread.Sleep(10000)
                Attempts = Attempts + 1
                If Attempts = 4 Then
                    If MsgBox("A record appears to be locked by another user" _
                              & vbNewLine & vbNewLine & "Do you want to try again?", vbYes) = vbYes Then
                        'Back to start       
                        Attempts = 0
                    Else
                        Call TryRollBack()
                    End If

                End If

            Catch ex2 As Exception
                Throw
                Call TryRollBack()

            End Try

        Loop

    End Sub

    Private Sub TryRollBack()

        'Loop and try to RollBack a transaction with delay
        Dim Attempts As Integer = 0


        Do While Attempts <= 3

            Try
                OpenCon()
                CurrentTrans.Rollback()
                Exit Sub

            Catch ex As Exception
                Threading.Thread.Sleep(10000)
                Attempts = Attempts + 1
                If Attempts = 4 Then
                    Throw
                    Exit Sub
                End If

            End Try

        Loop

    End Sub

    Public Sub CreateDataSet(SQLCode As String, BindSource As BindingSource, ctl As Object)
        'Create a new dataset, set a bindining source and object to that binding source


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
            If BindSource IsNot Nothing Then BindSource.DataSource = CurrentDataSet.Tables(0)
            If ctl IsNot Nothing Then ctl.DataSource = BindSource

        Catch ex As Exception
            CloseCon()
            Throw

        Finally

            'Close off & Clean up
            CloseCon()

        End Try

    End Sub

    Public Sub UpdateBackend(ctl As Object, Optional DisplayMessage As Boolean = True)
        'Saving function to update access backend

        Dim TempFilter As String = CurrentDataSet.Tables(0).DefaultView.RowFilter

        If ReadOnlyUser = True Then
            MsgBox("Read only permissions have been granted")
            Exit Sub
        End If

        'Is the data dirty / has errors that have auto-undone
        If CurrentDataSet.HasChanges() = False Then
            If DisplayMessage = True Then MsgBox("Errors present/No changes to upload")
            Call Refresher(ctl)
            CurrentDataSet.Tables(0).DefaultView.RowFilter = TempFilter
            Exit Sub
        End If

        'Open Connection & Set transaction
        OpenCon()
        If CurrentTrans Is Nothing Then CurrentTrans = con.BeginTransaction(IsolationLevel.ReadCommitted)


        If IsNothing(CurrentDataAdapter.UpdateCommand) And
        IsNothing(CurrentDataAdapter.InsertCommand) And
        IsNothing(CurrentDataAdapter.DeleteCommand) Then
            Call TryRollBack()
            CloseCon()
            Throw New Exception("Attempted to 'UpdateBackend' with no commands set. 
                                Possible autogenerated commands for multiple tables - Control " & ctl.name)

        End If

        If Not IsNothing(CurrentDataAdapter.UpdateCommand) Then CurrentDataAdapter.UpdateCommand.Transaction = CurrentTrans
        If Not IsNothing(CurrentDataAdapter.InsertCommand) Then CurrentDataAdapter.InsertCommand.Transaction = CurrentTrans
        If Not IsNothing(CurrentDataAdapter.DeleteCommand) Then CurrentDataAdapter.DeleteCommand.Transaction = CurrentTrans

        Try

            CurrentBindingSource.EndEdit()

            'AUDIT
            Dim Operation As String = vbNullString
            Dim Table As String = vbNullString
            Dim Person As String = GetUserName()
            Dim AuditValues As String = vbNullString
            Dim Version As System.Data.DataRowVersion


            For Each row As DataRow In CurrentDataSet.Tables(0).Rows

                If row.RowState = DataRowState.Detached _
                    Or row.RowState = DataRowState.Unchanged Then Continue For

                If row.RowState = DataRowState.Added Then
                    Operation = "INSERT"
                    Call GetSQLAudit(CurrentDataAdapter.InsertCommand.CommandText.ToUpper,, Table)
                    Version = DataRowVersion.Current

                ElseIf row.RowState = DataRowState.Modified Then
                    Operation = "UPDATE"
                    Call GetSQLAudit(CurrentDataAdapter.UpdateCommand.CommandText.ToUpper,, Table)
                    Version = DataRowVersion.Current

                ElseIf row.RowState = DataRowState.Deleted Then
                    Operation = "DELETE"
                    Call GetSQLAudit(CurrentDataAdapter.DeleteCommand.CommandText.ToUpper,, Table)
                    Version = DataRowVersion.Original

                End If

                For Each col As DataColumn In CurrentDataSet.Tables(0).Columns
                    AuditValues = AuditValues & Replace(col.ColumnName.ToString, "'", "") & "="
                    AuditValues = AuditValues & Replace(row.Item(col, Version).ToString, "'", "") & ","
                Next

                Dim CombineInsert As String = "'" & Person & "','" & Operation &
                    "','" & Table & "','" & AuditValues & "'"
                AddToMassSQL("INSERT INTO " & AuditTable & " ([Person], [Action], [TName], [NValue]) VALUES (" & CombineInsert & ")", False)
                AuditValues = vbNullString

            Next

            'Add to transaction only
            ExecuteMassSQL(True)
            'Use dataadapter to update the backend (Commands already set)
            CurrentDataAdapter.Update(CurrentDataSet)
            Call TryCommit()

            If DisplayMessage = True Then MsgBox("Table Updated")
            'Remove any error messages & accept changes
            CurrentDataSet.AcceptChanges()
            Call Refresher(ctl)
            CurrentDataSet.Tables(0).DefaultView.RowFilter = TempFilter

        Catch ex As Exception
            Call TryRollBack()
            CloseCon()
            Throw


        Finally
            'Close off & clean up
            CloseCon()
            CurrentTrans = Nothing

        End Try

    End Sub

    Protected Sub Quitter()

        On Error Resume Next

        CloseCon()
        Application.Exit()

    End Sub

    Public Function UnloadData() As Boolean
        'Close down currnt dataset, dataadapter & bindingsource

        'Variable if user wants to save
        Dim Cancel As Boolean = False

        'Is there currently a dataset to close?
        If IsNothing(CurrentDataSet) Then
            UnloadData = False
            Exit Function
        End If

        Try

            'Is the dataset dirty?
            If CurrentDataSet.HasChanges() Then

                'Ask user if they want to proceed and lose data?
                If (MsgBox("Changes To data will be lost unless saved first. Do you wish To discard changes?", vbYesNo) = vbNo) Then Cancel = True

            End If


            'If want to continue, clear all current data items
            If Cancel = False Then
                CurrentDataSet = Nothing
                CurrentDataAdapter = Nothing
                CurrentBindingSource = Nothing
            End If

        Catch ex As Exception
            CloseCon()
            Throw
        Finally
            'Pass back whether clean up happened
            UnloadData = Cancel
        End Try

    End Function

    Public Function MultiTempDataTable(ParamArray SQLCode() As String) As DataTable()
        'Create a temporary dataset for things such as combo box which arent based on the initial query

        Try
            'Open connection
            OpenCon()
            'New temporary data adapter and dataset
            Dim i As Integer = 0
            Dim TempDT() As DataTable
            Do While i < SQLCode.Count
                ReDim Preserve TempDT(i)
                TempDT(i) = New DataTable
                Dim TempDataAdapter = New OleDb.OleDbDataAdapter(SQLCode(i), con)
                TempDataAdapter.Fill(TempDT(i))
                i += 1
            Loop
            ReDim Preserve TempDT(i)
            MultiTempDataTable = TempDT

        Catch ex As Exception
            CloseCon()
            Throw

        Finally

            'Close off & Clean up
            CloseCon()

        End Try

    End Function

    Public Function TempDataTable(SQLCode As String) As DataTable
        'Create a temporary dataset for things such as combo box which arent based on the initial query

        Try
            'Open connection
            OpenCon()
            'New temporary data adapter and dataset
            Dim TempDataAdapter = New OleDb.OleDbDataAdapter(SQLCode, con)
            TempDataTable = New DataTable
            'Use temp adapter to fill temp dataset
            TempDataAdapter.Fill(TempDataTable)

        Catch ex As Exception
            CloseCon()
            Throw

        Finally

            'Close off & Clean up
            CloseCon()

        End Try

    End Function

    Public Function CreateCSVString(SQLCode As String) As String

        Dim da As New OleDb.OleDbDataAdapter(SQLCode, ConnectString)
        Dim dt As New DataTable
        Dim Output As String = vbNullString

        Try
            da.Fill(dt)

            For Each row As DataRow In dt.Rows

                If Not IsNothing(row.Item(0)) And Not IsDBNull(row.Item(0)) Then
                    Output = Output & row.Item(0).ToString & ", "
                End If

            Next

            If Output <> vbNullString Then Output = Left(Output, Len(Output) - 1)

        Catch ex As Exception
            CloseCon()
            Throw

        Finally
            CreateCSVString = Output
            dt = Nothing
            da = Nothing

        End Try

    End Function

    Public Sub Refresher(DataItem As Object)

        Try
            Call CreateDataSet(CurrentDataAdapter.SelectCommand.CommandText, CurrentBindingSource, DataItem)
            DataItem.Parent.Refresh()
        Catch ex As Exception
            CloseCon()
            Throw
        End Try

    End Sub

    Public Sub LoginCheck()

        Dim SQLString As String = "SELECT * FROM " & UserTable & " WHERE " & UserField & "='" & GetUserName() & "'"
        Dim ErrorMessage As String = "You do not have permission to use this database. Please contact David Burnside or " & Contact

        If SELECTCount(SQLString) = 0 Then
            MsgBox(ErrorMessage)
            Call Quitter()
        Else
            SQLString = "SELECT * FROM " & UserTable & " WHERE " & UserField & "='" & GetUserName() & "'" &
                    " AND [Read]=True"
            If SELECTCount(SQLString) <> 0 Then
                ReadOnlyUser = True
            Else
                ReadOnlyUser = False
            End If
        End If

    End Sub

    Public Sub LockCheck()

        Dim SQLString As String = "SELECT * FROM " & LockTable
        Dim ErrorMessage As String = "The database is currently locked. Please contact David Burnside"

        If SELECTCount(SQLString) <> 0 Then
            If GetUserName() <> "d.burnside" Then
                MsgBox(ErrorMessage)
                Call Quitter()
            Else
                MsgBox("Database is locked")

            End If
        End If

    End Sub

    Public Function GetUserName() As String

        Dim iReturn As Integer
        Dim userName As String
        userName = New String(CChar(" "), 50)
        iReturn = GetUserName(userName, 50)
        GetUserName = userName.Substring(0, userName.IndexOf(Chr(0)))

    End Function

    Public Sub SetPrivate(UserTbl As String,
                          UserFld As String,
                          LockTbl As String,
                          ContactPerson As String,
                          ConnectionString As String,
                          AuditTbl As String)

        AuditTable = AuditTbl
        UserTable = UserTbl
        UserField = UserFld
        LockTable = LockTbl
        Contact = ContactPerson
        ConnectString = ConnectionString
        con = New OleDb.OleDbConnection(ConnectString)

    End Sub

    Public Sub SetCommandConnection(Optional Command As OleDb.OleDbCommand = Nothing)

        On Error Resume Next
        CurrentDataAdapter.InsertCommand.Connection = con
        CurrentDataAdapter.UpdateCommand.Connection = con
        CurrentDataAdapter.DeleteCommand.Connection = con
        Command.Connection = con

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

    Public Function SQLDate(varDate As Object) As String

        If IsDate(varDate) Then
            If DateValue(varDate) = varDate Then
                SQLDate = Format$(varDate, "\#MM\/dd\/yyyy\#")
            Else
                SQLDate = Format$(varDate, "\#MM\/dd\/yyyy HH\:mm\:ss\#")
            End If
        Else
            SQLDate = ""
        End If

        'ALWAYS SQLCOMMAND date as a string like #1/1/2000# - The # tells it that is it american format
    End Function

    Private Sub GetSQLAudit(ByVal SQLCode As String,
                                 Optional ByRef ActionVariable As String = vbNullString,
                                 Optional ByRef TableVariable As String = vbNullString,
                                 Optional ByRef PersonVariable As String = vbNullString,
                                 Optional ByRef ValuesVariable As String = vbNullString)

        SQLCode = Replace(SQLCode, "'", "")

        PersonVariable = GetUserName()
        ActionVariable = Left(SQLCode, 6)

        'Get Table info
        Select Case ActionVariable

            Case "SELECT"
                Exit Sub

            Case "INSERT"
                SQLCode = Replace(SQLCode.ToUpper, "INSERT INTO", "")
                SQLCode = Trim(SQLCode)
                Dim FirstLocation As Long = InStr(SQLCode, " ")
                If FirstLocation > InStr(SQLCode, "(") Then FirstLocation = InStr(SQLCode, "(")
                TableVariable = Trim(Left(SQLCode, FirstLocation - 1))

            Case "UPDATE"
                SQLCode = Replace(SQLCode.ToUpper, "UPDATE", "")
                Dim SetLocation As Long = InStr(SQLCode, "SET")
                TableVariable = Trim(Left(SQLCode, SetLocation - 1))

            Case "DELETE"
                SQLCode = Replace(SQLCode.ToUpper, "DELETE FROM", "")
                Dim WhereLocation As Long = InStr(SQLCode, "WHERE")
                TableVariable = Trim(Left(SQLCode, WhereLocation - 1))

        End Select

        'Get Values info
        SQLCode = Trim(Replace(SQLCode, TableVariable, ""))
        ValuesVariable = SQLCode


    End Sub

    Protected Sub ComboKeyDown(sender As Object, e As KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub

    Protected Sub DataGridViewDataError(sender As Object, e As DataGridViewDataErrorEventArgs)
        Try
            e.Cancel = False
            Call ErrorHandler(sender, e)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Protected Sub GridComboBox(sender As Object, e As DataGridViewEditingControlShowingEventArgs)

        If e.Control.GetType IsNot GetType(DataGridViewComboBoxEditingControl) Then Exit Sub
        SendKeys.Send("{F4}")

        Dim cmbBx As ComboBox = e.Control

        If cmbBx IsNot Nothing Then
            RemoveHandler cmbBx.DropDownClosed, AddressOf ComboBoxCell_DropDownClosed
            AddHandler cmbBx.DropDownClosed, AddressOf ComboBoxCell_DropDownClosed
        End If

    End Sub

    Public Sub ComboBoxCell_DropDownClosed(sender As Object, e As EventArgs)

        Dim cmbBx As DataGridViewComboBoxEditingControl = sender
        SendKeys.Send("{TAB}")

    End Sub

    Protected Sub FormClosing(sender As Object, e As FormClosingEventArgs)
        If UnloadData() = True Then e.Cancel = True
        Call Quitter()
    End Sub

    Protected Sub ErrorHandler(sender As Object, e As Object)

        If Not e.RowIndex > 0 Then Exit Sub

        Dim Obj As Object

        Try
            If TypeOf (sender) Is DataGridView Then
                Obj = CType(sender, DataGridView)
                Obj.Rows(e.RowIndex).Cells(e.ColumnIndex).ErrorText = e.exception.message
            End If
        Catch ex As Exception
            Throw
        End Try

    End Sub

    Public Sub ResetCollection()

        For Each cmb As FilterCombo In ComboCollection
            cmb.CmbPointer = Nothing
        Next

        ComboCollection.Clear()
        ComboColumnCollection.Clear()

        For Each control In DataItemCollection
            If (TypeOf control Is DataGridView) Then
                Try
                    control.Columns.Clear()
                    control.DataSource = Nothing
                Catch ex As Exception
                End Try
            End If

        Next

    End Sub


    Protected Sub ControlWheelScroll(sender As DataGridView, e As MouseEventArgs)

        sender.Focus()

    End Sub

    Public Function CSVColumn(ColumnName As DataColumn)

        Dim TempString As String = ""

        For Each drv As DataRowView In ColumnName.Table.DefaultView
            If drv.Row.RowState = DataRowState.Deleted Then Continue For
            If InStr(TempString, drv.Row.Item(ColumnName, DataRowVersion.Current).ToString()) <> 0 Then Continue For
            If IsNumeric(drv.Row.Item(ColumnName, DataRowVersion.Current)) Then
                TempString = TempString & drv.Row.Item(ColumnName, DataRowVersion.Current).ToString() & ","
            Else
                TempString = TempString & "'" & drv.Row.Item(ColumnName, DataRowVersion.Current).ToString() & "',"
            End If
        Next

        If Right(TempString, 1) = "," Then TempString = Left(TempString, Len(TempString) - 1)

        Return TempString

    End Function

    Public Function SetUpNewComboColumn(SQLCode As String, RelyingComboBox As ComboBox,
                   ValMember As String, DispMember As String, DattaPropertyName As String,
                   WhatHeader As String, Parent As DataGridView, clmName As String) As MyCmbColumn

        Dim MyColumn As New MyCmbColumn

        MyColumn.SQLString = SQLCode
        MyColumn.ReliantComboBox = RelyingComboBox
        MyColumn.ValueMember = ValMember
        MyColumn.DisplayMember = DispMember
        MyColumn.DataPropertyName = DattaPropertyName
        MyColumn.HeaderText = WhatHeader
        MyColumn.TempOverclass = Me
        MyColumn.Parent = Parent
        MyColumn.Name = clmName
        MyColumn.ClmName = clmName
        Parent.Columns.Add(MyColumn)
        MyColumn.RefreshDataTable()

        Try
            ComboColumnCollection.Remove(MyColumn.HeaderText)
        Catch ex As Exception
        End Try
        ComboColumnCollection.Add(MyColumn, MyColumn.HeaderText)

        Return MyColumn

    End Function
End Class