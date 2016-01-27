Option Explicit On
Imports System.Data
Imports System.Data.OleDb.OleDbConnection


Public Class CentralFunctions
    Declare Function GetUserName Lib "advapi32.dll" Alias _
        "GetUserNameA" (ByVal lpBuffer As String,
        ByRef nSize As Integer) As Integer
    Public CurrentDataSet As DataSet = Nothing
    Public CurrentDataAdapter As OleDb.OleDbDataAdapter = Nothing
    Public CurrentBindingSource As BindingSource = Nothing
    Private ConnectString As String = Nothing
    Private UserTable As String = Nothing
    Private UserField As String = Nothing
    Private LockTable As String = Nothing
    Private ActiveUsersTable As String = Nothing
    Private Contact As String = Nothing
    Private con As OleDb.OleDbConnection
    Public CmdList As New List(Of OleDb.OleDbCommand)
    Private CurrentTrans As OleDb.OleDbTransaction = Nothing
    Public DataItemCollection As New Collection

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
            DefaultError(ex)

        Finally
            'Close Off & Clean up
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
            Dim AuditTable As String = vbNullString
            Dim AuditValues As String = vbNullString
            Dim AuditAction As String = vbNullString

            GetSQLAudit(Cmd.CommandText, AuditAction, AuditTable, AuditPerson, AuditValues)
            Dim AuditSQLCode As String = "'" & AuditPerson & "','" & AuditAction &
                    "','" & AuditTable & "','" & Left(AuditValues, 254) & "'"
            AuditSQLCode = "INSERT INTO AUDIT ([Person], [Action], [TName], [NValue]) VALUES (" & AuditSQLCode & ")"
            Dim AuditCmd = New OleDb.OleDbCommand(AuditSQLCode, con)
            CmdList.Add(AuditCmd)
        End If

    End Sub

    Public Sub ExecuteMassSQL(Optional AddTransactionOnly As Boolean = False)
        'Executes all commands in CmdList
        'Option to not commit - For simulatenous transactions

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

        Dim Cmd As OleDb.OleDbCommand = Nothing

        'Create connection & Command
        If TypeOf (SQLCodeOrCmd) Is String Then
            Cmd = New OleDb.OleDbCommand(SQLCodeOrCmd, con)
        ElseIf TypeOf (SQLCodeOrCmd) Is OleDb.OleDbCommand Then
            Cmd = New OleDb.OleDbCommand(SQLCodeOrCmd.CommandText, con)
        End If

        Dim Attempts As Integer = 0
        Dim AuditPerson As String = vbNullString
        Dim AuditTable As String = vbNullString
        Dim AuditValues As String = vbNullString
        Dim AuditAction As String = vbNullString

        'Open connection - assign a transaction
        OpenCon()
        If CurrentTrans Is Nothing Then CurrentTrans = con.BeginTransaction(IsolationLevel.ReadCommitted)
        Cmd.Transaction = CurrentTrans


        Try
            'Audit
            GetSQLAudit(Cmd.CommandText, AuditAction, AuditTable, AuditPerson, AuditValues)
            Dim AuditSQLCode As String = "'" & AuditPerson & "','" & AuditAction &
                    "','" & AuditTable & "','" & Left(AuditValues, 254) & "'"
            AuditSQLCode = "INSERT INTO AUDIT ([Person], [Action], [TName], [NValue]) VALUES (" & AuditSQLCode & ")"
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
            BindSource.DataSource = CurrentDataSet.Tables(0)
            ctl.DataSource = BindSource

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

        Dim ErrorMessage As Exception = Nothing

        'Is the data dirty / has errors that have auto-undone
        If CurrentDataSet.HasChanges() = False Then
            If DisplayMessage = True Then MsgBox("Errors present/No changes to upload")
            Call Refresher(ctl)
            Exit Sub
        End If

        'Open Connection & Set transaction
        OpenCon()
        If CurrentTrans Is Nothing Then CurrentTrans = con.BeginTransaction(IsolationLevel.ReadCommitted)

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
                    "','" & Table & "','" & Left(AuditValues, 255) & "'"
                AddToMassSQL("INSERT INTO AUDIT ([Person], [Action], [TName], [NValue]) VALUES (" & CombineInsert & ")", False)
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

        Catch ex As Exception
            ErrorMessage = ex
            Call TryRollBack()

        Finally
            'Close off & clean up
            CloseCon()
            CurrentTrans = Nothing
            If Not ErrorMessage Is Nothing Then DefaultError(ErrorMessage)

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
            Throw
        Finally
            'Pass back whether clean up happened
            UnloadData = Cancel
        End Try

    End Function

    Public Function TempDataTable(SQLCode As String) As DataTable
        'Create a temporary dataset for things such as combo box which arent based on the initial query

        Try
            'Open connection
            OpenCon()
            'New temporary data adapter and dataset
            Dim TempDataAdapter = New OleDb.OleDbDataAdapter(SQLCode, con)
            TempDataTable = New DataTable()
            'Use temp adapter to fill temp dataset
            TempDataAdapter.Fill(TempDataTable)

        Catch ex As Exception
            CloseCon()
            Throw
            TempDataTable = Nothing


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
            Throw
        End Try

    End Sub

    Public Sub LoginCheck()

        Dim SQLString As String = "SELECT * FROM " & UserTable & " WHERE " & UserField & "='" & GetUserName() & "'"
        Dim ErrorMessage As String = "You do not have permission to use this database. Please contact David Burnside or " & Contact

        If SELECTCount(SQLString) = 0 Then
            MsgBox(ErrorMessage)
            Call Quitter()
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

    Public Sub DefaultError(ex As Exception)

        Dim st As StackTrace = New StackTrace(ex, True)


        MsgBox("An application error has occured -" _
            & vbNewLine _
            & vbNewLine _
            & "Method: " & st.GetFrame(st.FrameCount - 1).GetMethod().Name.ToString _
            & vbNewLine _
            & "Line: " & st.GetFrame(st.FrameCount - 1).GetFileLineNumber().ToString _
            & vbNewLine _
            & vbNewLine _
            & "Error: " & ex.Message, , "Application Error")


    End Sub

    Protected Sub ComboKeyDown(sender As Object, e As KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub

    Protected Sub DataGridViewDataError(sender As Object, e As DataGridViewDataErrorEventArgs)
        Try
            e.Cancel = False
            Call ErrorHandler(sender, e)
        Catch ex As Exception
            Call DefaultError(ex)
        End Try
    End Sub

    Protected Sub GridComboEnter(sender As Object, e As DataGridViewCellEventArgs)

        If Not e.RowIndex > 0 Then Exit Sub
        On Error Resume Next
        Dim dgv As DataGridView = CType(sender, DataGridView)

        If dgv(e.ColumnIndex, e.RowIndex).EditType.ToString() = "System.Windows.Forms.DataGridViewComboBoxEditingControl" Then
            SendKeys.Send("{F4}")
        End If
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
            DefaultError(ex)
        End Try

    End Sub

    Public Sub SetupFilterCombo(Ctl As ComboBox,
                                CSV_1_Value_2_Display As String,
                                Optional DefaultEmpty As Boolean = True)

        Try

            Dim StringArray As String() = CSV_1_Value_2_Display.Split(",")
            If StringArray.Count = 1 Then Throw New ArgumentException("Only single arguement entered for SetupFilterCombo")

            Ctl.ValueMember = StringArray(0).ToString
            Ctl.DisplayMember = StringArray(1).ToString

            AddHandler Ctl.SelectionChangeCommitted, AddressOf FilterCombo
            AddHandler Ctl.Enter, AddressOf RefreshCombo

            If DefaultEmpty <> True Then RefreshCombo(Ctl, Ctl, DefaultEmpty)

        Catch ex As Exception
            DefaultError(ex)
        End Try

    End Sub

    Protected Sub RefreshCombo(ctl As ComboBox, e As Object, Optional DefaultEmpty As Boolean = True)

        Dim CurrentChoice As String = ctl.SelectedValue

        Dim Dt As DataTable

        Dim View As DataView = New DataView(CurrentDataSet.Tables(0),
        ctl.ValueMember & "<>'' AND " & ctl.DisplayMember & "<>''",
                                            "", DataViewRowState.CurrentRows)

        If ctl.ValueMember = ctl.DisplayMember Then
            Dt = View.ToTable(True, ctl.ValueMember.ToString)
        Else
            Dt = View.ToTable(True, ctl.ValueMember, ctl.DisplayMember)
        End If

        Dim DtRow As DataRow = Dt.NewRow
        DtRow(ctl.ValueMember) = ""
        If ctl.ValueMember.ToString <> ctl.DisplayMember Then DtRow(ctl.DisplayMember) = ""
        Dt.Rows.Add(DtRow)


        Dim dataView As New DataView(Dt)
        dataView.Sort = ctl.DisplayMember & " ASC"
        Dim Dt2 As DataTable = dataView.ToTable()
        Dt = Nothing
        View = Nothing

        ctl.DataSource = Dt2

        If DefaultEmpty = False Then
            ctl.SelectedValue = Dt2.Rows(1).Item(0)
            FilterCombo(ctl, New EventArgs)
        End If

        If CurrentChoice <> "" Then ctl.SelectedValue = CurrentChoice

    End Sub

    Protected Sub FilterCombo(sender As ComboBox, e As EventArgs)

        Dim CtlValue As String = ""

        If IsNumeric(sender.SelectedValue) = True Then
            CtlValue = sender.SelectedValue
        Else
            If sender.SelectedValue <> "" Then CtlValue = "'" & sender.SelectedValue & "'"
        End If

        Dim FilterString As String = vbNullString
        Dim RowFilter As String = CurrentDataSet.Tables(0).DefaultView.RowFilter
        Dim PrevLoc As Long = InStr(RowFilter, sender.ValueMember & "=")
        Dim ANDLoc As Long
        Dim ReplaceString As String = ""


        If PrevLoc <> 0 Then
            ANDLoc = InStr(PrevLoc, RowFilter, "AND", 0)
            If ANDLoc = 0 Then ANDLoc = Len(RowFilter) + 1
            ANDLoc = ANDLoc - PrevLoc
            ReplaceString = Mid(RowFilter, PrevLoc, ANDLoc)
            RowFilter = Replace(RowFilter, ReplaceString, "")
        End If


        RowFilter = Trim(RowFilter)
        If InStr(RowFilter, "AND AND") <> 0 Then RowFilter = Replace(RowFilter, "AND AND", "AND")
        If InStr(RowFilter, "ANDAND") <> 0 Then RowFilter = Replace(RowFilter, "ANDAND", "AND")
        If InStr(RowFilter, "  ") <> 0 Then RowFilter = Replace(RowFilter, "  ", " ")
        If Left(RowFilter, 3) = "AND" Then RowFilter = Mid(RowFilter, 4)
        If RowFilter.Length > 3 Then If RowFilter.Substring(RowFilter.Length - 3) = "AND" Then RowFilter = Left(RowFilter, RowFilter.Length - 3)
        RowFilter = Trim(RowFilter)

        If CtlValue <> "" And Len(RowFilter) <> 0 Then FilterString = " AND "

        If CtlValue <> "" Then FilterString = FilterString & sender.ValueMember & "=" & CtlValue

        FilterString = RowFilter & FilterString

        CurrentDataSet.Tables(0).DefaultView.RowFilter = FilterString


    End Sub

    Public Sub ResetCollection()

        For Each control In DataItemCollection
            If (TypeOf control Is DataGridView) Then
                control.Columns.Clear()
                control.DataSource = Nothing
            End If

        Next

    End Sub

End Class