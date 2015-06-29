Module LockModule

    Public Sub LockCheck()

        Dim TableName As String = "[Locker]"
        Dim SQLString As String = "SELECT * FROM " & TableName
        Dim ErrorMessage As String = "The database is currently locked. Please contact David Burnside"

        If QueryTest(SQLString) <> 0 Then
            MsgBox(ErrorMessage)
            Call Quitter(True)
        End If


    End Sub

    Private Sub LockUnlock()

        Dim TableName As String = "[Locker]"
        Dim SQLTest As String = "SELECT * FROM " & TableName
        Dim SQLInsert As String = "INSERT INTO " & TableName & " Values ('" & GetUserName() & "')"
        Dim SQLDelete As String = "DELETE * FROM" & TableName
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

End Module
