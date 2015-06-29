Module LoginModule
    Public Login As Boolean

    Public Sub LoginCheck()

        Dim TableName As String = "[Users]"
        Dim FieldName As String = "[UserName]"
        Dim ContactName As String = "Skye Firminger"
        Dim SQLString As String = "SELECT * FROM " & TableName & " WHERE " & FieldName & "='" & GetUserName() & "'"
        Dim ErrorMessage As String = "You do not have permission to use this database. Please contact David Burnside or " & ContactName

        
        If QueryTest(SQLString) = 0 Then
            MsgBox(ErrorMessage)
            Call Quitter(True)
        Else
            Login = True
        End If


    End Sub
End Module
