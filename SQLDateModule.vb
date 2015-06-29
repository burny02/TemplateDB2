Module SQLDateModule

    Public Function SQLDate(varDate As Date) As String

        Dim ErrorMessage As String = "SQLDate - Date conversion failed"

        Try
            If DateValue(varDate) = varDate Then
                SQLDate = Format$(varDate, "\#mm\/dd\/yyyy\#")
            Else
                SQLDate = Format$(varDate, "\#mm\/dd\/yyyy hh\:nn\:ss\#")
            End If
        Catch ex As Exception
            Throw New System.Exception(ErrorMessage)
        End Try

        'ALWAYS use in SQL Command - The # tells it that is it american format

    End Function

End Module
