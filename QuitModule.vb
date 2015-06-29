Module QuitModule

    Public Sub Quitter(Optional CloseAnyway As Boolean = False)

        On Error Resume Next

        If Login = False Or CloseAnyway = True Then
            Application.Exit()
            End
        End If

    End Sub

End Module
