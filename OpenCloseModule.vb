Module OpenCloseModule

    Public Sub OpenFrm(WhichForm As Form, Optional CloseCurrent As Form = Nothing)

        Dim nForm As Form

        Try
            Dim Fullname As String = Application.ProductName & "." & WhichForm.Name
            nForm = Activator.CreateInstance(Type.GetType(Fullname, True, True))
            nForm.Show()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            nForm = Nothing
            If Not IsNothing(CloseCurrent) Then Call CloseFrm(CloseCurrent)
        End Try




    End Sub

    Public Sub CloseFrm(WhichForm As Form)

        Dim OpenFrmCount As Long

        Try
            WhichForm.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            OpenFrmCount = My.Application.OpenForms.Count
            MsgBox(OpenFrmCount)
            If OpenFrmCount = 0 Then Call Quitter(True)
        End Try



    End Sub
End Module
