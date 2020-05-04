Public Class forexcel
    Private Sub forexcel_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)
        Const WM_SYSCOMMAND As Integer = &H112
        Const SC_MOVE As Integer = &HF010

        Select Case m.Msg
            Case WM_SYSCOMMAND
                Dim command As Integer = m.WParam.ToInt32() And &HFFF0
                If command = SC_MOVE Then Return
        End Select
        MyBase.WndProc(m)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Close()
    End Sub
End Class