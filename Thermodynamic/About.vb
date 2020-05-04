Public Class About

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

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object,
   ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) _
   Handles LinkLabel1.LinkClicked
        Try
            VisitLink()
        Catch ex As Exception
            ' The error message
            MessageBox.Show("Unable to open link that was clicked.")
        End Try
    End Sub

    Sub VisitLink()
        ' Change the color of the link text by setting LinkVisited 
        ' to True.
        LinkLabel1.LinkVisited = True
        ' Call the Process.Start method to open the default browser 
        ' with a URL:
        System.Diagnostics.Process.Start("http://chefranov.name")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Close()
    End Sub

    Private Sub About_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = "Thermodynamics " & Application.ProductVersion
        CheckBox1.Checked = My.Settings.checkupdate
    End Sub

    Private Sub About_Closed(sender As Object, e As EventArgs) Handles MyBase.Closed
        My.Settings.checkupdate = CheckBox1.Checked
    End Sub

End Class