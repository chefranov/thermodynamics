Public Class BaseEditor

    Private Function chk(ByVal T As TextBox) As Boolean
        Dim s As String = T.Text.Replace(".", ",") : T.BackColor = Color.White : T.Text = s : chk = True
        If Not IsNumeric(s) Then
            MessageBox.Show("В поле ввода следует вводить числа.", "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            T.Clear() : T.Focus() : T.BackColor = Color.Bisque
            Return False
        End If
    End Function

    Private Sub BaseEditor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button4.Enabled = False
        Button2.Enabled = False
        Загрузить2(ListView1, "termodynamics3.txt")
    End Sub
    Private Sub BaseEditor_Closed(sender As Object, e As EventArgs) Handles MyBase.Closed
        Form1.ComboBoxBase.SelectedIndex = 1
        Form1.ComboBoxBase.SelectedIndex = 2
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        e.Handled = ("|").Contains(e.KeyChar)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        For k As Integer = 2 To 6
            Me.GroupBox1.Controls("TextBox" & k).Text = "0"
        Next
        TextBox1.Text = "Item 1"
        Button4.Enabled = False
        Button2.Enabled = False
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        УдалитьВыбранный(ListView1)
        Сохранить(ListView1, Application.StartupPath & "\termodynamics3.txt")
        Button4.Enabled = False
        Button2.Enabled = False
        Dim t As String = "termodynamics3.txt"
        Dim delete As String = System.IO.File.ReadAllText(t, System.Text.Encoding.Unicode)
        If delete.Substring(delete.Length - 2, vbCrLf.Length) = vbCrLf Then
            System.IO.File.WriteAllText(t, delete.Remove(delete.Length - vbCrLf.Length), System.Text.Encoding.Unicode)
        End If
    End Sub

    Public Sub Добавить(_LV As ListView, i1 As String, i2 As String, i3 As String, i4 As String, i5 As String, i6 As String)
        'Добавляем в список строку и одновременно заполняем ее первый столбец:
        _LV.Items.Add(i1)
        'Добавляем в строку еще два элемента:
        With _LV.Items(_LV.Items.Count - 1).SubItems
            .Add(i2)
            .Add(i3)
            .Add(i4)
            .Add(i5)
            .Add(i6)
        End With
        _LV.Items(_LV.Items.Count - 1).Selected = True
        _LV.Focus()

    End Sub

    Public Sub Загрузить2(_LV As ListView, fName As String)
        _LV.Items.Clear() 'Очищаем список

        'Загружаем список:
        Using Чтение As New System.IO.StreamReader(fName, System.Text.Encoding.Unicode)
            Dim i As Integer = 0
            Do Until Чтение.EndOfStream
                'Добавляем в список очередную строку
                Dim drop() As String = Split(Чтение.ReadLine, "|")
                _LV.Items.Add(drop(0))
                With _LV.Items(i).SubItems
                    .Add(drop(1))
                    .Add(drop(2))
                    .Add(drop(3))
                    .Add(drop(4))
                    .Add(drop(5))
                End With
                i = i + 1
            Loop
        End Using
    End Sub

    Public Sub УдалитьВыбранный(_LV As ListView)
        If _LV.Items.Count <> 0 Then
            If _LV.FocusedItem.Selected Then _LV.FocusedItem.Remove() 'Удалить выделенный
        End If
    End Sub

    Public Sub Сохранить(_LV As ListView, fName As String)
        Dim ItemsArray As New ArrayList
        For Each item As ListViewItem In _LV.Items
            ItemsArray.Add(item.Text & "|" &
                           item.SubItems(1).Text & "|" &
                           item.SubItems(2).Text & "|" &
                           item.SubItems(3).Text & "|" &
                           item.SubItems(4).Text & "|" &
                           item.SubItems(5).Text)
        Next
        IO.File.WriteAllLines(fName, ItemsArray.ToArray.Cast(Of String), System.Text.Encoding.Unicode)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        For k As Integer = 1 To 6
            If Len(Me.GroupBox1.Controls("TextBox" & k).Text) = 0 Then
                MessageBox.Show("Поля не могут быть пустыми.", "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                DirectCast(Me.GroupBox1.Controls("TextBox" & k), TextBox).Clear() : Me.GroupBox1.Controls("TextBox" & k).Focus() : Me.GroupBox1.Controls("TextBox" & k).BackColor = Color.Bisque
                Exit Sub
            End If
        Next

        For k As Integer = 2 To 6
            If Not chk(Me.GroupBox1.Controls("TextBox" & k)) Then Exit Sub
        Next
        Добавить(ListView1, TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox4.Text, TextBox5.Text, TextBox6.Text)
        Сохранить(ListView1, Application.StartupPath & "\termodynamics3.txt")

        Dim t As String = "termodynamics3.txt"
        Dim delete As String = System.IO.File.ReadAllText(t, System.Text.Encoding.Unicode)
        If delete.Substring(delete.Length - 2, vbCrLf.Length) = vbCrLf Then
            System.IO.File.WriteAllText(t, delete.Remove(delete.Length - vbCrLf.Length), System.Text.Encoding.Unicode)
        End If
    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged


        Dim index As Integer = 0

        If ListView1.SelectedItems.Count > 0 Then

            If ListView1.SelectedItems(0).Index < 9 Then
                Button4.Enabled = False
                Button2.Enabled = False
            Else
                Button4.Enabled = True
                Button2.Enabled = True
            End If

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim name As String = InputBox("Введите название соединения")
        Dim H As String = InputBox("Введите ΔH")
        Dim S As String = InputBox("Введите ΔS")
        Dim a As String = InputBox("Введие a")
        Dim b As String = InputBox("Введие b")
        Dim c As String = InputBox("Введие c")
        For i = 0 To ListView1.Items.Count - 1
            If ListView1.SelectedIndices.Contains(i) Then
                ListView1.Items(i).SubItems(0).Text = name
                ListView1.Items(i).SubItems(1).Text = H
                ListView1.Items(i).SubItems(2).Text = S
                ListView1.Items(i).SubItems(3).Text = a
                ListView1.Items(i).SubItems(4).Text = b
                ListView1.Items(i).SubItems(5).Text = c
            End If
        Next
        Сохранить(ListView1, Application.StartupPath & "\termodynamics3.txt")
    End Sub
End Class