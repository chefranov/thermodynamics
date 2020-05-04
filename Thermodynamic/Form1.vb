Imports System.Drawing.Drawing2D
Imports System.Math
Imports System.Windows.Forms.DataVisualization.Charting


Public Class Form1
    'Thermodynamics - программа для расчета значений энергии Гиббса при различных температур
    'Автор программы: Евгений Чефранов
    'Сайт автора: chefranov.name
    'Страница программы: chefranov.name/thermodynamics
    '© 2016 Евгений Чефранов. Все права защищены.


    'Глобальная переменная для радиокнопок
    Dim Var4184 As Double = 1

    'Создаем класс, поля которого - строка и массив.
    Dim Data As New List(Of MyData)
    Class MyData
        Public Combo, T(4) As String
        ReadOnly Property Text()
            Get
                Return Combo
            End Get
        End Property
    End Class

    'Задаем параметры для радиокнопок
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked Then
            LabelH.Text = "∆H, кДж" : LabelS.Text = "∆S, Дж"
            For I As Integer = 1 To 40
                Controls("TextBox" & I).Text = Format(Controls("TextBox" & I).Text * 4.184, "0.##")
            Next
            Var4184 = 4.184
        Else
            LabelH.Text = "∆H, ккал" : LabelS.Text = "∆S, кал"
            For I As Integer = 1 To 40
                Controls("TextBox" & I).Text = Format(Controls("TextBox" & I).Text / 4.184, "0.##")
            Next
            Var4184 = 1
        End If
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = "Thermodynamics " & Application.ProductVersion & " - расчет термодинамики"
        If My.Settings.checkupdate = True Then
            'Проверка на доступность сайта с обновлениями, если есть обновление, то предлагаем обновиться
            'If My.Computer.Network.Ping("195.216.243.180") Then
            WBC.ScriptErrorsSuppressed() = True
            Dim url As String = "http://chefranov.name/thermodynamics" ' адрес страницы, где написан номер последней версии 
                WBC.Navigate(url)
            End If
        'End If
        'Задаем название осей для графика
        Me.Chart1.ChartAreas("ChartArea1").AxisY.Title = "Энергия Гиббса, кДж"
        Me.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Температура, К"

        ComboBoxBase.SelectedIndex = 0
        For I As Integer = 1 To 8
            AddHandler DirectCast(Controls("ComboBox" & I), ComboBox).SelectedIndexChanged, AddressOf ComboBox_SelectedIndexChanged
        Next
    End Sub
    'Делаем вставку значений в ячейки для соотвествующих ComboBox
    Private Sub ComboBox_SelectedIndexChanged(ByVal sender As ComboBox, ByVal e As System.EventArgs)
        Dim Ind As Integer = CInt(sender.Name.Replace("ComboBox", "")) - 1
        For I As Integer = Ind * 5 To Ind * 5 + 4
            Controls("TextBox" & I + 1).Text = (Format((Data(sender.SelectedIndex).T(I - (Ind * 5)) * Var4184), "0.##"))
        Next
    End Sub

    'Оброботка полей на наличие и замену точек на запятую
    Private Function chk(ByVal T As TextBox) As Boolean
        Dim separator As String = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator(0)
        Dim s As String = T.Text.Replace(".", separator) : T.BackColor = Color.White : T.Text = s : chk = True
        If Not IsNumeric(s) Then
            MessageBox.Show("В поле ввода следует вводить числа.", "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            T.Clear() : T.Focus() : T.BackColor = Color.Bisque
            Return False
        End If
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Делаем проверку полей ввода на правильность введенных данных
        For k As Integer = 1 To 40
            If Not chk(Me.Controls("TextBox" & k)) Then Exit Sub
        Next
        For k As Integer = 1 To 3
            If Not chk(Me.GroupBox2.Controls("TextBox4" & k)) Then Exit Sub
        Next
        For k As Integer = 1 To 8
            If Not chk(Me.Controls("TextBox0" & k)) Then Exit Sub
        Next

        Dim TempNach, TempKonech, TempInterval As Single
        TempNach = (TextBox41.Text)
        TempKonech = (TextBox42.Text)
        TempInterval = (TextBox43.Text)

        If TempNach < 0 Or TempNach > TempKonech Then
            MessageBox.Show("В поле ввода следует вводить положительные числа и начальная температура не может превышать конечную.", "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TextBox41.Clear() : TextBox41.Focus() : TextBox41.BackColor = Color.Bisque : TextBox42.Clear() : TextBox42.BackColor = Color.Bisque
            Exit Sub
        End If

        If TempKonech < 0 Then
            MessageBox.Show("В поле ввода следует вводить положительные числа.", "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TextBox42.Clear() : TextBox42.Focus() : TextBox42.BackColor = Color.Bisque
            Exit Sub
        End If

        If TempInterval < 1 Or TempInterval > TempKonech Then
            MessageBox.Show("Температурный шаг не может иметь отрицательное значение и превышать значение конечной температуры", "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TextBox43.Clear() : TextBox43.Focus() : TextBox43.BackColor = Color.Bisque
            Exit Sub
        End If

        If RadioButton4.Checked Then
            Me.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Температура, °C"
        Else
            Me.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Температура, K"
        End If

        Dim zn As Single = 0


        For i As Integer = 1 To 40
            zn += (Me.Controls("TextBox" & i).Text)
        Next

        If Math.Abs(zn) = 0 Then
            MessageBox.Show("Расчет невозможен, т.к. данные всех соединений пусты.", "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        For i As Integer = 1 To 8
            If Me.Controls("TextBox0" & i).Text < 0.0001 Then
                MessageBox.Show("Расчет невозможен, т.к. стехиометрические коэффициенты имееют отрицательное значение или равны нулю.", "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                DirectCast(Controls("TextBox0" & i), TextBox).Clear() : Me.Controls("TextBox0" & i).Focus() : Me.Controls("TextBox0" & i).BackColor = Color.Bisque
                Exit Sub
            End If
        Next

        'Очищаем значения в окне экспорта в Excel
        forexcel.TextBoxG.Text = ""
        forexcel.TextBoxH.Text = ""
        forexcel.TextBoxS.Text = ""
        forexcel.TextBoxTemp.Text = ""
        forexcel.TextBoxTempK.Text = ""

        'Делаем очистку диаграммы от предыдуших значений
        Me.Chart1.Series("Series1").Points.Clear()

        'Производим сам расчет энергии Гиббса (весь расчет считаем с использованием данных в калориях и градусах Кельвина)       
        Dim H298, S298, G298, Ht, St, Gt, a, b, c, HtOt, StOt, GtOt, Tn, Tk As Double

        'Данные (H,S,a,b,c) Al2O3 и SiO2 для полиморфных превращений (ккал, кал)
        Dim hgammaAl2O3, sgammaAl2O3, agammaAl2O3, bgammaAl2O3, cgammaAl2O3, halphaAl2O3, salphaAl2O3, aalphaAl2O3, balphaAl2O3, calphaAl2O3, halphaSiO2, salphaSiO2, aalphaSiO2, balphaSiO2, calphaSiO2, hbetaSiO2, sbetaSiO2, abetaSiO2, bbetaSiO2, cbetaSiO2, htridimitSiO2, stridimitSiO2, atridimitSiO2, btridimitSiO2, ctridimitSiO2, hkristobalitSiO2, skristobalitSiO2, akristobalitSiO2, bkristobalitSiO2, ckristobalitSiO2 As Single

        hgammaAl2O3 = -395 : sgammaAl2O3 = 12.55 : agammaAl2O3 = 16.37 : bgammaAl2O3 = 11.1 : cgammaAl2O3 = 0
        halphaAl2O3 = -400.48 : salphaAl2O3 = 12.17 : aalphaAl2O3 = 27.43 : balphaAl2O3 = 3.06 : calphaAl2O3 = -8.47

        halphaSiO2 = -217.6 : salphaSiO2 = 0 : aalphaSiO2 = 14.41 : balphaSiO2 = 1.94 : calphaSiO2 = 0
        hbetaSiO2 = -217.75 : sbetaSiO2 = 10 : abetaSiO2 = 11.22 : bbetaSiO2 = 8.2 : cbetaSiO2 = -2.7
        htridimitSiO2 = -216.5 : stridimitSiO2 = 10.4 : atridimitSiO2 = 13.64 : btridimitSiO2 = 2.64 : ctridimitSiO2 = 0
        hkristobalitSiO2 = -215.95 : skristobalitSiO2 = 10.19 : akristobalitSiO2 = 4.28 : bkristobalitSiO2 = 21.06 : ckristobalitSiO2 = 0

        'Начальная и конечная температура, К
        Tn = (TextBox41.Text)
        Tk = (TextBox42.Text)

        For z As Integer = Tn To Tk Step (TextBox43.Text)

            If CheckBox3.Checked Then
                If z <= 1500 Then
                    If ComboBox1.Text = "αAl2O3" Or ComboBox1.Text = "γAl2O3" Or ComboBox1.Text = "Al2O3(к)" Or ComboBox1.Text = "Al2O3 (полиморф.)" Then
                        TextBox1.Text = hgammaAl2O3 : TextBox2.Text = sgammaAl2O3 : TextBox3.Text = agammaAl2O3 : TextBox4.Text = bgammaAl2O3 : TextBox5.Text = cgammaAl2O3
                    End If
                    If ComboBox2.Text = "αAl2O3" Or ComboBox2.Text = "γAl2O3" Or ComboBox2.Text = "Al2O3(к)" Or ComboBox2.Text = "Al2O3 (полиморф.)" Then
                        TextBox6.Text = hgammaAl2O3 : TextBox7.Text = sgammaAl2O3 : TextBox8.Text = agammaAl2O3 : TextBox9.Text = bgammaAl2O3 : TextBox10.Text = cgammaAl2O3
                    End If
                    If ComboBox3.Text = "αAl2O3" Or ComboBox3.Text = "γAl2O3" Or ComboBox3.Text = "Al2O3(к)" Or ComboBox3.Text = "Al2O3 (полиморф.)" Then
                        TextBox11.Text = hgammaAl2O3 : TextBox12.Text = sgammaAl2O3 : TextBox13.Text = agammaAl2O3 : TextBox14.Text = bgammaAl2O3 : TextBox15.Text = cgammaAl2O3
                    End If
                    If ComboBox4.Text = "αAl2O3" Or ComboBox4.Text = "γAl2O3" Or ComboBox4.Text = "Al2O3(к)" Or ComboBox4.Text = "Al2O3 (полиморф.)" Then
                        TextBox16.Text = hgammaAl2O3 : TextBox17.Text = sgammaAl2O3 : TextBox18.Text = agammaAl2O3 : TextBox19.Text = bgammaAl2O3 : TextBox20.Text = cgammaAl2O3
                    End If
                    If ComboBox5.Text = "αAl2O3" Or ComboBox5.Text = "γAl2O3" Or ComboBox5.Text = "Al2O3(к)" Or ComboBox5.Text = "Al2O3 (полиморф.)" Then
                        TextBox21.Text = hgammaAl2O3 : TextBox22.Text = sgammaAl2O3 : TextBox23.Text = agammaAl2O3 : TextBox24.Text = bgammaAl2O3 : TextBox25.Text = cgammaAl2O3
                    End If
                    If ComboBox6.Text = "αAl2O3" Or ComboBox6.Text = "γAl2O3" Or ComboBox6.Text = "Al2O3(к)" Or ComboBox6.Text = "Al2O3 (полиморф.)" Then
                        TextBox26.Text = hgammaAl2O3 : TextBox27.Text = sgammaAl2O3 : TextBox28.Text = agammaAl2O3 : TextBox29.Text = bgammaAl2O3 : TextBox30.Text = cgammaAl2O3
                    End If
                    If ComboBox7.Text = "αAl2O3" Or ComboBox7.Text = "γAl2O3" Or ComboBox7.Text = "Al2O3(к)" Or ComboBox7.Text = "Al2O3 (полиморф.)" Then
                        TextBox31.Text = hgammaAl2O3 : TextBox32.Text = sgammaAl2O3 : TextBox33.Text = agammaAl2O3 : TextBox34.Text = bgammaAl2O3 : TextBox35.Text = cgammaAl2O3
                    End If
                    If ComboBox8.Text = "αAl2O3" Or ComboBox8.Text = "γAl2O3" Or ComboBox8.Text = "Al2O3(к)" Or ComboBox8.Text = "Al2O3 (полиморф.)" Then
                        TextBox36.Text = hgammaAl2O3 : TextBox37.Text = sgammaAl2O3 : TextBox38.Text = agammaAl2O3 : TextBox39.Text = bgammaAl2O3 : TextBox40.Text = cgammaAl2O3
                    End If
                End If

                If z > 1500 Then
                    If ComboBox1.Text = "αAl2O3" Or ComboBox1.Text = "γAl2O3" Or ComboBox1.Text = "Al2O3(к)" Or ComboBox1.Text = "Al2O3 (полиморф.)" Then
                        TextBox1.Text = halphaAl2O3 : TextBox2.Text = salphaAl2O3 : TextBox3.Text = aalphaAl2O3 : TextBox4.Text = balphaAl2O3 : TextBox5.Text = calphaAl2O3
                    End If
                    If ComboBox2.Text = "αAl2O3" Or ComboBox2.Text = "γAl2O3" Or ComboBox2.Text = "Al2O3(к)" Or ComboBox2.Text = "Al2O3 (полиморф.)" Then
                        TextBox6.Text = halphaAl2O3 : TextBox7.Text = salphaAl2O3 : TextBox8.Text = aalphaAl2O3 : TextBox9.Text = balphaAl2O3 : TextBox10.Text = calphaAl2O3
                    End If
                    If ComboBox3.Text = "αAl2O3" Or ComboBox3.Text = "γAl2O3" Or ComboBox3.Text = "Al2O3(к)" Or ComboBox3.Text = "Al2O3 (полиморф.)" Then
                        TextBox11.Text = halphaAl2O3 : TextBox12.Text = salphaAl2O3 : TextBox13.Text = aalphaAl2O3 : TextBox14.Text = balphaAl2O3 : TextBox15.Text = calphaAl2O3
                    End If
                    If ComboBox4.Text = "αAl2O3" Or ComboBox4.Text = "γAl2O3" Or ComboBox4.Text = "Al2O3(к)" Or ComboBox4.Text = "Al2O3 (полиморф.)" Then
                        TextBox16.Text = halphaAl2O3 : TextBox17.Text = salphaAl2O3 : TextBox18.Text = aalphaAl2O3 : TextBox19.Text = balphaAl2O3 : TextBox20.Text = calphaAl2O3
                    End If
                    If ComboBox5.Text = "αAl2O3" Or ComboBox5.Text = "γAl2O3" Or ComboBox5.Text = "Al2O3(к)" Or ComboBox5.Text = "Al2O3 (полиморф.)" Then
                        TextBox21.Text = halphaAl2O3 : TextBox22.Text = salphaAl2O3 : TextBox23.Text = aalphaAl2O3 : TextBox24.Text = balphaAl2O3 : TextBox25.Text = calphaAl2O3
                    End If
                    If ComboBox6.Text = "αAl2O3" Or ComboBox6.Text = "γAl2O3" Or ComboBox6.Text = "Al2O3(к)" Or ComboBox6.Text = "Al2O3 (полиморф.)" Then
                        TextBox26.Text = halphaAl2O3 : TextBox27.Text = salphaAl2O3 : TextBox28.Text = aalphaAl2O3 : TextBox29.Text = balphaAl2O3 : TextBox30.Text = calphaAl2O3
                    End If
                    If ComboBox7.Text = "αAl2O3" Or ComboBox7.Text = "γAl2O3" Or ComboBox7.Text = "Al2O3(к)" Or ComboBox7.Text = "Al2O3 (полиморф.)" Then
                        TextBox31.Text = halphaAl2O3 : TextBox32.Text = salphaAl2O3 : TextBox33.Text = aalphaAl2O3 : TextBox34.Text = balphaAl2O3 : TextBox35.Text = calphaAl2O3
                    End If
                    If ComboBox8.Text = "αAl2O3" Or ComboBox8.Text = "γAl2O3" Or ComboBox8.Text = "Al2O3(к)" Or ComboBox8.Text = "Al2O3 (полиморф.)" Then
                        TextBox36.Text = halphaAl2O3 : TextBox37.Text = salphaAl2O3 : TextBox38.Text = aalphaAl2O3 : TextBox39.Text = balphaAl2O3 : TextBox40.Text = calphaAl2O3
                    End If
                End If
            End If


            If CheckBox2.Checked Then
                If z <= 846.15 Then
                    If ComboBox1.Text = "SiO2 (β-кварц)" Or ComboBox1.Text = "SiO2 (α-кварц)" Or ComboBox1.Text = "SiO2 (α-тридимит)" Or ComboBox1.Text = "SiO2 (β-кристобалит)" Or ComboBox1.Text = "SiO2(к) кварц" Or ComboBox1.Text = "SiO2(к) тридимит" Or ComboBox1.Text = "SiO2(к) кристобалит" Or ComboBox1.Text = "SiO2 (полиморф.)" Then
                        TextBox1.Text = hbetaSiO2 : TextBox2.Text = sbetaSiO2 : TextBox3.Text = abetaSiO2 : TextBox4.Text = bbetaSiO2 : TextBox5.Text = cbetaSiO2
                    End If
                    If ComboBox2.Text = "SiO2 (β-кварц)" Or ComboBox2.Text = "SiO2 (α-кварц)" Or ComboBox2.Text = "SiO2 (α-тридимит)" Or ComboBox2.Text = "SiO2 (β-кристобалит)" Or ComboBox2.Text = "SiO2(к) кварц" Or ComboBox2.Text = "SiO2(к) тридимит" Or ComboBox2.Text = "SiO2(к) кристобалит" Or ComboBox2.Text = "SiO2 (полиморф.)" Then
                        TextBox6.Text = hbetaSiO2 : TextBox7.Text = sbetaSiO2 : TextBox8.Text = abetaSiO2 : TextBox9.Text = bbetaSiO2 : TextBox10.Text = cbetaSiO2
                    End If
                    If ComboBox3.Text = "SiO2 (β-кварц)" Or ComboBox3.Text = "SiO2 (α-кварц)" Or ComboBox3.Text = "SiO2 (α-тридимит)" Or ComboBox3.Text = "SiO2 (β-кристобалит)" Or ComboBox3.Text = "SiO2(к) кварц" Or ComboBox3.Text = "SiO2(к) тридимит" Or ComboBox3.Text = "SiO2(к) кристобалит" Or ComboBox3.Text = "SiO2 (полиморф.)" Then
                        TextBox11.Text = hbetaSiO2 : TextBox12.Text = sbetaSiO2 : TextBox13.Text = abetaSiO2 : TextBox14.Text = bbetaSiO2 : TextBox15.Text = cbetaSiO2
                    End If
                    If ComboBox4.Text = "SiO2 (β-кварц)" Or ComboBox4.Text = "SiO2 (α-кварц)" Or ComboBox4.Text = "SiO2 (α-тридимит)" Or ComboBox4.Text = "SiO2 (β-кристобалит)" Or ComboBox4.Text = "SiO2(к) кварц" Or ComboBox4.Text = "SiO2(к) тридимит" Or ComboBox4.Text = "SiO2(к) кристобалит" Or ComboBox4.Text = "SiO2 (полиморф.)" Then
                        TextBox16.Text = hbetaSiO2 : TextBox17.Text = sbetaSiO2 : TextBox18.Text = abetaSiO2 : TextBox19.Text = bbetaSiO2 : TextBox20.Text = cbetaSiO2
                    End If
                    If ComboBox5.Text = "SiO2 (β-кварц)" Or ComboBox5.Text = "SiO2 (α-кварц)" Or ComboBox5.Text = "SiO2 (α-тридимит)" Or ComboBox5.Text = "SiO2 (β-кристобалит)" Or ComboBox5.Text = "SiO2(к) кварц" Or ComboBox5.Text = "SiO2(к) тридимит" Or ComboBox5.Text = "SiO2(к) кристобалит" Or ComboBox5.Text = "SiO2 (полиморф.)" Then
                        TextBox21.Text = hbetaSiO2 : TextBox22.Text = sbetaSiO2 : TextBox23.Text = abetaSiO2 : TextBox24.Text = bbetaSiO2 : TextBox25.Text = cbetaSiO2
                    End If
                    If ComboBox6.Text = "SiO2 (β-кварц)" Or ComboBox6.Text = "SiO2 (α-кварц)" Or ComboBox6.Text = "SiO2 (α-тридимит)" Or ComboBox6.Text = "SiO2 (β-кристобалит)" Or ComboBox6.Text = "SiO2(к) кварц" Or ComboBox6.Text = "SiO2(к) тридимит" Or ComboBox6.Text = "SiO2(к) кристобалит" Or ComboBox6.Text = "SiO2 (полиморф.)" Then
                        TextBox26.Text = hbetaSiO2 : TextBox27.Text = sbetaSiO2 : TextBox28.Text = abetaSiO2 : TextBox29.Text = bbetaSiO2 : TextBox30.Text = cbetaSiO2
                    End If
                    If ComboBox7.Text = "SiO2 (β-кварц)" Or ComboBox7.Text = "SiO2 (α-кварц)" Or ComboBox7.Text = "SiO2 (α-тридимит)" Or ComboBox7.Text = "SiO2 (β-кристобалит)" Or ComboBox7.Text = "SiO2(к) кварц" Or ComboBox7.Text = "SiO2(к) тридимит" Or ComboBox7.Text = "SiO2(к) кристобалит" Or ComboBox7.Text = "SiO2 (полиморф.)" Then
                        TextBox31.Text = hbetaSiO2 : TextBox32.Text = sbetaSiO2 : TextBox33.Text = abetaSiO2 : TextBox34.Text = bbetaSiO2 : TextBox35.Text = cbetaSiO2
                    End If
                    If ComboBox8.Text = "SiO2 (β-кварц)" Or ComboBox8.Text = "SiO2 (α-кварц)" Or ComboBox8.Text = "SiO2 (α-тридимит)" Or ComboBox8.Text = "SiO2 (β-кристобалит)" Or ComboBox8.Text = "SiO2(к) кварц" Or ComboBox8.Text = "SiO2(к) тридимит" Or ComboBox8.Text = "SiO2(к) кристобалит" Or ComboBox8.Text = "SiO2 (полиморф.)" Then
                        TextBox36.Text = hbetaSiO2 : TextBox37.Text = sbetaSiO2 : TextBox38.Text = abetaSiO2 : TextBox39.Text = bbetaSiO2 : TextBox40.Text = cbetaSiO2
                    End If
                End If

                If z > 846.15 Then
                    If ComboBox1.Text = "SiO2 (β-кварц)" Or ComboBox1.Text = "SiO2 (α-кварц)" Or ComboBox1.Text = "SiO2 (α-тридимит)" Or ComboBox1.Text = "SiO2 (β-кристобалит)" Or ComboBox1.Text = "SiO2(к) кварц" Or ComboBox1.Text = "SiO2(к) тридимит" Or ComboBox1.Text = "SiO2(к) кристобалит" Or ComboBox1.Text = "SiO2 (полиморф.)" Then
                        TextBox1.Text = halphaSiO2 : TextBox2.Text = salphaSiO2 : TextBox3.Text = aalphaSiO2 : TextBox4.Text = balphaSiO2 : TextBox5.Text = calphaSiO2
                    End If
                    If ComboBox2.Text = "SiO2 (β-кварц)" Or ComboBox2.Text = "SiO2 (α-кварц)" Or ComboBox2.Text = "SiO2 (α-тридимит)" Or ComboBox2.Text = "SiO2 (β-кристобалит)" Or ComboBox2.Text = "SiO2(к) кварц" Or ComboBox2.Text = "SiO2(к) тридимит" Or ComboBox2.Text = "SiO2(к) кристобалит" Or ComboBox2.Text = "SiO2 (полиморф.)" Then
                        TextBox6.Text = halphaSiO2 : TextBox7.Text = salphaSiO2 : TextBox8.Text = aalphaSiO2 : TextBox9.Text = balphaSiO2 : TextBox10.Text = calphaSiO2
                    End If
                    If ComboBox3.Text = "SiO2 (β-кварц)" Or ComboBox3.Text = "SiO2 (α-кварц)" Or ComboBox3.Text = "SiO2 (α-тридимит)" Or ComboBox3.Text = "SiO2 (β-кристобалит)" Or ComboBox3.Text = "SiO2(к) кварц" Or ComboBox3.Text = "SiO2(к) тридимит" Or ComboBox3.Text = "SiO2(к) кристобалит" Or ComboBox3.Text = "SiO2 (полиморф.)" Then
                        TextBox11.Text = halphaSiO2 : TextBox12.Text = salphaSiO2 : TextBox13.Text = aalphaSiO2 : TextBox14.Text = balphaSiO2 : TextBox15.Text = calphaSiO2
                    End If
                    If ComboBox4.Text = "SiO2 (β-кварц)" Or ComboBox4.Text = "SiO2 (α-кварц)" Or ComboBox4.Text = "SiO2 (α-тридимит)" Or ComboBox4.Text = "SiO2 (β-кристобалит)" Or ComboBox4.Text = "SiO2(к) кварц" Or ComboBox4.Text = "SiO2(к) тридимит" Or ComboBox4.Text = "SiO2(к) кристобалит" Or ComboBox4.Text = "SiO2 (полиморф.)" Then
                        TextBox16.Text = halphaSiO2 : TextBox17.Text = salphaSiO2 : TextBox18.Text = aalphaSiO2 : TextBox19.Text = balphaSiO2 : TextBox20.Text = calphaSiO2
                    End If
                    If ComboBox5.Text = "SiO2 (β-кварц)" Or ComboBox5.Text = "SiO2 (α-кварц)" Or ComboBox5.Text = "SiO2 (α-тридимит)" Or ComboBox5.Text = "SiO2 (β-кристобалит)" Or ComboBox5.Text = "SiO2(к) кварц" Or ComboBox5.Text = "SiO2(к) тридимит" Or ComboBox5.Text = "SiO2(к) кристобалит" Or ComboBox5.Text = "SiO2 (полиморф.)" Then
                        TextBox21.Text = halphaSiO2 : TextBox22.Text = salphaSiO2 : TextBox23.Text = aalphaSiO2 : TextBox24.Text = balphaSiO2 : TextBox25.Text = calphaSiO2
                    End If
                    If ComboBox6.Text = "SiO2 (β-кварц)" Or ComboBox6.Text = "SiO2 (α-кварц)" Or ComboBox6.Text = "SiO2 (α-тридимит)" Or ComboBox6.Text = "SiO2 (β-кристобалит)" Or ComboBox6.Text = "SiO2(к) кварц" Or ComboBox6.Text = "SiO2(к) тридимит" Or ComboBox6.Text = "SiO2(к) кристобалит" Or ComboBox6.Text = "SiO2 (полиморф.)" Then
                        TextBox26.Text = halphaSiO2 : TextBox27.Text = salphaSiO2 : TextBox28.Text = aalphaSiO2 : TextBox29.Text = balphaSiO2 : TextBox30.Text = calphaSiO2
                    End If
                    If ComboBox7.Text = "SiO2 (β-кварц)" Or ComboBox7.Text = "SiO2 (α-кварц)" Or ComboBox7.Text = "SiO2 (α-тридимит)" Or ComboBox7.Text = "SiO2 (β-кристобалит)" Or ComboBox7.Text = "SiO2(к) кварц" Or ComboBox7.Text = "SiO2(к) тридимит" Or ComboBox7.Text = "SiO2(к) кристобалит" Or ComboBox7.Text = "SiO2 (полиморф.)" Then
                        TextBox31.Text = halphaSiO2 : TextBox32.Text = salphaSiO2 : TextBox33.Text = aalphaSiO2 : TextBox34.Text = balphaSiO2 : TextBox35.Text = calphaSiO2
                    End If
                    If ComboBox8.Text = "SiO2 (β-кварц)" Or ComboBox8.Text = "SiO2 (α-кварц)" Or ComboBox8.Text = "SiO2 (α-тридимит)" Or ComboBox8.Text = "SiO2 (β-кристобалит)" Or ComboBox8.Text = "SiO2(к) кварц" Or ComboBox8.Text = "SiO2(к) тридимит" Or ComboBox8.Text = "SiO2(к) кристобалит" Or ComboBox8.Text = "SiO2 (полиморф.)" Then
                        TextBox36.Text = halphaSiO2 : TextBox37.Text = salphaSiO2 : TextBox38.Text = aalphaSiO2 : TextBox39.Text = balphaSiO2 : TextBox40.Text = calphaSiO2
                    End If
                End If

                If z > 1143.15 Then
                    If ComboBox1.Text = "SiO2 (β-кварц)" Or ComboBox1.Text = "SiO2 (α-кварц)" Or ComboBox1.Text = "SiO2 (α-тридимит)" Or ComboBox1.Text = "SiO2 (β-кристобалит)" Or ComboBox1.Text = "SiO2(к) кварц" Or ComboBox1.Text = "SiO2(к) тридимит" Or ComboBox1.Text = "SiO2(к) кристобалит" Or ComboBox1.Text = "SiO2 (полиморф.)" Then
                        TextBox1.Text = htridimitSiO2 : TextBox2.Text = stridimitSiO2 : TextBox3.Text = atridimitSiO2 : TextBox4.Text = btridimitSiO2 : TextBox5.Text = ctridimitSiO2
                    End If
                    If ComboBox2.Text = "SiO2 (β-кварц)" Or ComboBox2.Text = "SiO2 (α-кварц)" Or ComboBox2.Text = "SiO2 (α-тридимит)" Or ComboBox2.Text = "SiO2 (β-кристобалит)" Or ComboBox2.Text = "SiO2(к) кварц" Or ComboBox2.Text = "SiO2(к) тридимит" Or ComboBox2.Text = "SiO2(к) кристобалит" Or ComboBox2.Text = "SiO2 (полиморф.)" Then
                        TextBox6.Text = htridimitSiO2 : TextBox7.Text = stridimitSiO2 : TextBox8.Text = atridimitSiO2 : TextBox9.Text = btridimitSiO2 : TextBox10.Text = ctridimitSiO2
                    End If
                    If ComboBox3.Text = "SiO2 (β-кварц)" Or ComboBox3.Text = "SiO2 (α-кварц)" Or ComboBox3.Text = "SiO2 (α-тридимит)" Or ComboBox3.Text = "SiO2 (β-кристобалит)" Or ComboBox3.Text = "SiO2(к) кварц" Or ComboBox3.Text = "SiO2(к) тридимит" Or ComboBox3.Text = "SiO2(к) кристобалит" Or ComboBox3.Text = "SiO2 (полиморф.)" Then
                        TextBox11.Text = htridimitSiO2 : TextBox12.Text = stridimitSiO2 : TextBox13.Text = atridimitSiO2 : TextBox14.Text = btridimitSiO2 : TextBox15.Text = ctridimitSiO2
                    End If
                    If ComboBox4.Text = "SiO2 (β-кварц)" Or ComboBox4.Text = "SiO2 (α-кварц)" Or ComboBox4.Text = "SiO2 (α-тридимит)" Or ComboBox4.Text = "SiO2 (β-кристобалит)" Or ComboBox4.Text = "SiO2(к) кварц" Or ComboBox4.Text = "SiO2(к) тридимит" Or ComboBox4.Text = "SiO2(к) кристобалит" Or ComboBox4.Text = "SiO2 (полиморф.)" Then
                        TextBox16.Text = htridimitSiO2 : TextBox17.Text = stridimitSiO2 : TextBox18.Text = atridimitSiO2 : TextBox19.Text = btridimitSiO2 : TextBox20.Text = ctridimitSiO2
                    End If
                    If ComboBox5.Text = "SiO2 (β-кварц)" Or ComboBox5.Text = "SiO2 (α-кварц)" Or ComboBox5.Text = "SiO2 (α-тридимит)" Or ComboBox5.Text = "SiO2 (β-кристобалит)" Or ComboBox5.Text = "SiO2(к) кварц" Or ComboBox5.Text = "SiO2(к) тридимит" Or ComboBox5.Text = "SiO2(к) кристобалит" Or ComboBox5.Text = "SiO2 (полиморф.)" Then
                        TextBox21.Text = htridimitSiO2 : TextBox22.Text = stridimitSiO2 : TextBox23.Text = atridimitSiO2 : TextBox24.Text = btridimitSiO2 : TextBox25.Text = ctridimitSiO2
                    End If
                    If ComboBox6.Text = "SiO2 (β-кварц)" Or ComboBox6.Text = "SiO2 (α-кварц)" Or ComboBox6.Text = "SiO2 (α-тридимит)" Or ComboBox6.Text = "SiO2 (β-кристобалит)" Or ComboBox6.Text = "SiO2(к) кварц" Or ComboBox6.Text = "SiO2(к) тридимит" Or ComboBox6.Text = "SiO2(к) кристобалит" Or ComboBox6.Text = "SiO2 (полиморф.)" Then
                        TextBox26.Text = htridimitSiO2 : TextBox27.Text = stridimitSiO2 : TextBox28.Text = atridimitSiO2 : TextBox29.Text = btridimitSiO2 : TextBox30.Text = ctridimitSiO2
                    End If
                    If ComboBox7.Text = "SiO2 (β-кварц)" Or ComboBox7.Text = "SiO2 (α-кварц)" Or ComboBox7.Text = "SiO2 (α-тридимит)" Or ComboBox7.Text = "SiO2 (β-кристобалит)" Or ComboBox7.Text = "SiO2(к) кварц" Or ComboBox7.Text = "SiO2(к) тридимит" Or ComboBox7.Text = "SiO2(к) кристобалит" Or ComboBox7.Text = "SiO2 (полиморф.)" Then
                        TextBox31.Text = htridimitSiO2 : TextBox32.Text = stridimitSiO2 : TextBox33.Text = atridimitSiO2 : TextBox34.Text = btridimitSiO2 : TextBox35.Text = ctridimitSiO2
                    End If
                    If ComboBox8.Text = "SiO2 (β-кварц)" Or ComboBox8.Text = "SiO2 (α-кварц)" Or ComboBox8.Text = "SiO2 (α-тридимит)" Or ComboBox8.Text = "SiO2 (β-кристобалит)" Or ComboBox8.Text = "SiO2(к) кварц" Or ComboBox8.Text = "SiO2(к) тридимит" Or ComboBox8.Text = "SiO2(к) кристобалит" Or ComboBox8.Text = "SiO2 (полиморф.)" Then
                        TextBox36.Text = htridimitSiO2 : TextBox37.Text = stridimitSiO2 : TextBox38.Text = atridimitSiO2 : TextBox39.Text = btridimitSiO2 : TextBox40.Text = ctridimitSiO2
                    End If
                End If

                If z >= 1743.15 Then
                    If ComboBox1.Text = "SiO2 (β-кварц)" Or ComboBox1.Text = "SiO2 (α-кварц)" Or ComboBox1.Text = "SiO2 (α-тридимит)" Or ComboBox1.Text = "SiO2 (β-кристобалит)" Or ComboBox1.Text = "SiO2(к) кварц" Or ComboBox1.Text = "SiO2(к) тридимит" Or ComboBox1.Text = "SiO2(к) кристобалит" Or ComboBox1.Text = "SiO2 (полиморф.)" Then
                        TextBox1.Text = hkristobalitSiO2 : TextBox2.Text = skristobalitSiO2 : TextBox3.Text = akristobalitSiO2 : TextBox4.Text = bkristobalitSiO2 : TextBox5.Text = ckristobalitSiO2
                    End If
                    If ComboBox2.Text = "SiO2 (β-кварц)" Or ComboBox2.Text = "SiO2 (α-кварц)" Or ComboBox2.Text = "SiO2 (α-тридимит)" Or ComboBox2.Text = "SiO2 (β-кристобалит)" Or ComboBox2.Text = "SiO2(к) кварц" Or ComboBox2.Text = "SiO2(к) тридимит" Or ComboBox2.Text = "SiO2(к) кристобалит" Or ComboBox2.Text = "SiO2 (полиморф.)" Then
                        TextBox6.Text = hkristobalitSiO2 : TextBox7.Text = skristobalitSiO2 : TextBox8.Text = akristobalitSiO2 : TextBox9.Text = bkristobalitSiO2 : TextBox10.Text = ckristobalitSiO2
                    End If
                    If ComboBox3.Text = "SiO2 (β-кварц)" Or ComboBox3.Text = "SiO2 (α-кварц)" Or ComboBox3.Text = "SiO2 (α-тридимит)" Or ComboBox3.Text = "SiO2 (β-кристобалит)" Or ComboBox3.Text = "SiO2(к) кварц" Or ComboBox3.Text = "SiO2(к) тридимит" Or ComboBox3.Text = "SiO2(к) кристобалит" Or ComboBox3.Text = "SiO2 (полиморф.)" Then
                        TextBox11.Text = hkristobalitSiO2 : TextBox12.Text = skristobalitSiO2 : TextBox13.Text = akristobalitSiO2 : TextBox14.Text = bkristobalitSiO2 : TextBox15.Text = ckristobalitSiO2
                    End If
                    If ComboBox4.Text = "SiO2 (β-кварц)" Or ComboBox4.Text = "SiO2 (α-кварц)" Or ComboBox4.Text = "SiO2 (α-тридимит)" Or ComboBox4.Text = "SiO2 (β-кристобалит)" Or ComboBox4.Text = "SiO2(к) кварц" Or ComboBox4.Text = "SiO2(к) тридимит" Or ComboBox4.Text = "SiO2(к) кристобалит" Or ComboBox4.Text = "SiO2 (полиморф.)" Then
                        TextBox16.Text = hkristobalitSiO2 : TextBox17.Text = skristobalitSiO2 : TextBox18.Text = akristobalitSiO2 : TextBox19.Text = bkristobalitSiO2 : TextBox20.Text = ckristobalitSiO2
                    End If
                    If ComboBox5.Text = "SiO2 (β-кварц)" Or ComboBox5.Text = "SiO2 (α-кварц)" Or ComboBox5.Text = "SiO2 (α-тридимит)" Or ComboBox5.Text = "SiO2 (β-кристобалит)" Or ComboBox5.Text = "SiO2(к) кварц" Or ComboBox5.Text = "SiO2(к) тридимит" Or ComboBox5.Text = "SiO2(к) кристобалит" Or ComboBox5.Text = "SiO2 (полиморф.)" Then
                        TextBox21.Text = hkristobalitSiO2 : TextBox22.Text = skristobalitSiO2 : TextBox23.Text = akristobalitSiO2 : TextBox24.Text = bkristobalitSiO2 : TextBox25.Text = ckristobalitSiO2
                    End If
                    If ComboBox6.Text = "SiO2 (β-кварц)" Or ComboBox6.Text = "SiO2 (α-кварц)" Or ComboBox6.Text = "SiO2 (α-тридимит)" Or ComboBox6.Text = "SiO2 (β-кристобалит)" Or ComboBox6.Text = "SiO2(к) кварц" Or ComboBox6.Text = "SiO2(к) тридимит" Or ComboBox6.Text = "SiO2(к) кристобалит" Or ComboBox6.Text = "SiO2 (полиморф.)" Then
                        TextBox26.Text = hkristobalitSiO2 : TextBox27.Text = skristobalitSiO2 : TextBox28.Text = akristobalitSiO2 : TextBox29.Text = bkristobalitSiO2 : TextBox30.Text = ckristobalitSiO2
                    End If
                    If ComboBox7.Text = "SiO2 (β-кварц)" Or ComboBox7.Text = "SiO2 (α-кварц)" Or ComboBox7.Text = "SiO2 (α-тридимит)" Or ComboBox7.Text = "SiO2 (β-кристобалит)" Or ComboBox7.Text = "SiO2(к) кварц" Or ComboBox7.Text = "SiO2(к) тридимит" Or ComboBox7.Text = "SiO2(к) кристобалит" Or ComboBox7.Text = "SiO2 (полиморф.)" Then
                        TextBox31.Text = hkristobalitSiO2 : TextBox32.Text = skristobalitSiO2 : TextBox33.Text = akristobalitSiO2 : TextBox34.Text = bkristobalitSiO2 : TextBox35.Text = ckristobalitSiO2
                    End If
                    If ComboBox8.Text = "SiO2 (β-кварц)" Or ComboBox8.Text = "SiO2 (α-кварц)" Or ComboBox8.Text = "SiO2 (α-тридимит)" Or ComboBox8.Text = "SiO2 (β-кристобалит)" Or ComboBox8.Text = "SiO2(к) кварц" Or ComboBox8.Text = "SiO2(к) тридимит" Or ComboBox8.Text = "SiO2(к) кристобалит" Or ComboBox8.Text = "SiO2 (полиморф.)" Then
                        TextBox36.Text = hkristobalitSiO2 : TextBox37.Text = skristobalitSiO2 : TextBox38.Text = akristobalitSiO2 : TextBox39.Text = bkristobalitSiO2 : TextBox40.Text = ckristobalitSiO2
                    End If
                End If
            End If

            'Считаем константы теплоемкости a,b,c
            a = (((TextBox05.Text) * (TextBox23.Text) + (TextBox06.Text) * (TextBox28.Text) + (TextBox07.Text) * (TextBox33.Text) + (TextBox08.Text) * (TextBox38.Text)) - ((TextBox01.Text) * (TextBox3.Text) + (TextBox02.Text) * (TextBox8.Text) + (TextBox03.Text) * (TextBox13.Text) + (TextBox04.Text) * (TextBox18.Text)))
            b = ((((TextBox05.Text) * (TextBox24.Text) + (TextBox06.Text) * (TextBox29.Text) + (TextBox07.Text) * (TextBox34.Text) + (TextBox08.Text) * (TextBox39.Text)) - ((TextBox01.Text) * (TextBox4.Text) + (TextBox02.Text) * (TextBox9.Text) + (TextBox03.Text) * (TextBox14.Text) + (TextBox04.Text) * (TextBox19.Text))) / 1000)
            c = ((((TextBox05.Text) * (TextBox25.Text) + (TextBox06.Text) * (TextBox30.Text) + (TextBox07.Text) * (TextBox35.Text) + (TextBox08.Text) * (TextBox40.Text)) - ((TextBox01.Text) * (TextBox5.Text) + (TextBox02.Text) * (TextBox10.Text) + (TextBox03.Text) * (TextBox15.Text) + (TextBox04.Text) * (TextBox20.Text))) * 100000)

            'Считаем ∆H, ∆S, ∆G при температуре 298K
            H298 = ((TextBox05.Text) * (TextBox21.Text) + (TextBox06.Text) * (TextBox26.Text) + (TextBox07.Text) * (TextBox31.Text) + (TextBox08.Text) * (TextBox36.Text)) - ((TextBox01.Text) * (TextBox1.Text) + (TextBox02.Text) * (TextBox6.Text) + (TextBox03.Text) * (TextBox11.Text) + (TextBox04.Text) * (TextBox16.Text))
            S298 = ((TextBox05.Text) * (TextBox22.Text) + (TextBox06.Text) * (TextBox27.Text) + (TextBox07.Text) * (TextBox32.Text) + (TextBox08.Text) * (TextBox37.Text)) - ((TextBox01.Text) * (TextBox2.Text) + (TextBox02.Text) * (TextBox7.Text) + (TextBox03.Text) * (TextBox12.Text) + (TextBox04.Text) * (TextBox17.Text))
            G298 = ((H298) - 298 * (S298))
            If RadioButton2.Checked Then
                G298 = G298 / 4.184
            End If

            If RadioButton4.Checked Then
                'Считаем ∆H, ∆S, ∆G при заданной температуре
                Ht = ((H298 * 1000) + ((a) * (z + 273.15)) - (298 * (a)) + ((b) * ((z + 273.15) ^ 2)) / 2 - ((298 ^ 2) * (b)) / 2 - ((c) / (z + 273.15)) + ((c) / 298))
                St = ((S298) + (a) * (Math.Log(z + 273.15)) - (a) * (Math.Log(298)) + (b) * (z + 273.15) - 298 * (b) - (c) / (2 * (z + 273.15) ^ 2) + (c) / (2 * 298 ^ 2))
                Gt = (((Ht - (z + 273.15) * St) / 1000) * 4.184)
            Else

                'Считаем ∆H, ∆S, ∆G при заданной температуре
                Ht = ((H298 * 1000) + ((a) * (z)) - (298 * (a)) + ((b) * (z ^ 2)) / 2 - ((298 ^ 2) * (b)) / 2 - ((c) / (z)) + ((c) / 298))
                St = ((S298) + (a) * (Math.Log(z)) - (a) * (Math.Log(298)) + (b) * (z) - 298 * (b) - (c) / (2 * (z) ^ 2) + (c) / (2 * 298 ^ 2))
                Gt = (((Ht - z * St) / 1000) * 4.184)
            End If

            If RadioButton2.Checked Then
                Gt = Gt / 4.184
            End If

            'Рисуем точки на графике
            If RadioButton4.Checked Then
                Me.Chart1.Series("Series1").Points.AddXY((z), Format(Gt, "#"))
            Else
                Me.Chart1.Series("Series1").Points.AddXY(z, Format(Gt, "#"))
            End If

            'Печатаем данные в виде таблицы в окне экспорта в Excel
            If RadioButton4.Checked Then
                forexcel.TextBoxG.AppendText((Format(Gt, "0.#")) & vbCrLf)
                forexcel.TextBoxH.AppendText((Format(((Ht / 1000) * 4.164), "0.#")) & vbCrLf)
                forexcel.TextBoxS.AppendText((Format((St * 4.164), "0.#")) & vbCrLf)
                forexcel.TextBoxTemp.AppendText((z) & vbCrLf)
                forexcel.TextBoxTempK.AppendText(z + 273.15 & vbCrLf)
            Else
                forexcel.TextBoxG.AppendText((Format(Gt, "0.#")) & vbCrLf)
                forexcel.TextBoxH.AppendText((Format(((Ht / 1000) * 4.164), "0.#")) & vbCrLf)
                forexcel.TextBoxS.AppendText((Format((St * 4.164), "0.#")) & vbCrLf)
                forexcel.TextBoxTemp.AppendText((z - 273.15) & vbCrLf)
                forexcel.TextBoxTempK.AppendText(z & vbCrLf)
            End If

        Next

        'Считаем первое значение ∆G для вывода в поле "∆G от..."
        If RadioButton4.Checked Then
            HtOt = ((H298 * 1000) + ((a) * ((TextBox41.Text) + 273.15)) - (298 * (a)) + ((b) * ((TextBox41.Text + 273.15) ^ 2)) / 2 - ((298 ^ 2) * (b)) / 2 - ((c) / ((TextBox41.Text) + 273.15)) + ((c) / 298))
            StOt = ((S298) + (a) * (Math.Log(TextBox41.Text + 273.15)) - (a) * (Math.Log(298)) + (b) * (TextBox41.Text + 273.15) - 298 * (b) - (c) / (2 * (TextBox41.Text + 273.15) ^ 2) + (c) / (2 * 298 ^ 2))
            GtOt = (((HtOt - (TextBox41.Text + 273.15) * StOt) / 1000) * 4.184)
        Else
            HtOt = ((H298 * 1000) + ((a) * ((TextBox41.Text))) - (298 * (a)) + ((b) * ((TextBox41.Text) ^ 2)) / 2 - ((298 ^ 2) * (b)) / 2 - ((c) / ((TextBox41.Text))) + ((c) / 298))
            StOt = ((S298) + (a) * (Math.Log(TextBox41.Text)) - (a) * (Math.Log(298)) + (b) * (TextBox41.Text) - 298 * (b) - (c) / (2 * (TextBox41.Text) ^ 2) + (c) / (2 * 298 ^ 2))
            GtOt = (((HtOt - (TextBox41.Text) * StOt) / 1000) * 4.184)
        End If
        If RadioButton2.Checked Then
            GtOt = GtOt / 4.184
        End If

        'Определяем и выводим возможность протекания реакции
        Label12.Text = "∆G = " & "от " & (Format(GtOt, "0.#")) & " до " & (Format(Gt, "0.#")) & " кДж"
        If (GtOt) >= 0 And (Gt) >= 0 Then
            Label13.Text = "Реакция невозможна"
            Label13.ForeColor = Color.DarkRed
        ElseIf (GtOt) < 0 And (Gt) < 0 Then
            Label13.Text = "Реакция возможна"
            Label13.ForeColor = Color.DarkGreen
        ElseIf (GtOt) < 0 And (Gt) > 0 Then
            Label13.Text = "Реакция ограничена"
            Label13.ForeColor = Color.Orange
        ElseIf (GtOt) > 0 And (Gt) < 0 Then
            Label13.Text = "Реакция ограничена"
            Label13.ForeColor = Color.Orange
        End If


        'Прописываем реакцию
        Dim r1, r2, r3, r4, r5, r6, r7, r8 As String
        If ComboBox1.SelectedIndex > 0 Then r1 = TextBox01.Text & "(" & ComboBox1.SelectedItem & ")" Else r1 = ""
        If ComboBox2.SelectedIndex > 0 Then r2 = " + " & TextBox02.Text & "(" & ComboBox2.SelectedItem & ")" Else r2 = ""
        If ComboBox3.SelectedIndex > 0 Then r3 = " + " & TextBox03.Text & "(" & ComboBox3.SelectedItem & ")" Else r3 = ""
        If ComboBox4.SelectedIndex > 0 Then r4 = " + " & TextBox04.Text & "(" & ComboBox4.SelectedItem & ")" Else r4 = ""
        If ComboBox5.SelectedIndex > 0 Then r5 = TextBox05.Text & "(" & ComboBox5.SelectedItem & ")" Else r5 = ""
        If ComboBox6.SelectedIndex > 0 Then r6 = " + " & TextBox06.Text & "(" & ComboBox6.SelectedItem & ")" Else r6 = ""
        If ComboBox7.SelectedIndex > 0 Then r7 = " + " & TextBox07.Text & "(" & ComboBox7.SelectedItem & ")" Else r7 = ""
        If ComboBox8.SelectedIndex > 0 Then r8 = " + " & TextBox08.Text & "(" & ComboBox8.SelectedItem & ")" Else r8 = ""


        LabelReaction.Text = r1 + r2 + r3 + r4 + " = " + r5 + r6 + r7 + r8


    End Sub

    'Обработка кнопки для сохранения диаграммы
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        'Создаем новый SaveFileDialog
        Dim saveFileDialog1 As New SaveFileDialog()

        saveFileDialog1.Filter = "Bitmap (*.bmp)|*.bmp|JPEG (*.jpg)|*.jpg|EMF (*.emf)|*.emf|PNG (*.png)|*.png|GIF (*.gif)|*.gif|TIFF (*.tif)|*.tif"
        saveFileDialog1.FilterIndex = 2
        saveFileDialog1.RestoreDirectory = True

        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim format As ChartImageFormat = ChartImageFormat.Bmp

            If saveFileDialog1.FileName.EndsWith("bmp") Then
                format = ChartImageFormat.Bmp
            Else
                If saveFileDialog1.FileName.EndsWith("jpg") Then
                    format = ChartImageFormat.Jpeg
                Else
                    If saveFileDialog1.FileName.EndsWith("emf") Then
                        format = ChartImageFormat.Emf
                    Else
                        If saveFileDialog1.FileName.EndsWith("gif") Then
                            format = ChartImageFormat.Gif
                        Else
                            If saveFileDialog1.FileName.EndsWith("png") Then
                                format = ChartImageFormat.Png
                            Else
                                If saveFileDialog1.FileName.EndsWith("tif") Then
                                    format = ChartImageFormat.Tiff
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            Chart1.SaveImage(saveFileDialog1.FileName, format)
        End If

    End Sub

    Private Sub ComboBoxBase_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ComboBoxBase.SelectedIndexChanged

        'Кнопка редактирования персональной базы соединений
        If ComboBoxBase.SelectedItem = "Персональная база соединений" Then PictureBox2.Visible = True Else PictureBox2.Visible = False

        For i As Integer = 1 To 40
            Me.Controls("TextBox" & i).Text = "0"
        Next

        For i As Integer = 1 To 8
            Me.Controls("TextBox0" & i).Text = "1"
        Next

        LabelReaction.Text = "Введите реакцию и нажмите «Расчёт»"
        Label12.Text = "∆G = 0"
        Label13.Text = ""
        forexcel.TextBoxG.Text = ""
        forexcel.TextBoxH.Text = ""
        forexcel.TextBoxS.Text = ""
        forexcel.TextBoxTemp.Text = ""
        forexcel.TextBoxTempK.Text = ""
        Me.Chart1.Series("Series1").Points.Clear()

        Data.Clear() '//////
        For Each Line As String In IO.File.ReadAllLines("termodynamics" & ComboBoxBase.SelectedIndex + 1 & ".txt", System.Text.Encoding.Default)
            Dim NewData As New MyData
            With NewData
                .Combo = Line.Split(New String() {"|"}, StringSplitOptions.RemoveEmptyEntries)(0)
                For I As Integer = 0 To 4
                    .T(I) = Line.Split(New String() {"|"}, StringSplitOptions.RemoveEmptyEntries)(I + 1)
                Next
            End With
            Data.Add(NewData)
        Next
        For I As Integer = 1 To 8
            DirectCast(Controls("ComboBox" & I), ComboBox).Items.Clear() '///
            DirectCast(Controls("ComboBox" & I), ComboBox).Items.AddRange((From s In Data Select s.Combo).ToArray)
        Next
    End Sub
    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        forexcel.ShowDialog()
        Show()
    End Sub
    Private Sub ВыходToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВыходToolStripMenuItem.Click
        Close()
    End Sub
    Private Sub ОПрограммеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОПрограммеToolStripMenuItem.Click
        About.ShowDialog()
        Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        For i As Integer = 1 To 40
            Me.Controls("TextBox" & i).Text = "0"
        Next

        For i As Integer = 1 To 8
            Me.Controls("TextBox0" & i).Text = "1"
        Next

        For i As Integer = 1 To 8
            DirectCast(Me.Controls("ComboBox" & i), ComboBox).SelectedIndex = 0
        Next

        LabelReaction.Text = "Введите реакцию и нажмите «Расчёт»"
        Label12.Text = "∆G = 0"
        Label13.Text = ""
        forexcel.TextBoxG.Text = ""
        forexcel.TextBoxH.Text = ""
        forexcel.TextBoxS.Text = ""
        forexcel.TextBoxTemp.Text = ""
        forexcel.TextBoxTempK.Text = ""
        Me.Chart1.Series("Series1").Points.Clear()
    End Sub

    Private Sub СправкаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СправкаToolStripMenuItem.Click
        Help.ShowDialog()
        Show()
    End Sub

    Private Sub ComboBox_SelectedIndexChanged2(sender As Object, e As EventArgs) Handles MyBase.Load, ComboBox1.SelectedIndexChanged, ComboBox2.SelectedIndexChanged, ComboBox3.SelectedIndexChanged, ComboBox4.SelectedIndexChanged, ComboBox5.SelectedIndexChanged, ComboBox6.SelectedIndexChanged, ComboBox7.SelectedIndexChanged, ComboBox8.SelectedIndexChanged

        If ComboBox1.SelectedIndex > 0 Then
            ComboBox2.Enabled = True : TextBox02.Enabled = True : TextBox6.Enabled = True : TextBox7.Enabled = True : TextBox8.Enabled = True : TextBox9.Enabled = True : TextBox10.Enabled = True
        Else
            ComboBox2.Enabled = False : TextBox02.Enabled = False : TextBox6.Enabled = False : TextBox7.Enabled = False : TextBox8.Enabled = False : TextBox9.Enabled = False : TextBox10.Enabled = False
        End If

        If ComboBox2.SelectedIndex > 0 Then
            ComboBox3.Enabled = True : TextBox03.Enabled = True : TextBox11.Enabled = True : TextBox12.Enabled = True : TextBox13.Enabled = True : TextBox14.Enabled = True : TextBox15.Enabled = True
        Else
            ComboBox3.Enabled = False : TextBox03.Enabled = False : TextBox11.Enabled = False : TextBox12.Enabled = False : TextBox13.Enabled = False : TextBox14.Enabled = False : TextBox15.Enabled = False
        End If

        If ComboBox3.SelectedIndex > 0 Then
            ComboBox4.Enabled = True : TextBox04.Enabled = True : TextBox16.Enabled = True : TextBox17.Enabled = True : TextBox18.Enabled = True : TextBox19.Enabled = True : TextBox20.Enabled = True
        Else
            ComboBox4.Enabled = False : TextBox04.Enabled = False : TextBox16.Enabled = False : TextBox17.Enabled = False : TextBox18.Enabled = False : TextBox19.Enabled = False : TextBox20.Enabled = False
        End If

        If ComboBox5.SelectedIndex > 0 Then
            ComboBox6.Enabled = True : TextBox06.Enabled = True : TextBox26.Enabled = True : TextBox27.Enabled = True : TextBox28.Enabled = True : TextBox29.Enabled = True : TextBox30.Enabled = True
        Else
            ComboBox6.Enabled = False : TextBox06.Enabled = False : TextBox26.Enabled = False : TextBox27.Enabled = False : TextBox28.Enabled = False : TextBox29.Enabled = False : TextBox30.Enabled = False
        End If

        If ComboBox6.SelectedIndex > 0 Then
            ComboBox7.Enabled = True : TextBox07.Enabled = True : TextBox31.Enabled = True : TextBox32.Enabled = True : TextBox33.Enabled = True : TextBox34.Enabled = True : TextBox35.Enabled = True
        Else
            ComboBox7.Enabled = False : TextBox07.Enabled = False : TextBox31.Enabled = False : TextBox32.Enabled = False : TextBox33.Enabled = False : TextBox34.Enabled = False : TextBox35.Enabled = False
        End If

        If ComboBox7.SelectedIndex > 0 Then
            ComboBox8.Enabled = True : TextBox08.Enabled = True : TextBox36.Enabled = True : TextBox37.Enabled = True : TextBox38.Enabled = True : TextBox39.Enabled = True : TextBox40.Enabled = True
        Else
            ComboBox8.Enabled = False : TextBox08.Enabled = False : TextBox36.Enabled = False : TextBox37.Enabled = False : TextBox38.Enabled = False : TextBox39.Enabled = False : TextBox40.Enabled = False
        End If

        If ComboBox1.Text = "SiO2 (β-кварц)" Or ComboBox1.Text = "SiO2 (α-кварц)" Or ComboBox1.Text = "SiO2 (α-тридимит)" Or ComboBox1.Text = "SiO2 (β-кристобалит)" Or ComboBox1.Text = "SiO2(к) кварц" Or ComboBox1.Text = "SiO2(к) тридимит" Or ComboBox1.Text = "SiO2(к) кристобалит" Or ComboBox1.Text = "SiO2 (полиморф.)" Then
            CheckBox2.Enabled = True
        ElseIf ComboBox2.Text = "SiO2 (β-кварц)" Or ComboBox2.Text = "SiO2 (α-кварц)" Or ComboBox2.Text = "SiO2 (α-тридимит)" Or ComboBox2.Text = "SiO2 (β-кристобалит)" Or ComboBox2.Text = "SiO2(к) кварц" Or ComboBox2.Text = "SiO2(к) тридимит" Or ComboBox2.Text = "SiO2(к) кристобалит" Or ComboBox2.Text = "SiO2 (полиморф.)" Then
            CheckBox2.Enabled = True
        ElseIf ComboBox3.Text = "SiO2 (β-кварц)" Or ComboBox3.Text = "SiO2 (α-кварц)" Or ComboBox3.Text = "SiO2 (α-тридимит)" Or ComboBox3.Text = "SiO2 (β-кристобалит)" Or ComboBox3.Text = "SiO2(к) кварц" Or ComboBox3.Text = "SiO2(к) тридимит" Or ComboBox3.Text = "SiO2(к) кристобалит" Or ComboBox3.Text = "SiO2 (полиморф.)" Then
            CheckBox2.Enabled = True
        ElseIf ComboBox4.Text = "SiO2 (β-кварц)" Or ComboBox4.Text = "SiO2 (α-кварц)" Or ComboBox4.Text = "SiO2 (α-тридимит)" Or ComboBox4.Text = "SiO2 (β-кристобалит)" Or ComboBox4.Text = "SiO2(к) кварц" Or ComboBox4.Text = "SiO2(к) тридимит" Or ComboBox4.Text = "SiO2(к) кристобалит" Or ComboBox4.Text = "SiO2 (полиморф.)" Then
            CheckBox2.Enabled = True
        ElseIf ComboBox5.Text = "SiO2 (β-кварц)" Or ComboBox5.Text = "SiO2 (α-кварц)" Or ComboBox5.Text = "SiO2 (α-тридимит)" Or ComboBox5.Text = "SiO2 (β-кристобалит)" Or ComboBox5.Text = "SiO2(к) кварц" Or ComboBox5.Text = "SiO2(к) тридимит" Or ComboBox5.Text = "SiO2(к) кристобалит" Or ComboBox5.Text = "SiO2 (полиморф.)" Then
            CheckBox2.Enabled = True
        ElseIf ComboBox6.Text = "SiO2 (β-кварц)" Or ComboBox6.Text = "SiO2 (α-кварц)" Or ComboBox6.Text = "SiO2 (α-тридимит)" Or ComboBox6.Text = "SiO2 (β-кристобалит)" Or ComboBox6.Text = "SiO2(к) кварц" Or ComboBox6.Text = "SiO2(к) тридимит" Or ComboBox6.Text = "SiO2(к) кристобалит" Or ComboBox6.Text = "SiO2 (полиморф.)" Then
            CheckBox2.Enabled = True
        ElseIf ComboBox7.Text = "SiO2 (β-кварц)" Or ComboBox7.Text = "SiO2 (α-кварц)" Or ComboBox7.Text = "SiO2 (α-тридимит)" Or ComboBox7.Text = "SiO2 (β-кристобалит)" Or ComboBox7.Text = "SiO2(к) кварц" Or ComboBox7.Text = "SiO2(к) тридимит" Or ComboBox7.Text = "SiO2(к) кристобалит" Or ComboBox7.Text = "SiO2 (полиморф.)" Then
            CheckBox2.Enabled = True
        ElseIf ComboBox8.Text = "SiO2 (β-кварц)" Or ComboBox8.Text = "SiO2 (α-кварц)" Or ComboBox8.Text = "SiO2 (α-тридимит)" Or ComboBox8.Text = "SiO2 (β-кристобалит)" Or ComboBox8.Text = "SiO2(к) кварц" Or ComboBox8.Text = "SiO2(к) тридимит" Or ComboBox8.Text = "SiO2(к) кристобалит" Or ComboBox8.Text = "SiO2 (полиморф.)" Then
            CheckBox2.Enabled = True
        Else CheckBox2.Enabled = False
        End If

        If ComboBox1.Text = "αAl2O3" Or ComboBox1.Text = "γAl2O3" Or ComboBox1.Text = "Al2O3(к)" Or ComboBox1.Text = "Al2O3 (полиморф.)" Then
            CheckBox3.Enabled = True
        ElseIf ComboBox2.Text = "αAl2O3" Or ComboBox2.Text = "γAl2O3" Or ComboBox2.Text = "Al2O3(к)" Or ComboBox2.Text = "Al2O3 (полиморф.)" Then
            CheckBox3.Enabled = True
        ElseIf ComboBox3.Text = "αAl2O3" Or ComboBox3.Text = "γAl2O3" Or ComboBox3.Text = "Al2O3(к)" Or ComboBox3.Text = "Al2O3 (полиморф.)" Then
            CheckBox3.Enabled = True
        ElseIf ComboBox4.Text = "αAl2O3" Or ComboBox4.Text = "γAl2O3" Or ComboBox4.Text = "Al2O3(к)" Or ComboBox4.Text = "Al2O3 (полиморф.)" Then
            CheckBox3.Enabled = True
        ElseIf ComboBox5.Text = "αAl2O3" Or ComboBox5.Text = "γAl2O3" Or ComboBox5.Text = "Al2O3(к)" Or ComboBox5.Text = "Al2O3 (полиморф.)" Then
            CheckBox3.Enabled = True
        ElseIf ComboBox6.Text = "αAl2O3" Or ComboBox6.Text = "γAl2O3" Or ComboBox6.Text = "Al2O3(к)" Or ComboBox6.Text = "Al2O3 (полиморф.)" Then
            CheckBox3.Enabled = True
        ElseIf ComboBox7.Text = "αAl2O3" Or ComboBox7.Text = "γAl2O3" Or ComboBox7.Text = "Al2O3(к)" Or ComboBox7.Text = "Al2O3 (полиморф.)" Then
            CheckBox3.Enabled = True
        ElseIf ComboBox8.Text = "αAl2O3" Or ComboBox8.Text = "γAl2O3" Or ComboBox8.Text = "Al2O3(к)" Or ComboBox8.Text = "Al2O3 (полиморф.)" Then
            CheckBox2.Enabled = True
        Else CheckBox3.Enabled = False

        End If
    End Sub
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        BaseEditor.ShowDialog()
        Show()
    End Sub


    Dim WithEvents WBC As New WebBrowser
    Dim AvVersion As String ' сюда запишем инфу о новой версии 
    Private Sub ОбновлениеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОбновлениеToolStripMenuItem.Click
        'Проверка на доступность сайта с обновлениями, если есть обновление, то предлагаем обновиться
        'If My.Computer.Network.Ping("195.216.243.180") Then
        WBC.ScriptErrorsSuppressed() = True
        Dim url As String = "http://chefranov.name/thermodynamics" ' адрес страницы, где написан номер последней версии 
            WBC.Navigate(url)
        'Else
        '    MsgBox("Проверить обновления для программы не удалось. Проблемы с интернет-соединением или же сайтом с обновлениями недоступен.", vbOKOnly + vbCritical, "Ошибка")
        'End If

    End Sub

    Private Sub WBC_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WBC.DocumentCompleted
        AvVersion = WBC.Document.GetElementById("version").InnerText
        Dim TVersion As String = Application.ProductVersion 'Текущая версия программы
        If Val(AvVersion) > Val(TVersion) Then
            Dim Msg, Title, Response
            Msg = "Доступна новая версия программы - Thermodynamics " & AvVersion & ". Желаете обновить программу до последней версии?" 'сообщение
            Title = "Обновление программы" 'заголовок
            Response = MsgBox(Msg, vbYesNo + vbQuestion, Title)
            If Response = vbYes Then 'нажата кнопка "Да" (Yes)
                MsgBox("Сейчас в браузере откроется страница с последней версией программы. На открытой странице Вы сможете просмотреть примечания к выпуску и загрузить программу.", vbOKOnly + vbInformation, "Загрузка обновления")
                Process.Start("http://chefranov.name/thermodynamics")
            End If
            'Else
            '    MsgBox("Вы используете последнию версию программы.", vbOKOnly + vbInformation, "Обновлений нет")

        End If
    End Sub
End Class
