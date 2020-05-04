<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class forexcel
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(forexcel))
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TextBoxTemp = New System.Windows.Forms.TextBox()
        Me.TextBoxG = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBoxTempK = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.TextBoxH = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextBoxS = New System.Windows.Forms.TextBox()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.SystemColors.MenuBar
        Me.TextBox1.HideSelection = False
        Me.TextBox1.Location = New System.Drawing.Point(14, 161)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(454, 34)
        Me.TextBox1.TabIndex = 0
        Me.TextBox1.TabStop = False
        Me.TextBox1.Text = "Здесь вы можете получить значения, которые можно использовать в Excel." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Данные зн" &
    "ачения обычно требуются для построения графика средствами Excel."
        '
        'TextBoxTemp
        '
        Me.TextBoxTemp.BackColor = System.Drawing.SystemColors.MenuBar
        Me.TextBoxTemp.HideSelection = False
        Me.TextBoxTemp.Location = New System.Drawing.Point(12, 215)
        Me.TextBoxTemp.Multiline = True
        Me.TextBoxTemp.Name = "TextBoxTemp"
        Me.TextBoxTemp.ReadOnly = True
        Me.TextBoxTemp.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBoxTemp.Size = New System.Drawing.Size(90, 255)
        Me.TextBoxTemp.TabIndex = 1
        Me.TextBoxTemp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBoxG
        '
        Me.TextBoxG.BackColor = System.Drawing.SystemColors.MenuBar
        Me.TextBoxG.HideSelection = False
        Me.TextBoxG.Location = New System.Drawing.Point(204, 215)
        Me.TextBoxG.Multiline = True
        Me.TextBoxG.Name = "TextBoxG"
        Me.TextBoxG.ReadOnly = True
        Me.TextBoxG.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBoxG.Size = New System.Drawing.Size(84, 255)
        Me.TextBoxG.TabIndex = 2
        Me.TextBoxG.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(11, 199)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(91, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Температура, °C"
        '
        'TextBoxTempK
        '
        Me.TextBoxTempK.BackColor = System.Drawing.SystemColors.MenuBar
        Me.TextBoxTempK.HideSelection = False
        Me.TextBoxTempK.Location = New System.Drawing.Point(108, 215)
        Me.TextBoxTempK.Multiline = True
        Me.TextBoxTempK.Name = "TextBoxTempK"
        Me.TextBoxTempK.ReadOnly = True
        Me.TextBoxTempK.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBoxTempK.Size = New System.Drawing.Size(90, 255)
        Me.TextBoxTempK.TabIndex = 4
        Me.TextBoxTempK.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(111, 199)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(87, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Температура, K"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(221, 199)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(51, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "∆G, кДж"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(12, 476)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(456, 38)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "OK"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = Global.Thermodynamic.My.Resources.Resources.excel
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PictureBox1.Location = New System.Drawing.Point(14, 12)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(454, 143)
        Me.PictureBox1.TabIndex = 8
        Me.PictureBox1.TabStop = False
        '
        'TextBoxH
        '
        Me.TextBoxH.BackColor = System.Drawing.SystemColors.MenuBar
        Me.TextBoxH.HideSelection = False
        Me.TextBoxH.Location = New System.Drawing.Point(294, 215)
        Me.TextBoxH.Multiline = True
        Me.TextBoxH.Name = "TextBoxH"
        Me.TextBoxH.ReadOnly = True
        Me.TextBoxH.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBoxH.Size = New System.Drawing.Size(84, 255)
        Me.TextBoxH.TabIndex = 9
        Me.TextBoxH.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(308, 199)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(51, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "∆H, кДж"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(398, 199)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(44, 13)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "∆S, Дж"
        '
        'TextBoxS
        '
        Me.TextBoxS.BackColor = System.Drawing.SystemColors.MenuBar
        Me.TextBoxS.HideSelection = False
        Me.TextBoxS.Location = New System.Drawing.Point(384, 215)
        Me.TextBoxS.Multiline = True
        Me.TextBoxS.Name = "TextBoxS"
        Me.TextBoxS.ReadOnly = True
        Me.TextBoxS.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBoxS.Size = New System.Drawing.Size(84, 255)
        Me.TextBoxS.TabIndex = 11
        Me.TextBoxS.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'forexcel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(483, 529)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.TextBoxS)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TextBoxH)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBoxTempK)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBoxG)
        Me.Controls.Add(Me.TextBoxTemp)
        Me.Controls.Add(Me.TextBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "forexcel"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Экспорт в Excel"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents TextBoxTemp As TextBox
    Friend WithEvents TextBoxG As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents TextBoxTempK As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents TextBoxH As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents TextBoxS As TextBox
End Class
