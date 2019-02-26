<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.BtnProceed = New System.Windows.Forms.Button()
        Me.SN1 = New System.Windows.Forms.TextBox()
        Me.BtnCloseAdt = New System.Windows.Forms.Button()
        Me.CH1 = New System.Windows.Forms.CheckBox()
        Me.CH2 = New System.Windows.Forms.CheckBox()
        Me.CH3 = New System.Windows.Forms.CheckBox()
        Me.CH4 = New System.Windows.Forms.CheckBox()
        Me.CH5 = New System.Windows.Forms.CheckBox()
        Me.CH6 = New System.Windows.Forms.CheckBox()
        Me.ChAll = New System.Windows.Forms.CheckBox()
        Me.SN2 = New System.Windows.Forms.TextBox()
        Me.SN3 = New System.Windows.Forms.TextBox()
        Me.SN6 = New System.Windows.Forms.TextBox()
        Me.SN5 = New System.Windows.Forms.TextBox()
        Me.SN4 = New System.Windows.Forms.TextBox()
        Me.StFreq = New System.Windows.Forms.Button()
        Me.StVolt = New System.Windows.Forms.Button()
        Me.ACheck = New System.Windows.Forms.Button()
        Me.BtnInstrChk = New System.Windows.Forms.Button()
        Me.TxtPreLevFPath = New DevComponents.DotNetBar.Controls.TextBoxX()
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TxtTmrFormat = New DevComponents.DotNetBar.Controls.TextBoxX()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBoxISA = New DevComponents.DotNetBar.Controls.TextBoxX()
        Me.Browse = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ISAFolderDialog = New System.Windows.Forms.FolderBrowserDialog()
        Me.txtFreqAdd = New System.Windows.Forms.TextBox()
        Me.txtVoltAdd = New System.Windows.Forms.TextBox()
        Me.txtComNo = New System.Windows.Forms.TextBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'BtnProceed
        '
        Me.BtnProceed.Font = New System.Drawing.Font("SansSerif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.BtnProceed.Location = New System.Drawing.Point(439, 77)
        Me.BtnProceed.Name = "BtnProceed"
        Me.BtnProceed.Size = New System.Drawing.Size(154, 34)
        Me.BtnProceed.TabIndex = 1
        Me.BtnProceed.Text = "Proceed"
        Me.BtnProceed.UseVisualStyleBackColor = True
        '
        'SN1
        '
        Me.SN1.Enabled = False
        Me.SN1.Location = New System.Drawing.Point(28, 44)
        Me.SN1.Name = "SN1"
        Me.SN1.Size = New System.Drawing.Size(100, 20)
        Me.SN1.TabIndex = 2
        '
        'BtnCloseAdt
        '
        Me.BtnCloseAdt.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnCloseAdt.Location = New System.Drawing.Point(529, 360)
        Me.BtnCloseAdt.Name = "BtnCloseAdt"
        Me.BtnCloseAdt.Size = New System.Drawing.Size(125, 33)
        Me.BtnCloseAdt.TabIndex = 1
        Me.BtnCloseAdt.Text = "CLOSE ADT"
        Me.BtnCloseAdt.UseVisualStyleBackColor = True
        '
        'CH1
        '
        Me.CH1.AutoSize = True
        Me.CH1.Location = New System.Drawing.Point(45, 21)
        Me.CH1.Name = "CH1"
        Me.CH1.Size = New System.Drawing.Size(58, 17)
        Me.CH1.TabIndex = 6
        Me.CH1.Text = "PT/S2"
        Me.CH1.UseVisualStyleBackColor = True
        '
        'CH2
        '
        Me.CH2.AutoSize = True
        Me.CH2.Location = New System.Drawing.Point(45, 77)
        Me.CH2.Name = "CH2"
        Me.CH2.Size = New System.Drawing.Size(58, 17)
        Me.CH2.TabIndex = 7
        Me.CH2.Text = "PT/S2"
        Me.CH2.UseVisualStyleBackColor = True
        '
        'CH3
        '
        Me.CH3.AutoSize = True
        Me.CH3.Location = New System.Drawing.Point(187, 20)
        Me.CH3.Name = "CH3"
        Me.CH3.Size = New System.Drawing.Size(64, 17)
        Me.CH3.TabIndex = 8
        Me.CH3.Text = "PS2/S4"
        Me.CH3.UseVisualStyleBackColor = True
        '
        'CH4
        '
        Me.CH4.AutoSize = True
        Me.CH4.Location = New System.Drawing.Point(187, 77)
        Me.CH4.Name = "CH4"
        Me.CH4.Size = New System.Drawing.Size(64, 17)
        Me.CH4.TabIndex = 9
        Me.CH4.Text = "PS2/S3"
        Me.CH4.UseVisualStyleBackColor = True
        '
        'CH5
        '
        Me.CH5.AutoSize = True
        Me.CH5.Location = New System.Drawing.Point(312, 22)
        Me.CH5.Name = "CH5"
        Me.CH5.Size = New System.Drawing.Size(64, 17)
        Me.CH5.TabIndex = 10
        Me.CH5.Text = "PS1/S1"
        Me.CH5.UseVisualStyleBackColor = True
        '
        'CH6
        '
        Me.CH6.AutoSize = True
        Me.CH6.Location = New System.Drawing.Point(312, 77)
        Me.CH6.Name = "CH6"
        Me.CH6.Size = New System.Drawing.Size(64, 17)
        Me.CH6.TabIndex = 11
        Me.CH6.Text = "PS1/S1"
        Me.CH6.UseVisualStyleBackColor = True
        '
        'ChAll
        '
        Me.ChAll.AutoSize = True
        Me.ChAll.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChAll.Location = New System.Drawing.Point(471, 47)
        Me.ChAll.Name = "ChAll"
        Me.ChAll.Size = New System.Drawing.Size(92, 17)
        Me.ChAll.TabIndex = 12
        Me.ChAll.Text = "ALL TESTS"
        Me.ChAll.UseVisualStyleBackColor = True
        '
        'SN2
        '
        Me.SN2.Enabled = False
        Me.SN2.Location = New System.Drawing.Point(28, 100)
        Me.SN2.Name = "SN2"
        Me.SN2.Size = New System.Drawing.Size(100, 20)
        Me.SN2.TabIndex = 13
        '
        'SN3
        '
        Me.SN3.Enabled = False
        Me.SN3.Location = New System.Drawing.Point(163, 43)
        Me.SN3.Name = "SN3"
        Me.SN3.Size = New System.Drawing.Size(100, 20)
        Me.SN3.TabIndex = 14
        '
        'SN6
        '
        Me.SN6.Enabled = False
        Me.SN6.Location = New System.Drawing.Point(295, 104)
        Me.SN6.Name = "SN6"
        Me.SN6.Size = New System.Drawing.Size(100, 20)
        Me.SN6.TabIndex = 17
        '
        'SN5
        '
        Me.SN5.Enabled = False
        Me.SN5.Location = New System.Drawing.Point(295, 43)
        Me.SN5.Name = "SN5"
        Me.SN5.Size = New System.Drawing.Size(100, 20)
        Me.SN5.TabIndex = 16
        '
        'SN4
        '
        Me.SN4.Enabled = False
        Me.SN4.Location = New System.Drawing.Point(163, 104)
        Me.SN4.Name = "SN4"
        Me.SN4.Size = New System.Drawing.Size(100, 20)
        Me.SN4.TabIndex = 15
        '
        'StFreq
        '
        Me.StFreq.BackColor = System.Drawing.Color.OldLace
        Me.StFreq.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StFreq.Location = New System.Drawing.Point(513, 140)
        Me.StFreq.Name = "StFreq"
        Me.StFreq.Size = New System.Drawing.Size(125, 36)
        Me.StFreq.TabIndex = 18
        Me.StFreq.Text = "FREQ METER"
        Me.StFreq.UseVisualStyleBackColor = False
        '
        'StVolt
        '
        Me.StVolt.BackColor = System.Drawing.Color.OldLace
        Me.StVolt.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StVolt.Location = New System.Drawing.Point(513, 184)
        Me.StVolt.Name = "StVolt"
        Me.StVolt.Size = New System.Drawing.Size(125, 36)
        Me.StVolt.TabIndex = 19
        Me.StVolt.Text = "VOLT METER"
        Me.StVolt.UseVisualStyleBackColor = False
        '
        'ACheck
        '
        Me.ACheck.BackColor = System.Drawing.Color.OldLace
        Me.ACheck.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ACheck.Location = New System.Drawing.Point(513, 226)
        Me.ACheck.Name = "ACheck"
        Me.ACheck.Size = New System.Drawing.Size(125, 32)
        Me.ACheck.TabIndex = 20
        Me.ACheck.Text = "ADT"
        Me.ACheck.UseVisualStyleBackColor = False
        '
        'BtnInstrChk
        '
        Me.BtnInstrChk.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnInstrChk.Location = New System.Drawing.Point(514, 275)
        Me.BtnInstrChk.Name = "BtnInstrChk"
        Me.BtnInstrChk.Size = New System.Drawing.Size(154, 53)
        Me.BtnInstrChk.TabIndex = 21
        Me.BtnInstrChk.Text = "Check Instruments"
        Me.BtnInstrChk.UseVisualStyleBackColor = True
        '
        'TxtPreLevFPath
        '
        '
        '
        '
        Me.TxtPreLevFPath.Border.Class = "TextBoxBorder"
        Me.TxtPreLevFPath.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.TxtPreLevFPath.Location = New System.Drawing.Point(12, 167)
        Me.TxtPreLevFPath.Name = "TxtPreLevFPath"
        Me.TxtPreLevFPath.PreventEnterBeep = True
        Me.TxtPreLevFPath.Size = New System.Drawing.Size(489, 20)
        Me.TxtPreLevFPath.TabIndex = 22
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(422, 191)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 28)
        Me.Button1.TabIndex = 23
        Me.Button1.Text = "Browse"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TxtTmrFormat
        '
        '
        '
        '
        Me.TxtTmrFormat.Border.Class = "TextBoxBorder"
        Me.TxtTmrFormat.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.TxtTmrFormat.Location = New System.Drawing.Point(12, 250)
        Me.TxtTmrFormat.Name = "TxtTmrFormat"
        Me.TxtTmrFormat.PreventEnterBeep = True
        Me.TxtTmrFormat.Size = New System.Drawing.Size(489, 20)
        Me.TxtTmrFormat.TabIndex = 22
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(422, 275)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 29)
        Me.Button2.TabIndex = 23
        Me.Button2.Text = "Browse"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(25, 151)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(117, 13)
        Me.Label1.TabIndex = 24
        Me.Label1.Text = "Pressure Level file path"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(25, 234)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(71, 13)
        Me.Label2.TabIndex = 24
        Me.Label2.Text = "TMR file path"
        '
        'TextBoxISA
        '
        '
        '
        '
        Me.TextBoxISA.Border.Class = "TextBoxBorder"
        Me.TextBoxISA.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.TextBoxISA.Location = New System.Drawing.Point(12, 338)
        Me.TextBoxISA.Name = "TextBoxISA"
        Me.TextBoxISA.PreventEnterBeep = True
        Me.TextBoxISA.Size = New System.Drawing.Size(489, 20)
        Me.TextBoxISA.TabIndex = 25
        '
        'Browse
        '
        Me.Browse.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Browse.Location = New System.Drawing.Point(422, 361)
        Me.Browse.Name = "Browse"
        Me.Browse.Size = New System.Drawing.Size(75, 30)
        Me.Browse.TabIndex = 26
        Me.Browse.Text = "Browse"
        Me.Browse.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(25, 322)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 13)
        Me.Label3.TabIndex = 27
        Me.Label3.Text = "ISA file path"
        '
        'txtFreqAdd
        '
        Me.txtFreqAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFreqAdd.Location = New System.Drawing.Point(645, 145)
        Me.txtFreqAdd.Name = "txtFreqAdd"
        Me.txtFreqAdd.Size = New System.Drawing.Size(32, 23)
        Me.txtFreqAdd.TabIndex = 28
        '
        'txtVoltAdd
        '
        Me.txtVoltAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVoltAdd.Location = New System.Drawing.Point(645, 188)
        Me.txtVoltAdd.Name = "txtVoltAdd"
        Me.txtVoltAdd.Size = New System.Drawing.Size(32, 23)
        Me.txtVoltAdd.TabIndex = 28
        '
        'txtComNo
        '
        Me.txtComNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtComNo.Location = New System.Drawing.Point(645, 228)
        Me.txtComNo.Name = "txtComNo"
        Me.txtComNo.Size = New System.Drawing.Size(32, 23)
        Me.txtComNo.TabIndex = 28
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(568, 12)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(100, 20)
        Me.TextBox1.TabIndex = 29
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(590, 39)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 23)
        Me.Button3.TabIndex = 30
        Me.Button3.Text = "Button3"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(13, 382)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(75, 23)
        Me.Button4.TabIndex = 31
        Me.Button4.Text = "Button4"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(679, 417)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.txtComNo)
        Me.Controls.Add(Me.txtVoltAdd)
        Me.Controls.Add(Me.txtFreqAdd)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Browse)
        Me.Controls.Add(Me.TextBoxISA)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TxtTmrFormat)
        Me.Controls.Add(Me.TxtPreLevFPath)
        Me.Controls.Add(Me.BtnInstrChk)
        Me.Controls.Add(Me.ACheck)
        Me.Controls.Add(Me.StVolt)
        Me.Controls.Add(Me.StFreq)
        Me.Controls.Add(Me.SN6)
        Me.Controls.Add(Me.SN5)
        Me.Controls.Add(Me.SN4)
        Me.Controls.Add(Me.SN3)
        Me.Controls.Add(Me.SN2)
        Me.Controls.Add(Me.ChAll)
        Me.Controls.Add(Me.CH6)
        Me.Controls.Add(Me.CH5)
        Me.Controls.Add(Me.CH4)
        Me.Controls.Add(Me.CH3)
        Me.Controls.Add(Me.CH2)
        Me.Controls.Add(Me.CH1)
        Me.Controls.Add(Me.SN1)
        Me.Controls.Add(Me.BtnCloseAdt)
        Me.Controls.Add(Me.BtnProceed)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Transducer Test"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents BtnProceed As Button
    Friend WithEvents SN1 As TextBox
    Friend WithEvents BtnCloseAdt As Button
    Friend WithEvents CH1 As CheckBox
    Friend WithEvents CH2 As CheckBox
    Friend WithEvents CH3 As CheckBox
    Friend WithEvents CH4 As CheckBox
    Friend WithEvents CH5 As CheckBox
    Friend WithEvents CH6 As CheckBox
    Friend WithEvents ChAll As CheckBox
    Friend WithEvents SN2 As TextBox
    Friend WithEvents SN3 As TextBox
    Friend WithEvents SN6 As TextBox
    Friend WithEvents SN5 As TextBox
    Friend WithEvents SN4 As TextBox
    Friend WithEvents StFreq As Button
    Friend WithEvents StVolt As Button
    Friend WithEvents ACheck As Button
    Friend WithEvents BtnInstrChk As Button
    Friend WithEvents TxtPreLevFPath As DevComponents.DotNetBar.Controls.TextBoxX
    Friend WithEvents OpenFileDialog As OpenFileDialog
    Friend WithEvents Button1 As Button
    Friend WithEvents TxtTmrFormat As DevComponents.DotNetBar.Controls.TextBoxX
    Friend WithEvents Button2 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents TextBoxISA As DevComponents.DotNetBar.Controls.TextBoxX
    Friend WithEvents Browse As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents ISAFolderDialog As FolderBrowserDialog
    Friend WithEvents txtFreqAdd As TextBox
    Friend WithEvents txtVoltAdd As TextBox
    Friend WithEvents txtComNo As TextBox
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
End Class
