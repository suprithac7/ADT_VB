Imports System
Imports System.IO
Imports System.IO.Ports

Public Class Form1

    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles BtnProceed.Click

        Instrcheck()


        If Not ((StFr) And (StVl) And (StAdt)) Then
            MsgBox("Instruments are not connected. Exiting execution.")
            Exit Sub
        End If

        If (CH1.Checked) Or (CH2.Checked) Or (CH3.Checked) Or (CH4.Checked) Or (CH5.Checked) Or (CH6.Checked) Then
        Else
            MsgBox("Please Select atleast one ADT Port")
            Exit Sub
        End If
        Frm2Title = ""

        If CH1.Checked Then
            ExlRow = 15
            If SN1.Text = "" Then
                MsgBox("Please enter the Sl. No. Of transducer for port 1")
                Exit Sub

            End If
            If File.Exists(TextBoxISA.Text & "\C" & SN1.Text & ".ISA") Then
            Else
                MsgBox("file doesn't exists for port 1")
                Exit Sub
            End If
            Frm2Title += CH1.Text + ":" + SN1.Text + ";  "

        End If



        If CH2.Checked Then
            ExlRow = 15
            If SN1.Text = "" Then
                MsgBox("Please enter the Sl. No. Of transducer for port 2")
                Exit Sub

            End If
            If File.Exists(TextBoxISA.Text & "\C" & SN2.Text & ".ISA") Then
            Else
                MsgBox("file doesn't exists for port 2")
                Exit Sub
            End If
            Frm2Title += CH2.Text + ":" + SN2.Text + ";  "
        End If

        If CH3.Checked Then
            ExlRow = 39
            If SN3.Text = "" Then
                MsgBox("Please enter the Sl. No. Of transducer for port 3")
                Exit Sub

            End If
            If File.Exists(TextBoxISA.Text & "\C" & SN3.Text & ".ISA") Then
            Else
                MsgBox("file doesn't exists for port 3")
                Exit Sub
            End If

            Frm2Title += CH3.Text + ":" + SN3.Text + ";  "
        End If


        If CH4.Checked Then
            ExlRow = 27
            If SN4.Text = "" Then
                MsgBox("Please enter the Sl. No. Of transducer for port 4")
                Exit Sub

            End If
            If File.Exists(TextBoxISA.Text & "\C" & SN4.Text & ".ISA") Then
            Else
                MsgBox("file doesn't exists for port 4")
                Exit Sub
            End If

            Frm2Title += CH4.Text + ":" + SN4.Text + ";  "
        End If


        If CH5.Checked Then
            ExlRow = 3
            If SN5.Text = "" Then
                MsgBox("Please enter the Sl. No. Of transducer for port 5")
                Exit Sub

            End If
            If File.Exists(TextBoxISA.Text & "\C" & SN5.Text & ".ISA") Then
            Else
                MsgBox("file doesn't exists for port 5")
                Exit Sub
            End If

            Frm2Title += CH5.Text + ":" + SN5.Text + ";  "
        End If


        If CH6.Checked Then
            ExlRow = 3
            If SN6.Text = "" Then
                MsgBox("Please enter the Sl. No. Of transducer for port 6")
                Exit Sub

            End If
            If File.Exists(TextBoxISA.Text & "\C" & SN6.Text & ".ISA") Then
            Else
                MsgBox("file doesn't exists for port 6")
                Exit Sub
            End If

            Frm2Title += CH6.Text + ":" + SN6.Text + ";  "
        End If


        Sln1 = SN1.Text
        Sln2 = SN2.Text
        Sln3 = SN3.Text
        Sln4 = SN4.Text
        Sln5 = SN5.Text
        Sln6 = SN6.Text

        C1 = CH1.Checked
        C2 = CH2.Checked
        C3 = CH3.Checked
        C4 = CH4.Checked
        C5 = CH5.Checked
        C6 = CH6.Checked

        Form2.Show()
        Form2.Text = Frm2Title

    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles BtnCloseAdt.Click
        If ACheck.BackColor = Color.Red Then
            MsgBox("Cont perform. ADT is not available")
            Exit Sub
        End If
        Form2.ADTclose()
    End Sub



    Private Sub CH_CheckedChanged(sender As Object, e As EventArgs) Handles CH1.CheckedChanged, CH2.CheckedChanged, CH3.CheckedChanged, CH4.CheckedChanged, CH5.CheckedChanged, CH6.CheckedChanged
        SN1.Enabled = CH1.Checked
        SN2.Enabled = CH2.Checked
        SN3.Enabled = CH3.Checked
        SN4.Enabled = CH4.Checked
        SN5.Enabled = CH5.Checked
        SN6.Enabled = CH6.Checked

    End Sub

    Private Sub CHAll_CheckedChanged(sender As Object, e As EventArgs) Handles ChAll.CheckedChanged
        For Each chk As Control In Controls
            If TypeOf chk Is CheckBox Then
                DirectCast(chk, CheckBox).Checked = ChAll.Checked
            End If
        Next

    End Sub

    Private Sub Instrcheck()
        Dim fdata(35) As Char
        Dim fvolt(35) As Char
        Dim AdT As String




        ACheck.BackColor = Color.White
        StFreq.BackColor = Color.White
        StVolt.BackColor = Color.White



        StFr = dev_check(FreqMeter_Addr, My.Settings.FMeterAddr)
        If StFr Then
            StFreq.BackColor = Color.Green
        Else
            StFreq.BackColor = Color.Red

        End If
        StVl = dev_check(VoltMeter_Addr, My.Settings.VMeterAddr)
        If StVl Then
            StVolt.BackColor = Color.Green
        Else
            StVolt.BackColor = Color.Red

        End If

        If Not Form2.SP.IsOpen Then
            Try
                Form2.SP.PortName = "com" & My.Settings.ComAddr.ToString
                Form2.SP.Open()
            Catch ex As Exception
                MsgBox(ex.Message + vbNewLine + "Check RS232 Configuration!")
            End Try
        End If

        Form2.SP.Write("PA" & vbCr)
        Sleep(500)
        AdT = Form2.SP.ReadExisting

        If AdT.Length > 0 Then
            ACheck.BackColor = Color.Green
            StAdt = True
        Else
            ACheck.BackColor = Color.Red
            StAdt = False
        End If


        ' todo:check for relay circuit





    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        MsgBox(Sln1)

    End Sub

    Private Sub StFreq_Click(sender As Object, e As EventArgs) Handles StFreq.Click

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Dim result As DialogResult = OpenFileDialog.ShowDialog()
        If result = DialogResult.OK Then
            Dim path As String = OpenFileDialog.FileName
            TxtPreLevFPath.Text = path
            My.Settings.PLevelFilePath = path
            My.Settings.Save()
        End If

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TxtPreLevFPath.Text = My.Settings.PLevelFilePath
        TxtTmrFormat.Text = My.Settings.TmrFormatPath
        TextBoxISA.Text = My.Settings.ISAFilePath
        txtFreqAdd.Text = My.Settings.FMeterAddr
        txtVoltAdd.Text = My.Settings.VMeterAddr
        txtComNo.Text = My.Settings.ComAddr

        Register_Card(PCI_6208A, 0)
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Dim result As DialogResult = OpenFileDialog.ShowDialog()
        If result = DialogResult.OK Then
            Dim path As String = OpenFileDialog.FileName
            TxtTmrFormat.Text = path
            My.Settings.TmrFormatPath = path
            My.Settings.Save()
        End If

    End Sub

    Private Sub Browse_Click(sender As Object, e As EventArgs) Handles Browse.Click
        Dim result As DialogResult = ISAFolderDialog.ShowDialog()
        If result = DialogResult.OK Then
            Dim path As String = ISAFolderDialog.SelectedPath
            TextBoxISA.Text = path
            My.Settings.ISAFilePath = path
            My.Settings.Save()
        End If
    End Sub

    Private Sub BtnInstrChk_Click(sender As Object, e As EventArgs) Handles BtnInstrChk.Click
        If Val(txtFreqAdd.Text) > 0 And Val(txtFreqAdd.Text) < 31 Then
            My.Settings.FMeterAddr = Val(txtFreqAdd.Text)
        Else
            MsgBox(" Invalid GPIB adress")
            Exit Sub
        End If

        If Val(txtVoltAdd.Text) > 0 And Val(txtVoltAdd.Text) < 31 Then
            My.Settings.VMeterAddr = Val(txtVoltAdd.Text)

        Else
            MsgBox(" Invalid GPIB adress")
            Exit Sub
        End If
        My.Settings.ComAddr = Val(txtComNo.Text)

        My.Settings.Save()
            Instrcheck()
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        '  MsgBox(Val(TextBox1.Text))
        SwitchTo(Val(TextBox1.Text))
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Process.Start("EXCEL.EXE", curfile)
    End Sub
End Class
