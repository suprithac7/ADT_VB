
Imports System
Imports System.IO
Imports System.IO.Ports
Imports excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office



Public Class Form2
    Dim strm As IO.Stream
    Public ISAExists As Boolean
    Dim line As String
    Dim trimchars() As Char = ""
    Dim tmp As String
    Dim MainThread As Threading.Thread = New Threading.Thread(AddressOf processing)

    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)

    Private Sub BtnStart_Click(sender As Object, e As EventArgs) Handles BtnStart.Click


        CheckForIllegalCrossThreadCalls = 0
        MainThread = New Threading.Thread(AddressOf processing)
        MainThread.Start()
        '   BtnSave.Enabled = 1
    End Sub

    Public Sub Processing()
        TxtLog.AppendText("Started @   " & Date.Now.ToString("hh_mm_ss"))
        BtnAbort.Enabled = 1


        ' todo: loop following code  for all six channel (only selecte)

        'Dim x As Single = MeasVpp()
        'If (x < 2) Or (x > 10) Then        ' TODO: update spec for correct vpp
        '    MsgBox("Vp-p is out of range(2-10v). Exiting Measurement")
        '    Exit Sub
        'End If


        ADTinit()
        TxtLog.AppendText(vbNewLine)
        For i = 9 To 0 Step -1  '' TOdo: CHANGE STEP

            setADT(Dgv1.Rows(i).Cells(1).Value, Dgv3.Rows(i).Cells(1).Value, Dgv5.Rows(i).Cells(1).Value)

            ' todo: switch port 1
            If C1 Then
                SwitchTo(1)
                MeasnFillRow(Dgv1, i)
            End If

            ' todo: switch port 2
            If C2 Then
                SwitchTo(2)
                MeasnFillRow(Dgv2, i)
            End If

            ' todo: switch port 3
            If C3 Then
                SwitchTo(3)
                MeasnFillRow(Dgv3, i)
            End If

            ' todo: switch port 4
            If C4 Then
                SwitchTo(4)
                MeasnFillRow(Dgv4, i)
            End If

            ' todo: switch port 5
            If C5 Then
                SwitchTo(5)
                MeasnFillRow(Dgv5, i)
            End If

            ' todo: switch port 6
            If C6 Then
                SwitchTo(6)
                MeasnFillRow(Dgv6, i)
            End If

            TxtLog.AppendText(vbNewLine)

        Next


        ExlClose()
        SP.Close()
        BtnAbort.Enabled = 0
        TxtLog.AppendText("Stopped @   " & Date.Now.ToString("hh_mm_ss"))
    End Sub

    Public Sub MeasnFillRow(ByVal DGV As DataGridView, ByVal Row As Integer)
        Dim recheck As Boolean = 1

rechkLbl: DGV.Rows(Row).Cells(3).Value = MeasFreq.ToString
        DGV.Rows(Row).Cells(4).Value = MeasVolt.ToString


        DGV.Rows(Row).Cells(5).Value = ((DGV.Rows(Row).Cells(2).Value)) - ((DGV.Rows(Row).Cells(3).Value))
        If (Math.Abs(DGV.Rows(Row).Cells(5).Value) < 1) And ((Val(DGV.Rows(Row).Cells(4).Value) > 0.4) And (Val(DGV.Rows(Row).Cells(4).Value) < 0.6)) Then
            DGV.Rows(Row).Cells(6).Value = "Pass"
        Else
            DGV.Rows(Row).Cells(6).Value = "Fail"
            DGV.Rows(Row).DefaultCellStyle.BackColor = Color.LightCoral

            If recheck Then
                TxtLog.AppendText("  >> Rechecking for Current Setting   ")
                recheck = 0
                GoTo rechkLbl
            End If

        End If



    End Sub



    Public Function FillDgv(ByVal SerialNo As String, ByVal ExlRow As Integer, ByVal DGV As DataGridView)

        If Not File.Exists(Form1.TextBoxISA.Text & "\C" & SerialNo & ".ISA") Then
            MsgBox("file Doesn't Exists for  " & "Sl No." & SerialNo & "." & " Exiting Procedure")
            ISAExists = 1

            Exit Function
        End If

        strm = IO.File.OpenRead(Form1.TextBoxISA.Text & "\C" & SerialNo & ".ISA")
        Dim sr As New IO.StreamReader(strm)


        Do While sr.Peek <> -1
            line = sr.ReadLine()
            If line.TrimStart(trimchars).StartsWith("Temperature 25") Then
                TxtLog.AppendText(" >>" & DGV.Name & " Specs block found    ")
                Exit Do
            End If
        Loop

        For i = ExlRow To ExlRow + 9
            DGV.Rows.Add(xlApp.Range("b" & i).Value, " ", " ", " ", " ", " ", " ")
        Next

        For i = 0 To DGV.RowCount - 1
            tmp = Math.Round(DGV.Rows(i).Cells(0).Value)
            Do While sr.Peek <> -1

                line = sr.ReadLine()



                If i > 0 Then
                    If (Math.Round(DGV.Rows(i - 1).Cells(0).Value) = tmp) Then
                        DGV.Rows(i).Cells(2).Value = DGV.Rows(i - 1).Cells(2).Value
                        DGV.Rows(i).Cells(1).Value = DGV.Rows(i - 1).Cells(1).Value
                        GoTo extLoop

                    End If
                End If

                If (line.TrimStart(trimchars).StartsWith(tmp)) Or (line.TrimStart(trimchars).StartsWith(Val(tmp) - 1).ToString) Or (line.TrimStart(trimchars).StartsWith(Val(tmp) + 1).ToString) Then
                    line = line.Substring(3)
                    DGV.Rows(i).Cells(1).Value = line.Remove(line.IndexOf(vbTab))
                    line = line.Substring(line.IndexOf(vbTab) + 1)
                    line = line.Remove(line.IndexOf(vbTab))
                    DGV.Rows(i).Cells(2).Value = line
                    Exit Do
                ElseIf line.TrimStart(trimchars).StartsWith("Temperature") Then
                    DGV.Rows(i).Cells(2).Value = "0"
                    TxtLog.AppendText("  >>" & DGV.Name & " InComplete specs Not found")
                    Exit Do
                End If


extLoop:    Loop

        Next

        sr.Close()
    End Function




    Public Sub ADTclose()
        ADTinit()

        SP.Write("DT=900" & vbCr)
        SP.Write("ST=900" & vbCr)
        SP.Write("PT=900" & vbCr)

        SP.Write("GO" & vbCr)


RECHECK: SP.Write("PA" & vbCr)
        Dim A, B, C As String
        A = SP.ReadExisting
        SP.Write("SA" & vbCr)
        B = SP.ReadExisting
        SP.Write("DA" & vbCr)
        C = SP.ReadExisting


        If (Math.Abs(900 - Val(A)) < 1) Or (Math.Abs(900 - Val(B)) < 1) Or (Math.Abs(900 - Val(C)) < 1) Then
            SP.Write("PM=1" & vbCr)
            SP.Write("SM=1" & vbCr)
            SP.Write("DM=1" & vbCr)
            MsgBox("ADT can be switched off")
            Beep()
            Sleep(500)
            Beep()

        Else
            Sleep(5000)
            GoTo RECHECK
        End If

    End Sub


    Public Sub ADTinit()
        If Not SP.IsOpen Then
            Try
                SP.Open()

                TxtLog.AppendText("  >> Serial port open." + " on " + SP.PortName + " @ " + Str(SP.BaudRate))
            Catch ex As Exception
                MsgBox(ex.Message + vbNewLine + "Check RS232 Configuration!")

            End Try

        End If

        SP.Write("SG" & vbCr)
        SP.Write("GG" & vbCr)

        SP.Write("PM=3" & vbCr)
        SP.Write("SM=3" & vbCr)
        SP.Write("DM=3" & vbCr)
    End Sub


    Public Sub setADT(ByVal SetPt As String, ByVal SetS2 As String, ByVal SetS1 As String)

        Dim PtA As String
        Dim S2A As String
        Dim S1A As String
        Dim Count As Integer = 0

recheck:
        SP.ReadExisting()

        SP.Write("PT=" & SetPt & vbCr)
        Sleep(500)
        SP.Write("DT=" & SetS2 & vbCr)
        Sleep(500)
        SP.Write("ST=" & SetS1 & vbCr)
        Sleep(500)
        SP.Write("GO" & vbCr)
        Sleep(500)

        SP.Write("PA" & vbCr)
        Sleep(500)
        PtA = SP.ReadExisting
        TxtLog.AppendText(">>" & PtA & vbTab)
        SP.Write("DA" & vbCr)
        Sleep(500)
        S2A = SP.ReadExisting
        TxtLog.AppendText(">>" & S2A & vbTab)
        SP.Write("SA" & vbCr)
        Sleep(500)
        S1A = SP.ReadExisting
        TxtLog.AppendText(">>" & S1A & vbTab)

        Dim TAP, TAD, TAS As Boolean

        TAP = IIf((Val(SetPt) = 0), 1, Math.Abs(Val(SetPt) - Val(PtA)) < 1)
        TAD = IIf(Val(SetS2) = 0, 1, (Math.Abs(Val(SetS2) - Val(S2A)) < 1))
        TAS = IIf(Val(SetS1) = 0, 1, (Math.Abs(Val(SetS1) - Val(S1A)) < 1))


        If TAP And TAS And TAD Then
            TxtLog.AppendText("  >> Target Pressure Achieved")
        Else
            TxtLog.AppendText("  >> Achieving Press.")
            Sleep(2500)
            Count = Count + 1
            If Count < 100 Then
                GoTo recheck
            Else
                MainThread.Abort()
            End If
        End If

    End Sub
    Public Function MeasFreq() As Single
        ibwrt(FreqMeter_Addr, "*rst")
        Sleep(1000)
        ibwrt(FreqMeter_Addr, "MEAS:freq?")
        Sleep(1000)
        ibrd(FreqMeter_Addr, MeasuredArr)
        TxtLog.AppendText(" Freq: " & Val(MeasuredArr).ToString)
        Return (Val(MeasuredArr))
    End Function
    Public Function MeasVolt() As Single
        ibwrt(VoltMeter_Addr, "*rst")
        Sleep(1000)
        ibwrt(VoltMeter_Addr, "MEAS:volt?")
        Sleep(1000)
        ibrd(VoltMeter_Addr, MeasuredArr)
        TxtLog.AppendText(" volt: " & Val(MeasuredArr).ToString)
        Return (Val(MeasuredArr))
    End Function

    Private Sub BtnFillDgv_Click(sender As Object, e As EventArgs) Handles BtnFillDgv.Click
        ISAExists = 0
        Dgv1.Rows.Clear()
        Dgv2.Rows.Clear()
        Dgv3.Rows.Clear()
        Dgv4.Rows.Clear()
        Dgv5.Rows.Clear()
        Dgv6.Rows.Clear()

        If File.Exists(Form1.TxtPreLevFPath.Text) Then

            Try
                xlApp = New Microsoft.Office.Interop.Excel.Application
                xlWorkBook = xlApp.Workbooks.Open(Form1.TxtPreLevFPath.Text)
                xlWorkSheet = xlWorkBook.Sheets("sheet1")
            Catch ex As Exception
                MsgBox("cant find cpressure level sheets.")
                Exit Sub
            End Try
        Else
            MsgBox("File Doesn't Exists. Exiting Procedure")
            Exit Sub
        End If

        If C1 Then FillDgv(Sln1, 15, Dgv1)
        If C2 Then FillDgv(Sln2, 15, Dgv2)
        If C3 Then FillDgv(Sln3, 39, Dgv3)
        If C4 Then FillDgv(Sln4, 27, Dgv4)
        If C5 Then FillDgv(Sln5, 3, Dgv5)
        If C6 Then FillDgv(Sln6, 3, Dgv6)

        If ISAExists Then
            Exit Sub
        End If

        Dim Sdgv As DataGridView = SourceDGVchk()

        BlankDgvCheck(Dgv1, Sdgv)
        BlankDgvCheck(Dgv2, Sdgv)
        BlankDgvCheck(Dgv3, Sdgv)
        BlankDgvCheck(Dgv4, Sdgv)
        BlankDgvCheck(Dgv5, Sdgv)
        BlankDgvCheck(Dgv6, Sdgv)

        ExlClose()

        BtnSave.Enabled = 1
        BtnStart.Enabled = 1

    End Sub

    Public Function SourceDGVchk() As DataGridView

        If C1 Then Return Dgv1
        If C2 Then Return Dgv2
        If C3 Then Return Dgv3
        If C4 Then Return Dgv4
        If C5 Then Return Dgv5
        If C6 Then Return Dgv6


    End Function

    Public Sub BlankDgvCheck(ByVal DGV As DataGridView, ByVal DGVsource As DataGridView)

        If DGV.Rows.Count = 0 Then
            For i = 0 To 9
                DGV.Rows.Add(DGVsource.Rows(i).Cells(0).Value, DGVsource.Rows(i).Cells(1).Value, "", "", "", "", "")
            Next
        End If

    End Sub

    Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles BtnSave.Click
        xlApp = New Microsoft.Office.Interop.Excel.Application

        If File.Exists(Form1.TxtTmrFormat.Text) Then
            Try
                curfile = System.IO.Path.GetDirectoryName(Form1.TxtTmrFormat.Text) & "\Reports\" & Date.Now.ToString("dd_MM_yy-hh_mm_ss") & ".xlsx"
                My.Computer.FileSystem.CopyFile(Form1.TxtTmrFormat.Text, curfile)
                xlWorkBook = xlApp.Workbooks.Open(curfile)    'keep in try block
            Catch ex As Exception
                MsgBox("Error Occured while opening the path...   " & ex.Message)

                Exit Sub
            End Try
        Else
            MsgBox("File Doesn't Exists. Exiting the procedure")
            Exit Sub
        End If

        xlWorkBook.Application.DisplayAlerts = False

        If C1 Then SaveDgv(Sln1, "Sheet1", Dgv1, "4506 105 803 97", "R200-075B-2743") Else xlWorkBook.Sheets("Sheet1").delete
        If C2 Then SaveDgv(Sln2, "Sheet2", Dgv2, "4506 105 803 97", "R200-075B-2743") Else xlWorkBook.Sheets("Sheet2").delete
        If C3 Then SaveDgv(Sln3, "Sheet3", Dgv3, "4506 105 806 88", "R200-075B-2-5377") Else xlWorkBook.Sheets("Sheet3").delete
        If C4 Then SaveDgv(Sln4, "Sheet4", Dgv4, "4506 105 805 91", "R200-075B-1-5377") Else xlWorkBook.Sheets("Sheet4").delete
        If C5 Then SaveDgv(Sln5, "Sheet5", Dgv5, "4506 105 804 94", "R200-06B-5350") Else xlWorkBook.Sheets("Sheet5").delete
        If C6 Then SaveDgv(Sln6, "Sheet6", Dgv6, "4506 105 804 94", "R200-06B-5350") Else xlWorkBook.Sheets("Sheet6").delete

        xlWorkBook.Application.DisplayAlerts = True
        xlWorkBook.Save()
        xlWorkBook.ExportAsFixedFormat(0, curfile.Remove(curfile.Length - 5) & IIf(C1, " " & Sln1, "") & IIf(C2, " " & Sln2, "") & IIf(C3, " " & Sln3, "") & IIf(C4, " " & Sln4, "") & IIf(C5, Sln5 & " ", "") & IIf(C6, " " & Sln6, "") & ".pdf")

        BtnSave.Enabled = 0
        ExlClose()

        If MsgBox("Report Saved Successfully. Do you want to see Report ? ", vbYesNo) = vbYes Then
            Process.Start("EXCEL.EXE", curfile)
        End If




    End Sub
    Private sub SaveDgv(ByVal SerialNo As String, ByVal SheetName As String, ByVal DGV As DataGridView, ByVal PartNo As String, ByVal RefNo As String)

        xlWorkSheet = xlWorkBook.Worksheets(SheetName)
        xlWorkSheet.Activate()

        xlApp.Range("f10").Value = Date.Now.ToString("dd-MM-yyyy")
        xlApp.Range("b10").Value = SerialNo
        xlApp.Range("e4").Value = PartNo
        xlApp.Range("b9").Value = RefNo


        For i = 0 To 9
            xlApp.Range("b" & (i + 19).ToString).Value = DGV.Rows(i).Cells(1).Value
            xlApp.Range("c" & (i + 19).ToString).Value = DGV.Rows(i).Cells(2).Value
            xlApp.Range("d" & (i + 19).ToString).Value = DGV.Rows(i).Cells(3).Value
            xlApp.Range("e" & (i + 19).ToString).Value = DGV.Rows(i).Cells(4).Value
            xlApp.Range("f" & (i + 19).ToString).Value = DGV.Rows(i).Cells(6).Value
        Next
    End sub


    Private Sub Form2_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed, Me.Closing, Me.FormClosed, Me.FormClosed, BtnAbort.Click
        If Not (IsNothing(MainThread)) Then
            If MainThread.IsAlive Then
                MainThread.Abort()
            End If
        End If
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class