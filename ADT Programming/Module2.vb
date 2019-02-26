Module Module2

    Public xlApp As Microsoft.Office.Interop.Excel.Application
    Public xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Public xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
    Public curfile As String

    Public FreqMeter_Addr As Single = ildev(0, My.Settings.FMeterAddr, 0, T10s, 1, 0)
    Public VoltMeter_Addr As Single = ildev(0, My.Settings.VMeterAddr, 0, T10s, 1, 0)

    Public MeasuredArr(25) As Char

    Public ExlRow As Integer
    Public Frm2Title As String
    Public Sln1, Sln2, Sln3, Sln4, Sln5, Sln6 As String
    Public C1, C2, C3, C4, C5, C6 As Boolean

    Public StFr As Boolean = False
    Public StAdt As Boolean = False
    Public StVl As Boolean = False

    Public instr_desc(100) As Char
    Public instr_temp As String

    Public Function MeasVpp() As Single
        ibwrt(FreqMeter_Addr, "*rst")
        ibwrt(FreqMeter_Addr, "MEAS:VOLT:AC?")
        Sleep(1000)
        ibrd(FreqMeter_Addr, MeasuredArr)
        Return Math.Round(2 * Math.Sqrt(2) * Val(MeasuredArr), 2)
    End Function


    Public Function dev_check(ByVal addr As Integer, ByVal pad As Integer) As Boolean
        Dim rvalue As Integer
        Call ibln(addr, pad, 0, rvalue)
        If rvalue <> 0 Then
            Call corr_dev_check(pad)
            instr_temp = Left(instr_desc, 44) '36 for FSU
            If instr_temp.Length > 10 Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    Sub corr_dev_check(ByRef addr1 As Integer)
        Dim str As String
        str = "*IDN?"
        Call ibwrt(addr1, str)
        Call ibrd(addr1, instr_desc)
    End Sub

    Public Sub SwitchTo(ByVal relay As Integer)


        If relay < 1 Or relay > 6 Then
            MsgBox("pls Switch to valid relay")
            Exit Sub
        End If
        For i = 0 To 5
            AO_VWriteChannel(0, i, 0.0)
        Next
        ' MsgBox((relay - 1).ToString)
        AO_VWriteChannel(0, (relay - 1), 10.0)

    End Sub


    Public Sub ExlClose()
        '    xlWorkBook.Save()
        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

End Module
