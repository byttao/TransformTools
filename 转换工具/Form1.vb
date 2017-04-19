Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel
Imports System.IO
Imports System
Imports System.Management

Public Class Form1
    Dim DG1, DG2 As IWorkbook
    Dim O As Boolean
    Dim User As String
    Private Sub Writein(ByRef DG1 As IWorkbook, ByRef DG2 As IWorkbook, Ssh1 As String, Ssh2 As String, Srng1 As String, Srng2 As String)
        Dim SS1 As ISheet = DG1.GetSheet(Ssh1)
        Dim SS2 As ISheet = DG2.GetSheet(Ssh2)
        Dim SR1, SR2 As IRow
        Dim SC1, SC2 As ICell
        Dim a() As String
        Dim i, j, R1, R2, C1, C2, Rnum, Cnum As Integer
        '
        If InStr(Srng1, ":") > 0 Then
            a = Split(Srng1, ":")
            C1 = Asc(UCase(Mid(a(0), 2, 1))) - 65
            Cnum = Asc(UCase(Mid(a(1), 2, 1))) - 64 - C1
            R1 = Val(Mid(a(0), 4, Len(a(0)))) - 1
            Rnum = Val(Mid(a(1), 4, Len(a(1)))) - R1
        Else
            C1 = Asc(UCase(Mid(Srng1, 2, 1))) - 65
            Cnum = 1
            R1 = Val(Mid(Srng1, 4, Len(Srng1))) - 1
            Rnum = 1
        End If

        If InStr(Srng2, ":") > 0 Then
            a = Split(Srng2, ":")
            C2 = Asc(UCase(Mid(a(0), 2, 1))) - 65
            R2 = Val(Mid(a(0), 4, Len(a(0)))) - 1
            If Cnum <> Asc(UCase(Mid(a(1), 2, 1))) - 64 - C2 Or Rnum <> Val(Mid(a(1), 4, Len(a(1)))) - R2 Then
                MsgBox("错误对应：表1：" & Ssh1 & "，区域1：" & Srng1 & "；表2：" & Ssh2 & "，区域2：" & Srng2)
                O = False
                Exit Sub
            End If
        Else
            C2 = Asc(UCase(Mid(Srng2, 2, 1))) - 65
            R2 = Val(Mid(Srng2, 4, Len(Srng2))) - 1
            If Cnum <> 1 Or Rnum <> 1 Then
                MsgBox("错误对应：表1：" & Ssh1 & "，区域1：" & Srng1 & "；表2：" & Ssh2 & "，区域2：" & Srng2)
                O = False
                Exit Sub
            End If
        End If

        '
        For i = 0 To Rnum - 1
            SR1 = SS1.GetRow(R1 + i)
            SR2 = SS2.GetRow(R2 + i)
            If SR1 Is Nothing Then Exit For
            For j = 0 To Cnum - 1
                SC1 = SR1.GetCell(C1 + j, MissingCellPolicy.CREATE_NULL_AS_BLANK)
                SC2 = SR2.GetCell(C2 + j, MissingCellPolicy.CREATE_NULL_AS_BLANK)
                Select Case SC1.CellType
                    Case CellType.Numeric
                        SC2.SetCellValue(Math.Round(SC1.NumericCellValue, 2))
                    Case CellType.String
                        SC2.SetCellValue(SC1.StringCellValue)
                    Case CellType.Formula
                        'If Not IsError(SC1.ErrorCellValue) Then
                        Try
                            SC2.SetCellValue(Math.Round(SC1.NumericCellValue, 2))
                        Catch ex As Exception
                            Try
                                SC2.SetCellValue(SC1.StringCellValue)
                            Catch
                                SC2.SetCellValue("")
                            End Try
                        End Try
                        'End If
                    Case Else
                        SC2.SetCellValue("")
                End Select
            Next
        Next
    End Sub
    Private Sub 特殊(ByRef DG1 As IWorkbook, ByRef DG2 As IWorkbook, Ssh1 As String, Ssh2 As String, Srng1 As String, Srng2 As String)

    End Sub
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Dim Strinfo, a(), b(), c() As String
        Dim i, j, k, m, SSS As Integer
        'Dim R1, C1, Rnum, Cnum As Integer
        Dim SS1, SS2 As ISheet
        Dim SR1, SR2 As IRow
        Dim SC1, SC2 As ICell
        Dim CC As ICell
        Dim eval As XSSFFormulaEvaluator

        'Dim Fs1 As FileStream = File.Open(TextBox1.Text, FileMode.Open, FileAccess.ReadWrite)
        'DG1 = WorkbookFactory.Create(Fs1)
        'SS2 = DG1.GetSheet("Sheet3")
        'SR1 = SS2.GetRow(0)
        'If SR1 Then
        '    SC1 = SR1.GetCell(2, MissingCellPolicy.CREATE_NULL_AS_BLANK)
        '    SC1.SetCellFormula("Sheet2!C12()")
        '    DG1.Write(Fs1)

        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("文件未选择。")
            Exit Sub
        End If
        If MsgBox("是否进行转化？", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            O = True
            Dim Fs1 As FileStream = File.Open(TextBox1.Text, FileMode.Open, FileAccess.ReadWrite)
            Dim Fs2 As FileStream = File.Open(TextBox2.Text, FileMode.Open, FileAccess.ReadWrite)
            'Label6.Text = "正在打开 底稿文件..."
            'ProgressBar1.Value = 10
            'Using dia As New Dialog1("正在打开 底稿文件...")
            'dia.StartPosition = FormStartPosition.CenterParent
            'dia.Show()
            DG1 = WorkbookFactory.Create(Fs1)
            'dia.Close()
            'End Using
            'Label6.Text = "正在打开 报告文件..."
            'Using dia As New Dialog1("正在打开 报告文件...")
            '    dia.StartPosition = FormStartPosition.CenterParent
            '    dia.Show()
            DG2 = WorkbookFactory.Create(Fs2)
            'dia.Close()
            'End Using


            '读取配置文件
            Dim iRead As New System.IO.StreamReader(User, System.Text.Encoding.GetEncoding("GB2312"))
            If iRead IsNot Nothing Then

                'Using dia As New Dialog1("正在生成 报告文件...")
                '    dia.StartPosition = FormStartPosition.CenterParent
                '    dia.Show()
                Strinfo = iRead.ReadLine
                m = 0
                'Label6.Text = "正在生成 报告文件..."
                a = Strinfo.Split(" ")
                If a(0) = "事项说明" Then

                    SS2 = DG1.GetSheet("(二)附表-纳税调整额的审核")
                    m = 6

                    eval = New XSSFFormulaEvaluator(DG1)
                    For j = 6 To SS2.LastRowNum
                        SS2.GetRow(j).Cells.Clear()
                    Next
                    For i = 1 To Val(a(1))
                        Strinfo = iRead.ReadLine
                        b = Strinfo.Split(" ")
                        SR2 = SS2.GetRow(m)
                        SC2 = SR2.GetCell(3)


                        SC2.SetCellFormula(b(3))
                        SC2.SetCellValue(eval.Evaluate(SC2).NumberValue)
                        If SC2.NumericCellValue <> 0 Then
                            SC2 = SR2.GetCell(0)

                            SC2.SetCellValue(b(0))
                            SC2 = SR2.GetCell(1)
                            SC2.SetCellFormula(b(1))
                            If IsError(eval.Evaluate(SC2).NumberValue) Then
                                SC2.SetCellType(CellType.Numeric)
                                SC2.SetCellValue(0)
                            Else
                                SC2.SetCellValue(eval.Evaluate(SC2).NumberValue)
                            End If

                            SC2 = SR2.GetCell(2)
                            SC2.SetCellFormula(b(2))
                            If IsError(eval.Evaluate(SC2).NumberValue) Then
                                SC2.SetCellType(CellType.Numeric)
                                SC2.SetCellValue(0)
                            Else
                                SC2.SetCellValue(eval.Evaluate(SC2).NumberValue)
                            End If
                            SC2 = SR2.GetCell(4)
                            SC2.SetCellValue("税法规定")
                            m = m + 1
                        Else
                            SC2.SetCellValue("")
                        End If
                        '    k = Len(a(1)) - Len(Replace(a(1), ",", ""))
                        '    b = a(1).Split(",")

                        '    If InStr(b(0), ":") > 0 Then
                        '        c = Split(b(0), ":")
                        '        C1 = Asc(UCase(Mid(c(0), 2, 1))) - 65
                        '        R1 = Val(Mid(c(0), 4, Len(c(0)))) - 1
                        '        Rnum = Val(Mid(c(1), 4, Len(c(1)))) - R1
                        '    Else
                        '        C1 = Asc(UCase(Mid(b(0), 2, 1))) - 65
                        '        R1 = Val(Mid(b(0), 4, Len(b(0)))) - 1
                        '        Rnum = 1
                        '    End If
                        '    ReDim c(b.Length - 2)
                        '    For k = 0 To b.Length - 2
                        '        c(k) = Asc(UCase(Mid(b(0), 2, 1))) - 65
                        '    Next
                        '    SS1 = DG1.GetSheet(a(0))
                        '    SS2 = DG1.GetSheet("(二)附表-纳税调整额的审核")
                        '    For j = 6 To SS1.LastRowNum
                        '        SS1.GetRow(j).Cells.Clear()
                        '    Next
                        '    m = 6
                        '    SR2 = SS2.GetRow(m)
                        '    For j = 0 To Rnum
                        '        SR1 = SS1.GetRow(j + R1)
                        '        SC1 = SR1.GetCell(C1, MissingCellPolicy.CREATE_NULL_AS_BLANK)
                        '        If SC1.NumericCellValue <> 0 Then
                        '            SC2 = SR2.GetCell(3, MissingCellPolicy.CREATE_NULL_AS_BLANK)
                        '            SC2.SetCellValue(SC1.NumericCellValue)

                        '            SC1 = SR1.GetCell(c(1))
                        '            SC2 = SR2.GetCell(0, MissingCellPolicy.CREATE_NULL_AS_BLANK)
                        '            SC2.SetCellValue(SC1.StringCellValue)
                        '            SC1 = SR1.GetCell(c(2))
                        '            SC2 = SR2.GetCell(1, MissingCellPolicy.CREATE_NULL_AS_BLANK)
                        '            SC2.SetCellValue(SC1.NumericCellValue)
                        '            SC1 = SR1.GetCell(c(3))
                        '            SC2 = SR2.GetCell(2, MissingCellPolicy.CREATE_NULL_AS_BLANK)
                        '            SC2.SetCellValue(SC1.NumericCellValue)
                        '            SC1 = SR1.GetCell(c(4))
                        '            SC2 = SR2.GetCell(4, MissingCellPolicy.CREATE_NULL_AS_BLANK)
                        '            SC2.SetCellValue(SC1.StringCellValue)
                        '        End If
                        '    Next
                    Next

                    'Fs1 = File.Open(TextBox1.Text, FileMode.Open, FileAccess.ReadWrite)
                    'DG1.Write(Fs1)
                    'Fs1.Close()
                    Strinfo = iRead.ReadLine
                    a = Strinfo.Split(" ")
                End If

                If a(0) = "普通转换" Then
                    m = 0
                    SSS = Val(a(1))
                    Do Until iRead.EndOfStream
                        Strinfo = iRead.ReadLine
                        a = Strinfo.Split(" ")
                        k = Len(a(1)) - Len(Replace(a(1), ",", ""))
                        If k = Len(a(3)) - Len(Replace(a(3), ",", "")) Then
                            b = a(1).Split(",")
                            c = a(3).Split(",")
                            m = m + 1
                            If a(0) = "(二)附表-纳税调整额的审核" Then
                                SS2 = DG2.GetSheet("(二)附表-纳税调整额的审核")
                                For j = 6 To SS2.LastRowNum
                                    SS2.GetRow(j).Cells.Clear()
                                Next
                            End If
                            For i = 0 To k
                                'MsgBox(b(i))
                                Writein(DG1, DG2, a(0), a(2), b(i), c(i))
                                If Not O Then Exit Sub
                            Next
                            ProgressBar1.Value = Int(m / SSS * 100)
                        Else
                            MsgBox("错误对应：表1：" & a(0) & "，区域1：" & a(1) & "；表2：" & a(2) & "，区域2：" & a(3))
                            Exit Sub
                        End If
                    Loop
                End If

                Fs2 = File.Open(TextBox2.Text, FileMode.Open, FileAccess.ReadWrite)
                DG2.Write(Fs2)
                Fs2.Close()

                'dia.Close()
                'End Using
                'Label6.Text = "生成完成..."
                MsgBox("生成完成...")
            End If
        End If
    End Sub

    Private Sub PictureBox2_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox2.MouseEnter
        PictureBox2.Image = My.Resources._3

    End Sub

    Private Sub PictureBox2_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox2.MouseLeave
        PictureBox2.Image = My.Resources._2
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        With OpenFileDialog1
            .InitialDirectory = System.Environment.SpecialFolder.Desktop
            .Filter = "Excel Files (*.xls*)|*.xls*|All files (*.*)|*.*"
            .CheckFileExists = True
            .FileName = ""
            .ShowDialog()
            TextBox1.Text = OpenFileDialog1.FileName
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        With OpenFileDialog1
            .InitialDirectory = System.Environment.SpecialFolder.Desktop
            .Filter = "Excel Files (*.xls*)|*.xls*|All files (*.*)|*.*"
            .CheckFileExists = True
            .FileName = ""
            .ShowDialog()
            TextBox2.Text = OpenFileDialog1.FileName
        End With
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Label6.Parent = ProgressBar1
        User = Application.StartupPath & "\默认配置.ini"
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        With OpenFileDialog1
            .InitialDirectory = System.Environment.SpecialFolder.Desktop
            .Filter = "Excel Files (*.ini)|*.ini|All files (*.*)|*.*"
            .CheckFileExists = True
            .FileName = ""
            .ShowDialog()
            User = .FileName
            Label6.Text = "当前加载配置文件为 " & .SafeFileName
        End With
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        MsgBox("CPU代码：" & GetCUPID() & "；硬盘代码：" & GetHardDriveID())
    End Sub

    Private Function GetHardDriveID() As String
        Try
            GetHardDriveID = ""
            Dim info As ManagementBaseObject
            Dim query As New SelectQuery("Win32_DiskDrive")
            Dim search As New ManagementObjectSearcher(query)
            For Each info In search.Get()
                If info("Model") IsNot Nothing Then
                    Return info("Model").ToString
                Else
                    Return ""
                End If
            Next
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private Function GetCUPID() As String
        Try
            GetCUPID = ""
            Dim info As ManagementBaseObject
            Dim query As New SelectQuery("Win32_Processor")
            Dim search As New ManagementObjectSearcher(query)
            For Each info In search.Get
                If info("ProcessorId") IsNot Nothing Then
                    Return info("ProcessorId").ToString
                Else
                    Return ""
                End If
            Next
        Catch ex As Exception
            Return ""
        End Try
    End Function
End Class
