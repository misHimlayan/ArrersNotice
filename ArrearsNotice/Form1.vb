Imports System
Imports System.Data
Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class Form1
    Dim strmessage, DExcelConnPath, sqlQuery As String
    Dim ofdBrowse As New OpenFileDialog
    Dim errProvider As New ErrorProvider
    Dim xclDtable, xclColumnsDTable, xclDistinct As DataTable
    Dim olecon As OleDbConnection
    Dim oleAdpter As OleDbDataAdapter
    Dim oxclApp1, oxclApp2, oxclApp3, oxclApp4, oxclAppSummary As New Excel.Application
    Dim oxclSourceWorkbook As Excel.Workbook
    Dim intWorkSheetCount As Integer
    Dim xlSourceRange As Excel.Range = Nothing
    Dim oxclWorkbook1 As Excel.Workbook = Nothing
    Dim oxclWorkbook2 As Excel.Workbook = Nothing
    Dim oxclWorkbook3 As Excel.Workbook = Nothing
    Dim oxclWorkbook4 As Excel.Workbook = Nothing
    Dim oxclWorkbookSummary As Excel.Workbook = Nothing
    Dim oxclWorkSheet1 As Excel.Worksheet = Nothing
    Dim oxclWorkSheet2 As Excel.Worksheet = Nothing
    Dim oxclWorkSheet3 As Excel.Worksheet = Nothing
    Dim oxclWorkSheet4 As Excel.Worksheet = Nothing
    Dim oxclWorkSheetSFirst As Excel.Worksheet = Nothing
    Dim oxclWorkSheetSSecond As Excel.Worksheet = Nothing
    Dim oxclWorkSheetSThird As Excel.Worksheet = Nothing
    Dim noticeDate As Date
    Dim decTotal, decVat, decCents As Decimal
    Dim intApp1, intApp2, intApp3, intApp4, intRow1, intRow2, intRow3 As Integer
    Dim blnIntSumMore, blnSumMoreDate As Boolean

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblCompletedTask.Text = "" : lblCompletedTask.Visible = False : btnBrowse.Focus()
        blnIntSumMore = False : blnSumMoreDate = False
    End Sub

    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        ofdBrowse.InitialDirectory = "C:\"
        ofdBrowse.Title = "Select your Excel Data"
        ofdBrowse.Filter = "Excel Files|*.xls"
        ofdBrowse.ShowDialog()
        txtDataSource.Text = ofdBrowse.FileName
        'GetSourceWorksheetCount()
        DExcelConnPath = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ofdBrowse.FileName + ";Extended Properties=Excel 8.0;"
    End Sub

    Private Sub btnGenNotice_Click(sender As Object, e As EventArgs) Handles btnGenNotice.Click
        strmessage = ""
        If txtDataSource.Text.Trim = String.Empty Then strmessage = " - Browse for the excel file which contains the data." : errProvider.SetError(txtDataSource, "Cannot leave data source blank") Else errProvider.SetError(txtDataSource, "")
        If strmessage <> "" Then
            MessageBox.Show(vbCrLf & vbCrLf & strmessage, "MISSING IMPORTANT DETAIL", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            noticeDate = dtpDate.Value.Date : btnGenNotice.Enabled = False : btnBrowse.Enabled = False : dtpDate.Enabled = False : lblCompletedTask.Visible = True : SelectandSetProcedure() : PlottoExcelTemplate() : btnGenNotice.Enabled = True : btnBrowse.Enabled = True : dtpDate.Enabled = True : lblCompletedTask.Text = String.Empty : lblCompletedTask.Visible = False : txtDataSource.Text = String.Empty
        End If
    End Sub

#Region "Data Procedures"
    Private Function GetConnection() As OleDbConnection
        olecon = New OleDbConnection
        olecon.ConnectionString = DExcelConnPath
        olecon.Open()
        Return olecon
    End Function
    Private Sub SelectandSetProcedure()
        Try
            GetConnection()
            xclColumnsDTable = New DataTable
            xclDtable = New DataTable
            xclDistinct = New DataTable
            xclColumnsDTable = olecon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
            For i As Integer = 0 To xclColumnsDTable.Rows.Count - 1
                sqlQuery = "SELECT AGRTNO,OWNER,ARRCNT,ARREAR,SURCHR,CURDUE,ADDRS1,ADDRS2,ZIPCDE,COLCDE,TELNUM FROM [" & xclColumnsDTable.Rows(i).Item(2).ToString & "] ORDER BY OWNER ASC"
                oleAdpter = New OleDbDataAdapter(sqlQuery, olecon)
                oleAdpter.Fill(xclDtable)

                sqlQuery = "SELECT DISTINCT ARRCNT FROM [" & xclColumnsDTable.Rows(i).Item(2).ToString & "] ORDER BY ARRCNT ASC"
                oleAdpter = New OleDbDataAdapter(sqlQuery, olecon)
                oleAdpter.Fill(xclDistinct)
            Next
        Catch ex As Exception
            MessageBox.Show("An error has occured while selecting and setting data", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            GetConnection.Close()
        End Try
    End Sub
    Private Sub PlottoExcelTemplate()
        Try
            Dim decTot, decAdd As Decimal
            intApp1 = 0 : intApp2 = 0 : intApp3 = 0
            'oxclApp = New Excel.Application 'CreateObject("Excel.Application")
            If xclDistinct.Rows.Count <> 0 Then
                OpenWorkbook()
            End If
            If xclDtable.Rows.Count <> 0 Then
                decAdd = 0 : decTot = 0
                decAdd = 100 / xclDtable.Rows.Count
                For i As Integer = 0 To xclDtable.Rows.Count - 1
                    Select Case xclDtable.Rows(i).Item(2)
                        Case 3
                            intApp1 += 1
                            FirstNotice(i)
                        Case 4
                            intApp2 += 1
                            SecondNotice(i)
                        Case 5
                            intApp3 += 1
                            ThirdNotice(i)
                        Case Is > 5
                            intApp4 += 1
                            MoreThanNotice(i, CInt(xclDtable.Rows(i).Item(2).ToString))
                    End Select
                    decTot += decAdd
                    lblCompletedTask.Text = CStr(Math.Round(decTot)) & "% Completed"
                Next
            End If
            If intApp1 <> 0 Then oxclWorkSheetSFirst.Range("A5:M" & intRow1).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
            If intApp2 <> 0 Then oxclWorkSheetSSecond.Range("A5:M" & intRow2).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
            If intApp3 <> 0 Then oxclWorkSheetSThird.Range("A5:M" & intRow3).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
            OpenXclApp()
        Catch ex As Exception
            MessageBox.Show("Error in generating excel for printing." & vbCrLf & "Error Description : " & vbCrLf & Err.Description, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error) : Application.Exit()
        End Try
    End Sub

    Private Sub OpenWorkbook()
        Dim blnOpenMore As Boolean
        Try
            blnOpenMore = False
            For p As Integer = 0 To xclDistinct.Rows.Count - 1
                With xclDistinct.Rows(p)
                    Select Case .Item(0)
                        Case 3
                            oxclWorkbook1 = oxclApp1.Workbooks.Open(Mid(My.Application.Info.DirectoryPath, 1, (My.Application.Info.DirectoryPath).Length - 25) + "NOTICETEMPLATE\firstNotice.xlt")
                            oxclWorkSheet1 = oxclWorkbook1.Worksheets("Sheet1")
                        Case 4
                            oxclWorkbook2 = oxclApp2.Workbooks.Open(Mid(My.Application.Info.DirectoryPath, 1, (My.Application.Info.DirectoryPath).Length - 25) + "NOTICETEMPLATE\secondNotice.xlt")
                            oxclWorkSheet2 = oxclWorkbook2.Worksheets("Sheet1")
                        Case 5
                            oxclWorkbook3 = oxclApp3.Workbooks.Open(Mid(My.Application.Info.DirectoryPath, 1, (My.Application.Info.DirectoryPath).Length - 25) + "NOTICETEMPLATE\thirdfinalNotice.xlt")
                            oxclWorkSheet3 = oxclWorkbook3.Worksheets("Sheet1")
                        Case Is > 5
                            blnOpenMore = True
                    End Select
                End With
            Next
            If blnOpenMore Then OpenWorkbookMoreThan()
            oxclWorkbookSummary = oxclAppSummary.Workbooks.Open(Mid(My.Application.Info.DirectoryPath, 1, (My.Application.Info.DirectoryPath).Length - 25) + "NOTICETEMPLATE\Summary.xlt")
            oxclWorkSheetSFirst = oxclWorkbookSummary.Worksheets("FirstNotice")
            oxclWorkSheetSSecond = oxclWorkbookSummary.Worksheets("SecondNotice")
            oxclWorkSheetSThird = oxclWorkbookSummary.Worksheets("ThirdNotice")
        Catch ex As Exception
            MessageBox.Show("Error in opening workbook." & vbCrLf & "Error Description : " & vbCrLf & Err.Description, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error) : Application.Exit()
        End Try
    End Sub
    Private Sub OpenWorkbookMoreThan()
        oxclWorkbook4 = oxclApp4.Workbooks.Open(Mid(My.Application.Info.DirectoryPath, 1, (My.Application.Info.DirectoryPath).Length - 25) + "NOTICETEMPLATE\MorethanFive.xlt")
        oxclWorkSheet4 = oxclWorkbook4.Worksheets("Sheet1")
    End Sub
    Private Sub OpenXclApp()
        Try
            For p As Integer = 0 To xclDistinct.Rows.Count - 1
                With xclDistinct.Rows(p)
                    Select Case .Item(0)
                        Case 3
                            oxclApp1.Visible = True
                        Case 4
                            oxclApp2.Visible = True
                        Case 5
                            oxclApp3.Visible = True
                        Case Is > 5
                            oxclApp4.Visible = True
                    End Select
                End With
            Next
            oxclAppSummary.Visible = True
        Catch ex As Exception
            MessageBox.Show("Error in opening excel." & vbCrLf & "Error Description : " & vbCrLf & Err.Description, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error) : Application.Exit()
        End Try
    End Sub

    Private Sub FirstNotice(ByVal RowIndex As Integer)
        With xclDtable.Rows(RowIndex)
            decTotal = Math.Round(xclDtable.Rows(RowIndex).Item(3) + xclDtable.Rows(RowIndex).Item(4) + xclDtable.Rows(RowIndex).Item(5), 2, MidpointRounding.AwayFromZero)
            decVat = Math.Round(decTotal * 0.12, 2, MidpointRounding.AwayFromZero)
            decCents = 0
            decCents = decVat Mod 1
            decVat = IIf(decCents <> 0.0, SeparateWholeInt(decVat) + 1, decVat)
            If RowIndex = 0 Or intApp1 = 1 Then
                With oxclWorkSheet1
                    .Range("G6").Value = Format(noticeDate, "MMMM dd, yyyy")
                    .Range("D26").Value = Format(CDate(noticeDate.AddMonths(-1).Month & "/" & MonthEndDay() & "/" & noticeDate.Year), "MMMM dd, yyyy")
                    .Range("F12").Value = xclDtable.Rows(RowIndex).Item(0)
                    .Range("A7").Value = xclDtable.Rows(RowIndex).Item(1)
                    .Range("A8").Value = xclDtable.Rows(RowIndex).Item(6)
                    .Range("A9").Value = xclDtable.Rows(RowIndex).Item(7)
                    .Range("A10").Value = xclDtable.Rows(RowIndex).Item(8)

                    .Range("G18").Value = xclDtable.Rows(RowIndex).Item(3) 'FormatNumber(xclDtable.Rows(RowIndex).Item(3), 2, TriState.True)
                    .Range("G19").Value = xclDtable.Rows(RowIndex).Item(4) 'FormatNumber(xclDtable.Rows(RowIndex).Item(4), 2, TriState.True)
                    .Range("G20").Value = xclDtable.Rows(RowIndex).Item(5) 'FormatNumber(xclDtable.Rows(RowIndex).Item(5), 2, TriState.True)
                    .Range("G21").Value = decTotal 'FormatNumber(decTotal, 2, TriState.True)
                    .Range("G22").Value = decVat  'FormatNumber(decVat, 2, TriState.True)
                    .Range("G23").Value = decTotal + decVat 'FormatNumber(decTotal + decVat, 2, TriState.True)
                End With
            Else
                Application.DoEvents()
                Dim oxclNewWorkSheet1 As Excel.Worksheet = Nothing
                With oxclWorkbook1
                    oxclNewWorkSheet1 = .Worksheets.Add()
                    oxclNewWorkSheet1.Select()
                End With
                With oxclNewWorkSheet1
                    .Range("A1:I1").Merge()
                    .Range("A2:I2").Merge()
                    .Range("A4:I4").Merge()
                    .Range("G6:I6").Merge()
                    .Range("F12:G12").Merge()
                    .Range("G18:H18").Merge()
                    .Range("G19:H19").Merge()
                    .Range("G20:H20").Merge()
                    .Range("G21:H21").Merge()
                    .Range("G22:H22").Merge()
                    .Range("G23:H23").Merge()
                    .Range("D26:F26").Merge()
                    .Range("A16:I17").Merge()
                    .Range("B28:I29").Merge()
                    .Range("B30:I31").Merge()
                    .Range("A37:I37").Merge()
                    .Range("A41:I42").Merge()
                    oxclWorkSheet1.Cells.Range("A1:I56").Copy()
                    .Paste()
                    .Range("A1:A3").RowHeight = 12.75
                    .Range("A5:A56").RowHeight = 12.75
                    .Range("A5").RowHeight = 6.75
                    .Range("A11").RowHeight = 6.75
                    .Range("A13").RowHeight = 6.75
                    .Range("A15").RowHeight = 6.75
                    .Range("A24").RowHeight = 6.75
                    .Range("A32").RowHeight = 6.75
                    .Range("A36").RowHeight = 6.75
                    .Range("A40").RowHeight = 6.75
                    .Range("A55").RowHeight = 6.75
                    .Range("A1").ColumnWidth = 4.57
                    .Range("E1").ColumnWidth = 6.14
                    With .PageSetup
                        .TopMargin = oxclApp1.InchesToPoints(1.0)
                        .LeftMargin = oxclApp1.InchesToPoints(1.0)
                        .RightMargin = oxclApp1.InchesToPoints(1.0)
                        .BottomMargin = oxclApp1.InchesToPoints(1.0)
                        .FooterMargin = oxclApp1.InchesToPoints(0)
                        .HeaderMargin = oxclApp1.InchesToPoints(0)
                    End With
                    .Range("G6").Value = Format(noticeDate, "MMMM dd, yyyy")
                    .Range("D26").Value = Format(CDate(noticeDate.AddMonths(-1).Month & "/" & MonthEndDay() & "/" & noticeDate.Year), "MMMM dd, yyyy")
                    .Range("F12").Value = xclDtable.Rows(RowIndex).Item(0)
                    .Range("A7").Value = xclDtable.Rows(RowIndex).Item(1)
                    .Range("A8").Value = xclDtable.Rows(RowIndex).Item(6)
                    .Range("A9").Value = xclDtable.Rows(RowIndex).Item(7)
                    .Range("A10").Value = xclDtable.Rows(RowIndex).Item(8)

                    .Range("G18").Value = xclDtable.Rows(RowIndex).Item(3) 'FormatNumber(xclDtable.Rows(RowIndex).Item(3), 2, TriState.True)
                    .Range("G19").Value = xclDtable.Rows(RowIndex).Item(4) 'FormatNumber(xclDtable.Rows(RowIndex).Item(4), 2, TriState.True)
                    .Range("G20").Value = xclDtable.Rows(RowIndex).Item(5) 'FormatNumber(xclDtable.Rows(RowIndex).Item(5), 2, TriState.True)
                    .Range("G21").Value = decTotal 'FormatNumber(decTotal, 2, TriState.True)
                    .Range("G22").Value = decVat   'FormatNumber(decVat, 2, TriState.True)
                    .Range("G23").Value = decTotal + decVat 'FormatNumber(decTotal + decVat, 2, TriState.True)
                End With
            End If
        End With
        Summary(3, RowIndex)
    End Sub

    Private Sub SecondNotice(ByVal RowIndex As Integer)
        With xclDtable.Rows(RowIndex)
            decTotal = Math.Round(xclDtable.Rows(RowIndex).Item(3) + xclDtable.Rows(RowIndex).Item(4) + xclDtable.Rows(RowIndex).Item(5), 2, MidpointRounding.AwayFromZero)
            decVat = Math.Round(decTotal * 0.12, 2, MidpointRounding.AwayFromZero)
            decCents = 0
            decCents = decVat Mod 1
            decVat = IIf(decCents <> 0.0, SeparateWholeInt(decVat) + 1, decVat)
            If RowIndex = 0 Or intApp2 = 1 Then
                With oxclWorkSheet2
                    .Range("G6").Value = Format(noticeDate, "MMMM dd, yyyy")
                    .Range("F12").Value = xclDtable.Rows(RowIndex).Item(0)
                    .Range("A7").Value = xclDtable.Rows(RowIndex).Item(1)
                    .Range("A8").Value = xclDtable.Rows(RowIndex).Item(6)
                    .Range("A9").Value = xclDtable.Rows(RowIndex).Item(7)
                    .Range("A10").Value = xclDtable.Rows(RowIndex).Item(8)

                    .Range("G19").Value = xclDtable.Rows(RowIndex).Item(3) 'FormatNumber(xclDtable.Rows(RowIndex).Item(3), 2, TriState.True)
                    .Range("G20").Value = xclDtable.Rows(RowIndex).Item(4) 'FormatNumber(xclDtable.Rows(RowIndex).Item(4), 2, TriState.True)
                    .Range("G21").Value = xclDtable.Rows(RowIndex).Item(5) 'FormatNumber(xclDtable.Rows(RowIndex).Item(5), 2, TriState.True)
                    .Range("G22").Value = decTotal 'FormatNumber(decTotal, 2, TriState.True)
                    .Range("G23").Value = decVat  'FormatNumber(decVat, 2, TriState.True, TriState.UseDefault, TriState.True)
                    .Range("G24").Value = decTotal + decVat 'FormatNumber(decTotal + decVat, 2, TriState.True)
                End With
            Else
                Application.DoEvents()
                Dim oxclNewWorkSheet2 As Excel.Worksheet = Nothing
                With oxclWorkbook2
                    oxclNewWorkSheet2 = .Worksheets.Add()
                    oxclNewWorkSheet2.Select()
                End With
                With oxclNewWorkSheet2
                    .Range("A1:I1").Merge()
                    .Range("A2:I2").Merge()
                    .Range("A4:I4").Merge()
                    .Range("A46:I46").Merge()
                    .Range("A16:I17").Merge()
                    .Range("A26:I30").Merge()
                    .Range("A32:I34").Merge()
                    .Range("A36:I37").Merge()
                    .Range("G3:H3").Merge()
                    .Range("F9:G9").Merge()
                    .Range("G19:H19").Merge()
                    .Range("G20:H20").Merge()
                    .Range("G21:H21").Merge()
                    .Range("G22:H22").Merge()
                    .Range("G23:H23").Merge()
                    .Range("G24:H24").Merge()
                    .Range("G6:I6").Merge()
                    oxclWorkSheet2.Cells.Range("A1:I49").Copy()
                    .Paste()
                    .Range("A2:A3").RowHeight = 14.25
                    .Range("A5:A11").RowHeight = 14.25
                    .Range("A13:A41").RowHeight = 14.25
                    .Range("A43:A49").RowHeight = 14.25
                    .Range("A18").RowHeight = 6.75
                    .Range("A31").RowHeight = 6.75
                    .Range("A35").RowHeight = 6.75
                    .Range("A38").RowHeight = 6.75
                    .Range("A40").RowHeight = 6.75
                    With .PageSetup
                        .TopMargin = oxclApp2.InchesToPoints(1.0)
                        .LeftMargin = oxclApp2.InchesToPoints(1.0)
                        .RightMargin = oxclApp2.InchesToPoints(1.0)
                        .BottomMargin = oxclApp2.InchesToPoints(1.0)
                        .FooterMargin = oxclApp2.InchesToPoints(0)
                        .HeaderMargin = oxclApp2.InchesToPoints(0)
                    End With
                    .Range("G6").Value = Format(noticeDate, "MMMM dd, yyyy")
                    .Range("F12").Value = xclDtable.Rows(RowIndex).Item(0)
                    .Range("A7").Value = xclDtable.Rows(RowIndex).Item(1)
                    .Range("A8").Value = xclDtable.Rows(RowIndex).Item(6)
                    .Range("A9").Value = xclDtable.Rows(RowIndex).Item(7)
                    .Range("A10").Value = xclDtable.Rows(RowIndex).Item(8)

                    .Range("G19").Value = xclDtable.Rows(RowIndex).Item(3) 'FormatNumber(xclDtable.Rows(RowIndex).Item(3), 2, TriState.True)
                    .Range("G20").Value = xclDtable.Rows(RowIndex).Item(4) 'FormatNumber(xclDtable.Rows(RowIndex).Item(4), 2, TriState.True)
                    .Range("G21").Value = xclDtable.Rows(RowIndex).Item(5) 'FormatNumber(xclDtable.Rows(RowIndex).Item(5), 2, TriState.True)
                    .Range("G22").Value = decTotal 'FormatNumber(decTotal, 2, TriState.True)
                    .Range("G23").Value = decVat   'FormatNumber(decVat, 2, TriState.True, TriState.UseDefault, TriState.True)
                    .Range("G24").Value = decTotal + decVat 'FormatNumber(decTotal + decVat, 2, TriState.True)
                End With
            End If
        End With
        Summary(4, RowIndex)
    End Sub

    Private Sub ThirdNotice(ByVal RowIndex As Integer)
        With xclDtable.Rows(RowIndex)
            decTotal = Math.Round(xclDtable.Rows(RowIndex).Item(3) + xclDtable.Rows(RowIndex).Item(4) + xclDtable.Rows(RowIndex).Item(5), 2, MidpointRounding.AwayFromZero)
            decVat = Math.Round(decTotal * 0.12, 2, MidpointRounding.AwayFromZero)
            decCents = 0
            decCents = decVat Mod 1
            decVat = IIf(decCents <> 0.0, SeparateWholeInt(decVat) + 1, decVat)
            If RowIndex = 0 Or intApp3 = 1 Then
                With oxclWorkSheet3
                    .Range("G6").Value = Format(noticeDate, "MMMM dd, yyyy")
                    .Range("F12").Value = xclDtable.Rows(RowIndex).Item(0)
                    .Range("A7").Value = xclDtable.Rows(RowIndex).Item(1)
                    .Range("A8").Value = xclDtable.Rows(RowIndex).Item(6)
                    .Range("A9").Value = xclDtable.Rows(RowIndex).Item(7)
                    .Range("A10").Value = xclDtable.Rows(RowIndex).Item(8)

                    '.Range("A18").Value = Format(noticeDate.AddMonths(-2), "MMMM yyyy")
                    '.Range("F18").Value = Format(noticeDate.AddMonths(-1), "MMMM yyyy")
                    '.Range("G20").Value = xclDtable.Rows(RowIndex).Item(3) 'FormatNumber(xclDtable.Rows(RowIndex).Item(3), 2, TriState.True)
                    '.Range("G21").Value = xclDtable.Rows(RowIndex).Item(4) 'FormatNumber(xclDtable.Rows(RowIndex).Item(4), 2, TriState.True)
                    '.Range("G22").Value = xclDtable.Rows(RowIndex).Item(5) 'FormatNumber(xclDtable.Rows(RowIndex).Item(5), 2, TriState.True)
                    '.Range("G23").Value = decTotal 'FormatNumber(decTotal, 2, TriState.True)
                    '.Range("G24").Value = decVat  'FormatNumber(decVat, 2, TriState.True)
                    '.Range("G25").Value = decTotal + decVat 'FormatNumber(decTotal + decVat, 2, TriState.True)
                End With
            Else
                Application.DoEvents()
                Dim oxclNewWorkSheet3 As Excel.Worksheet = Nothing
                With oxclWorkbook3
                    oxclNewWorkSheet3 = .Worksheets.Add()
                    oxclNewWorkSheet3.Select()
                End With
                With oxclNewWorkSheet3
                    .Range("A1:I1").Merge()
                    .Range("A2:I2").Merge()
                    .Range("A4:I4").Merge()
                    .Range("G6:I6").Merge()
                    .Range("F12:G12").Merge()
                    .Range("A16:I18").Merge()
                    .Range("A20:I21").Merge()
                    .Range("A23:I24").Merge()
                   
                    oxclWorkSheet3.Cells.Range("A1:I36").Copy()
                    .Paste()
                    .Range("A2:A3").RowHeight = 14.25
                    .Range("A5:A11").RowHeight = 14.25
                    .Range("A13:A37").RowHeight = 14.25
                    .Range("A13").RowHeight = 14.25
                    .Range("A15:A36").RowHeight = 14.25
                    With .PageSetup
                        .TopMargin = oxclApp3.InchesToPoints(1.0)
                        .LeftMargin = oxclApp3.InchesToPoints(1.0)
                        .RightMargin = oxclApp3.InchesToPoints(1.0)
                        .BottomMargin = oxclApp3.InchesToPoints(1.0)
                        .FooterMargin = oxclApp3.InchesToPoints(0)
                        .HeaderMargin = oxclApp3.InchesToPoints(0)
                    End With
                    .Range("G6").Value = Format(noticeDate, "MMMM dd, yyyy")
                    .Range("F12").Value = xclDtable.Rows(RowIndex).Item(0)
                    .Range("A7").Value = xclDtable.Rows(RowIndex).Item(1)
                    .Range("A8").Value = xclDtable.Rows(RowIndex).Item(6)
                    .Range("A9").Value = xclDtable.Rows(RowIndex).Item(7)
                    .Range("A10").Value = xclDtable.Rows(RowIndex).Item(8)

                    '.Range("A18").Value = Format(noticeDate.AddMonths(-2), "MMMM yyyy")
                    '.Range("F18").Value = Format(noticeDate.AddMonths(-1), "MMMM yyyy")
                    '.Range("G20").Value = xclDtable.Rows(RowIndex).Item(3) 'FormatNumber(xclDtable.Rows(RowIndex).Item(3), 2, TriState.True)
                    '.Range("G21").Value = xclDtable.Rows(RowIndex).Item(4) 'FormatNumber(xclDtable.Rows(RowIndex).Item(4), 2, TriState.True)
                    '.Range("G22").Value = xclDtable.Rows(RowIndex).Item(5) 'FormatNumber(xclDtable.Rows(RowIndex).Item(5), 2, TriState.True)
                    '.Range("G23").Value = decTotal 'FormatNumber(decTotal, 2, TriState.True)
                    '.Range("G24").Value = decVat  'FormatNumber(decVat, 2, TriState.True)
                    '.Range("G25").Value = decTotal + decVat 'FormatNumber(decTotal + decVat, 2, TriState.True)
                End With
            End If
        End With
        Summary(5, RowIndex)
    End Sub

    Private Sub MoreThanNotice(ByVal RowIndex As Integer, ByVal ArrCnt As Integer)
        With xclDtable.Rows(RowIndex)
            decTotal = Math.Round(xclDtable.Rows(RowIndex).Item(3) + xclDtable.Rows(RowIndex).Item(4) + xclDtable.Rows(RowIndex).Item(5), 2, MidpointRounding.AwayFromZero)
            decVat = Math.Round(decTotal * 0.12, 2, MidpointRounding.AwayFromZero)
            decCents = 0
            decCents = decVat Mod 1
            decVat = IIf(decCents <> 0.0, SeparateWholeInt(decVat) + 1, decVat)
            If RowIndex = 0 Or intApp4 = 1 Then
                With oxclWorkSheet4
                    .Range("G6").Value = Format(noticeDate, "MMMM dd, yyyy")
                    .Range("F12").Value = xclDtable.Rows(RowIndex).Item(0)
                    .Range("A7").Value = xclDtable.Rows(RowIndex).Item(1)
                    .Range("A8").Value = xclDtable.Rows(RowIndex).Item(6)
                    .Range("A9").Value = xclDtable.Rows(RowIndex).Item(7)
                    .Range("A10").Value = xclDtable.Rows(RowIndex).Item(8)

                    '.Range("A18").Value = Format(noticeDate.AddMonths(-2), "MMMM yyyy")
                    '.Range("F18").Value = Format(noticeDate.AddMonths(-1), "MMMM yyyy")
                    '.Range("G20").Value = xclDtable.Rows(RowIndex).Item(3) 'FormatNumber(xclDtable.Rows(RowIndex).Item(3), 2, TriState.True)
                    '.Range("G21").Value = xclDtable.Rows(RowIndex).Item(4) 'FormatNumber(xclDtable.Rows(RowIndex).Item(4), 2, TriState.True)
                    '.Range("G22").Value = xclDtable.Rows(RowIndex).Item(5) 'FormatNumber(xclDtable.Rows(RowIndex).Item(5), 2, TriState.True)
                    '.Range("G23").Value = decTotal 'FormatNumber(decTotal, 2, TriState.True)
                    '.Range("G24").Value = decVat  'FormatNumber(decVat, 2, TriState.True)
                    '.Range("G25").Value = decTotal + decVat 'FormatNumber(decTotal + decVat, 2, TriState.True)
                End With
            Else
                Application.DoEvents()
                Dim oxclNewWorkSheet4 As Excel.Worksheet = Nothing
                With oxclWorkbook4
                    oxclNewWorkSheet4 = .Worksheets.Add()
                    oxclNewWorkSheet4.Select()
                End With
                With oxclNewWorkSheet4
                    .Range("A1:I1").Merge()
                    .Range("A2:I2").Merge()
                    .Range("A4:I4").Merge()
                    .Range("G6:I6").Merge()
                    .Range("F12:G12").Merge()
                    .Range("A16:I18").Merge()
                    .Range("A20:I21").Merge()
                    .Range("A23:I24").Merge()

                    oxclWorkSheet4.Cells.Range("A1:I36").Copy()
                    .Paste()
                    .Range("A2:A3").RowHeight = 14.25
                    .Range("A5:A11").RowHeight = 14.25
                    .Range("A13:A37").RowHeight = 14.25
                    .Range("A13").RowHeight = 14.25
                    .Range("A15:A36").RowHeight = 14.25
                    With .PageSetup
                        .TopMargin = oxclApp4.InchesToPoints(1.0)
                        .LeftMargin = oxclApp4.InchesToPoints(1.0)
                        .RightMargin = oxclApp4.InchesToPoints(1.0)
                        .BottomMargin = oxclApp4.InchesToPoints(1.0)
                        .FooterMargin = oxclApp4.InchesToPoints(0)
                        .HeaderMargin = oxclApp4.InchesToPoints(0)
                    End With
                    .Range("G6").Value = Format(noticeDate, "MMMM dd, yyyy")
                    .Range("F12").Value = xclDtable.Rows(RowIndex).Item(0)
                    .Range("A7").Value = xclDtable.Rows(RowIndex).Item(1)
                    .Range("A8").Value = xclDtable.Rows(RowIndex).Item(6)
                    .Range("A9").Value = xclDtable.Rows(RowIndex).Item(7)
                    .Range("A10").Value = xclDtable.Rows(RowIndex).Item(8)
                End With
            End If
        End With
        Summary(ArrCnt, RowIndex)
    End Sub

    Private Sub Summary(ByVal ArrCnt As Integer, ByVal RowIndex As Integer)
        Select Case ArrCnt
            Case 3
                If intApp1 = 1 Then intRow1 = 6 Else intRow1 += 1
                If intApp1 = 1 Then
                    oxclWorkSheetSFirst.Range("A3").Value = Format(noticeDate, "MMMM yyyy")
                End If
                SetSummary(oxclWorkSheetSFirst, RowIndex, intApp1, intRow1)
            Case 4
                If intApp2 = 1 Then intRow2 = 6 Else intRow2 += 1
                If intApp2 = 1 Then
                    oxclWorkSheetSSecond.Range("A3").Value = Format(noticeDate, "MMMM yyyy")
                End If
                SetSummary(oxclWorkSheetSSecond, RowIndex, intApp2, intRow2)
            Case 5, Is > 5
                If (intApp3 = 1 Or intApp4 = 1) And blnIntSumMore = False Then intRow3 = 6 : blnIntSumMore = True Else intRow3 += 1
                If (intApp3 = 1 Or intApp4 = 1) And blnSumMoreDate = False Then
                    blnSumMoreDate = True : oxclWorkSheetSThird.Range("A3").Value = Format(noticeDate, "MMMM yyyy")
                End If
                If ArrCnt > 5 Then intApp3 = intApp4
                SetSummary(oxclWorkSheetSThird, RowIndex, intApp3, intRow3)
        End Select
    End Sub

    Private Sub SetSummary(ByVal wrksheet As Excel.Worksheet, ByVal RowIndex As Integer, ByVal intApp As Integer, ByVal intExcelRow As Integer)
        With wrksheet
            .Range("A" & intExcelRow).Value = intApp
            .Range("B" & intExcelRow).Value = xclDtable.Rows(RowIndex).Item(9).ToString.ToUpper
            .Range("C" & intExcelRow).Value = xclDtable.Rows(RowIndex).Item(0).ToString.ToUpper
            .Range("D" & intExcelRow).Value = xclDtable.Rows(RowIndex).Item(1).ToString.ToUpper
            .Range("E" & intExcelRow).Value = xclDtable.Rows(RowIndex).Item(10).ToString.ToUpper
            .Range("F" & intExcelRow).Value = xclDtable.Rows(RowIndex).Item(6).ToString.ToUpper & " " & xclDtable.Rows(RowIndex).Item(7).ToString.ToUpper & " " & xclDtable.Rows(RowIndex).Item(8).ToString.ToUpper
            .Range("G" & intExcelRow).Value = xclDtable.Rows(RowIndex).Item(3) 'FormatNumber(xclDtable.Rows(RowIndex).Item(3), 2, TriState.True)
            .Range("H" & intExcelRow).Value = xclDtable.Rows(RowIndex).Item(4) 'FormatNumber(xclDtable.Rows(RowIndex).Item(4), 2, TriState.True)
            .Range("I" & intExcelRow).Value = xclDtable.Rows(RowIndex).Item(5) 'FormatNumber(xclDtable.Rows(RowIndex).Item(5), 2, TriState.True)
            .Range("J" & intExcelRow).Value = decTotal 'FormatNumber(decTotal, 2, TriState.True)
            .Range("K" & intExcelRow).Value = decVat 'FormatNumber(decVat, 2, TriState.True)
            .Range("L" & intExcelRow).Value = decTotal + decVat 'FormatNumber(decTotal + decVat, 2, TriState.True)
        End With
    End Sub

    Private Function MonthEndDay() As Integer
        Dim month As Integer = noticeDate.AddMonths(-1).Month
        Select Case month
            Case 1, 3, 5, 7, 8, 10, 12
                MonthEndDay = 31
            Case 4, 6, 9, 11
                MonthEndDay = 30
            Case Else
                MonthEndDay = IIf(Date.IsLeapYear(noticeDate.Year), 29, 28)
        End Select
        Return MonthEndDay
    End Function
    Private Function SeparateWholeInt(ByVal dec As Decimal) As Integer
        Dim strSplit As String()
        strSplit = CStr(dec).Split(".")
        Return CInt(strSplit(0))
    End Function
#End Region

End Class
