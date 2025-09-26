Imports OfficeOpenXml
Imports System.Configuration
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net.Mail
Imports System.Data.Common
Imports System.Data.OleDb
Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types
Imports OfficeOpenXml.FormulaParsing.Excel.Functions.Math
Imports System.ComponentModel
Imports System.Globalization



Module ExceptionReport


    Dim DBUtil As New DBUtil

    'Dim connBOS As New SqlClient.SqlConnection(DBUtil.KPIConnStr("BOS")) 
    Dim connOracle As New Oracle.ManagedDataAccess.Client.OracleConnection(DBUtil.OraConnStr())
    Dim objCmd As OracleCommand
    Dim objReader As OracleDataReader
    Dim strResponceConnSql As String

    Dim strPathFile As String = ConfigurationManager.AppSettings("strPathFile")
    Dim strBUcode As String
    Dim strBUName As String
    Dim strBUH As String
    Dim strMailTo As String
    Dim strMailCc As String
    Dim strPathFileNameReport1 As String
    Dim strPathFileNameReport2 As String
    Dim strPathFileNameReport3 As String
    Dim strDataAsofDate As String
    Dim strDateMonth As String 'MM
    Dim strDateYear As String 'YYYY
    Dim strDateSnap As String 'YYMM
    Dim queryExc014 As String
    Dim queryExc015 As String
    Dim queryExc016 As String
    'Count'
    Dim strExc14 As String
    Dim strExc15 As String
    Dim strExc16 As String
    Dim dataDate As DateTime
    Dim formattedDataDate As String
    Dim today As Date
    Dim endOfMonth As Date
    Dim endOfMonthStr As String
    Dim strMsg As String
    Dim nameTo As String
    Dim nameToWithFormatted As String
    Dim strQuerynameTo As String
    Dim asAtDate As String
    Dim excPassword As String
    Dim dateFromExcTable As String
    Dim asAtDateReferdtExcTable As String


    Sub Main()


        connOracle.Open()
        If connOracle.State = ConnectionState.Open Then
            strResponceConnSql = "Oracle Server Connected"
            Debug.WriteLine(strResponceConnSql)


            ' Print environment details
            Debug.WriteLine("Connection String: " & connOracle.ConnectionString)
            Debug.WriteLine("Data Source: " & connOracle.DataSource)
            Debug.WriteLine("Server Version: " & connOracle.ServerVersion)
            Debug.WriteLine("State: " & connOracle.State.ToString())

        Else
            strResponceConnSql = "Oracle Server Connect Failed"
            Debug.WriteLine(strResponceConnSql)
        End If

        'ปีเดือนวันเปลี่ยนตามวันที่ออกแบบ 20 ,25,27 ปีก็เปลียน'
        'Dim currentYear, currentMonth

        'currentYear = Year(Now)
        'currentMonth = Month(Now)

        'excPassword = "RDT" & currentYear & currentMonth

        'getDataTable - Mail List

        Dim strSQL As String = "Select * FROM rm_ma_mail WHERE batch_job = 'RDT_EXCEPTION_REPORT' "
        dateFromExcTable = "SELECT MAX(data_date) AS data_date FROM t_rdt_exc_report_011"
        objCmd = connOracle.CreateCommand
        objCmd.CommandText = dateFromExcTable
        objReader = objCmd.ExecuteReader()
        Dim dtMaxDate As DataTable = New DataTable("Max_Date")
        dtMaxDate.Load(objReader)
        Dim inputMaxDate As DateTime = DBUtil.chkNull(dtMaxDate.Rows(0)("data_date"))
        Debug.WriteLine("test input data date 1 => " + inputMaxDate)

        'export to excel

        'getDataTable - Mail List
        objCmd = connOracle.CreateCommand
        objCmd.CommandText = strSQL
        objReader = objCmd.ExecuteReader()
        'Create a new DataTable.
        Dim dtMailList As DataTable = New DataTable("MailList")
        'Load DataReader into the DataTable.
        dtMailList.Load(objReader)
        Dim intNumRecordBUList As Integer = dtMailList.Rows.Count
        today = Date.Today.AddDays(-1)
        Dim thaiCulture As New CultureInfo("th-TH")

        thaiCulture.DateTimeFormat.Calendar = New GregorianCalendar()

        asAtDate = today.ToString("dd/MM/yyyy", thaiCulture)
        asAtDateReferdtExcTable = inputMaxDate.ToString("dd/MM/yyyy", thaiCulture)

        Dim inputDate As String = inputMaxDate.ToString("dd MMM yyyy", thaiCulture)


        Debug.WriteLine("test input data date 2 => " + inputDate)
        'thaiCulture.DateTimeFormat.Calendar = New GregorianCalendar()'
        Dim endOfDayThaiDate As String = today.ToString("dd MMM yyyy", thaiCulture)
        Debug.WriteLine("currentDate => " + endOfDayThaiDate)

        For i As Integer = 0 To intNumRecordBUList - 1
            strBUcode = dtMailList.Rows(i)("BU_CODE")
            strMailTo = DBUtil.chkNull(dtMailList.Rows(i)("MAIL_TO"))
            strMailCc = DBUtil.chkNull(dtMailList.Rows(i)("MAIL_CC"))
            'ExportToExcelReport1(strBUcode, endOfDayThaiDate)
            ExportToExcelReport1(strBUcode, inputDate)
            'SendMail(strBUcode, strMailTo, strMailCc, asAtDate)
            SendMail(strBUcode, strMailTo, strMailCc, asAtDateReferdtExcTable)
        Next

        'objReader.Close()
        'objConn.Close()

    End Sub
    Sub ExportToExcelReport1(ByVal strBUcode As String, ByVal endOfDayDateThai As String)


        'ใน table mail ถ้าเปลี่ยนเป็นอย่างอื่นนอกจาก  CC ก็ปรับได้'
        If strBUcode <> "CC" Then
            Dim strWhereCondition As String = ""

            'Template ที่จะใช้
            Dim strPathTemplate1 As String = ConfigurationManager.AppSettings("strPathTemplate1")
            Dim buParam As String = strBUcode
            Dim buList() As String = buParam.Split(","c)
            Dim strBUCodeFinal = "('" & String.Join("','", buList) & "')"
            Dim strBUcodeOriginal As String = strBUcode
            If InStr("SW,SCP,SCM,SBP,SBM", strBUcode) > 0 Then
                strBUcode = "SAM"
            End If
            If strBUcode <> "CC" Then
                strWhereCondition = $" WHERE BU_CODE IN {strBUCodeFinal} "
            End If
            strPathFileNameReport1 = strPathFile & $"Exception Report - {strBUcode}.xlsx"
            Dim strReport1SheetExc022 As String = "EXC_022"
            Dim strReport1SheetExc003 As String = "EXC_003"
            Dim strReport1Sheet1 As String = "EXC_011"
            Dim strReport1Sheet2 As String = "EXC_012"
            Dim strReport1SheetExc014 As String = "EXC_014"
            Dim strReport1SheetExc015 As String = "EXC_015"
            Dim strReport1SheetExc016 As String = "EXC_016"
            'copy template to new excel file 
            File.Copy(strPathTemplate1, strPathFileNameReport1, True)
            'open excel file
            Dim fileExcel As New FileInfo(strPathFileNameReport1)

            Using excelPackage As New ExcelPackage(fileExcel)
                objCmd = connOracle.CreateCommand
                '-----------------------Sheet1-----------------------
                Dim ws1 As ExcelWorksheet = excelPackage.Workbook.Worksheets(strReport1Sheet1)

                Debug.WriteLine("bufinal => " + strBUCodeFinal)

                objCmd.CommandText = $"SELECT     
                                         DATA_DATE,
                                         LINKAGE_NO,
                                         COLLATERAL_TYPE,
                                         COLLATERAL_TYPE_DESC,
                                         COLLATERAL_SUB_TYPE1,
                                         STOCK_CODE,
                                         APPRASIAL_VALUE,
                                         APPRASIAL_VALUE_CURRENCY,
                                         APPRASIAL_DATE,
                                         ISSUER_NAME,
                                         RM_CUST_NR,
                                         BOS_CUST_ID,
                                         CUST_NAME,
                                         BU_CODE,
                                         DIV_BC_CODE,
                                         DIV_BC_NAME,
                                         RO_ARM_AO_ALLOC_CODE,
                                         RO_ARM_AO_NAME,
                                         RM_AM_ALLOCA_CODE,
                                         RM_AM_NAME,
                                         M_TL_NAME
                                      FROM crrnew.t_rdt_exc_report_011
                                      WHERE linkage_no not in (select linkage_no from T_RDT_EXC_EXCLUDE_011_012) AND BU_CODE IN {strBUCodeFinal} AND DATA_DATE = '{endOfDayDateThai}' " & "ORDER BY DIV_BC_CODE,RO_ARM_AO_ALLOC_CODE,BOS_CUST_ID"

                objReader = objCmd.ExecuteReader()
                'Create a new DataTable.
                Dim dtEXC011 As DataTable = New DataTable("EXC011")
                'Load DataReader into the DataTable.
                dtEXC011.Load(objReader)
                'export to excel file
                Dim intNumRecordSheet1 As Integer = dtEXC011.Rows.Count
                If intNumRecordSheet1 > 0 Then
                    Dim TempRow1 As Integer = 3
                    Dim lastRow As Long


                    For i As Integer = 0 To intNumRecordSheet1 - 1
                        ws1.Cells("A" & i + TempRow1 & "").Value = dtEXC011.Rows(i)("DATA_DATE")
                        ws1.Cells("B" & i + TempRow1 & "").Value = dtEXC011.Rows(i)("LINKAGE_NO")
                        ws1.Cells("C" & i + TempRow1 & "").Value = dtEXC011.Rows(i)("COLLATERAL_TYPE")
                        ws1.Cells("D" & i + TempRow1 & "").Value = dtEXC011.Rows(i)("COLLATERAL_TYPE_DESC")
                        ws1.Cells("E" & i + TempRow1 & "").Value = dtEXC011.Rows(i)("COLLATERAL_SUB_TYPE1")
                        ws1.Cells("F" & i + TempRow1 & "").Value = dtEXC011.Rows(i)("STOCK_CODE")
                        ws1.Cells("G" & i + TempRow1 & "").Value = dtEXC011.Rows(i)("APPRASIAL_VALUE")
                        ws1.Cells("H" & i + TempRow1 & "").Value = dtEXC011.Rows(i)("APPRASIAL_VALUE_CURRENCY")
                        ws1.Cells("I" & i + TempRow1 & "").Value = dtEXC011.Rows(i)("APPRASIAL_DATE")
                        ws1.Cells("J" & i + TempRow1 & "").Value = dtEXC011.Rows(i)("ISSUER_NAME")
                        ws1.Cells("K" & i + TempRow1 & "").Value = dtEXC011.Rows(i)("RM_CUST_NR")
                        ws1.Cells("L" & i + TempRow1 & "").Value = dtEXC011.Rows(i)("BOS_CUST_ID")
                        ws1.Cells("M" & i + TempRow1 & "").Value = dtEXC011.Rows(i)("CUST_NAME")
                        ws1.Cells("N" & i + TempRow1 & "").Value = dtEXC011.Rows(i)("BU_CODE")
                        ws1.Cells("O" & i + TempRow1 & "").Value = dtEXC011.Rows(i)("DIV_BC_CODE")
                        ws1.Cells("P" & i + TempRow1 & "").Value = dtEXC011.Rows(i)("DIV_BC_NAME")
                        ws1.Cells("Q" & i + TempRow1 & "").Value = dtEXC011.Rows(i)("RO_ARM_AO_ALLOC_CODE")
                        ws1.Cells("R" & i + TempRow1 & "").Value = dtEXC011.Rows(i)("RO_ARM_AO_NAME")
                        ws1.Cells("S" & i + TempRow1 & "").Value = dtEXC011.Rows(i)("RM_AM_ALLOCA_CODE")
                        ws1.Cells("T" & i + TempRow1 & "").Value = dtEXC011.Rows(i)("RM_AM_NAME")
                        ws1.Cells("U" & i + TempRow1 & "").Value = dtEXC011.Rows(i)("M_TL_NAME")
                        'format font size'
                        lastRow = i + TempRow1
                        Dim cellRange As ExcelRange = ws1.Cells("A3:U" & lastRow)
                        cellRange.Style.Font.Size = 10



                    Next
                End If

                'Save ไฟล์
                excelPackage.Save()
                '-----------------------Sheet2-----------------------
                Dim ws2 As ExcelWorksheet = excelPackage.Workbook.Worksheets(strReport1Sheet2)
                objCmd.CommandText = $"SELECT     
                                         DATA_DATE,
                                         LINKAGE_NO,
                                         COLLATERAL_TYPE,
                                         COLLATERAL_TYPE_DESC,
                                         COLLATERAL_SUB_TYPE1,
                                         STOCK_CODE,
                                         APPRASIAL_VALUE,
                                         APPRASIAL_VALUE_CURRENCY,
                                         APPRASIAL_DATE,
                                         ISSUER_NAME,
                                         RM_CUST_NR,
                                         BOS_CUST_ID,
                                         CUST_NAME,
                                         BU_CODE,
                                         DIV_BC_CODE,
                                         DIV_BC_NAME,
                                         RO_ARM_AO_ALLOC_CODE,
                                         RO_ARM_AO_NAME,
                                         RM_AM_ALLOCA_CODE,
                                         RM_AM_NAME,
                                         M_TL_NAME
                                      FROM crrnew.t_rdt_exc_report_012
                                      WHERE linkage_no not in (select linkage_no from T_RDT_EXC_EXCLUDE_011_012) AND BU_CODE IN {strBUCodeFinal} AND DATA_DATE = '{endOfDayDateThai}' " & "ORDER BY DIV_BC_CODE,RO_ARM_AO_ALLOC_CODE,BOS_CUST_ID"
                objReader = objCmd.ExecuteReader()
                'Create data Table
                Dim dtEXC012 As DataTable = New DataTable("EXC012")
                dtEXC012.Load(objReader)
                'export to excel file
                Dim intNumRecordSheet2 As Integer = dtEXC012.Rows.Count
                If intNumRecordSheet2 > 0 Then
                    Dim TempRow2 As Integer = 3
                    Dim lastRow2 As Long


                    For i As Integer = 0 To intNumRecordSheet2 - 1
                        ws2.Cells("A" & i + TempRow2 & "").Value = dtEXC012.Rows(i)("DATA_DATE")
                        ws2.Cells("B" & i + TempRow2 & "").Value = dtEXC012.Rows(i)("LINKAGE_NO")
                        ws2.Cells("C" & i + TempRow2 & "").Value = dtEXC012.Rows(i)("COLLATERAL_TYPE")
                        ws2.Cells("D" & i + TempRow2 & "").Value = dtEXC012.Rows(i)("COLLATERAL_TYPE_DESC")
                        ws2.Cells("E" & i + TempRow2 & "").Value = dtEXC012.Rows(i)("COLLATERAL_SUB_TYPE1")
                        ws2.Cells("F" & i + TempRow2 & "").Value = dtEXC012.Rows(i)("STOCK_CODE")
                        ws2.Cells("G" & i + TempRow2 & "").Value = dtEXC012.Rows(i)("APPRASIAL_VALUE")
                        ws2.Cells("H" & i + TempRow2 & "").Value = dtEXC012.Rows(i)("APPRASIAL_VALUE_CURRENCY")
                        ws2.Cells("I" & i + TempRow2 & "").Value = dtEXC012.Rows(i)("APPRASIAL_DATE")
                        ws2.Cells("J" & i + TempRow2 & "").Value = dtEXC012.Rows(i)("ISSUER_NAME")
                        ws2.Cells("K" & i + TempRow2 & "").Value = dtEXC012.Rows(i)("RM_CUST_NR")
                        ws2.Cells("L" & i + TempRow2 & "").Value = dtEXC012.Rows(i)("BOS_CUST_ID")
                        ws2.Cells("M" & i + TempRow2 & "").Value = dtEXC012.Rows(i)("CUST_NAME")
                        ws2.Cells("N" & i + TempRow2 & "").Value = dtEXC012.Rows(i)("BU_CODE")
                        ws2.Cells("O" & i + TempRow2 & "").Value = dtEXC012.Rows(i)("DIV_BC_CODE")
                        ws2.Cells("P" & i + TempRow2 & "").Value = dtEXC012.Rows(i)("DIV_BC_NAME")
                        ws2.Cells("Q" & i + TempRow2 & "").Value = dtEXC012.Rows(i)("RO_ARM_AO_ALLOC_CODE")
                        ws2.Cells("R" & i + TempRow2 & "").Value = dtEXC012.Rows(i)("RO_ARM_AO_NAME")
                        ws2.Cells("S" & i + TempRow2 & "").Value = dtEXC012.Rows(i)("RM_AM_ALLOCA_CODE")
                        ws2.Cells("T" & i + TempRow2 & "").Value = dtEXC012.Rows(i)("RM_AM_NAME")
                        ws2.Cells("U" & i + TempRow2 & "").Value = dtEXC012.Rows(i)("M_TL_NAME")
                        'format font size'
                        lastRow2 = i + TempRow2
                        Dim cellRange As ExcelRange = ws2.Cells("A3:U" & lastRow2)
                        cellRange.Style.Font.Size = 10



                    Next

                End If

                'Save ไฟล์
                excelPackage.Save()
                '-----------------------SheetEXC_022-----------------------
                Dim wsExc022 As ExcelWorksheet = excelPackage.Workbook.Worksheets(strReport1SheetExc022)
                objCmd.CommandText = $"SELECT
                                    DATA_DATE,
                                    ACCT_NR,
                                    SYSTEM_NAME,
                                    CUST_NR,
                                    ACCT_TYPE,
                                    ACCT_PROD,
                                    BOT_PURPOSE_CODE,
                                    ISIC_CODE,
                                    RM_CUST_NR,
                                    BOS_CUST_ID,
                                    CUSTOMER_NAME,
                                    BU_CODE,
                                    DIV_BC_CODE,
                                    DIV_BC_NAME,
                                    RO_ARM_AO_ALLOC_CODE,
                                    RO_ARM_AO_NAME,
                                    RM_AM_ALLOC_CODE,
                                    RM_AM_NAME,
                                    M_TL_NAME
                                  FROM t_rdt_exc_report_022
                                  WHERE BU_CODE IN {strBUCodeFinal} AND DATA_DATE = '{endOfDayDateThai}' " & "ORDER BY DIV_BC_CODE,RO_ARM_AO_ALLOC_CODE,BOS_CUST_ID"
                objReader = objCmd.ExecuteReader()
                'Create Data Table'
                Dim dtEXC022 As DataTable = New DataTable("EXC022")
                dtEXC022.Load(objReader)
                'Export To Excel File'
                Dim intNumRecordSheetExc022 As Integer = dtEXC022.Rows.Count
                If intNumRecordSheetExc022 > 0 Then
                    Dim TempRow022 As Integer = 3
                    Dim lastRow022 As Long
                    For i As Integer = 0 To intNumRecordSheetExc022 - 1
                        wsExc022.Cells("A" & i + TempRow022 & "").Value = dtEXC022.Rows(i)("DATA_DATE")
                        wsExc022.Cells("B" & i + TempRow022 & "").Value = dtEXC022.Rows(i)("ACCT_NR")
                        wsExc022.Cells("C" & i + TempRow022 & "").Value = dtEXC022.Rows(i)("SYSTEM_NAME")
                        wsExc022.Cells("D" & i + TempRow022 & "").Value = dtEXC022.Rows(i)("CUST_NR")
                        wsExc022.Cells("E" & i + TempRow022 & "").Value = dtEXC022.Rows(i)("ACCT_TYPE")
                        wsExc022.Cells("F" & i + TempRow022 & "").Value = dtEXC022.Rows(i)("ACCT_PROD")
                        wsExc022.Cells("G" & i + TempRow022 & "").Value = dtEXC022.Rows(i)("BOT_PURPOSE_CODE")
                        wsExc022.Cells("H" & i + TempRow022 & "").Value = dtEXC022.Rows(i)("ISIC_CODE")
                        wsExc022.Cells("I" & i + TempRow022 & "").Value = dtEXC022.Rows(i)("RM_CUST_NR")
                        wsExc022.Cells("J" & i + TempRow022 & "").Value = dtEXC022.Rows(i)("BOS_CUST_ID")
                        wsExc022.Cells("K" & i + TempRow022 & "").Value = dtEXC022.Rows(i)("CUSTOMER_NAME")
                        wsExc022.Cells("L" & i + TempRow022 & "").Value = dtEXC022.Rows(i)("BU_CODE")
                        wsExc022.Cells("M" & i + TempRow022 & "").Value = dtEXC022.Rows(i)("DIV_BC_CODE")
                        wsExc022.Cells("N" & i + TempRow022 & "").Value = dtEXC022.Rows(i)("DIV_BC_NAME")
                        wsExc022.Cells("O" & i + TempRow022 & "").Value = dtEXC022.Rows(i)("RO_ARM_AO_ALLOC_CODE")
                        wsExc022.Cells("P" & i + TempRow022 & "").Value = dtEXC022.Rows(i)("RO_ARM_AO_NAME")
                        wsExc022.Cells("Q" & i + TempRow022 & "").Value = dtEXC022.Rows(i)("RM_AM_ALLOC_CODE")
                        wsExc022.Cells("R" & i + TempRow022 & "").Value = dtEXC022.Rows(i)("RM_AM_NAME")
                        wsExc022.Cells("S" & i + TempRow022 & "").Value = dtEXC022.Rows(i)("M_TL_NAME")
                        'format font size'
                        lastRow022 = i + TempRow022
                        Dim cellRange As ExcelRange = wsExc022.Cells("A3:U" & lastRow022)
                        cellRange.Style.Font.Size = 10
                    Next

                End If
                'Save File'
                excelPackage.Save()
                '-----------------------SheetEXC_003-----------------------
                Dim wsExc003 As ExcelWorksheet = excelPackage.Workbook.Worksheets(strReport1SheetExc003)
                objCmd.CommandText = $"SELECT 
                                  DATA_DATE,
                                  ACCT_NR,
                                  SYSTEM_NAME,
                                  CUST_NR,
                                  ACCT_TYPE,
                                  ACCT_PROD,
                                  BOT_PURPOSE_CODE,
                                  ISIC_CODE,
                                  RM_CUST_NR,
                                  BOS_CUST_ID,
                                  CUSTOMER_NAME,
                                  BU_CODE,
                                  DIV_BC_CODE,
                                  DIV_BC_NAME,
                                  RO_ARM_AO_ALLOC_CODE,
                                  RO_ARM_AO_NAME,
                                  RM_AM_ALLOC_CODE,
                                  RM_AM_NAME,
                                  M_TL_NAME  
                                  FROM t_rdt_exc_report_003
                                  WHERE BU_CODE IN {strBUCodeFinal} AND DATA_DATE = '{endOfDayDateThai}' " & "ORDER BY DIV_BC_CODE,RO_ARM_AO_ALLOC_CODE,BOS_CUST_ID"
                objReader = objCmd.ExecuteReader()
                'Create Data Table'
                Dim dtEXC003 As DataTable = New DataTable("EXC003")
                dtEXC003.Load(objReader)
                'Export TO Excel File'
                Dim intNumRecordSheetExc003 As Integer = dtEXC003.Rows.Count
                If intNumRecordSheetExc003 > 0 Then
                    Dim TempRow003 As Integer = 3
                    Dim lastRow003 As Long
                    For i As Integer = 0 To intNumRecordSheetExc003 - 1
                        wsExc003.Cells("A" & i + TempRow003 & "").Value = dtEXC003.Rows(i)("DATA_DATE")
                        wsExc003.Cells("B" & i + TempRow003 & "").Value = dtEXC003.Rows(i)("ACCT_NR")
                        wsExc003.Cells("C" & i + TempRow003 & "").Value = dtEXC003.Rows(i)("SYSTEM_NAME")
                        wsExc003.Cells("D" & i + TempRow003 & "").Value = dtEXC003.Rows(i)("CUST_NR")
                        wsExc003.Cells("E" & i + TempRow003 & "").Value = dtEXC003.Rows(i)("ACCT_TYPE")
                        wsExc003.Cells("F" & i + TempRow003 & "").Value = dtEXC003.Rows(i)("ACCT_PROD")
                        wsExc003.Cells("G" & i + TempRow003 & "").Value = dtEXC003.Rows(i)("BOT_PURPOSE_CODE")
                        wsExc003.Cells("H" & i + TempRow003 & "").Value = dtEXC003.Rows(i)("ISIC_CODE")
                        wsExc003.Cells("I" & i + TempRow003 & "").Value = dtEXC003.Rows(i)("RM_CUST_NR")
                        wsExc003.Cells("J" & i + TempRow003 & "").Value = dtEXC003.Rows(i)("BOS_CUST_ID")
                        wsExc003.Cells("K" & i + TempRow003 & "").Value = dtEXC003.Rows(i)("CUSTOMER_NAME")
                        wsExc003.Cells("L" & i + TempRow003 & "").Value = dtEXC003.Rows(i)("BU_CODE")
                        wsExc003.Cells("M" & i + TempRow003 & "").Value = dtEXC003.Rows(i)("DIV_BC_CODE")
                        wsExc003.Cells("N" & i + TempRow003 & "").Value = dtEXC003.Rows(i)("DIV_BC_NAME")
                        wsExc003.Cells("O" & i + TempRow003 & "").Value = dtEXC003.Rows(i)("RO_ARM_AO_ALLOC_CODE")
                        wsExc003.Cells("P" & i + TempRow003 & "").Value = dtEXC003.Rows(i)("RO_ARM_AO_NAME")
                        wsExc003.Cells("Q" & i + TempRow003 & "").Value = dtEXC003.Rows(i)("RM_AM_ALLOC_CODE")
                        wsExc003.Cells("R" & i + TempRow003 & "").Value = dtEXC003.Rows(i)("RM_AM_NAME")
                        wsExc003.Cells("S" & i + TempRow003 & "").Value = dtEXC003.Rows(i)("M_TL_NAME")
                        'format font size'
                        lastRow003 = i + TempRow003
                        Dim cellRange As ExcelRange = wsExc003.Cells("A3:U" & lastRow003)
                        cellRange.Style.Font.Size = 10
                    Next
                End If
                'Save File'
                excelPackage.Save()
                '-----------------------SheetEXC_014-----------------------
                Dim wsExc014 As ExcelWorksheet = excelPackage.Workbook.Worksheets(strReport1SheetExc014)
                If strBUcode <> "CC" Then
                    queryExc014 = $"SELECT
                                  DATA_DATE,
                                  ISIC_CODE,
                                  TOTAL_INCOME_BAHT,
                                  DATE_OF_TOTAL_INCOME_BAHT,
                                  DOMESTIC_INCOME_BAHT,
                                  EXPORT_INCOME_BAHT,
                                  LABOR,
                                  DATE_OF_LABOR,
                                  GROUP_DATA,
                                  RM_CUST_NR,
                                  BOS_CUST_ID,
                                  CUST_NAME,
                                  BU_CODE,
                                  DIV_BC_CODE,
                                  DIV_BC_NAME,
                                  RO_ARM_AO_ALLOC_CODE,
                                  RO_ARM_AO_NAME,
                                  RM_AM_ALLOCA_CODE,
                                  RM_AM_NAME,
                                  M_TL_NAME
                                  FROM t_rdt_exc_report_014
                                  " & strWhereCondition & $"And GROUP_DATA = 'BU' AND DATA_DATE = '{endOfDayDateThai}' " & "ORDER BY DIV_BC_CODE,RO_ARM_AO_ALLOC_CODE,BOS_CUST_ID"
                End If
                objCmd.CommandText = queryExc014

                objReader = objCmd.ExecuteReader()
                'Create Data Table'
                Dim dtEXC014 As DataTable = New DataTable("EXC014")
                dtEXC014.Load(objReader)
                'Export To Excel File'
                Dim intNumRecordSheetExc014 As Integer = dtEXC014.Rows.Count
                If intNumRecordSheetExc014 > 0 Then
                    Dim TempRow014 As Integer = 3
                    Dim lastRow014 As Long
                    For i As Integer = 0 To intNumRecordSheetExc014 - 1
                        wsExc014.Cells("A" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("DATA_DATE")
                        wsExc014.Cells("B" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("RM_CUST_NR")
                        wsExc014.Cells("C" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("BOS_CUST_ID")
                        wsExc014.Cells("D" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("CUST_NAME")
                        wsExc014.Cells("E" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("ISIC_CODE")
                        wsExc014.Cells("F" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("TOTAL_INCOME_BAHT")
                        wsExc014.Cells("G" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("DOMESTIC_INCOME_BAHT")
                        wsExc014.Cells("H" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("EXPORT_INCOME_BAHT")
                        wsExc014.Cells("I" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("DATE_OF_TOTAL_INCOME_BAHT")
                        wsExc014.Cells("J" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("LABOR")
                        wsExc014.Cells("K" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("DATE_OF_LABOR")
                        wsExc014.Cells("L" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("BU_CODE")
                        wsExc014.Cells("M" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("DIV_BC_CODE")
                        wsExc014.Cells("N" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("DIV_BC_NAME")
                        wsExc014.Cells("O" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("RO_ARM_AO_ALLOC_CODE")
                        wsExc014.Cells("P" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("RO_ARM_AO_NAME")
                        wsExc014.Cells("Q" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("RM_AM_ALLOCA_CODE")
                        wsExc014.Cells("R" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("RM_AM_NAME")
                        wsExc014.Cells("S" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("M_TL_NAME")
                        'format font size'
                        lastRow014 = i + TempRow014
                        Dim cellRange As ExcelRange = wsExc014.Cells("A3:U" & lastRow014)
                        cellRange.Style.Font.Size = 10
                    Next

                End If
                excelPackage.Save()

                '-----------------------SheetEXC_015-----------------------
                Dim wsExc015 As ExcelWorksheet = excelPackage.Workbook.Worksheets(strReport1SheetExc015)
                If strBUcode <> "CC" Then
                    queryExc015 = $"SELECT
                                    DATA_DATE,
                                    ISIC_CODE,
                                    TOTAL_INCOME_BAHT,
                                    DATE_OF_TOTAL_INCOME_BAHT,
                                    DOMESTIC_INCOME_BAHT,
                                    EXPORT_INCOME_BAHT,
                                    LABOR,
                                    DATE_OF_LABOR,
                                    GROUP_DATA,
                                    RM_CUST_NR,
                                    BOS_CUST_ID,
                                    CUST_NAME,
                                    BU_CODE,
                                    DIV_BC_CODE,
                                    DIV_BC_NAME,
                                    RO_ARM_AO_ALLOC_CODE,
                                    RO_ARM_AO_NAME,
                                    RM_AM_ALLOCA_CODE,
                                    RM_AM_NAME,
                                    M_TL_NAME
                                  FROM t_rdt_exc_report_015
                                  " & strWhereCondition & $"And GROUP_DATA = 'BU' AND DATA_DATE = '{endOfDayDateThai}' " & "ORDER BY DIV_BC_CODE,RO_ARM_AO_ALLOC_CODE,BOS_CUST_ID"
                End If
                objCmd.CommandText = queryExc015
                objReader = objCmd.ExecuteReader()
                'Create Data Table'
                Dim dtEXC015 As DataTable = New DataTable("EXC015")
                dtEXC015.Load(objReader)
                'Export To Excel File'
                Dim intNumRecordSheetExc015 As Integer = dtEXC015.Rows.Count
                If intNumRecordSheetExc015 > 0 Then
                    Dim TempRow015 As Integer = 3
                    Dim lastRow015 As Long
                    For i As Integer = 0 To intNumRecordSheetExc015 - 1
                        wsExc015.Cells("A" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("DATA_DATE")
                        wsExc015.Cells("B" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("RM_CUST_NR")
                        wsExc015.Cells("C" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("BOS_CUST_ID")
                        wsExc015.Cells("D" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("CUST_NAME")
                        wsExc015.Cells("E" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("ISIC_CODE")
                        wsExc015.Cells("F" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("TOTAL_INCOME_BAHT")
                        wsExc015.Cells("G" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("DOMESTIC_INCOME_BAHT")
                        wsExc015.Cells("H" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("EXPORT_INCOME_BAHT")
                        wsExc015.Cells("I" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("DATE_OF_TOTAL_INCOME_BAHT")
                        wsExc015.Cells("J" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("LABOR")
                        wsExc015.Cells("K" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("DATE_OF_LABOR")
                        wsExc015.Cells("L" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("BU_CODE")
                        wsExc015.Cells("M" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("DIV_BC_CODE")
                        wsExc015.Cells("N" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("DIV_BC_NAME")
                        wsExc015.Cells("O" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("RO_ARM_AO_ALLOC_CODE")
                        wsExc015.Cells("P" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("RO_ARM_AO_NAME")
                        wsExc015.Cells("Q" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("RM_AM_ALLOCA_CODE")
                        wsExc015.Cells("R" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("RM_AM_NAME")
                        wsExc015.Cells("S" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("M_TL_NAME")
                        'format font size'
                        lastRow015 = i + TempRow015
                        Dim cellRange As ExcelRange = wsExc015.Cells("A3:U" & lastRow015)
                        cellRange.Style.Font.Size = 10
                    Next

                End If
                excelPackage.Save()
                '-----------------------SheetEXC_016-----------------------
                Dim wsExc016 As ExcelWorksheet = excelPackage.Workbook.Worksheets(strReport1SheetExc016)
                If strBUcode <> "CC" Then
                    queryExc016 = $"SELECT
                                    DATA_DATE,
                                    ISIC_CODE,
                                    TOTAL_INCOME_BAHT,
                                    DATE_OF_TOTAL_INCOME_BAHT,
                                    DOMESTIC_INCOME_BAHT,
                                    EXPORT_INCOME_BAHT,
                                    LABOR,
                                    DATE_OF_LABOR,
                                    GROUP_DATA,
                                    RM_CUST_NR,
                                    BOS_CUST_ID,
                                    CUST_NAME,
                                    BU_CODE,
                                    DIV_BC_CODE,
                                    DIV_BC_NAME,
                                    RO_ARM_AO_ALLOC_CODE,
                                    RO_ARM_AO_NAME,
                                    RM_AM_ALLOCA_CODE,
                                    RM_AM_NAME,
                                    M_TL_NAME
                                  FROM t_rdt_exc_report_016
                                  " & strWhereCondition & $"And GROUP_DATA = 'BU' AND DATA_DATE = '{endOfDayDateThai}' " & "ORDER BY DIV_BC_CODE,RO_ARM_AO_ALLOC_CODE,BOS_CUST_ID"
                End If
                objCmd.CommandText = queryExc016

                objReader = objCmd.ExecuteReader()
                'Create Data Table'
                Dim dtEXC016 As DataTable = New DataTable("EXC016")
                dtEXC016.Load(objReader)
                'Export To Excel File'
                Dim intNumRecordSheetExc016 As Integer = dtEXC016.Rows.Count
                If intNumRecordSheetExc016 > 0 Then
                    Dim TempRow016 As Integer = 3
                    Dim lastRow016 As Long
                    For i As Integer = 0 To intNumRecordSheetExc016 - 1
                        wsExc016.Cells("A" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("DATA_DATE")
                        wsExc016.Cells("B" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("RM_CUST_NR")
                        wsExc016.Cells("C" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("BOS_CUST_ID")
                        wsExc016.Cells("D" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("CUST_NAME")
                        wsExc016.Cells("E" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("ISIC_CODE")
                        wsExc016.Cells("F" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("TOTAL_INCOME_BAHT")
                        wsExc016.Cells("G" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("DOMESTIC_INCOME_BAHT")
                        wsExc016.Cells("H" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("EXPORT_INCOME_BAHT")
                        wsExc016.Cells("I" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("DATE_OF_TOTAL_INCOME_BAHT")
                        wsExc016.Cells("J" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("LABOR")
                        wsExc016.Cells("K" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("DATE_OF_LABOR")
                        wsExc016.Cells("L" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("BU_CODE")
                        wsExc016.Cells("M" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("DIV_BC_CODE")
                        wsExc016.Cells("N" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("DIV_BC_NAME")
                        wsExc016.Cells("O" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("RO_ARM_AO_ALLOC_CODE")
                        wsExc016.Cells("P" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("RO_ARM_AO_NAME")
                        wsExc016.Cells("Q" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("RM_AM_ALLOCA_CODE")
                        wsExc016.Cells("R" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("RM_AM_NAME")
                        wsExc016.Cells("S" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("M_TL_NAME")
                        'format font size'
                        lastRow016 = i + TempRow016
                        Dim cellRange As ExcelRange = wsExc016.Cells("A3:U" & lastRow016)
                        cellRange.Style.Font.Size = 10
                    Next

                End If
                'excelPackage.Workbook.Protection.SetPassword("ABC")
                'excelPackage.Encryption.IsEncrypted = True
                'excelPackage.SaveAs(fileExcel, "ABC")
                excelPackage.Save()
            End Using

        Else
            Debug.WriteLine("get in here?")
            Debug.WriteLine("bufinal => " + strBUcode)
            Dim strWhereCondition As String = ""

            'Template ที่จะใช้
            Dim strPathTemplate1 As String = ConfigurationManager.AppSettings("strPathTemplate2")
            Dim buParam As String = strBUcode
            Dim buList() As String = buParam.Split(","c)
            Dim strBUCodeFinal = "('" & String.Join("','", buList) & "')"
            Dim strBUcodeOriginal As String = strBUcode
            If InStr("SW,SCP,SCM,SBP,SBM", strBUcode) > 0 Then
                strBUcode = "SAM"
            End If
            If strBUcode = "CC" Then
                strWhereCondition = $" WHERE Group_Data = 'CC' AND DATA_DATE = '{endOfDayDateThai}'"
            End If
            strPathFileNameReport1 = strPathFile & $"Exception Report - {strBUcode}.xlsx"
            Dim strReport1SheetExc014 As String = "EXC_014"
            Dim strReport1SheetExc015 As String = "EXC_015"
            Dim strReport1SheetExc016 As String = "EXC_016"


            'copy template to new excel file 
            File.Copy(strPathTemplate1, strPathFileNameReport1, True)
            'open excel file
            Dim fileExcel As New FileInfo(strPathFileNameReport1)
            Using excelPackage As New ExcelPackage(fileExcel)
                objCmd = connOracle.CreateCommand

                '-----------------------SheetEXC_014-----------------------
                Dim wsExc014 As ExcelWorksheet = excelPackage.Workbook.Worksheets(strReport1SheetExc014)
                If strBUcode = "CC" Then
                    queryExc014 = $"SELECT
                                  DATA_DATE,
                                  ISIC_CODE,
                                  TOTAL_INCOME_BAHT,
                                  DATE_OF_TOTAL_INCOME_BAHT,
                                  DOMESTIC_INCOME_BAHT,
                                  EXPORT_INCOME_BAHT,
                                  LABOR,
                                  DATE_OF_LABOR,
                                  GROUP_DATA,
                                  RM_CUST_NR,
                                  BOS_CUST_ID,
                                  CUST_NAME,
                                  BU_CODE,
                                  DIV_BC_CODE,
                                  DIV_BC_NAME,
                                  RO_ARM_AO_ALLOC_CODE,
                                  RO_ARM_AO_NAME,
                                  RM_AM_ALLOCA_CODE,
                                  RM_AM_NAME,
                                  M_TL_NAME
                                  FROM t_rdt_exc_report_014
                                  " & strWhereCondition & "ORDER BY DIV_BC_CODE,RO_ARM_AO_ALLOC_CODE,BOS_CUST_ID"

                End If
                objCmd.CommandText = queryExc014

                objReader = objCmd.ExecuteReader()
                'Create Data Table'
                Dim dtEXC014 As DataTable = New DataTable("EXC014")
                dtEXC014.Load(objReader)
                'Export To Excel File'
                Dim intNumRecordSheetExc014 As Integer = dtEXC014.Rows.Count
                If intNumRecordSheetExc014 > 0 Then
                    Dim TempRow014 As Integer = 3
                    Dim lastRow014 As Long
                    For i As Integer = 0 To intNumRecordSheetExc014 - 1
                        wsExc014.Cells("A" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("DATA_DATE")
                        wsExc014.Cells("B" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("RM_CUST_NR")
                        wsExc014.Cells("C" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("BOS_CUST_ID")
                        wsExc014.Cells("D" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("CUST_NAME")
                        wsExc014.Cells("E" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("ISIC_CODE")
                        wsExc014.Cells("F" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("TOTAL_INCOME_BAHT")
                        wsExc014.Cells("G" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("DOMESTIC_INCOME_BAHT")
                        wsExc014.Cells("H" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("EXPORT_INCOME_BAHT")
                        wsExc014.Cells("I" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("DATE_OF_TOTAL_INCOME_BAHT")
                        wsExc014.Cells("J" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("LABOR")
                        wsExc014.Cells("K" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("DATE_OF_LABOR")
                        wsExc014.Cells("L" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("BU_CODE")
                        wsExc014.Cells("M" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("DIV_BC_CODE")
                        wsExc014.Cells("N" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("DIV_BC_NAME")
                        wsExc014.Cells("O" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("RO_ARM_AO_ALLOC_CODE")
                        wsExc014.Cells("P" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("RO_ARM_AO_NAME")
                        wsExc014.Cells("Q" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("RM_AM_ALLOCA_CODE")
                        wsExc014.Cells("R" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("RM_AM_NAME")
                        wsExc014.Cells("S" & i + TempRow014 & "").Value = dtEXC014.Rows(i)("M_TL_NAME")
                        'format font size'
                        lastRow014 = i + TempRow014
                        Dim cellRange As ExcelRange = wsExc014.Cells("A3:U" & lastRow014)
                        cellRange.Style.Font.Size = 10
                    Next

                End If
                excelPackage.Save()

                '-----------------------SheetEXC_015-----------------------
                Dim wsExc015 As ExcelWorksheet = excelPackage.Workbook.Worksheets(strReport1SheetExc015)
                If strBUcode = "CC" Then
                    queryExc015 = $"SELECT
                                    DATA_DATE,
                                    ISIC_CODE,
                                    TOTAL_INCOME_BAHT,
                                    DATE_OF_TOTAL_INCOME_BAHT,
                                    DOMESTIC_INCOME_BAHT,
                                    EXPORT_INCOME_BAHT,
                                    LABOR,
                                    DATE_OF_LABOR,
                                    GROUP_DATA,
                                    RM_CUST_NR,
                                    BOS_CUST_ID,
                                    CUST_NAME,
                                    BU_CODE,
                                    DIV_BC_CODE,
                                    DIV_BC_NAME,
                                    RO_ARM_AO_ALLOC_CODE,
                                    RO_ARM_AO_NAME,
                                    RM_AM_ALLOCA_CODE,
                                    RM_AM_NAME,
                                    M_TL_NAME
                                  FROM t_rdt_exc_report_015
                                  " & strWhereCondition & "ORDER BY DIV_BC_CODE,RO_ARM_AO_ALLOC_CODE,BOS_CUST_ID"

                End If
                objCmd.CommandText = queryExc015
                objReader = objCmd.ExecuteReader()
                'Create Data Table'
                Dim dtEXC015 As DataTable = New DataTable("EXC015")
                dtEXC015.Load(objReader)
                'Export To Excel File'
                Dim intNumRecordSheetExc015 As Integer = dtEXC015.Rows.Count
                If intNumRecordSheetExc015 > 0 Then
                    Dim TempRow015 As Integer = 3
                    Dim lastRow015 As Long
                    For i As Integer = 0 To intNumRecordSheetExc015 - 1
                        wsExc015.Cells("A" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("DATA_DATE")
                        wsExc015.Cells("B" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("RM_CUST_NR")
                        wsExc015.Cells("C" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("BOS_CUST_ID")
                        wsExc015.Cells("D" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("CUST_NAME")
                        wsExc015.Cells("E" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("ISIC_CODE")
                        wsExc015.Cells("F" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("TOTAL_INCOME_BAHT")
                        wsExc015.Cells("G" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("DOMESTIC_INCOME_BAHT")
                        wsExc015.Cells("H" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("EXPORT_INCOME_BAHT")
                        wsExc015.Cells("I" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("DATE_OF_TOTAL_INCOME_BAHT")
                        wsExc015.Cells("J" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("LABOR")
                        wsExc015.Cells("K" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("DATE_OF_LABOR")
                        wsExc015.Cells("L" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("BU_CODE")
                        wsExc015.Cells("M" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("DIV_BC_CODE")
                        wsExc015.Cells("N" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("DIV_BC_NAME")
                        wsExc015.Cells("O" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("RO_ARM_AO_ALLOC_CODE")
                        wsExc015.Cells("P" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("RO_ARM_AO_NAME")
                        wsExc015.Cells("Q" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("RM_AM_ALLOCA_CODE")
                        wsExc015.Cells("R" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("RM_AM_NAME")
                        wsExc015.Cells("S" & i + TempRow015 & "").Value = dtEXC015.Rows(i)("M_TL_NAME")
                        'format font size'
                        lastRow015 = i + TempRow015
                        Dim cellRange As ExcelRange = wsExc015.Cells("A3:U" & lastRow015)
                        cellRange.Style.Font.Size = 10
                    Next

                End If
                excelPackage.Save()
                '-----------------------SheetEXC_016-----------------------
                Dim wsExc016 As ExcelWorksheet = excelPackage.Workbook.Worksheets(strReport1SheetExc016)
                If strBUcode = "CC" Then
                    queryExc016 = $"SELECT
                                    DATA_DATE,
                                    ISIC_CODE,
                                    TOTAL_INCOME_BAHT,
                                    DATE_OF_TOTAL_INCOME_BAHT,
                                    DOMESTIC_INCOME_BAHT,
                                    EXPORT_INCOME_BAHT,
                                    LABOR,
                                    DATE_OF_LABOR,
                                    GROUP_DATA,
                                    RM_CUST_NR,
                                    BOS_CUST_ID,
                                    CUST_NAME,
                                    BU_CODE,
                                    DIV_BC_CODE,
                                    DIV_BC_NAME,
                                    RO_ARM_AO_ALLOC_CODE,
                                    RO_ARM_AO_NAME,
                                    RM_AM_ALLOCA_CODE,
                                    RM_AM_NAME,
                                    M_TL_NAME
                                  FROM t_rdt_exc_report_016
                                  " & strWhereCondition & "ORDER BY DIV_BC_CODE,RO_ARM_AO_ALLOC_CODE,BOS_CUST_ID"

                End If
                objCmd.CommandText = queryExc016

                objReader = objCmd.ExecuteReader()
                'Create Data Table'
                Dim dtEXC016 As DataTable = New DataTable("EXC016")
                dtEXC016.Load(objReader)
                'Export To Excel File'
                Dim intNumRecordSheetExc016 As Integer = dtEXC016.Rows.Count
                If intNumRecordSheetExc016 > 0 Then
                    Dim TempRow016 As Integer = 3
                    Dim lastRow016 As Long
                    For i As Integer = 0 To intNumRecordSheetExc016 - 1
                        wsExc016.Cells("A" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("DATA_DATE")
                        wsExc016.Cells("B" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("RM_CUST_NR")
                        wsExc016.Cells("C" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("BOS_CUST_ID")
                        wsExc016.Cells("D" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("CUST_NAME")
                        wsExc016.Cells("E" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("ISIC_CODE")
                        wsExc016.Cells("F" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("TOTAL_INCOME_BAHT")
                        wsExc016.Cells("G" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("DOMESTIC_INCOME_BAHT")
                        wsExc016.Cells("H" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("EXPORT_INCOME_BAHT")
                        wsExc016.Cells("I" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("DATE_OF_TOTAL_INCOME_BAHT")
                        wsExc016.Cells("J" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("LABOR")
                        wsExc016.Cells("K" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("DATE_OF_LABOR")
                        wsExc016.Cells("L" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("BU_CODE")
                        wsExc016.Cells("M" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("DIV_BC_CODE")
                        wsExc016.Cells("N" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("DIV_BC_NAME")
                        wsExc016.Cells("O" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("RO_ARM_AO_ALLOC_CODE")
                        wsExc016.Cells("P" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("RO_ARM_AO_NAME")
                        wsExc016.Cells("Q" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("RM_AM_ALLOCA_CODE")
                        wsExc016.Cells("R" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("RM_AM_NAME")
                        wsExc016.Cells("S" & i + TempRow016 & "").Value = dtEXC016.Rows(i)("M_TL_NAME")
                        'format font size'
                        lastRow016 = i + TempRow016
                        Dim cellRange As ExcelRange = wsExc016.Cells("A3:U" & lastRow016)
                        cellRange.Style.Font.Size = 10
                    Next

                End If
                'excelPackage.Workbook.Protection.SetPassword("ABC")
                'excelPackage.Encryption.IsEncrypted = True
                'excelPackage.SaveAs(fileExcel, "ABC")
                excelPackage.Save()
            End Using
        End If

    End Sub


    Sub SendMail(ByVal strBUcode As String, ByVal strMailTo As String, ByVal strMailCc As String, ByVal endOfDayDateThai As String)
        Dim insMail As New MailMessage
        Dim ClientMail As New SmtpClient
        Dim attachment1 As Attachment
        Dim attachment2 As Attachment
        Dim attachment3 As Attachment
        Dim buThaiName As String
        Dim buParam As String = strBUcode
        Dim buList() As String = buParam.Split(","c)
        Dim strBUCodeFinal = "('" & String.Join("','", buList) & "')"

        'lack of NBM (waiting for confirm data to insert it)
        If strBUcode = "NWM" Then
            buThaiName = "สายธุรกิจลูกค้ารายใหญ่"
        ElseIf strBUcode = "NCM" Then
            buThaiName = "สายธุรกิจลูกค้ารายกลาง นคล."
        ElseIf strBUcode = "NCP" Then
            buThaiName = "สายธุรกิจลูกค้ารายกลาง ตจว."
        ElseIf strBUcode = "NBP" Then
            buThaiName = "สายธุรกิจลูกค้ารายปลีก ตจว."
        ElseIf strBUcode = "CC" Then
            buThaiName = "สายบัตรเครดิต"
        ElseIf strBUcode = "NBM" Then
            buThaiName = "สายธุรกิจลูกค้ารายปลีก นคล."
        Else
            buThaiName = "สายบริหารสินเชื่อพิเศษ"
        End If

        ClientMail.Host = Trim(ConfigurationManager.AppSettings("mailstmp"))
        insMail.From = New MailAddress(Trim(ConfigurationManager.AppSettings("mailfrom")), Trim(ConfigurationManager.AppSettings("mailfromName")))
        Dim getDataDateString As String = "SELECT  DATA_DATE FROM t_rdt_exc_report_011 WHERE ROWNUM = 1"
        objCmd = connOracle.CreateCommand
        objCmd.CommandText = getDataDateString
        objReader = objCmd.ExecuteReader()
        If objReader.Read() Then
            If Not IsDBNull(objReader("DATA_DATE")) Then
                Dim dataDateStr As String = objReader("DATA_DATE").ToString()


                If DateTime.TryParse(dataDateStr, dataDate) Then
                    ' Subtract one day
                    dataDate = dataDate.AddDays(-1)
                    Dim yearInAD As Integer = dataDate.Year

                    ' Format back to "dd/MM/yyyy"
                    formattedDataDate = dataDate.ToString("dd/MM/") & yearInAD.ToString()

                    ' Output the result
                    Debug.WriteLine("The end of day date is " & formattedDataDate)
                Else
                    ' Handle the case where the date string is not in a valid format
                    Debug.WriteLine("The date string is not in a valid format.")
                End If


            Else
                Debug.WriteLine("DATA_DATE is null")
                dataDate = ""
            End If

        End If

        Dim strListMail As String = ""
        Dim strSubject As String = "RDT Exception Report – " & buThaiName & " as at " & endOfDayDateThai '& strDateMonth & "/" & strDateYear '12/2024 MM/YYYY
        Dim strExc11 As String = $"SELECT  COUNT(*) AS TOTAL_ROW FROM crrnew.t_rdt_exc_report_011 WHERE BU_CODE IN {strBUCodeFinal} and linkage_no not in (select linkage_no from T_RDT_EXC_EXCLUDE_011_012)"

        objCmd = connOracle.CreateCommand
        objCmd.CommandText = strExc11
        objReader = objCmd.ExecuteReader()
        Dim dtTotalRowExc11 As Integer = 0
        If objReader.Read() Then
            dtTotalRowExc11 = Convert.ToInt32(objReader("TOTAL_ROW"))
        End If

        Dim strExc12 As String = $"SELECT  COUNT(*) AS TOTAL_ROW FROM crrnew.t_rdt_exc_report_012 WHERE BU_CODE IN {strBUCodeFinal} and linkage_no not in (select linkage_no from T_RDT_EXC_EXCLUDE_011_012)"
        objCmd = connOracle.CreateCommand
        objCmd.CommandText = strExc12
        objReader = objCmd.ExecuteReader()
        Dim dtTotalRowExc12 As Integer = 0
        If objReader.Read() Then
            dtTotalRowExc12 = Convert.ToInt32(objReader("TOTAL_ROW"))
        End If

        If strBUcode = "CC" Then
            strExc14 = $"SELECT  COUNT(*) AS TOTAL_ROW FROM crrnew.t_rdt_exc_report_014  WHERE Group_Data = 'CC'"
            strExc15 = $"SELECT  COUNT(*) AS TOTAL_ROW FROM crrnew.t_rdt_exc_report_015  WHERE Group_Data = 'CC'"
            strExc16 = $"SELECT  COUNT(*) AS TOTAL_ROW FROM crrnew.t_rdt_exc_report_016  WHERE Group_Data = 'CC'"
        Else
            strExc14 = $"SELECT  COUNT(*) AS TOTAL_ROW FROM crrnew.t_rdt_exc_report_014  WHERE BU_CODE IN {strBUCodeFinal} And GROUP_DATA = 'BU' "
            strExc15 = $"SELECT  COUNT(*) AS TOTAL_ROW FROM crrnew.t_rdt_exc_report_015  WHERE BU_CODE IN {strBUCodeFinal} And GROUP_DATA = 'BU' "
            strExc16 = $"SELECT  COUNT(*) AS TOTAL_ROW FROM crrnew.t_rdt_exc_report_016  WHERE BU_CODE IN {strBUCodeFinal} And GROUP_DATA = 'BU' "
        End If

        objCmd = connOracle.CreateCommand
        objCmd.CommandText = strExc14
        objReader = objCmd.ExecuteReader()
        Dim dtTotalRowExc14 As Integer = 0
        If objReader.Read() Then
            dtTotalRowExc14 = Convert.ToInt32(objReader("TOTAL_ROW"))
        End If

        'Dim strExc15 As String = $"SELECT COUNT(*) AS TOTAL_ROW FROM crrnew.t_rdt_exc_report_015 WHERE BU_CODE IN {strBUCodeFinal}"
        objCmd = connOracle.CreateCommand
        objCmd.CommandText = strExc15
        objReader = objCmd.ExecuteReader()
        Dim dtTotalRowExc15 As Integer = 0
        If objReader.Read() Then
            dtTotalRowExc15 = Convert.ToInt32(objReader("TOTAL_ROW"))
        End If

        'Dim strExc16 As String = $"SELECT COUNT(*) AS TOTAL_ROW FROM crrnew.t_rdt_exc_report_016 WHERE BU_CODE IN {strBUCodeFinal}"
        objCmd = connOracle.CreateCommand
        objCmd.CommandText = strExc16
        objReader = objCmd.ExecuteReader()
        Dim dtTotalRowExc16 As Integer = 0
        If objReader.Read() Then
            dtTotalRowExc16 = Convert.ToInt32(objReader("TOTAL_ROW"))
        End If

        Dim strExc22 As String = $"SELECT COUNT(*) AS TOTAL_ROW FROM crrnew.t_rdt_exc_report_022 WHERE BU_CODE IN {strBUCodeFinal}"
        objCmd = connOracle.CreateCommand
        objCmd.CommandText = strExc22
        objReader = objCmd.ExecuteReader()
        Dim dtTotalRowExc22 As Integer = 0
        If objReader.Read() Then
            dtTotalRowExc22 = Convert.ToInt32(objReader("TOTAL_ROW"))
        End If

        Dim strExc03 As String = $"SELECT COUNT(*) AS TOTAL_ROW FROM crrnew.t_rdt_exc_report_003 WHERE BU_CODE IN {strBUCodeFinal}"
        objCmd = connOracle.CreateCommand
        objCmd.CommandText = strExc03
        objReader = objCmd.ExecuteReader()
        Dim dtTotalRowExc03 As Integer = 0
        If objReader.Read() Then
            dtTotalRowExc03 = Convert.ToInt32(objReader("TOTAL_ROW"))
        End If

        If strBUcode <> "CC" Then
            strMsg = "RDT Exception Report as at " & endOfDayDateThai & " ตามไฟล์แนบ โดยมีสรุปรายการดังนี้ 
                                <br>" &
                                "<style>table, th, td { border: 1px solid; font-family:CordiaUPC; font-size:14pt; vertical-align: text-top;}</style>" & vbNewLine &
                                "<table style='border-collapse: collapse; width: 800px'>" &
                                "<tr style = 'background-color: #d0e7f9;'>" &
                                "<th style='width: 10%;'>No.</th>" &
                                "<th style='width: 15%;text-align: left; padding-left: 20px'>Exception<br>Report No.</th>" &
                                "<th style='width: 60%; text-align: left; padding-left: 20px'>Exception Name</th>" &
                                "<th style='width: 15%;'>จำนวนรายการ</th>" &
                                "</tr>" &
                                "<tr >" &
                                "<td style = 'text-align:center;'>1</td>" &
                                "<td style = 'text-align:center;'>EXC_022</td>" &
                                "<td style='padding-left: 20px; padding-right: 20px'>สินเชื่อที่ BOT Purpose Code ไม่มีค่า</td>" &
                                "<td style = 'text-align:center;'>" & dtTotalRowExc22.ToString("N0") & "</td>" &
                                "</tr>" &
                                "<tr>" &
                                "<td style = 'text-align:center;'>2</td>" &
                                "<td style = 'text-align:center;'>EXC_003</td>" &
                                "<td style='padding-left: 20px; padding-right: 20px'>สินเชื่อที่ ISIC และ BOT Purpose Code ไม่สอดคล้องกัน</td>" &
                                "<td style = 'text-align:center;'>" & dtTotalRowExc03.ToString("N0") & "</td>" &
                                "</tr>" &
                                "<tr>" &
                                "<td style = 'text-align:center;'>3</td>" &
                                "<td style = 'text-align:center;'>EXC_011</td>" &
                                "<td style='padding-left: 20px; padding-right: 20px'>หุ้นนอกตลาดหลักทรัพย์ทั้งในประเทศ และต่างประเทศ ที่ไม่มีวันที่ประเมินมูลค่าหุ้น</td>" &
                                "<td style = 'text-align:center;'>" & dtTotalRowExc11.ToString("N0") & "</td>" &
                                "</tr>" &
                                "<tr>" &
                                "<td style = 'text-align:center;'>4</td>" &
                                "<td style = 'text-align:center;'>EXC_012</td>" &
                                "<td style='padding-left: 20px; padding-right: 20px'>หุ้นนอกตลาดหลักทรัพย์ทั้งในประเทศ และต่างประเทศ ที่ไม่มีราคาประเมินมูลค่าหุ้น</td>" &
                                "<td style = 'text-align:center;'>" & dtTotalRowExc12.ToString("N0") & "</td>" &
                                "</tr>" &
                                "<tr>" &
                                "<td style = 'text-align:center;'>5</td>" &
                                "<td style = 'text-align:center;'>EXC_014</td>" &
                                "<td style='padding-left: 20px; padding-right: 20px'>ลูกค้าสินเชื่อธุรกิจที่ไม่ผูก Allocation Code สินเชื่อ</td>" &
                                "<td style = 'text-align:center;'>" & dtTotalRowExc14.ToString("N0") & "</td>" &
                                "</tr>" &
                                "<tr>" &
                                "<td style = 'text-align:center;'>6</td>" &
                                "<td style = 'text-align:center;'>EXC_015</td>" &
                                "<td style='padding-left: 20px; padding-right: 20px'>ลูกค้าที่ไม่มีข้อมูลที่ไม่มีข้อมูลรายได้</td>" &
                                "<td style = 'text-align:center;'>" & dtTotalRowExc15.ToString("N0") & "</td>" &
                                "</tr>" &
                                "<tr>" &
                                "<td style = 'text-align:center;'>7</td>" &
                                "<td style = 'text-align:center;'>EXC_016</td>" &
                                "<td style='padding-left: 20px; padding-right: 20px'>ลูกค้าสินเชื่อธุรกิจที่ไม่มีข้อมูลจำนวนแรงงาน </td>" &
                                "<td style = 'text-align:center;'>" & dtTotalRowExc16.ToString("N0") & "</td>" &
                                "</tr>" &
                                "</table>"

        Else
            strMsg = "RDT Exception Report as at " & endOfDayDateThai & " ตามไฟล์แนบ โดยมีสรุปรายการดังนี้ 
                                <br>" &
                                "<style>table, th, td { border: 1px solid; font-family:CordiaUPC; font-size:14pt; vertical-align: text-top;}</style>" & vbNewLine &
                                "<table style='border-collapse: collapse; width: 800px'>" &
                                "<tr style = 'background-color: #d0e7f9;'>" &
                                "<th style='width: 10%;'>No.</th>" &
                                "<th style='width: 15%;text-align: left; padding-left: 20px'>Exception<br>Report No.</th>" &
                                "<th style='width: 60%; text-align: left; padding-left: 20px'>Exception Name</th>" &
                                "<th style='width: 15%;'>จำนวนรายการ</th>" &
                                "</tr>" &
                                "<td style = 'text-align:center;'>1</td>" &
                                "<td style = 'text-align:center;'>EXC_014</td>" &
                                "<td style='padding-left: 20px; padding-right: 20px'>ลูกค้าสินเชื่อธุรกิจที่ไม่ผูก Allocation Code สินเชื่อ</td>" &
                                "<td style = 'text-align:center;'>" & dtTotalRowExc14.ToString("N0") & "</td>" &
                                "</tr>" &
                                "<tr>" &
                                "<td style = 'text-align:center;'>2</td>" &
                                "<td style = 'text-align:center;'>EXC_015</td>" &
                                "<td style='padding-left: 20px; padding-right: 20px'>ลูกค้าที่ไม่มีข้อมูลที่ไม่มีข้อมูลรายได้</td>" &
                                "<td style = 'text-align:center;'>" & dtTotalRowExc15.ToString("N0") & "</td>" &
                                "</tr>" &
                                "<tr>" &
                                "<td style = 'text-align:center;'>3</td>" &
                                "<td style = 'text-align:center;'>EXC_016</td>" &
                                "<td style='padding-left: 20px; padding-right: 20px'>ลูกค้าสินเชื่อธุรกิจที่ไม่มีข้อมูลจำนวนแรงงาน</td>" &
                                "<td style = 'text-align:center;'>" & dtTotalRowExc16.ToString("N0") & "</td>" &
                                "</tr>" &
                                "</table>"

        End If


        '& strDateMonth & "/" & strDateYear & " โดยมีรายละเอียดตามไฟล์แนบ" '12/2024 MM/YYYY

        If ConfigurationManager.AppSettings("Mode") = "DEV" Then
            '----------Body for show Mail To,Cc----------
            strListMail = "<font style='font-family:CordiaUPC; font-size:14pt;'>------------------------------------------------------------------------------------------<br>"
            strListMail = strListMail & " Mail To: " & strMailTo & "<br>"
            strListMail = strListMail & " Mail Cc: " & strMailCc & "<br>"
            strListMail = strListMail & "------------------------------------------------------------------------------------------</font><br>"
            '----------To----------
            Dim strTo As String = ConfigurationManager.AppSettings("mailto")
            If Not IsNothing(strTo) AndAlso strTo <> "" Then
                insMail.To.Add(Trim(ConfigurationManager.AppSettings("mailto")))
                'insMail.To.Add(Trim(strMailCc))

            End If
            '----------Cc----------
            Dim strCC As String = ConfigurationManager.AppSettings("mailcc")
            If Not IsNothing(strCC) AndAlso strCC <> "" Then
                insMail.CC.Add(Trim(ConfigurationManager.AppSettings("mailcc")))
            End If

        Else 'Prod - select from database
            '----------To----------
            If Not IsNothing(strMailTo) AndAlso strMailTo <> "" Then
                insMail.To.Add(Trim(strMailTo))
            End If
            '----------Cc----------
            If Not IsNothing(strMailCc) AndAlso strMailCc <> "" Then
                insMail.CC.Add(Trim(strMailCc))
            End If
        End If

        insMail.Subject = strSubject
        insMail.IsBodyHtml = True
        Dim htmlString As String
        today = Date.Today
        Dim thaiCulture As New CultureInfo("th-TH")

        thaiCulture.DateTimeFormat.Calendar = New GregorianCalendar()

        today = today.ToString("dd/MM/yyyy", thaiCulture)
        Debug.WriteLine("today test: " + today)


        If Day(today) <= 20 Then
            endOfMonth = DateSerial(Year(today), Month(today), 20)
        ElseIf Day(today) > 20 And Day(today) <= 27 Then
            endOfMonth = DateSerial(Year(today), Month(today), 27)
        Else
            endOfMonth = DateSerial(Year(today), Month(today) + 1, 0)
        End If


        endOfMonthStr = endOfMonth.ToString("dd/MM/yyyy")
        Debug.WriteLine("endOfMonthStr: " + endOfMonthStr)

        If (strBUcode <> "CC") Then
            htmlString = "<font style='font-family:CordiaUPC; font-size:14pt;'>" & strListMail & "<p>เรียน ทุกท่าน</p>
                     " & strMsg & "
                     &nbsp;หมายเหตุ: EXC_014, EXC_015 และ EXC_016 อยู่ระหว่างการทบทวนเงื่อนไข และจัดส่งข้อมูลให้หากดำเนินการปรับเงื่อนไขเรียบร้อยแล้ว<br><br>ทั้งนี้ กรุณาแก้ไขข้อมูลให้ถูกต้องภายใน " & endOfMonthStr & " เพื่อให้สามารถนำส่งข้อมูลที่ถูกต้องให้ RDT <br>จึงเรียนมาเพื่อโปรดดำเนินการ<br>
                     RDT Gap Closure <br><br>Please DO NOT REPLY to this automated e-mail, any reply made to this message will NOT be collected nor responded. 
                     " & "</font>"
        Else
            htmlString = "<font style='font-family:CordiaUPC; font-size:14pt;'>" & strListMail & "<p>เรียน ทุกท่าน</p>
                     " & strMsg & "
                     &nbsp;<br>ทั้งนี้ กรุณาแก้ไขข้อมูลให้ถูกต้องภายใน " & endOfMonthStr & " เพื่อให้สามารถนำส่งข้อมูลที่ถูกต้องให้ RDT <br>จึงเรียนมาเพื่อโปรดดำเนินการ<br>
                     RDT Gap Closure <br><br>Please DO NOT REPLY to this automated e-mail, any reply made to this message will NOT be collected nor responded. 
                     " & "</font>"

        End If

        'used for 06/2568
        'htmlString = "<font style='font-family:CordiaUPC; font-size:14pt;'>" & strListMail & "<p>เรียน ทุกท่าน</p>
        '             " & strMsg & "
        '             &nbsp;<br>ทั้งนี้ กรุณาแก้ไขข้อมูลให้ถูกต้องภายใน " & "30/06/2025" & " เพื่อให้สามารถนำส่งข้อมูลที่ถูกต้องให้ RDT <br>จึงเรียนมาเพื่อโปรดดำเนินการ<br>
        '             RDT Gap Closure <br><br>Please DO NOT REPLY to this automated e-mail, any reply made to this message will NOT be collected nor responded. 
        '             " & "</font>"
        'Standard footer
        'htmlString = htmlString & "<br><br> Please DO NOT REPLY to this automated e-mail, any reply made to this message will NOT be collected nor responded."
        'htmlString = htmlString & "</small>"

        insMail.Body = htmlString
        attachment1 = New Attachment(strPathFileNameReport1)
        insMail.Attachments.Add(attachment1)

        ClientMail.Send(insMail)
        insMail.To.Clear()
    End Sub

End Module
