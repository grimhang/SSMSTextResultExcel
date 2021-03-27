using System;
using System.Diagnostics;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SSMSTextResultExcelLib;
using System.Data;
//using SSMSTextResultFormatter;

namespace SSMSTextResultFormatterTest
{
    [TestClass]
    public class MainUnitTest
    {
        [TestCategory("ColumnCategory")]
        [TestMethod]
        public void OneResult_ConInfo_PrivateTest()
        {
            TextFormatter target = new TextFormatter(2);

            object[] args = new object[] { File.ReadAllText("ResultSample1.txt") };

            PrivateObject obj = new PrivateObject(target);
            var retVal = (ColProperty)obj.Invoke("calcOneResultAfterColLen", args);

            Assert.AreEqual(retVal.colCnt, 2);
            Assert.AreEqual(retVal.rowCnt, 3);

        }


        #region DS테스트
        [TestCategory("DatatableCategory")]
        [TestMethod]
        public void OneResultDT_Test2x1()
        {
            TextFormatter target = new TextFormatter(1);

            var dt = target.OneResultFormatDT(File.ReadAllText("ResultSample1.txt"));

            Assert.AreEqual(dt.Columns.Count, 2);
            Assert.AreEqual(dt.Rows.Count, 1);

            Assert.AreEqual(dt.Columns[0].ColumnName, "Server Name");
            Assert.AreEqual(dt.Columns[1].ColumnName, "Report Date");

            Assert.AreEqual(dt.Rows[0][0], @"AWMSSQL\FXAW");
            Assert.AreEqual(dt.Rows[0][1], "2019-01-22 15:08:02");
        }

        [TestCategory("DatatableCategory")]
        [TestMethod]
        public void OneResultDT_Test2x7()
        {
            TextFormatter target = new TextFormatter(1);

            var dt = target.OneResultFormatDT(File.ReadAllText("ResultSample2.txt"));

            Assert.AreEqual(dt.Columns.Count, 4);
            Assert.AreEqual(dt.Rows.Count, 7);

            Assert.AreEqual(dt.Columns[0].ColumnName, @"SQL Server\Instance Name");
            Assert.AreEqual(dt.Columns[1].ColumnName, "Service Name");
            Assert.AreEqual(dt.Columns[2].ColumnName, "Service Status");
            Assert.AreEqual(dt.Columns[3].ColumnName, @"Status Date\Time");

            Assert.AreEqual(dt.Rows[0][0], @"FXAW");
            Assert.AreEqual(dt.Rows[0][1], @"MS SQL Server Service");
            Assert.AreEqual(dt.Rows[0][2], @"Running.");
            Assert.AreEqual(dt.Rows[0][3], @"2019-01-22 15:08:03.003");
        }

        [TestCategory("DatasetCategory")]
        [TestMethod]
        public void OneResultDS_Test()
        {
            TextFormatter target = new TextFormatter(1);

            var ds = target.TotalResultFormatDS(File.ReadAllText("ResultSampleDataset.txt"));

            Assert.AreEqual(ds.Tables.Count, 2);    // 총 테이블수

            // 첫번째 테이블 확인
            Assert.AreEqual(ds.Tables[0].Rows.Count, 1);
            Assert.AreEqual(ds.Tables[0].Columns.Count, 2);

            Assert.AreEqual(ds.Tables[0].Columns[0].ColumnName, "Server Name");
            Assert.AreEqual(ds.Tables[0].Columns[1].ColumnName, "Report Date");

            Assert.AreEqual(ds.Tables[0].Rows[0][0], @"AWMSSQL\FXAW");
            Assert.AreEqual(ds.Tables[0].Rows[0][1], "2019-01-22 15:08:02");

            // 두번째 테이블 확인
            Assert.AreEqual(ds.Tables[1].Rows.Count, 26);
            Assert.AreEqual(ds.Tables[1].Columns.Count, 2);

            Assert.AreEqual(ds.Tables[1].Columns[0].ColumnName, "KeyName");
            Assert.AreEqual(ds.Tables[1].Columns[1].ColumnName, "KeyVal");

            Assert.AreEqual(ds.Tables[1].Rows[3][0], @"Install Date");
            Assert.AreEqual(ds.Tables[1].Rows[3][1], "2016-04-11 11:11:31");

            Assert.AreEqual(ds.Tables[1].Rows[20][0], @"SQL Server Default Trace Location");
            Assert.AreEqual(ds.Tables[1].Rows[20][1], @"E:\MSSQL11.FXAW\MSSQL\Log\log.trc");
        }

        [TestCategory("DatasetCategory")]
        [TestMethod]
        public void OneResultDS_withTitle_Test()
        {
            TextFormatter target = new TextFormatter(1);

            var ds = target.TotalResultFormatDS(File.ReadAllText("ResultSampleDatasetTitle.txt"), true);

            Assert.AreEqual(ds.Tables.Count, 2);    // 총 테이블수

            // 첫번째 테이블 확인
            Assert.AreEqual(ds.Tables[0].TableName, "SQL Server Report Date - Version 2.0");
            Assert.AreEqual(ds.Tables[0].Rows.Count, 1);
            Assert.AreEqual(ds.Tables[0].Columns.Count, 2);

            Assert.AreEqual(ds.Tables[0].Columns[0].ColumnName, "Server Name");
            Assert.AreEqual(ds.Tables[0].Columns[1].ColumnName, "Report Date");

            Assert.AreEqual(ds.Tables[0].Rows[0][0], @"AWMSSQL\FXAW");
            Assert.AreEqual(ds.Tables[0].Rows[0][1], "2019-01-22 15:08:02");

            // 두번째 테이블 확인
            Assert.AreEqual(ds.Tables[1].TableName, "SQL Server Summary");
            Assert.AreEqual(ds.Tables[1].Rows.Count, 26);
            Assert.AreEqual(ds.Tables[1].Columns.Count, 2);

            Assert.AreEqual(ds.Tables[1].Columns[0].ColumnName, "KeyName");
            Assert.AreEqual(ds.Tables[1].Columns[1].ColumnName, "KeyVal");

            Assert.AreEqual(ds.Tables[1].Rows[3][0], @"Install Date");
            Assert.AreEqual(ds.Tables[1].Rows[3][1], "2016-04-11 11:11:31");

            Assert.AreEqual(ds.Tables[1].Rows[20][0], @"SQL Server Default Trace Location");
            Assert.AreEqual(ds.Tables[1].Rows[20][1], @"E:\MSSQL11.FXAW\MSSQL\Log\log.trc");
        }

        [TestCategory("DatasetCategory")]
        [TestMethod]
        public void OneResultDS_withTitle_ExcelTest()
        {
            var excelResult = new ExportFile();

            excelResult.ExcelResult(File.ReadAllText("ResultSampleDatasetTitle.txt"), "ExcelResult.xlsx");
        }

        [TestCategory("DatasetCategory")]
        [TestMethod]
        public void OneResultDS_withTitleLastBlank_ExcelTest()
        {
            TextFormatter target = new TextFormatter(1);

            var ds = target.TotalResultFormatDS(File.ReadAllText("ResultSampleDatasetTitle_LastBlankData.txt"), true);

            Assert.AreEqual(ds.Tables.Count, 4);    // 총 테이블수

            // 첫번째 테이블 확인
            Assert.AreEqual(ds.Tables[0].TableName, "SQL Server Report Date - Version 2.0");
            Assert.AreEqual(ds.Tables[0].Rows.Count, 1);
            Assert.AreEqual(ds.Tables[0].Columns.Count, 2);

            Assert.AreEqual(ds.Tables[0].Columns[0].ColumnName, "Server Name");
            Assert.AreEqual(ds.Tables[0].Columns[1].ColumnName, "Report Date");

            Assert.AreEqual(ds.Tables[0].Rows[0][0], @"AWMSSQL\FXAW");
            Assert.AreEqual(ds.Tables[0].Rows[0][1], "2019-01-22 15:08:02");

            // 두번째 테이블 확인
            Assert.AreEqual(ds.Tables[1].TableName, "SQL Server Summary");
            Assert.AreEqual(ds.Tables[1].Rows.Count, 26);
            Assert.AreEqual(ds.Tables[1].Columns.Count, 2);

            Assert.AreEqual(ds.Tables[1].Columns[0].ColumnName, "KeyName");
            Assert.AreEqual(ds.Tables[1].Columns[1].ColumnName, "KeyVal");

            Assert.AreEqual(ds.Tables[1].Rows[3][0], @"Install Date");
            Assert.AreEqual(ds.Tables[1].Rows[3][1], "2016-04-11 11:11:31");

            Assert.AreEqual(ds.Tables[1].Rows[20][0], @"SQL Server Default Trace Location");
            Assert.AreEqual(ds.Tables[1].Rows[20][1], @"E:\MSSQL11.FXAW\MSSQL\Log\log.trc");

            //Assert.AreEqual(ds.Tables.Count, 2);    // 총 테이블수

            // 세번째 테이블 확인
            Assert.AreEqual(ds.Tables[2].TableName, "Location of Database files");
            Assert.AreEqual(ds.Tables[2].Columns.Count, 4);
            Assert.AreEqual(ds.Tables[2].Rows.Count, 26);

            Assert.AreEqual(ds.Tables[2].Columns[0].ColumnName, "Database ID");
            Assert.AreEqual(ds.Tables[2].Columns[1].ColumnName, "Database Name");
            Assert.AreEqual(ds.Tables[2].Columns[2].ColumnName, "Physical Location");
            Assert.AreEqual(ds.Tables[2].Columns[3].ColumnName, "Type");

            Assert.AreEqual(ds.Tables[2].Rows[13][0], @"7");
            Assert.AreEqual(ds.Tables[2].Rows[13][1], @"IM_WEBUI_DATALog");
            Assert.AreEqual(ds.Tables[2].Rows[13][2], @"E:\MSSQL11.FXAW\MSSQL\DATA\IM_WEBUI_DATA.ldf");

            // 네번째 테이블 확인
            Assert.AreEqual(ds.Tables[3].TableName, "permissions of the users for each database");
            Assert.AreEqual(ds.Tables[3].Columns.Count, 6);
            Assert.AreEqual(ds.Tables[3].Rows.Count, 42);

            //Assert.AreEqual(ds.Tables[3].Columns[0].ColumnName, "Database ID");
            //Assert.AreEqual(ds.Tables[3].Columns[1].ColumnName, "Database Name");
            //Assert.AreEqual(ds.Tables[3].Columns[2].ColumnName, "Physical Location");
            //Assert.AreEqual(ds.Tables[3].Columns[3].ColumnName, "Type");

            Assert.AreEqual(ds.Tables[3].Rows[41][0], @"SDM_DATA");
            Assert.AreEqual(ds.Tables[3].Rows[41][1], @"public");
            Assert.AreEqual(ds.Tables[3].Rows[41][2], @"DATABASE_ROLE");
            Assert.AreEqual(ds.Tables[3].Rows[41][3], @"2003-04-08 09:10:42.317");
            Assert.AreEqual(ds.Tables[3].Rows[41][4], @"2009-04-13 12:59:14.467");
            Assert.AreEqual(ds.Tables[3].Rows[41][5], @"");
        }

        [TestCategory("DatasetCategory")]
        [TestMethod]
        public void OneResultDS_withTitleLastBlankDataOnly_Test()
        {
            TextFormatter target = new TextFormatter(1);

            var ds = target.TotalResultFormatDS(File.ReadAllText("ResultSampleDatasetTitle_LastBlankDataOnly.txt"), true);

            Assert.AreEqual(ds.Tables.Count, 1);    // 총 테이블수

            // 테이블 확인
            Assert.AreEqual(ds.Tables[0].TableName, "permissions of the users for each database");
            Assert.AreEqual(ds.Tables[0].Columns.Count, 6);
            Assert.AreEqual(ds.Tables[0].Rows.Count, 42);

            Assert.AreEqual(ds.Tables[0].Rows[17][0], @"msdb");
            Assert.AreEqual(ds.Tables[0].Rows[17][1], @"db_ssisadmin");
            Assert.AreEqual(ds.Tables[0].Rows[17][2], @"DATABASE_ROLE");
            Assert.AreEqual(ds.Tables[0].Rows[17][3], @"2012-02-10 21:05:40.257");
            Assert.AreEqual(ds.Tables[0].Rows[17][4], @"2012-02-10 21:05:40.257");
            Assert.AreEqual(ds.Tables[0].Rows[17][5], @"");

            Assert.AreEqual(ds.Tables[0].Rows[25][0], @"msdb");
            Assert.AreEqual(ds.Tables[0].Rows[25][1], @"PolicyAdministratorRole");
            Assert.AreEqual(ds.Tables[0].Rows[25][2], @"DATABASE_ROLE");
            Assert.AreEqual(ds.Tables[0].Rows[25][3], @"2012-02-10 21:07:43.647");
            Assert.AreEqual(ds.Tables[0].Rows[25][4], @"2012-02-10 21:07:43.647");
            Assert.AreEqual(ds.Tables[0].Rows[25][5], @"SQLAgentOperatorRole");

            Assert.AreEqual(ds.Tables[0].Rows[41][0], @"SDM_DATA");
            Assert.AreEqual(ds.Tables[0].Rows[41][1], @"public");
            Assert.AreEqual(ds.Tables[0].Rows[41][2], @"DATABASE_ROLE");
            Assert.AreEqual(ds.Tables[0].Rows[41][3], @"2003-04-08 09:10:42.317");
            Assert.AreEqual(ds.Tables[0].Rows[41][4], @"2009-04-13 12:59:14.467");
            Assert.AreEqual(ds.Tables[0].Rows[41][5], @"");
        }

        #endregion

        [TestCategory("ExtensionMethod")]
        [TestMethod]
        public void ExtensionMethod_RemovePostNewLineTest()
        {

            TextFormatter target = new TextFormatter(1);

            var str = target.OneResultFormat(File.ReadAllText("ResultSample_RemovePostNewLine.txt"));

            var str2 = str.RemovePostNewLine();

            string[] str2Arr = str2.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);

            Assert.AreEqual(str2Arr[0], "Server Name  Report Date");
            Assert.AreEqual(str2Arr[2], @"AWMSSQL\FXAW 2019-01-22 15:08:02");

        }

        [TestCategory("TryCheckText")]
        [TestMethod]
        public void TryCheckTextPartial_Test()
        {

            string str;
            bool bResult;

            TextFormatter target = new TextFormatter(1);

            bResult = target.TryCheckTextPartial(File.ReadAllText("ResultSample1.txt"), out str);
            Assert.AreEqual(bResult, true);

            bResult = target.TryCheckTextPartial(File.ReadAllText("ResultSample2.txt"), out str);
            Assert.AreEqual(bResult, true);

        }
    }
}