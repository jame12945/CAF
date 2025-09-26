using CAFGenerator.Model;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Data;

namespace CAFGenerator.Classes
{
    public class HeaderGenerator
    {
        private readonly string _caseID;
        private readonly string _userID;
        private readonly CreditApplicationType _creditApplicationType;
        private readonly ExcelPackage _excelPackage;

        public HeaderGenerator(string caseID, string userID, CreditApplicationType creditApplicationType, ExcelPackage excelPackage)
        {
            _caseID = caseID;
            _userID = userID;
            _creditApplicationType = creditApplicationType;
            _excelPackage = excelPackage;
            GenerateSheet();
        }

        private void GenerateSheet()
        {
            ExcelWorksheet ws1 = _excelPackage.Workbook.Worksheets[1];
            ExcelWorksheet ws2 = _excelPackage.Workbook.Worksheets[2];
            ExcelWorksheet ws3 = _excelPackage.Workbook.Worksheets[3];
            ExcelWorksheet ws4 = _excelPackage.Workbook.Worksheets[4];
            ExcelWorksheet ws5 = _excelPackage.Workbook.Worksheets[5];
            ExcelWorksheet ws6 = _excelPackage.Workbook.Worksheets[6];
            ExcelWorksheet ws7 = _excelPackage.Workbook.Worksheets[7];
            ExcelWorksheet ws8 = _excelPackage.Workbook.Worksheets[8];
            ExcelWorksheet ws9 = _excelPackage.Workbook.Worksheets[9];
            ExcelWorksheet ws10 = _excelPackage.Workbook.Worksheets[10];
            ExcelWorksheet ws11 = _excelPackage.Workbook.Worksheets[11];

            DataTable dtHeader = GetHeaderData();
            if (dtHeader != null)
            {

                // Group Name, Customer Name
                string strHeader_Name_Group = " " + dtHeader.Rows[0]["group_customer_name"].ToString();

                if (_creditApplicationType == CreditApplicationType.SingleCustomer)
                {
                    var boldRichText = ws1.Cells[$"A3"].RichText.Add("ลูกค้า:"); boldRichText.Bold = true;
                    boldRichText = ws2.Cells[$"A3"].RichText.Add("ลูกค้า:"); boldRichText.Bold = true;
                    boldRichText = ws3.Cells[$"A3"].RichText.Add("ลูกค้า:"); boldRichText.Bold = true;
                    boldRichText = ws4.Cells[$"A3"].RichText.Add("ลูกค้า:"); boldRichText.Bold = true;
                    boldRichText = ws5.Cells[$"A3"].RichText.Add("ลูกค้า:"); boldRichText.Bold = true;
                    boldRichText = ws6.Cells[$"A3"].RichText.Add("ลูกค้า:"); boldRichText.Bold = true;
                    boldRichText = ws7.Cells[$"A3"].RichText.Add("ลูกค้า:"); boldRichText.Bold = true;
                    boldRichText = ws8.Cells[$"A3"].RichText.Add("ลูกค้า:"); boldRichText.Bold = true;
                    boldRichText = ws9.Cells[$"A3"].RichText.Add("ลูกค้า:"); boldRichText.Bold = true;
                    boldRichText = ws10.Cells[$"A3"].RichText.Add("ลูกค้า:"); boldRichText.Bold = true;
                    boldRichText = ws11.Cells[$"A3"].RichText.Add("ลูกค้า:"); boldRichText.Bold = true;

                }
                else
                {
                    var boldRichText = ws1.Cells[$"A3"].RichText.Add("ชื่อกลุ่ม:"); boldRichText.Bold = true;
                    boldRichText = ws2.Cells[$"A3"].RichText.Add("ชื่อกลุ่ม:"); boldRichText.Bold = true;
                    boldRichText = ws3.Cells[$"A3"].RichText.Add("ชื่อกลุ่ม:"); boldRichText.Bold = true;
                    boldRichText = ws4.Cells[$"A3"].RichText.Add("ชื่อกลุ่ม:"); boldRichText.Bold = true;
                    boldRichText = ws5.Cells[$"A3"].RichText.Add("ชื่อกลุ่ม:"); boldRichText.Bold = true;
                    boldRichText = ws6.Cells[$"A3"].RichText.Add("ชื่อกลุ่ม:"); boldRichText.Bold = true;
                    boldRichText = ws7.Cells[$"A3"].RichText.Add("ชื่อกลุ่ม:"); boldRichText.Bold = true;
                    boldRichText = ws8.Cells[$"A3"].RichText.Add("ชื่อกลุ่ม:"); boldRichText.Bold = true;
                    boldRichText = ws9.Cells[$"A3"].RichText.Add("ชื่อกลุ่ม:"); boldRichText.Bold = true;
                    boldRichText = ws10.Cells[$"A3"].RichText.Add("ชื่อกลุ่ม:"); boldRichText.Bold = true;
                    boldRichText = ws11.Cells[$"A3"].RichText.Add("ชื่อกลุ่ม:"); boldRichText.Bold = true;

                }

                var normalRichText = ws1.Cells["A3"].RichText.Add(strHeader_Name_Group); normalRichText.Bold = false;
                normalRichText = ws2.Cells["A3"].RichText.Add(strHeader_Name_Group); normalRichText.Bold = false;
                normalRichText = ws3.Cells["A3"].RichText.Add(strHeader_Name_Group); normalRichText.Bold = false;
                normalRichText = ws4.Cells["A3"].RichText.Add(strHeader_Name_Group); normalRichText.Bold = false;
                normalRichText = ws5.Cells["A3"].RichText.Add(strHeader_Name_Group); normalRichText.Bold = false;
                normalRichText = ws6.Cells["A3"].RichText.Add(strHeader_Name_Group); normalRichText.Bold = false;
                normalRichText = ws7.Cells["A3"].RichText.Add(strHeader_Name_Group); normalRichText.Bold = false;
                normalRichText = ws8.Cells["A3"].RichText.Add(strHeader_Name_Group); normalRichText.Bold = false;
                normalRichText = ws9.Cells["A3"].RichText.Add(strHeader_Name_Group); normalRichText.Bold = false;
                normalRichText = ws10.Cells["A3"].RichText.Add(strHeader_Name_Group); normalRichText.Bold = false;
                normalRichText = ws11.Cells["A3"].RichText.Add(strHeader_Name_Group); normalRichText.Bold = false;


                //BU, BC/DIV
                string strBUName = " " + dtHeader.Rows[0]["business_unit_name"].ToString() + "    ";
                string strDivorBC = " " + dtHeader.Rows[0]["business_center"].ToString();
                string strBUShortName = dtHeader.Rows[0]["business_unit_shortname"].ToString();

                var boldRichText2 = ws1.Cells[$"Q2"].RichText.Add("สายงาน:"); boldRichText2.Bold = true;
                boldRichText2 = ws2.Cells[$"O2"].RichText.Add("สายงาน:"); boldRichText2.Bold = true;
                boldRichText2 = ws3.Cells[$"S2"].RichText.Add("สายงาน:"); boldRichText2.Bold = true;
                boldRichText2 = ws4.Cells[$"P2"].RichText.Add("สายงาน:"); boldRichText2.Bold = true;
                boldRichText2 = ws5.Cells[$"I2"].RichText.Add("สายงาน:"); boldRichText2.Bold = true;
                boldRichText2 = ws6.Cells[$"G2"].RichText.Add("สายงาน:"); boldRichText2.Bold = true;
                boldRichText2 = ws7.Cells[$"R2"].RichText.Add("สายงาน:"); boldRichText2.Bold = true;
                boldRichText2 = ws8.Cells[$"S2"].RichText.Add("สายงาน:"); boldRichText2.Bold = true;
                boldRichText2 = ws9.Cells[$"S2"].RichText.Add("สายงาน:"); boldRichText2.Bold = true;
                boldRichText2 = ws10.Cells[$"E2"].RichText.Add("สายงาน:"); boldRichText2.Bold = true;
                boldRichText2 = ws11.Cells[$"L2"].RichText.Add("สายงาน:"); boldRichText2.Bold = true;

                normalRichText = ws1.Cells[$"Q2"].RichText.Add(strBUName); normalRichText.Bold = false;
                normalRichText = ws2.Cells[$"O2"].RichText.Add(strBUName); normalRichText.Bold = false;
                normalRichText = ws3.Cells[$"S2"].RichText.Add(strBUName); normalRichText.Bold = false;
                normalRichText = ws4.Cells[$"P2"].RichText.Add(strBUName); normalRichText.Bold = false;
                normalRichText = ws5.Cells[$"I2"].RichText.Add(strBUName); normalRichText.Bold = false;
                normalRichText = ws6.Cells[$"G2"].RichText.Add(strBUName); normalRichText.Bold = false;
                normalRichText = ws7.Cells[$"R2"].RichText.Add(strBUName); normalRichText.Bold = false;
                normalRichText = ws8.Cells[$"S2"].RichText.Add(strBUName); normalRichText.Bold = false;
                normalRichText = ws9.Cells[$"S2"].RichText.Add(strBUName); normalRichText.Bold = false;
                normalRichText = ws10.Cells[$"E2"].RichText.Add(strBUName); normalRichText.Bold = false;
                normalRichText = ws11.Cells[$"L2"].RichText.Add(strBUName); normalRichText.Bold = false;

                if (strBUShortName == "NBM")
                {
                    boldRichText2 = ws1.Cells[$"Q2"].RichText.Add("ธุรกิจ:"); boldRichText2.Bold = true;
                    boldRichText2 = ws2.Cells[$"O2"].RichText.Add("ธุรกิจ:"); boldRichText2.Bold = true;
                    boldRichText2 = ws3.Cells[$"S2"].RichText.Add("ธุรกิจ:"); boldRichText2.Bold = true;
                    boldRichText2 = ws4.Cells[$"P2"].RichText.Add("ธุรกิจ:"); boldRichText2.Bold = true;
                    boldRichText2 = ws5.Cells[$"I2"].RichText.Add("ธุรกิจ:"); boldRichText2.Bold = true;
                    boldRichText2 = ws6.Cells[$"G2"].RichText.Add("ธุรกิจ:"); boldRichText2.Bold = true;
                    boldRichText2 = ws7.Cells[$"R2"].RichText.Add("ธุรกิจ:"); boldRichText2.Bold = true;
                    boldRichText2 = ws8.Cells[$"S2"].RichText.Add("ธุรกิจ:"); boldRichText2.Bold = true;
                    boldRichText2 = ws9.Cells[$"S2"].RichText.Add("ธุรกิจ:"); boldRichText2.Bold = true;
                    boldRichText2 = ws10.Cells[$"E2"].RichText.Add("ธุรกิจ:"); boldRichText2.Bold = true;
                    boldRichText2 = ws11.Cells[$"L2"].RichText.Add("ธุรกิจ:"); boldRichText2.Bold = true;
                }
                else
                {
                    boldRichText2 = ws1.Cells[$"Q2"].RichText.Add("สำนักธุรกิจ:"); boldRichText2.Bold = true;
                    boldRichText2 = ws2.Cells[$"O2"].RichText.Add("สำนักธุรกิจ:"); boldRichText2.Bold = true;
                    boldRichText2 = ws3.Cells[$"S2"].RichText.Add("สำนักธุรกิจ:"); boldRichText2.Bold = true;
                    boldRichText2 = ws4.Cells[$"P2"].RichText.Add("สำนักธุรกิจ:"); boldRichText2.Bold = true;
                    boldRichText2 = ws5.Cells[$"I2"].RichText.Add("สำนักธุรกิจ:"); boldRichText2.Bold = true;
                    boldRichText2 = ws6.Cells[$"G2"].RichText.Add("สำนักธุรกิจ:"); boldRichText2.Bold = true;
                    boldRichText2 = ws7.Cells[$"R2"].RichText.Add("สำนักธุรกิจ:"); boldRichText2.Bold = true;
                    boldRichText2 = ws8.Cells[$"S2"].RichText.Add("สำนักธุรกิจ:"); boldRichText2.Bold = true;
                    boldRichText2 = ws9.Cells[$"S2"].RichText.Add("สำนักธุรกิจ:"); boldRichText2.Bold = true;
                    boldRichText2 = ws10.Cells[$"E2"].RichText.Add("สำนักธุรกิจ:"); boldRichText2.Bold = true;
                    boldRichText2 = ws11.Cells[$"L2"].RichText.Add("สำนักธุรกิจ:"); boldRichText2.Bold = true;
                }



                normalRichText = ws1.Cells[$"Q2"].RichText.Add(strDivorBC); normalRichText.Bold = false;
                normalRichText = ws2.Cells[$"O2"].RichText.Add(strDivorBC); normalRichText.Bold = false;
                normalRichText = ws3.Cells[$"S2"].RichText.Add(strDivorBC); normalRichText.Bold = false;
                normalRichText = ws4.Cells[$"P2"].RichText.Add(strDivorBC); normalRichText.Bold = false;
                normalRichText = ws5.Cells[$"I2"].RichText.Add(strDivorBC); normalRichText.Bold = false;
                normalRichText = ws6.Cells[$"G2"].RichText.Add(strDivorBC); normalRichText.Bold = false;
                normalRichText = ws7.Cells[$"R2"].RichText.Add(strDivorBC); normalRichText.Bold = false;
                normalRichText = ws8.Cells[$"S2"].RichText.Add(strDivorBC); normalRichText.Bold = false;
                normalRichText = ws9.Cells[$"S2"].RichText.Add(strDivorBC); normalRichText.Bold = false;
                normalRichText = ws10.Cells[$"E2"].RichText.Add(strDivorBC); normalRichText.Bold = false;
                normalRichText = ws11.Cells[$"L2"].RichText.Add(strDivorBC); normalRichText.Bold = false;


                //BOS Case ID,  BLWF Case ID
                string strBOSCaseID = " " + dtHeader.Rows[0]["case_id"].ToString() + "    ";
                string strBLWFCaseID = " " + dtHeader.Rows[0]["blwf_case_id"].ToString();

                boldRichText2 = ws1.Cells[$"Q3"].RichText.Add("BOS Case ID:"); boldRichText2.Bold = true;
                boldRichText2 = ws2.Cells[$"O3"].RichText.Add("BOS Case ID:"); boldRichText2.Bold = true;
                boldRichText2 = ws3.Cells[$"S3"].RichText.Add("BOS Case ID:"); boldRichText2.Bold = true;
                boldRichText2 = ws4.Cells[$"P3"].RichText.Add("BOS Case ID:"); boldRichText2.Bold = true;
                boldRichText2 = ws5.Cells[$"I3"].RichText.Add("BOS Case ID:"); boldRichText2.Bold = true;
                boldRichText2 = ws6.Cells[$"G3"].RichText.Add("BOS Case ID:"); boldRichText2.Bold = true;
                boldRichText2 = ws7.Cells[$"R3"].RichText.Add("BOS Case ID:"); boldRichText2.Bold = true;
                boldRichText2 = ws8.Cells[$"S3"].RichText.Add("BOS Case ID:"); boldRichText2.Bold = true;
                boldRichText2 = ws9.Cells[$"S3"].RichText.Add("BOS Case ID:"); boldRichText2.Bold = true;
                boldRichText2 = ws10.Cells[$"E3"].RichText.Add("BOS Case ID:"); boldRichText2.Bold = true;
                boldRichText2 = ws11.Cells[$"L3"].RichText.Add("BOS Case ID:"); boldRichText2.Bold = true;

                normalRichText = ws1.Cells[$"Q3"].RichText.Add(strBOSCaseID); normalRichText.Bold = false;
                normalRichText = ws2.Cells[$"O3"].RichText.Add(strBOSCaseID); normalRichText.Bold = false;
                normalRichText = ws3.Cells[$"S3"].RichText.Add(strBOSCaseID); normalRichText.Bold = false;
                normalRichText = ws4.Cells[$"P3"].RichText.Add(strBOSCaseID); normalRichText.Bold = false;
                normalRichText = ws5.Cells[$"I3"].RichText.Add(strBOSCaseID); normalRichText.Bold = false;
                normalRichText = ws6.Cells[$"G3"].RichText.Add(strBOSCaseID); normalRichText.Bold = false;
                normalRichText = ws7.Cells[$"R3"].RichText.Add(strBOSCaseID); normalRichText.Bold = false;
                normalRichText = ws8.Cells[$"S3"].RichText.Add(strBOSCaseID); normalRichText.Bold = false;
                normalRichText = ws9.Cells[$"S3"].RichText.Add(strBOSCaseID); normalRichText.Bold = false;
                normalRichText = ws10.Cells[$"E3"].RichText.Add(strBOSCaseID); normalRichText.Bold = false;
                normalRichText = ws11.Cells[$"L3"].RichText.Add(strBOSCaseID); normalRichText.Bold = false;

                boldRichText2 = ws1.Cells[$"Q3"].RichText.Add("BLWF Case ID:"); boldRichText2.Bold = true;
                boldRichText2 = ws2.Cells[$"O3"].RichText.Add("BLWF Case ID:"); boldRichText2.Bold = true;
                boldRichText2 = ws3.Cells[$"S3"].RichText.Add("BLWF Case ID:"); boldRichText2.Bold = true;
                boldRichText2 = ws4.Cells[$"P3"].RichText.Add("BLWF Case ID:"); boldRichText2.Bold = true;
                boldRichText2 = ws5.Cells[$"I3"].RichText.Add("BLWF Case ID:"); boldRichText2.Bold = true;
                boldRichText2 = ws6.Cells[$"G3"].RichText.Add("BLWF Case ID:"); boldRichText2.Bold = true;
                boldRichText2 = ws7.Cells[$"R3"].RichText.Add("BLWF Case ID:"); boldRichText2.Bold = true;
                boldRichText2 = ws8.Cells[$"S3"].RichText.Add("BLWF Case ID:"); boldRichText2.Bold = true;
                boldRichText2 = ws9.Cells[$"S3"].RichText.Add("BLWF Case ID:"); boldRichText2.Bold = true;
                boldRichText2 = ws10.Cells[$"E3"].RichText.Add("BLWF Case ID:"); boldRichText2.Bold = true;
                boldRichText2 = ws11.Cells[$"L3"].RichText.Add("BLWF Case ID:"); boldRichText2.Bold = true;

                normalRichText = ws1.Cells[$"Q3"].RichText.Add(strBLWFCaseID); normalRichText.Bold = false;
                normalRichText = ws2.Cells[$"O3"].RichText.Add(strBLWFCaseID); normalRichText.Bold = false;
                normalRichText = ws3.Cells[$"S3"].RichText.Add(strBLWFCaseID); normalRichText.Bold = false;
                normalRichText = ws4.Cells[$"P3"].RichText.Add(strBLWFCaseID); normalRichText.Bold = false;
                normalRichText = ws5.Cells[$"I3"].RichText.Add(strBLWFCaseID); normalRichText.Bold = false;
                normalRichText = ws6.Cells[$"G3"].RichText.Add(strBLWFCaseID); normalRichText.Bold = false;
                normalRichText = ws7.Cells[$"R3"].RichText.Add(strBLWFCaseID); normalRichText.Bold = false;
                normalRichText = ws8.Cells[$"S3"].RichText.Add(strBLWFCaseID); normalRichText.Bold = false;
                normalRichText = ws9.Cells[$"S3"].RichText.Add(strBLWFCaseID); normalRichText.Bold = false;
                normalRichText = ws10.Cells[$"E3"].RichText.Add(strBLWFCaseID); normalRichText.Bold = false;
                normalRichText = ws11.Cells[$"L3"].RichText.Add(strBLWFCaseID); normalRichText.Bold = false;

            }
        }
        private DataTable GetHeaderData()
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@itemid", _caseID);
            parameters.Add("@userid", _userID);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[param].[SpBBL_rpt_CAF_Header_Main]", parameters);
            return dt;
        }
    }
}