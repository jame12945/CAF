using CAFGenerator.Model;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection.Metadata.Ecma335;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.UI.WebControls;
using System.Xml;

namespace CAFGenerator.Classes
{
    public class CurrentSheetGenerator
    {


        private int minLimitIndex = 0;
        private int minRequestIndex = 0;
        private int requestIndex = 9;
        private int existingLimitIndex = 0;
        private int outstandingIndex = 0;
        private int requestAmountIndex = 0;
        private int mortagePledgeIndex = 0;
        private int counterIndex = 0;
        private int collateralRequestNumberIndex = 0;
        private int facilityNoIndex = 0;
        private int interestSuspenseIndex = 0;
        private int principleIndex = 0;
        private int interestIndex = 0;
        private int interestOrFeeRateIndex = 0;
        private int termIndex = 0;
        private int risklevelIndex = 0;

        private int minCollateralIndex = 0;
        private int collateralIndex = 9;
        private int collateralSTWIndex = 0;
        private int mortageIndex = 0;
        private int apprisalValueIndex = 0;
        private int appraiserNameIndex = 0;

        private int countRound = 1;

        private int countGrandTotalIndex = 0;
        private int minTotalIndex = 0;
        private int newRoleIndex = 0;
        private int minNewRoleIndex = 0;
        private int rankIndex = 0;
        private int interestOrFeeRateStringArrayIndex = 0;

        private List<int> shareLimitContent = new List<int>();
        private List<int> singleLimitContent = new List<int>();
        private List<int> totalLimitContent = new List<int>();
        private List<string> mergeSecurityContent = new List<string>();

        private bool isSingleLimitAdded = false;
        private bool isShareLimitAdded = false;
        private int requestComponentIndex = 0;


        private readonly string _caseID;
        private readonly string _userID;
        private readonly CreditApplicationType _creditApplicationType;
        private readonly string _unitCaption;
        private readonly ExcelPackage _excelPackage;
        private readonly int requestLenght;



        public CurrentSheetGenerator(string caseID, string userTD, CreditApplicationType creditApplicationType, string unitCaption, ExcelPackage excelPackage)
        {
            _caseID = caseID;
            _userID = userTD;
            _creditApplicationType = creditApplicationType;
            _unitCaption = unitCaption;
            _excelPackage = excelPackage;
            requestLenght = 55;
            GenerateSheet();
        }

        private void GenerateSheet()
        {
            ExcelWorksheet ws = _excelPackage.Workbook.Worksheets["วงเงินภาระหนี้ปัจจุบัน"];
            DataTable dtCaseCustomerGroup = GetCaseCustomerGroup();
            //DataTable dtHeader = dtHeader;
            ws.Cells["S5"].RichText.Add("หน่วย: ").Bold = true;
            ws.Cells["S5"].RichText.Add(_unitCaption).Bold = false;
            ws.Cells["S5"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            ws.Cells["D7"].Value = "วันที่";
            //add 19/6/2568
            DataTable dtHeader = GetHeaderData();
            //fix 19/6/2568
            string outstandingDate;
            string formattedOutstandingDate = "";
            string custGroupExistingTable = "";
            if (dtHeader != null && dtHeader.Rows.Count > 0)
            {
                outstandingDate = dtHeader.Rows[0]["outstanding_date"].ToString();
                custGroupExistingTable = dtHeader.Rows[0]["cust_group_existing_table"].ToString();
                if (!string.IsNullOrEmpty(outstandingDate))
                {
                    formattedOutstandingDate = DateTime.ParseExact(outstandingDate, "yy/MM/dd", CultureInfo.InvariantCulture).ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                }
            }
            // ws.Cells["D8"].Value = dtHeader.Rows.Count > 0 ? DateTime.ParseExact(dtHeader.Rows[0]["outstanding_date"].ToString(), "yy/MM/dd", CultureInfo.InvariantCulture).ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) : null;
            ws.Cells["D8"].Value = formattedOutstandingDate;
            ws.Cells["D8"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            if (dtCaseCustomerGroup != null && dtCaseCustomerGroup.Rows.Count > 0)
            {

                foreach (DataRow row in dtCaseCustomerGroup.Rows)
                {
                    minRequestIndex = requestIndex;
                    ws.Cells[$"A{requestIndex}:S{requestIndex}"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    ws.Cells[$"A{requestIndex}:S{requestIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    ws.Cells[$"A{requestIndex}:S{requestIndex}"].Style.Font.Bold = true;
                    ws.Cells[$"A{requestIndex}:S{requestIndex}"].Merge = true;
                    if (dtCaseCustomerGroup != null && dtCaseCustomerGroup.Rows.Count > 0)
                    {
                        ws.Cells[$"A{requestIndex}"].Value = row["CUST_DETAILS"].ToString();
                        ws.Cells[$"A{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        requestIndex++;
                        collateralIndex++;

                        //DataTable dtCurrentFacilityCollateral = GetCurrentFacilityCollateralDetail(row["COUNTERPARTY_ID"].ToString());

                        //Change layout to function 12/6/2568
                        GetFacilityCollateralDetail(
                            ws,
                            row["COUNTERPARTY_ID"].ToString(),
                            row["EXISTING_LIMIT"].ToString(),
                            row["OUTSTANDING"].ToString(),
                            row["INTEREST_SUSPENSE"].ToString(),
                            custGroupExistingTable
                            );



                    }
                    // ถ้าทำการทำให้แสดงแค่ User Case สุดท้าย
                }

                GetFacilityCollateralSummary(ws, custGroupExistingTable);

                //requestIndex++;
                for (int i = 9; i <= requestIndex; i++)
                {
                    ws.Row(i).Height = 24;
                }
            }
            else
            {
                Debug.WriteLine("case no data return from sp");
                DataTable checkCustomerGroup = CheckCustomerGroup();
                if (checkCustomerGroup != null && checkCustomerGroup.Rows.Count > 0)
                {
                    foreach (DataRow row in checkCustomerGroup.Rows)
                    {
                        Debug.WriteLine(row["count_case_custid"].ToString());
                        Debug.WriteLine(row["case_id"].ToString());
                        Debug.WriteLine(row["counterparty_id"].ToString());
                        Debug.WriteLine(row["grouping_id_child"] == DBNull.Value ? "Null JA" : row["grouping_id_child"].ToString());
                        Debug.WriteLine(dtHeader.Rows[0]["cust_group_existing_table"]);

                        if ((custGroupExistingTable == "true") /*&& (string.IsNullOrEmpty(row["grouping_id_child"].ToString())) */&& (Convert.ToInt32(row["count_case_custid"]) > 0))
                        {

                            Debug.WriteLine("push value for counterparty");

                            //foreach (DataRow row in dtCaseCustomerGroup.Rows)
                            //{
                            minRequestIndex = requestIndex;


                            //Change layout to function 12/6/2568
                            GetFacilityCollateralDetail(
                                ws,
                                row["counterparty_id"].ToString(),
                                "",
                                "",
                                "",
                                custGroupExistingTable
                                );



                            //}
                            // ถ้าทำการทำให้แสดงแค่ User Case สุดท้าย
                            //}

                            GetFacilityCollateralSummary(ws, custGroupExistingTable);

                            //requestIndex++;
                            for (int i = 9; i <= requestIndex; i++)
                            {
                                ws.Row(i).Height = 24;
                            }

                        }
                        else
                        {
                            Debug.WriteLine("grouping_id is not null why coming here");
                        }


                    }
                }
            }

            ws.Cells[$"A9:A{requestIndex}"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"A9:A{requestIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"B9:B{requestIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"C9:C{requestIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"D9:D{requestIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"E9:E{requestIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"F9:F{requestIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"G9:G{requestIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"H9:H{requestIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"I9:I{requestIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"J9:J{requestIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"K9:K{requestIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"P9:P{requestIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"Q9:Q{requestIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"R9:R{requestIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"S9:S{requestIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"A{requestIndex}:S{requestIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

        }





        /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //check customer 1 ราย , no group

        private DataTable CheckCustomerGroup()
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@transaction_id", _caseID);


            DataTable dt = SqlServerDBUtil.getDataTableFromQuery("SELECT  COUNT(t2.transaction_id) AS count_case_custid,  t.transaction_id AS case_id, t2.grouping_id AS grouping_id_child , t2.counterparty_id FROM [transaction] t LEFT JOIN transaction_hierarchy th ON t.transaction_id = th.transaction_id LEFT JOIN [transaction] t2 ON th.child_transaction_id = t2.transaction_id WHERE t.transaction_id = @transaction_id and t2.grouping_id is null GROUP BY   t.transaction_id,   t2.grouping_id,t2.counterparty_id", parameters);
            return dt;
        }

        //change layout to function out if --block 2
        private void GetFacilityCollateralSummary(ExcelWorksheet ws, string custGroupExistingTable)
        {
            if (custGroupExistingTable == "false")
            {
                ws.Cells[$"A{requestIndex}:B{requestIndex}"].Merge = true;
                ws.Cells[$"A{requestIndex}"].Style.Font.Bold = true;
                ws.Cells[$"A{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                ws.Cells[$"A{requestIndex}:J{requestIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
                ws.Cells[$"A{requestIndex}"].Value = "รวมทั้งหมด (Grand Total)";
                ws.Cells[$"C{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                // ws.Cells[$"C{requestIndex}"].Value = GetTotalExistingCustomerGroup().Rows.Count > 0 ? GetTotalExistingCustomerGroup().Rows[0]["EXISTING_LIMIT"] : null;
                double doubleExistingLimit_2;
                string existingLimitStr_2 = null;

                DataTable groupTable = GetTotalExistingCustomerGroup();
                if (groupTable.Rows.Count > 0 && groupTable != null)
                {
                    existingLimitStr_2 = groupTable.Rows[0]["EXISTING_LIMIT"].ToString();
                }

                if (double.TryParse(existingLimitStr_2, out doubleExistingLimit_2))
                {
                    ws.Cells[$"C{requestIndex}"].Value = doubleExistingLimit_2;
                }
                else
                {
                    Console.WriteLine($"Unable to convert '{existingLimitStr_2}' to a double.");
                }

                ws.Cells[$"D{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                // ws.Cells[$"D{requestIndex}"].Value = GetTotalExistingCustomerGroup().Rows.Count > 0 ? GetTotalExistingCustomerGroup().Rows[0]["OUTSTANDING"] : null;


                double doubleTotalOutstanding;
                string outstandingTotalStr = null;

                if (groupTable.Rows.Count > 0)
                {
                    outstandingTotalStr = groupTable.Rows[0]["OUTSTANDING"].ToString();
                }

                if (double.TryParse(outstandingTotalStr, out doubleTotalOutstanding))
                {
                    ws.Cells[$"D{requestIndex}"].Value = doubleTotalOutstanding;
                }
                else
                {
                    // Handle the case where the string is not a valid double
                    Console.WriteLine($"Unable to convert '{outstandingTotalStr}' to a double.");
                }

                ws.Cells[$"E{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                ws.Cells[$"E{requestIndex}"].Value = GetTotalExistingCustomerGroup().Rows.Count > 0 ? GetTotalExistingCustomerGroup().Rows[0]["INTEREST_SUSPENSE"] : null;
                ws.Cells[$"K{requestIndex}:P{requestIndex}"].Merge = true;
                ws.Cells[$"K{requestIndex}:P{requestIndex}"].Style.Font.Bold = true;
                ws.Cells[$"K{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                ws.Cells[$"K{requestIndex}"].Value = "รวมหลักประกันที่เป็นทรัพย์สิน (Grand Total)";
                ws.Cells[$"Q{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                // ws.Cells[$"Q{requestIndex}"].Value = GetExistingAllCustomerCollateralTotalGroup().Rows.Count > 0 ? GetExistingAllCustomerCollateralTotalGroup().Rows[0]["MORTAGE_PLEDGE_VALUE"] : null;
                double doubleTotalMortagePledgeValue;
                string mortagePledgeValueTotalStr = null;
                var groupTable_2 = GetExistingAllCustomerCollateralTotalGroup();
                if (groupTable_2.Rows.Count > 0)
                {
                    mortagePledgeValueTotalStr = groupTable_2.Rows[0]["MORTAGE_PLEDGE_VALUE"].ToString();
                }

                if (double.TryParse(mortagePledgeValueTotalStr, out doubleTotalMortagePledgeValue))
                {
                    ws.Cells[$"Q{requestIndex}"].Value = doubleTotalMortagePledgeValue;
                }
                else
                {
                    // Handle the case where the string is not a valid double
                    Console.WriteLine($"Unable to convert '{mortagePledgeValueTotalStr}' to a double.");
                }
                ws.Cells[$"R{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                //ws.Cells[$"R{requestIndex}"].Value = GetExistingAllCustomerCollateralTotalGroup().Rows.Count > 0 ? GetExistingAllCustomerCollateralTotalGroup().Rows[0]["APPRAISAL_VALUE"] : null;
                double doubleTotalAppraisalValue;
                string totalApprasalValueStr = null;
                if (groupTable_2.Rows.Count > 0)
                {
                    totalApprasalValueStr = groupTable_2.Rows[0]["APPRAISAL_VALUE"].ToString();
                }

                if (double.TryParse(totalApprasalValueStr, out doubleTotalAppraisalValue))
                {
                    ws.Cells[$"R{requestIndex}"].Value = doubleTotalAppraisalValue;
                }
                else
                {
                    // Handle the case where the string is not a valid double
                    Console.WriteLine($"Unable to convert '{totalApprasalValueStr}' to a double.");
                }
                ws.Cells[$"K{requestIndex}:R{requestIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells[$"S{requestIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;

                requestIndex++;
                ws.Cells[$"A{requestIndex}:B{requestIndex}"].Merge = true;
                ws.Cells[$"A{requestIndex}:J{requestIndex}"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
                ws.Cells[$"A{requestIndex}:S{requestIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells[$"K{requestIndex}:P{requestIndex}"].Merge = true;
                ws.Cells[$"K{requestIndex}:P{requestIndex}"].Style.Font.Bold = true;
                ws.Cells[$"K{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                ws.Cells[$"K{requestIndex}"].Value = "รวมการค้ำประกันโดยบุคคล/ นิติบุคคล (Grand Total)";
                ws.Cells[$"Q{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                //ws.Cells[$"Q{requestIndex}"].Value = GetExistingAllCustomerGuaranteeTotalGroup().Rows.Count > 0 ? GetExistingAllCustomerGuaranteeTotalGroup().Rows[0]["MORTAGE_PLEDGE_VALUE"] : null;
                double doubleTotalMortagePledgeValue2;
                string mortagePledgeValueTotalStr2 = null;
                var groupTable3 = GetExistingAllCustomerGuaranteeTotalGroup();
                if (groupTable3.Rows.Count > 0 && groupTable3 != null)
                {
                    mortagePledgeValueTotalStr2 = groupTable3.Rows[0]["MORTAGE_PLEDGE_VALUE"].ToString();
                }

                if (double.TryParse(mortagePledgeValueTotalStr2, out doubleTotalMortagePledgeValue2))
                {
                    ws.Cells[$"Q{requestIndex}"].Value = doubleTotalMortagePledgeValue2;
                }
                else
                {
                    // Handle the case where the string is not a valid double
                    Console.WriteLine($"Unable to convert '{mortagePledgeValueTotalStr2}' to a double.");
                }
                ws.Cells[$"R{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                //ws.Cells[$"R{requestIndex}"].Value = GetExistingAllCustomerGuaranteeTotalGroup().Rows.Count > 0 ? GetExistingAllCustomerGuaranteeTotalGroup().Rows[0]["APPRAISAL_VALUE"] : null;
                double doubleTotalApprialValue2;
                string totalApprialValueStr2 = null;
                if (groupTable3.Rows.Count > 0 && groupTable3 != null)
                {
                    totalApprialValueStr2 = groupTable3.Rows[0]["APPRAISAL_VALUE"].ToString();
                }

                if (double.TryParse(totalApprialValueStr2, out doubleTotalApprialValue2))
                {
                    ws.Cells[$"R{requestIndex}"].Value = doubleTotalApprialValue2;
                }
                else
                {
                    // Handle the case where the string is not a valid double
                    Console.WriteLine($"Unable to convert '{totalApprialValueStr2}' to a double.");
                }
                requestIndex++;
            }


            ws.Cells[$"A{requestIndex}:B{requestIndex}"].Merge = true;
            ws.Cells[$"A{requestIndex}"].Style.Font.Bold = true;
            ws.Cells[$"A{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells[$"A{requestIndex}"].Value = "(1) รวมสินเชื่อและภาระผูกพัน (ไม่รวม FWC)";
            ws.Cells[$"A{requestIndex}:B{requestIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
            ws.Cells[$"C{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            // ws.Cells[$"C{requestIndex}"].Value = GetCurrentApprovedFacilityGroup().Rows.Count > 0 ? GetCurrentApprovedFacilityGroup().Rows[0]["LIMIT"] : null;
            double doubleLimitValue;
            string lmitStr = null;
            DataTable approveFacilityTable = GetCurrentApprovedFacilityGroup();
            DataTable dtGetSummaryCurrentApprovedFacility = GetSummaryCurrentApprovedFacility();
            if (approveFacilityTable.Rows.Count > 0 && custGroupExistingTable == "false")
            {
                lmitStr = approveFacilityTable.Rows[0]["LIMIT"].ToString();
            }
            else
            {
                lmitStr = dtGetSummaryCurrentApprovedFacility.Rows.Count > 0 && dtGetSummaryCurrentApprovedFacility != null ? dtGetSummaryCurrentApprovedFacility.Rows[0]["limit"].ToString() : "";
            }

            if (double.TryParse(lmitStr, out doubleLimitValue))
            {
                ws.Cells[$"C{requestIndex}"].Value = doubleLimitValue;
            }
            else
            {
                Console.WriteLine($"Unable to convert '{lmitStr}' to a double.");
            }
            ws.Cells[$"D{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            //ws.Cells[$"D{requestIndex}"].Value = GetCurrentApprovedFacilityGroup().Rows.Count > 0 ? GetCurrentApprovedFacilityGroup().Rows[0]["OUTSTANDING"] : null;
            double doubleOustStanding3;
            string outStandingStr3 = null;
            if (approveFacilityTable.Rows.Count > 0 && custGroupExistingTable == "false")
            {
                outStandingStr3 = approveFacilityTable.Rows[0]["OUTSTANDING"].ToString();
            }
            else
            {
                outStandingStr3 = dtGetSummaryCurrentApprovedFacility.Rows.Count > 0 && dtGetSummaryCurrentApprovedFacility != null ? dtGetSummaryCurrentApprovedFacility.Rows[0]["outstanding"].ToString() : "";
            }

            if (double.TryParse(outStandingStr3, out doubleOustStanding3))
            {
                ws.Cells[$"D{requestIndex}"].Value = doubleOustStanding3;
            }
            else
            {
                Console.WriteLine($"Unable to convert '{outStandingStr3}' to a double.");
            }
            ws.Cells[$"E{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            if (custGroupExistingTable == "false")
            {
                ws.Cells[$"E{requestIndex}"].Value = approveFacilityTable.Rows.Count > 0 && approveFacilityTable != null ? approveFacilityTable.Rows[0]["INTEREST_SUSPENSE"] : "";
            }
            else
            {
                ws.Cells[$"E{requestIndex}"].Value = dtGetSummaryCurrentApprovedFacility.Rows.Count > 0 && dtGetSummaryCurrentApprovedFacility != null ? dtGetSummaryCurrentApprovedFacility.Rows[0]["INTEREST_SUSPENSE"].ToString() : "";
            }
            ws.Cells[$"F{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            if (custGroupExistingTable == "false")
            {
                ws.Cells[$"F{requestIndex}"].Value = approveFacilityTable.Rows.Count > 0 && approveFacilityTable != null ? approveFacilityTable.Rows[0]["PRINCIPLE"] : "";
            }
            else
            {
                ws.Cells[$"F{requestIndex}"].Value = dtGetSummaryCurrentApprovedFacility.Rows.Count > 0 && dtGetSummaryCurrentApprovedFacility != null ? dtGetSummaryCurrentApprovedFacility.Rows[0]["PRINCIPLE"].ToString() : "";
            }
            ws.Cells[$"G{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            if (custGroupExistingTable == "false")
            {
                ws.Cells[$"G{requestIndex}"].Value = approveFacilityTable.Rows.Count > 0 && approveFacilityTable != null ? approveFacilityTable.Rows[0]["INTEREST"] : "";
            }
            else
            {
                ws.Cells[$"G{requestIndex}"].Value = dtGetSummaryCurrentApprovedFacility.Rows.Count > 0 && dtGetSummaryCurrentApprovedFacility != null ? dtGetSummaryCurrentApprovedFacility.Rows[0]["INTEREST"].ToString() : "";
            }
            ws.Cells[$"K{requestIndex}:S{requestIndex}"].Merge = true;
            ws.Cells[$"C{requestIndex}:S{requestIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            requestIndex++;
            ws.Cells[$"A{requestIndex}:B{requestIndex}"].Merge = true;
            ws.Cells[$"A{requestIndex}"].Style.Font.Bold = true;
            ws.Cells[$"A{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells[$"A{requestIndex}"].Value = "(2) FWC - Notional Amount";
            ws.Cells[$"A{requestIndex}:B{requestIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
            ws.Cells[$"A{requestIndex}:B{requestIndex}"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
            ws.Cells[$"C{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            //ws.Cells[$"C{requestIndex}"].Value = GetCurrentApprovedFacilityGroup().Rows.Count > 0 ? GetCurrentApprovedFacilityGroup().Rows[1]["LIMIT"] : null;
            double doubleLimitValue2;
            string lmitStr2 = null;
            if (approveFacilityTable.Rows.Count > 0 && custGroupExistingTable == "false")
            {
                lmitStr2 = approveFacilityTable.Rows[1]["LIMIT"].ToString();
            }
            else
            {
                lmitStr2 = dtGetSummaryCurrentApprovedFacility.Rows.Count > 0 && dtGetSummaryCurrentApprovedFacility != null ? dtGetSummaryCurrentApprovedFacility.Rows[1]["LIMIT"].ToString(): "";
            }

            if (double.TryParse(lmitStr2, out doubleLimitValue2))
            {
                ws.Cells[$"C{requestIndex}"].Value = doubleLimitValue2;
            }
            else
            {
                Console.WriteLine($"Unable to convert '{lmitStr2}' to a double.");
            }
            ws.Cells[$"D{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            //ws.Cells[$"D{requestIndex}"].Value = GetCurrentApprovedFacilityGroup().Rows.Count > 0 ? GetCurrentApprovedFacilityGroup().Rows[1]["OUTSTANDING"] : null;
            double doubleOustStanding4;
            string outStandingStr4 = null;
            if (approveFacilityTable.Rows.Count > 0 && custGroupExistingTable == "false")
            {
                outStandingStr4 = approveFacilityTable.Rows[1]["OUTSTANDING"].ToString();
            }
            else
            {
                outStandingStr4 = dtGetSummaryCurrentApprovedFacility.Rows.Count > 0 && dtGetSummaryCurrentApprovedFacility != null ? dtGetSummaryCurrentApprovedFacility.Rows[1]["outstanding"].ToString():"";
            }

            if (double.TryParse(outStandingStr4, out doubleOustStanding4))
            {
                ws.Cells[$"D{requestIndex}"].Value = doubleOustStanding4;
            }
            else
            {
                Console.WriteLine($"Unable to convert '{outStandingStr4}' to a double.");
            }
            ws.Cells[$"E{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            if (custGroupExistingTable == "false")
            {
                ws.Cells[$"E{requestIndex}"].Value = approveFacilityTable.Rows.Count > 0 && approveFacilityTable != null ? approveFacilityTable.Rows[1]["INTEREST_SUSPENSE"] : "";
            }
            else
            {
                ws.Cells[$"E{requestIndex}"].Value = dtGetSummaryCurrentApprovedFacility.Rows.Count > 0 && dtGetSummaryCurrentApprovedFacility != null ? dtGetSummaryCurrentApprovedFacility.Rows[1]["INTEREST_SUSPENSE"].ToString():"";
            }

            ws.Cells[$"F{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            if (custGroupExistingTable == "false")
            {
                ws.Cells[$"F{requestIndex}"].Value = approveFacilityTable.Rows.Count > 0 && approveFacilityTable != null ? approveFacilityTable.Rows[1]["PRINCIPLE"] : "";
            }
            else
            {
                ws.Cells[$"F{requestIndex}"].Value = dtGetSummaryCurrentApprovedFacility.Rows.Count > 0 && dtGetSummaryCurrentApprovedFacility != null ? dtGetSummaryCurrentApprovedFacility.Rows[1]["PRINCIPLE"].ToString():"";
            }
            ws.Cells[$"G{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            if (custGroupExistingTable == "false")
            {
                ws.Cells[$"G{requestIndex}"].Value = approveFacilityTable.Rows.Count > 0 && approveFacilityTable != null ? approveFacilityTable.Rows[1]["INTEREST"] : "";
            }
            else
            {
                ws.Cells[$"G{requestIndex}"].Value = dtGetSummaryCurrentApprovedFacility.Rows.Count > 0 && dtGetSummaryCurrentApprovedFacility != null ? dtGetSummaryCurrentApprovedFacility.Rows[1]["INTEREST"].ToString():"";
            }
            ws.Cells[$"K{requestIndex}:S{requestIndex}"].Merge = true;
            ws.Cells[$"C{requestIndex}:S{requestIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            requestIndex++;
            ws.Cells[$"A{requestIndex}:B{requestIndex}"].Merge = true;
            ws.Cells[$"A{requestIndex}"].Style.Font.Bold = true;
            ws.Cells[$"A{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells[$"A{requestIndex}"].Value = "(3) FWC - 5%, IRS (1%), CCS (1%) Weighted Amount";
            ws.Cells[$"A{requestIndex}:B{requestIndex}"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
            ws.Cells[$"C{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            // ws.Cells[$"C{requestIndex}"].Value = GetCurrentApprovedFacilityGroup().Rows.Count > 0 ? GetCurrentApprovedFacilityGroup().Rows[2]["LIMIT"] : null;
            double doubleLimitValue3;
            string lmitStr3 = null;
            if (approveFacilityTable.Rows.Count > 0 && custGroupExistingTable == "false")
            {
                lmitStr3 = approveFacilityTable.Rows[2]["LIMIT"].ToString();
            }
            else
            {
                lmitStr3 = dtGetSummaryCurrentApprovedFacility.Rows.Count > 0 && dtGetSummaryCurrentApprovedFacility != null ? dtGetSummaryCurrentApprovedFacility.Rows[2]["LIMIT"].ToString(): "";
            }

            if (double.TryParse(lmitStr3, out doubleLimitValue3))
            {
                ws.Cells[$"C{requestIndex}"].Value = doubleLimitValue3;
            }
            else
            {
                Console.WriteLine($"Unable to convert '{lmitStr3}' to a double.");
            }
            ws.Cells[$"D{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            //ws.Cells[$"D{requestIndex}"].Value = GetCurrentApprovedFacilityGroup().Rows.Count > 0 ? GetCurrentApprovedFacilityGroup().Rows[2]["OUTSTANDING"] : null;
            double doubleOustStanding5;
            string outStandingStr5 = null;
            if (approveFacilityTable.Rows.Count > 0 && custGroupExistingTable == "false")
            {
                outStandingStr5 = approveFacilityTable.Rows[2]["OUTSTANDING"].ToString();
            }
            else
            {
                outStandingStr5 = dtGetSummaryCurrentApprovedFacility.Rows.Count > 0 && dtGetSummaryCurrentApprovedFacility != null ? dtGetSummaryCurrentApprovedFacility.Rows[2]["OUTSTANDING"].ToString() : "";
            }

            if (double.TryParse(outStandingStr5, out doubleOustStanding5))
            {
                ws.Cells[$"D{requestIndex}"].Value = doubleOustStanding5;
            }
            else
            {
                Console.WriteLine($"Unable to convert '{outStandingStr5}' to a double.");
            }
            ws.Cells[$"E{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            if (custGroupExistingTable == "false")
            {
                ws.Cells[$"E{requestIndex}"].Value = approveFacilityTable.Rows.Count > 0 && approveFacilityTable != null ? approveFacilityTable.Rows[2]["INTEREST_SUSPENSE"] : "";
            }
            else
            {
                ws.Cells[$"E{requestIndex}"].Value = dtGetSummaryCurrentApprovedFacility.Rows.Count > 0 && dtGetSummaryCurrentApprovedFacility != null ? dtGetSummaryCurrentApprovedFacility.Rows[2]["INTEREST_SUSPENSE"].ToString(): "";
            }
            ws.Cells[$"F{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            if (custGroupExistingTable == "false")
            {
                ws.Cells[$"F{requestIndex}"].Value = approveFacilityTable.Rows.Count > 0 && approveFacilityTable != null ? approveFacilityTable.Rows[2]["PRINCIPLE"] : "";
            }
            else
            {
                ws.Cells[$"F{requestIndex}"].Value = dtGetSummaryCurrentApprovedFacility.Rows.Count > 0 && dtGetSummaryCurrentApprovedFacility != null ? dtGetSummaryCurrentApprovedFacility.Rows[2]["PRINCIPLE"].ToString():"";
            }
            ws.Cells[$"G{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            if (custGroupExistingTable == "false")
            {
                ws.Cells[$"G{requestIndex}"].Value = approveFacilityTable.Rows.Count > 0 && approveFacilityTable != null ? approveFacilityTable.Rows[2]["INTEREST"] : "";
            }
            else
            {
                ws.Cells[$"G{requestIndex}"].Value = dtGetSummaryCurrentApprovedFacility.Rows.Count > 0 && dtGetSummaryCurrentApprovedFacility != null ? dtGetSummaryCurrentApprovedFacility.Rows[2]["INTEREST"].ToString():"";
            }
            ws.Cells[$"K{requestIndex}:S{requestIndex}"].Merge = true;
            ws.Cells[$"C{requestIndex}:J{requestIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"K{requestIndex}:S{requestIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;

            requestIndex++;
            ws.Cells[$"A{requestIndex}:B{requestIndex}"].Merge = true;
            ws.Cells[$"A{requestIndex}"].Style.Font.Bold = true;
            ws.Cells[$"A{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells[$"A{requestIndex}"].Value = "รวมสินเชื่อและภาระผูกพัน (1)+(2)";
            ws.Cells[$"A{requestIndex}:B{requestIndex}"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"A{requestIndex}:B{requestIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
            ws.Cells[$"C{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            // ws.Cells[$"C{requestIndex}"].Value = GetCurrentApprovedFacilityTotalGroup().Rows.Count > 0 ? GetCurrentApprovedFacilityTotalGroup().Rows[0]["LIMIT"] : null;
            double doubleLimitValue4;
            string lmitStr4 = null;
            var approveFacilityTotalTable = GetCurrentApprovedFacilityTotalGroup();
            if (approveFacilityTable.Rows.Count > 0 && custGroupExistingTable == "false")
            {
                lmitStr4 = approveFacilityTotalTable.Rows[0]["LIMIT"].ToString();
            }
            else
            {
                lmitStr4 = GetSummaryCurrentApprovedFacilityTotal().Rows[0]["LIMIT"].ToString();
            }

            if (double.TryParse(lmitStr4, out doubleLimitValue4))
            {
                ws.Cells[$"C{requestIndex}"].Value = doubleLimitValue4;
            }
            else
            {
                Console.WriteLine($"Unable to convert '{lmitStr4}' to a double.");
            }
            ws.Cells[$"D{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            //ws.Cells[$"D{requestIndex}"].Value = GetCurrentApprovedFacilityTotalGroup().Rows.Count > 0 ? GetCurrentApprovedFacilityTotalGroup().Rows[0]["OUTSTANDING"] : null;
            double doubleOustStanding6;
            string outStandingStr6 = null;
            if (approveFacilityTable.Rows.Count > 0 && custGroupExistingTable == "false")
            {
                outStandingStr6 = approveFacilityTotalTable.Rows[0]["OUTSTANDING"].ToString();
            }
            else
            {
                outStandingStr6 = GetSummaryCurrentApprovedFacilityTotal().Rows[0]["OUTSTANDING"].ToString();
            }

            if (double.TryParse(outStandingStr6, out doubleOustStanding6))
            {
                ws.Cells[$"D{requestIndex}"].Value = doubleOustStanding6;
            }
            else
            {
                Console.WriteLine($"Unable to convert '{outStandingStr6}' to a double.");
            }
            ws.Cells[$"E{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            DataTable dtGetCurrentApprovedFacilityTotalGroup = GetCurrentApprovedFacilityTotalGroup();
            DataTable dtGetSummaryCurrentApprovedFacilityTotal = GetSummaryCurrentApprovedFacilityTotal();
            if (custGroupExistingTable == "false")
            {

                ws.Cells[$"E{requestIndex}"].Value = dtGetCurrentApprovedFacilityTotalGroup.Rows.Count > 0  && dtGetCurrentApprovedFacilityTotalGroup != null ? dtGetCurrentApprovedFacilityTotalGroup.Rows[0]["INTEREST_SUSPENSE"] : "";
            }
            else
            {
                ws.Cells[$"E{requestIndex}"].Value = dtGetSummaryCurrentApprovedFacilityTotal.Rows.Count > 0 && dtGetSummaryCurrentApprovedFacilityTotal != null ? dtGetSummaryCurrentApprovedFacilityTotal.Rows[0]["INTEREST_SUSPENSE"].ToString() : "";
            }
            ws.Cells[$"F{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            if (custGroupExistingTable == "false")
            {
                ws.Cells[$"F{requestIndex}"].Value = dtGetCurrentApprovedFacilityTotalGroup.Rows.Count > 0  && dtGetCurrentApprovedFacilityTotalGroup  != null ? dtGetCurrentApprovedFacilityTotalGroup.Rows[0]["PRINCIPLE"] : "";
            }
            else
            {
                ws.Cells[$"F{requestIndex}"].Value = dtGetSummaryCurrentApprovedFacilityTotal.Rows.Count > 0 && dtGetSummaryCurrentApprovedFacilityTotal != null ? dtGetSummaryCurrentApprovedFacilityTotal.Rows[0]["PRINCIPLE"].ToString() : "";
            }
            ws.Cells[$"G{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            if (custGroupExistingTable == "false")
            {
                ws.Cells[$"G{requestIndex}"].Value = dtGetCurrentApprovedFacilityTotalGroup.Rows.Count > 0 && dtGetCurrentApprovedFacilityTotalGroup != null  ? dtGetCurrentApprovedFacilityTotalGroup.Rows[0]["INTEREST"] : "";
            }
            else
            {
                ws.Cells[$"G{requestIndex}"].Value = dtGetSummaryCurrentApprovedFacilityTotal.Rows.Count > 0 && dtGetSummaryCurrentApprovedFacilityTotal != null ? dtGetSummaryCurrentApprovedFacilityTotal.Rows[0]["INTEREST"].ToString() : "";
            }
            ws.Cells[$"K{requestIndex}:S{requestIndex}"].Merge = true;
            ws.Cells[$"C{requestIndex}:J{requestIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"K{requestIndex}:S{requestIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;

            requestIndex++;
            ws.Cells[$"A{requestIndex}:B{requestIndex}"].Merge = true;
            ws.Cells[$"A{requestIndex}"].Style.Font.Bold = true;
            ws.Cells[$"A{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells[$"A{requestIndex}"].Value = "รวมสินเชื่อและภาระผูกพัน (1)+(3)";
            ws.Cells[$"A{requestIndex}:B{requestIndex}"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
            ws.Cells[$"A{requestIndex}:B{requestIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"C{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            //ws.Cells[$"C{requestIndex}"].Value = GetCurrentApprovedFacilityTotalGroup().Rows.Count > 0 ? GetCurrentApprovedFacilityTotalGroup().Rows[1]["LIMIT"] : null;
            double doubleLimitValue5;
            string lmitStr5 = null;
            if (approveFacilityTable.Rows.Count > 0 && custGroupExistingTable == "false")
            {
                lmitStr5 = approveFacilityTotalTable.Rows[1]["LIMIT"].ToString();
            }
            else
            {
                lmitStr5 = dtGetSummaryCurrentApprovedFacilityTotal.Rows[1]["LIMIT"].ToString();
            }

            if (double.TryParse(lmitStr5, out doubleLimitValue5))
            {
                ws.Cells[$"C{requestIndex}"].Value = doubleLimitValue5;
            }
            else
            {
                Console.WriteLine($"Unable to convert '{lmitStr5}' to a double.");
            }
            ws.Cells[$"D{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            //ws.Cells[$"D{requestIndex}"].Value = GetCurrentApprovedFacilityTotalGroup().Rows.Count > 0 ? GetCurrentApprovedFacilityTotalGroup().Rows[1]["OUTSTANDING"] : null;
            double doubleOustStanding7;
            string outStandingStr7 = null;
            if (approveFacilityTable.Rows.Count > 0 && custGroupExistingTable == "false")
            {
                outStandingStr7 = approveFacilityTotalTable.Rows[1]["OUTSTANDING"].ToString();
            }
            else
            {
                outStandingStr7 = dtGetSummaryCurrentApprovedFacilityTotal.Rows[1]["OUTSTANDING"].ToString();
            }

            if (double.TryParse(outStandingStr7, out doubleOustStanding7))
            {
                ws.Cells[$"D{requestIndex}"].Value = doubleOustStanding7;
            }
            else
            {
                Console.WriteLine($"Unable to convert '{outStandingStr7}' to a double.");
            }
            ws.Cells[$"E{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            if (custGroupExistingTable == "false")
            {
                ws.Cells[$"E{requestIndex}"].Value = dtGetCurrentApprovedFacilityTotalGroup.Rows.Count > 0 && dtGetCurrentApprovedFacilityTotalGroup != null  ? dtGetCurrentApprovedFacilityTotalGroup.Rows[1]["INTEREST_SUSPENSE"] : "";
            }
            else
            {
                ws.Cells[$"E{requestIndex}"].Value = dtGetSummaryCurrentApprovedFacilityTotal.Rows.Count > 0 && dtGetSummaryCurrentApprovedFacilityTotal != null ? dtGetSummaryCurrentApprovedFacilityTotal.Rows[1]["INTEREST_SUSPENSE"].ToString() : "";
            }
            ws.Cells[$"F{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            if (custGroupExistingTable == "false")
            {
                ws.Cells[$"F{requestIndex}"].Value = dtGetCurrentApprovedFacilityTotalGroup.Rows.Count > 0 && dtGetCurrentApprovedFacilityTotalGroup != null ? dtGetCurrentApprovedFacilityTotalGroup.Rows[1]["PRINCIPLE"] : "";
            }
            else
            {
                ws.Cells[$"F{requestIndex}"].Value = dtGetSummaryCurrentApprovedFacilityTotal.Rows.Count > 0 && dtGetSummaryCurrentApprovedFacilityTotal != null ? dtGetSummaryCurrentApprovedFacilityTotal.Rows[1]["PRINCIPLE"].ToString() : "";
            }
            ws.Cells[$"G{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            if (custGroupExistingTable == "false")
            {
                ws.Cells[$"G{requestIndex}"].Value = dtGetCurrentApprovedFacilityTotalGroup.Rows.Count > 0 && dtGetCurrentApprovedFacilityTotalGroup != null ? dtGetCurrentApprovedFacilityTotalGroup.Rows[1]["INTEREST"] : "";
            }
            else
            {
                ws.Cells[$"G{requestIndex}"].Value = dtGetSummaryCurrentApprovedFacilityTotal.Rows.Count > 0 && dtGetSummaryCurrentApprovedFacilityTotal != null ? dtGetSummaryCurrentApprovedFacilityTotal.Rows[1]["INTEREST"].ToString() : "";
            }
            ws.Cells[$"K{requestIndex}:S{requestIndex}"].Merge = true;
            ws.Cells[$"K{requestIndex}:S{requestIndex}"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
            ws.Cells[$"C{requestIndex}:S{requestIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        }
        //change layout to function in if -- block 1
        private void GetFacilityCollateralDetail(
            ExcelWorksheet ws,
            string counterpartyID,
            string existingLimit,
            string outstanding,
            string interestSuspense,
            string custGroupExistingTable
)
        {


            Debug.WriteLine($"counterpartyID: {ws}");
            Debug.WriteLine($"counterpartyID: {counterpartyID}");
            Debug.WriteLine($"existingLimit: {existingLimit}");
            Debug.WriteLine($"interestSuspense: {interestSuspense}");

            DataTable dtCurrentFacilityCollateral = GetCurrentFacilityCollateralDetail(counterpartyID);
            if (dtCurrentFacilityCollateral != null && dtCurrentFacilityCollateral.Rows.Count > 0)
            {

                foreach (DataRow eachCurrentFacilityCollateralRow in dtCurrentFacilityCollateral.Rows)
                {

                    minRequestIndex = requestIndex;
                    ws.Cells[$"A{requestIndex}"].Value = eachCurrentFacilityCollateralRow["FACILITY_NO"].ToString();

                    string creditFacilityTypeValues = eachCurrentFacilityCollateralRow["CREDIT_FACILITY_TYPE"].ToString();
                    //ลบตั้งแต่ต้นทาง
                    if (creditFacilityTypeValues.Contains("เงื่อนไขเบิกใช:"))
                    {
                        string[] parts = creditFacilityTypeValues.Split(new string[] { "เงื่อนไขเบิกใช:" }, StringSplitOptions.None);
                        if (parts.Length > 1 && string.IsNullOrWhiteSpace(parts[1]))
                        {
                            creditFacilityTypeValues = parts[0].Trim();
                        }
                    }
                    string trimmedCreditFacilityTypeValues = Regex.Replace(creditFacilityTypeValues, @"[ \t\u00A0]+", " ");
                    string addBulletToTrimmedCreditFacilityTypeValues = Regex.Replace(trimmedCreditFacilityTypeValues, @"(\[)", "\u2022[");

                    IEnumerable<string> creditFacilityTypeStrings = Utility.NewLineSplitText(addBulletToTrimmedCreditFacilityTypeValues, 37);
                    if (eachCurrentFacilityCollateralRow["LIMIT"].ToString() == "SHARE_LIMIT")
                    {
                        //ใส่ลูกนำ้ใน eachCurrentFacilityCollateralRow["INTEREST RATE / FEE RATE"].ToString()
                        string originalText = eachCurrentFacilityCollateralRow["INTEREST RATE / FEE RATE"].ToString();
                        //string trimSpaceOriginalText = Regex.Replace(originalText, @"\s{2,}", "");
                        //string replacedText = Regex.Replace(originalText, @"\s{5,}", ",");
                        //Debug.WriteLine("replaceText: " + replacedText);
                        //List<string> interestOrFeeRateStringArray = replacedText.Split(',').ToList();
                        int maxSpaceStartIndex = -1;
                        int maxSpaceLength = 0;
                        int currentSpaceStartIndex = -1;
                        int currentSpaceLength = 0;
                        for (int i = 0; i < originalText.Length; i++)
                        {
                            if (char.IsWhiteSpace(originalText[i]))
                            {
                                if (currentSpaceStartIndex == -1)
                                {
                                    currentSpaceStartIndex = i;
                                }
                                currentSpaceLength++;
                            }
                            else
                            {
                                if (currentSpaceLength > maxSpaceLength)
                                {
                                    maxSpaceStartIndex = currentSpaceStartIndex;
                                    maxSpaceLength = currentSpaceLength;
                                }
                                currentSpaceStartIndex = -1;
                                currentSpaceLength = 0;
                            }
                        }

                        //add for ครบกำหนด/ ระยะเวลา 19/5/2568
                        int maxSpaceStartIndex_2 = -1;
                        int maxSpaceLength_2 = 0;
                        int currentSpaceStartIndex_2 = -1;
                        int currentSpaceLength_2 = 0;
                        string termText = eachCurrentFacilityCollateralRow["TERM"].ToString();

                        for (int i = 0; i < termText.Length; i++)
                        {
                            if (char.IsWhiteSpace(termText[i]))
                            {
                                if (currentSpaceStartIndex_2 == -1)
                                {
                                    currentSpaceStartIndex_2 = i;
                                }
                                currentSpaceLength_2++;
                            }
                            else
                            {
                                if (currentSpaceLength_2 > maxSpaceLength_2)
                                {
                                    maxSpaceStartIndex_2 = currentSpaceStartIndex_2;
                                    maxSpaceLength_2 = currentSpaceLength_2;
                                }
                                currentSpaceStartIndex_2 = -1;
                                currentSpaceLength_2 = 0;
                            }
                        }

                        if (maxSpaceStartIndex_2 != -1 && maxSpaceLength_2 > 1)
                        {
                            termText = termText.Remove(maxSpaceStartIndex_2, maxSpaceLength_2).Insert(maxSpaceStartIndex_2, ",");
                        }
                        /////////////////////////////////////

                        if (maxSpaceStartIndex != -1 && maxSpaceLength > 1)
                        {
                            originalText = originalText.Remove(maxSpaceStartIndex, maxSpaceLength).Insert(maxSpaceStartIndex, ",");
                        }

                        List<string> interestOrFeeRateStringArray = originalText.Split(',').ToList();
                        int interestOrFeeRateStringArrayIndex = 1;
                        //Debug.WriteLine($"interestOrFeeRateStringArray  <==> {string.Join(", ", interestOrFeeRateStringArray)}");
                        List<string> termStringArray = termText.Split(',').ToList();
                        termStringArray = termStringArray.Select(s => s.Trim()).ToList();
                        if (string.IsNullOrEmpty(termStringArray[0]))
                        {
                            termStringArray.RemoveAt(0);
                        }

                        //Debug.WriteLine($"termStringArray  <==> {string.Join(", ", termStringArray)}");
                        int termArrayIndex = 1;

                        foreach (string str in creditFacilityTypeStrings)
                        {
                            if (!string.IsNullOrEmpty(str))
                            {
                                if (str == "-" || str == "-,")
                                {
                                    continue;
                                }
                                if (str.Contains('\u2022'))
                                {
                                    if (interestOrFeeRateStringArrayIndex < interestOrFeeRateStringArray.Count)
                                    {
                                        interestOrFeeRateIndex = requestIndex;
                                        IEnumerable<string> interestOrFeeRateStrings = Utility.NewLineSplitText(interestOrFeeRateStringArray[interestOrFeeRateStringArrayIndex], 12);
                                        foreach (string interestOrFeeRateStr in interestOrFeeRateStrings)
                                        {
                                            ws.Cells[$"H{interestOrFeeRateIndex}"].Value = Regex.Replace(interestOrFeeRateStr, @"[ \t\u00A0]+", " ");//interestOrFeeRateStr;
                                            ws.Cells[$"H{interestOrFeeRateIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                                            interestOrFeeRateIndex++;
                                        }
                                        interestOrFeeRateStringArrayIndex++;
                                    }
                                    if (termArrayIndex < termStringArray.Count)
                                    {
                                        termIndex = requestIndex;
                                        IEnumerable<string> termStrings = Utility.NewLineSplitText(termStringArray[termArrayIndex], 10);
                                        foreach (string termStr in termStrings)
                                        {
                                            ws.Cells[$"I{termIndex}"].Value = termStr;
                                            ws.Cells[$"I{termIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                                            termIndex++;
                                        }
                                        termArrayIndex++;


                                    }

                                    if (str.Contains('[') || str.Contains(']'))
                                    {
                                        string strWithoutBracket = str.Replace("[", "").Replace("]", "");
                                        ws.Cells[$"B{requestIndex}"].Value = strWithoutBracket;
                                    }
                                }
                                else
                                {

                                    if (str.Contains('[') || str.Contains(']'))
                                    {
                                        string strWithoutBracket = str.Replace("[", "").Replace("]", "");
                                        ws.Cells[$"B{requestIndex}"].Value = strWithoutBracket;
                                    }
                                    else
                                    {
                                        ws.Cells[$"B{requestIndex}"].Value = str;
                                    }
                                }
                                requestIndex++;
                            }
                        }
                        requestIndex = Math.Max(requestIndex, interestIndex);

                    }
                    else
                    {

                        //ลองปรับเป็น requestComponent function
                        requestComponentIndex = RequestComponent(ws, eachCurrentFacilityCollateralRow, requestIndex);
                        foreach (string str in creditFacilityTypeStrings)
                        {
                            if (!string.IsNullOrEmpty(str))
                            {
                                if (str == "-" || str == "-,")
                                {
                                    continue;
                                }
                                //if(str.Contains("เงื่อนไขเบิกใช:")) // other condition สามารถ Add ได้
                                //{
                                //    string[] parts = str.Split(':');
                                //    if(parts.Length > 1 && string.IsNullOrWhiteSpace(parts[1]))
                                //    {
                                //        break;
                                //    }
                                //}
                                ws.Cells[$"B{requestIndex}"].Value = str;
                                //add
                                requestIndex++;
                            }
                        }
                        requestIndex = Math.Max(requestIndex, requestComponentIndex);
                    }



                    existingLimitIndex = minRequestIndex;
                    IEnumerable<string> ExistingLimitStrings = Utility.NewLineSplitText(eachCurrentFacilityCollateralRow["EXISTING_LIMIT"].ToString(), 10);
                    foreach (string str in ExistingLimitStrings)
                    {
                        if (!string.IsNullOrEmpty(str))
                        {
                            double doubleStr;
                            if (double.TryParse(str, out doubleStr))
                            {
                                ws.Cells[$"C{existingLimitIndex}"].Value = doubleStr;
                                ws.Cells[$"C{existingLimitIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                                existingLimitIndex++;
                            }
                            else
                            {
                                Console.WriteLine($"Unable to convert '{str}' to a double.");
                            }
                        }
                    }

                    outstandingIndex = minRequestIndex;
                    string outstandingValues = eachCurrentFacilityCollateralRow["OUTSTANDING"].ToString();
                    List<string> outstandingValuesStrings = new List<string> { outstandingValues };
                    foreach (string str in outstandingValuesStrings)
                    {
                        if (!string.IsNullOrEmpty(str))
                        {
                            double doubleStr;
                            //ws.Cells[$"D{outstandingIndex}"].Value = str;
                            if (double.TryParse(str, out doubleStr))
                            {
                                ws.Cells[$"D{outstandingIndex}"].Value = doubleStr;
                                ws.Cells[$"D{outstandingIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                                outstandingIndex++;
                            }
                            else
                            {
                                Console.WriteLine($"Unable to convert '{str}' to a double.");
                            }
                        }
                    }

                    interestSuspenseIndex = minRequestIndex;
                    IEnumerable<string> interestSuspenseStrings = Utility.NewLineSplitText(eachCurrentFacilityCollateralRow["INTEREST_SUSPENSE"].ToString(), 10);
                    foreach (string str in interestSuspenseStrings)
                    {
                        ws.Cells[$"E{interestSuspenseIndex}"].Value = str;
                        ws.Cells[$"E{interestSuspenseIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        interestSuspenseIndex++;
                    }

                    //Principle
                    principleIndex = minRequestIndex;
                    IEnumerable<string> pricipleStrings = Utility.NewLineSplitText(eachCurrentFacilityCollateralRow["PRINCIPLE"].ToString(), 10);
                    foreach (string str in pricipleStrings)
                    {
                        ws.Cells[$"F{principleIndex}"].Value = str;
                        ws.Cells[$"F{principleIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        principleIndex++;
                    }
                    //Interest
                    interestIndex = minRequestIndex;
                    IEnumerable<string> interestStrings = Utility.NewLineSplitText(eachCurrentFacilityCollateralRow["INTEREST"].ToString(), 10);
                    foreach (string str in interestStrings)
                    {
                        ws.Cells[$"G{interestIndex}"].Value = str;
                        ws.Cells[$"G{interestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        interestIndex++;
                    }


                    //Term
                    //termIndex = minRequestIndex;
                    //string termValue = eachCurrentFacilityCollateralRow["TERM"].ToString();
                    //string trimmedTermValue = Regex.Replace(termValue, ",", "");

                    //IEnumerable<string> termStrings = Utility.NewLineSplitText(trimmedTermValue, 10);
                    //foreach (string str in termStrings)
                    //{
                    //    if (!string.IsNullOrEmpty(str))
                    //    {
                    //        ws.Cells[$"I{termIndex}"].Value = str;
                    //        ws.Cells[$"I{termIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    //        termIndex++;
                    //    }
                    //}
                    //Risk Level
                    risklevelIndex = minRequestIndex;
                    IEnumerable<string> risklevelStrings = Utility.NewLineSplitText(eachCurrentFacilityCollateralRow["RISK LEVEL"].ToString(), 10);
                    foreach (string str in risklevelStrings)
                    {
                        ws.Cells[$"J{risklevelIndex}"].Value = str;
                        ws.Cells[$"J{risklevelIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        risklevelIndex++;
                    }

                    minCollateralIndex = collateralIndex;
                    int collateralRequestNoIndex = collateralIndex;
                    string collateralRequestNumberStrings = eachCurrentFacilityCollateralRow["COLLATERAL REQUEST NO."].ToString();
                    List<string> uniqueCollateralRequestNumber = collateralRequestNumberStrings
                                                                 .Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                                                                 .Distinct()
                                                                 .ToList();
                    string finalUniqueCollateralRequestNumber = string.Join(", ", uniqueCollateralRequestNumber);
                    List<string> finalUniqueCollateralRequestNumberList = finalUniqueCollateralRequestNumber.Split(new[] { ", " }, StringSplitOptions.RemoveEmptyEntries).ToList();
                    string finalUniqueCollateralRequestNumberString = string.Join(", ", finalUniqueCollateralRequestNumberList);
                    //foreach (string str in finalUniqueCollateralRequestNumberList)
                    IEnumerable<string> collateralRequestNumberStr = Utility.NewLineSplitText(finalUniqueCollateralRequestNumberString, 4);
                    foreach (string str in collateralRequestNumberStr)
                    {
                        ws.Cells[$"K{collateralRequestNoIndex}"].Value = str;
                        ws.Cells[$"K{collateralRequestNoIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        collateralRequestNoIndex++;
                    }


                    //collateral
                    //collateralIndex--;
                    string collateralSTWvalue = eachCurrentFacilityCollateralRow["COLLATERAL"].ToString();
                    string trimmedCollateralSTWvalue = Regex.Replace(collateralSTWvalue, @"[ \t\u00A0]+", " ");

                    IEnumerable<string> collateralSTWStrings = Utility.NewLineSplitText(trimmedCollateralSTWvalue, 38);
                    foreach (string str in collateralSTWStrings)
                    {
                        if (!string.IsNullOrEmpty(str))
                        {
                            //ws.Cells[$"M{collateralIndex}:O{collateralIndex}"].Merge = true;
                            ws.Cells[$"L{collateralIndex}"].Value = str;
                            collateralIndex++;
                        }
                    }

                    string[] collTypes = { "05", "06", "07", "23" };
                    string[] legalTypes = { "LAT_24", "LAT_26", "LAT_50", "LAT_65", "LAT_68", "LAT_102", "LAT_152" }; //can you tt.formula instead of lat but need to pull field from sp
                    if (eachCurrentFacilityCollateralRow["COLL_TYPE"] != null && eachCurrentFacilityCollateralRow["LEGAL_TYPE"] != null)
                    {
                        if (collTypes.Contains(eachCurrentFacilityCollateralRow["COLL_TYPE"].ToString()) && legalTypes.Contains(eachCurrentFacilityCollateralRow["LEGAL_TYPE"].ToString()))
                        //if (eachCurrentFacilityCollateralRow["MORTGAGE_TABLE"].ToString() == "YES")
                        {
                            DataTable dtCurrentMortgage = GetCurrentMortgage(eachCurrentFacilityCollateralRow["LINKAGE_ID"].ToString(), eachCurrentFacilityCollateralRow["COLLATERAL_ID"].ToString());
                            if (dtCurrentMortgage != null && dtCurrentMortgage.Rows.Count > 0)
                            {
                                ws.Cells[$"M{collateralIndex}"].Value = "ลำดับ";
                                ws.Cells[$"N{collateralIndex}"].Value = "มูลจำนอง";
                                ws.Cells[$"O{collateralIndex}"].Value = "ค้ำประกัน";
                                ws.Cells[$"M{collateralIndex}:O{collateralIndex}"].Style.Font.Bold = true;
                                ws.Cells[$"M{collateralIndex}:O{collateralIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                ws.Cells[$"M{collateralIndex}:O{collateralIndex}"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                ws.Cells[$"M{collateralIndex}:O{collateralIndex}"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                ws.Cells[$"M{collateralIndex}:O{collateralIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                collateralIndex++;

                                newRoleIndex = collateralIndex;
                                foreach (DataRow row2 in dtCurrentMortgage.Rows)
                                {
                                    IEnumerable<string> rankStrings = Utility.NewLineSplitText(row2["RANK"].ToString(), 10);
                                    foreach (string str in rankStrings)
                                    {
                                        ws.Cells[$"M{newRoleIndex}"].Value = str;
                                        ws.Cells[$"M{newRoleIndex}"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        ws.Cells[$"M{newRoleIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        ws.Cells[$"M{newRoleIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                    }


                                    IEnumerable<string> mortgageValueStrings = Utility.NewLineSplitText(row2["MORTGAGE_VALUE"].ToString(), 20);
                                    foreach (string str in mortgageValueStrings)
                                    {
                                        ws.Cells[$"N{newRoleIndex}"].Value = str;
                                        ws.Cells[$"N{newRoleIndex}"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        ws.Cells[$"N{newRoleIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        ws.Cells[$"N{newRoleIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;


                                    }
                                    string securityValue = row2["SECURITY"].ToString();
                                    string trimmedSecurityValue = Regex.Replace(securityValue, @"[ \t\u00A0]+", " ");
                                    // IEnumerable<string> securityStrings = Utility.NewLineSplitText(row2["SECURITY"].ToString(), 22);
                                    IEnumerable<string> securityStrings = Utility.NewLineSplitText(trimmedSecurityValue, 22);
                                    foreach (string str in securityStrings)
                                    {
                                        if (!string.IsNullOrEmpty(str))
                                        {
                                            //mergeSecurityContent.Add(str);
                                            ws.Cells[$"O{newRoleIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                            ws.Cells[$"O{newRoleIndex}"].Value = str;
                                            ws.Cells[$"O{newRoleIndex}"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                            ws.Cells[$"O{newRoleIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                            newRoleIndex++;
                                        }
                                    }
                                    ws.Cells[$"M{newRoleIndex - 1}:O{newRoleIndex - 1}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                }
                                //Debug.WriteLine("mergeSecurityContent: " + mergeSecurityContent);
                                ws.Cells[$"M{collateralIndex}:O{Math.Max(rankIndex - 1, newRoleIndex - 1)}"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws.Cells[$"M{collateralIndex}:O{Math.Max(rankIndex - 1, newRoleIndex - 1)}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                ws.Cells[$"M{collateralIndex}:O{Math.Max(rankIndex - 1, newRoleIndex - 1)}"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                ws.Cells[$"M{newRoleIndex}:O{newRoleIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                ws.Cells[$"M{newRoleIndex}:O{newRoleIndex}"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                ws.Cells[$"M{newRoleIndex}:O{newRoleIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                ws.Cells[$"M{newRoleIndex}:O{newRoleIndex}"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                ws.Cells[$"M{newRoleIndex}:O{newRoleIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                ws.Cells[$"M{newRoleIndex}"].Value = "Total";
                                ws.Cells[$"N{newRoleIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                ws.Cells[$"N{newRoleIndex}"].Value = GetTotalCurrentMortgage(eachCurrentFacilityCollateralRow["LINKAGE_ID"].ToString(), eachCurrentFacilityCollateralRow["COLLATERAL_ID"].ToString()).Rows.Count > 0 ? GetTotalCurrentMortgage(eachCurrentFacilityCollateralRow["LINKAGE_ID"].ToString(), eachCurrentFacilityCollateralRow["COLLATERAL_ID"].ToString()).Rows[0]["MORTGAGE_VALUE"].ToString() : null;
                                collateralIndex = Math.Max(rankIndex, newRoleIndex);
                                collateralIndex++;
                            }
                        }
                    }
                    mortageIndex = minCollateralIndex;
                    IEnumerable<string> mortagePledgeStrings = Utility.NewLineSplitText(eachCurrentFacilityCollateralRow["MORTAGE_PLEDGE VALUE"].ToString(), 10);
                    foreach (string str in mortagePledgeStrings)
                    {
                        if (!string.IsNullOrEmpty(str))
                        {
                            //ws.Cells[$"Q{mortageIndex}"].Value = str;
                            double doubleStr;
                            if (double.TryParse(str, out doubleStr))
                            {
                                ws.Cells[$"Q{mortageIndex}"].Value = doubleStr;
                                ws.Cells[$"Q{mortageIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                                mortageIndex++;
                            }
                            else
                            {
                                Console.WriteLine($"Unable to convert '{str}' to a double.");
                                ws.Cells[$"Q{mortageIndex}"].Value = str;
                                mortageIndex++;
                            }
                        }
                    }

                    apprisalValueIndex = minCollateralIndex;
                    IEnumerable<string> apprisalValueStrings = Utility.NewLineSplitText(eachCurrentFacilityCollateralRow["APPRISAL VALUE"].ToString(), 10);
                    foreach (string str in apprisalValueStrings)
                    {
                        if (!string.IsNullOrEmpty(str))
                        {
                            //ws.Cells[$"R{apprisalValueIndex}"].Value = str;
                            double doubleStr;
                            if (double.TryParse(str, out doubleStr))
                            {
                                ws.Cells[$"R{apprisalValueIndex}"].Value = doubleStr;
                                ws.Cells[$"R{apprisalValueIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                                apprisalValueIndex++;
                            }
                            else
                            {
                                Console.WriteLine($"Unable to convert '{str}' to a double.");
                                ws.Cells[$"R{apprisalValueIndex}"].Value = str;
                                apprisalValueIndex++;

                            }
                        }
                    }
                    appraiserNameIndex = minCollateralIndex;
                    IEnumerable<string> appraiserNameStrings = Utility.NewLineSplitText(eachCurrentFacilityCollateralRow["APPRAISER NAME"].ToString(), 8);
                    foreach (string str in appraiserNameStrings)
                    {
                        if (!string.IsNullOrEmpty(str))
                        {
                            ws.Cells[$"S{appraiserNameIndex}"].Value = str;
                            ws.Cells[$"S{appraiserNameIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            appraiserNameIndex++;
                        }
                    }

                    collateralIndex = Math.Max(collateralIndex, Math.Max(mortageIndex, Math.Max(apprisalValueIndex, appraiserNameIndex)));
                    requestIndex = Math.Max(requestIndex, Math.Max(existingLimitIndex, Math.Max(outstandingIndex, Math.Max(interestSuspenseIndex, Math.Max(principleIndex, Math.Max(interestIndex, Math.Max(interestOrFeeRateIndex, Math.Max(termIndex, risklevelIndex))))))));

                    // }



                }


            }


            requestIndex = Math.Max(requestIndex, collateralIndex);
            //facility หลังจบ loop : Index ถูก +1 (ขึ้นบรรทัดใหม่)
            ws.Cells[$"A{requestIndex}:B{requestIndex}"].Merge = true;
            ws.Cells[$"A{requestIndex}:S{requestIndex}"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"A{requestIndex}:J{requestIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
            //ws.Cells[$"A{requestIndex}:B{requestIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
            ws.Cells[$"A{requestIndex}:B{requestIndex}"].Style.Font.Bold = true;
            ws.Cells[$"A{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells[$"A{requestIndex}"].Value = "รวมรายลูกค้า";
            //ws.Cells[$"C{requestIndex}"].Value = row["EXISTING_LIMIT"].ToString();
            double doubleExistingLimit;
            DataTable dtGetCurrentApprovedFacility = GetCurrentApprovedFacility();
            //fix here
            if (custGroupExistingTable == "false")
            {
                Debug.WriteLine("Congratulation your data in group");
            }
            else
            {
                existingLimit = dtGetCurrentApprovedFacility.Rows.Count > 0 && dtGetCurrentApprovedFacility != null ? dtGetCurrentApprovedFacility.Rows[0]["EXISTING_LIMIT"].ToString() : "";
            }
            string existingLimitStr = existingLimit;
            if (double.TryParse(existingLimitStr, out doubleExistingLimit))
            {
                ws.Cells[$"C{requestIndex}"].Value = doubleExistingLimit;
            }
            else
            {
                Console.WriteLine($"Unable to convert '{existingLimitStr}' to a double.");
            }
            ws.Cells[$"C{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            // ws.Cells[$"D{requestIndex}"].Value = row["OUTSTANDING"].ToString();
            double doubleOutstanding;
            if (custGroupExistingTable == "false")
            {
                Debug.WriteLine("Congratulation your data in group");
            }
            else
            {
                outstanding = dtGetCurrentApprovedFacility.Rows.Count > 0 && dtGetCurrentApprovedFacility != null ? dtGetCurrentApprovedFacility.Rows[0]["OUTSTANDING"].ToString() : "";
            }
            string outstandingStr = outstanding;
            if (double.TryParse(outstandingStr, out doubleOutstanding))
            {
                ws.Cells[$"D{requestIndex}"].Value = doubleOutstanding;
            }
            else
            {
                Console.WriteLine($"Unable to convert '{outstandingStr}' to a double.");
            }
            ws.Cells[$"D{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            if (custGroupExistingTable == "false")
            {
                Debug.WriteLine("Congratulation your data in group");
            }
            else
            {
                interestSuspense = dtGetCurrentApprovedFacility.Rows.Count > 0 && dtGetCurrentApprovedFacility != null ? dtGetCurrentApprovedFacility.Rows[0]["INTEREST_SUSPENSE"].ToString(): "";
            }
            ws.Cells[$"E{requestIndex}"].Value = interestSuspense;
            ws.Cells[$"E{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            ws.Cells[$"K{requestIndex}:P{requestIndex}"].Merge = true;
            ws.Cells[$"K{requestIndex}:P{requestIndex}"].Style.Font.Bold = true;
            ws.Cells[$"K{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells[$"K{requestIndex}"].Value = "รวมหลักประกันที่เป็นทรัพย์สินรายลูกค้า";
            double doubleMortgagePledgeValue;
            string mortgagePledgeValueStr = null;
            // Check if the DataTable has rows and the value is not null
            var collateralTable = GetTotalExistingCustomerCollateral(counterpartyID);
            DataTable dtGetExistingAllCustomerCollateralTotal = GetExistingAllCustomerCollateralTotal();
            if (collateralTable.Rows.Count > 0)
            {
                if (custGroupExistingTable == "true")
                {
                    mortgagePledgeValueStr = dtGetExistingAllCustomerCollateralTotal.Rows.Count > 0 && dtGetExistingAllCustomerCollateralTotal != null ? dtGetExistingAllCustomerCollateralTotal.Rows[0]["MORTAGE_PLEDGE_VALUE"].ToString() : "";
                }
                else
                {
                    mortgagePledgeValueStr = collateralTable.Rows[0]["MORTAGE_PLEDGE_VALUE"].ToString();
                }
            }

            if (double.TryParse(mortgagePledgeValueStr, out doubleMortgagePledgeValue))
            {
                ws.Cells[$"Q{requestIndex}"].Value = doubleMortgagePledgeValue;
            }
            else
            {
                Console.WriteLine($"Unable to convert '{mortgagePledgeValueStr}' to a double.");
            }
            // ws.Cells[$"Q{requestIndex}"].Value = GetTotalExistingCustomerCollateral(row["COUNTERPARTY_ID"].ToString()).Rows.Count > 0 ? GetTotalExistingCustomerCollateral(row["COUNTERPARTY_ID"].ToString()).Rows[0]["MORTAGE_PLEDGE_VALUE"] : null;
            ws.Cells[$"Q{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            double doubleAppraisalValue;
            string appraisalValueStr = null;

            // Check if the DataTable has rows and the value is not null
            var collateralTable_2 = GetTotalExistingCustomerCollateral(counterpartyID);
            if (collateralTable_2.Rows.Count > 0)
            {

                if (custGroupExistingTable == "true")
                {
                    appraisalValueStr = dtGetExistingAllCustomerCollateralTotal.Rows.Count > 0 && dtGetExistingAllCustomerCollateralTotal != null ? dtGetExistingAllCustomerCollateralTotal.Rows[0]["APPRAISAL_VALUE"].ToString() : "";
                }
                else
                {
                    appraisalValueStr = collateralTable_2.Rows[0]["APPRAISAL_VALUE"].ToString();
                }
            }

            if (double.TryParse(appraisalValueStr, out doubleAppraisalValue))
            {
                ws.Cells[$"R{requestIndex}"].Value = doubleAppraisalValue;
            }
            else
            {
                // Handle the case where the string is not a valid double
                Console.WriteLine($"Unable to convert '{appraisalValueStr}' to a double.");
            }

            //ws.Cells[$"R{requestIndex}"].Value = GetTotalExistingCustomerCollateral(row["COUNTERPARTY_ID"].ToString()).Rows.Count > 0 ? GetTotalExistingCustomerCollateral(row["COUNTERPARTY_ID"].ToString()).Rows[0]["APPRAISAL_VALUE"] : null;
            ws.Cells[$"R{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            ws.Cells[$"K{requestIndex}:R{requestIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"S{requestIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
            requestIndex++;
            ws.Cells[$"A{requestIndex}:B{requestIndex}"].Merge = true;
            ws.Cells[$"A{requestIndex}:S{requestIndex}"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
            ws.Cells[$"A{requestIndex}:S{requestIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[$"K{requestIndex}:P{requestIndex}"].Merge = true;
            ws.Cells[$"K{requestIndex}:P{requestIndex}"].Style.Font.Bold = true;
            ws.Cells[$"K{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells[$"K{requestIndex}"].Value = "รวมการค้ำประกันโดยบุคคล/ นิติบุคคลรายลูกค้า";
            double doubleTotalMortgagePledgeValue;
            string mortgageTotalPledgeValueStr = null;

            // Check if the DataTable has rows and the value is not null
            var guaranteeTable = GetTotalExistingCustomerGurantee(counterpartyID);
            if (guaranteeTable.Rows.Count > 0 && guaranteeTable.Rows[0]["MORTAGE_PLEDGE_VALUE"] != DBNull.Value)
            {
                mortgageTotalPledgeValueStr = guaranteeTable.Rows[0]["MORTAGE_PLEDGE_VALUE"].ToString();
            }

            if (double.TryParse(mortgageTotalPledgeValueStr, out doubleTotalMortgagePledgeValue))
            {
                ws.Cells[$"Q{requestIndex}"].Value = doubleTotalMortgagePledgeValue;
            }
            else
            {
                // Handle the case where the string is not a valid double
                Console.WriteLine($"Unable to convert '{mortgagePledgeValueStr}' to a double.");
            }
            DataTable dtGetTotalExistingCustomerGurantee = GetTotalExistingCustomerGurantee(counterpartyID);
            //ws.Cells[$"Q{requestIndex}"].Value = GetTotalExistingCustomerGurantee(row["COUNTERPARTY_ID"].ToString()).Rows.Count > 0 ? GetTotalExistingCustomerGurantee(row["COUNTERPARTY_ID"].ToString()).Rows[0]["MORTAGE_PLEDGE_VALUE"] : null;
            ws.Cells[$"Q{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            ws.Cells[$"R{requestIndex}"].Value = dtGetTotalExistingCustomerGurantee.Rows.Count > 0 && dtGetTotalExistingCustomerGurantee != null ? dtGetTotalExistingCustomerGurantee.Rows[0]["APPRAISAL_VALUE"] : "";
           
            ws.Cells[$"R{requestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            ws.Cells[$"S{requestIndex}"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.None;
            requestIndex++;
            collateralIndex = requestIndex;

        }

        //Single Customer
        private DataTable GetCurrentApprovedFacility()
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@itemid", _caseID);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[dbo].[SpBBL_rpt_EXISTING_ALL_CUSTOMER_TOTAL]", parameters);
            return dt;

        }
        private DataTable GetExistingAllCustomerCollateralTotal()
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@itemid", _caseID);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[dbo].[SpBBL_rpt_EXISTING_ALL_CUSTOMER_COLLATERAL_TOTAL]", parameters);
            return dt;

        }
        private DataTable GetSummaryCurrentApprovedFacility()
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@itemid", _caseID);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[dbo].[SpBBL_rpt_CAF_CURRENT_APPROVED_FACILITY]", parameters);
            return dt;

        }
        private DataTable GetSummaryCurrentApprovedFacilityTotal()
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@itemid", _caseID);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[dbo].[SpBBL_rpt_CAF_CURRENT_APPROVED_FACILITY_TOTAL ]", parameters);
            return dt;

        }
        // Group Start//
        //fixed
        private DataTable GetHeaderData()
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@itemid", _caseID);
            parameters.Add("@userid", _userID);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[dbo].[SpBBL_rpt_CAF_Header]", parameters);
            return dt;

        }
        private DataTable GetCaseCustomerGroup()
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@itemid", _caseID);
            parameters.Add("@userid", _userID);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[param].[spBBL_rpt_GET_CAF_GROUP_CUST_LIST]", parameters);
            return dt;
        }
        // counterparty get field  for param from spBBL_rpt_GET_CAF_GROUP_CUST_LIST
        // ปรับ schema [dbo] => [paaram] เพื่อ add coll_type เข้าไป
        //ต้องเอาไปใส่ใน folder ปรับ create or alter
        private DataTable GetCurrentFacilityCollateralDetail(string counterpartyID)
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@counterparty_id", counterpartyID);
            parameters.Add("@itemid", _caseID);
            parameters.Add("@userid", _userID);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[param].[SpBBL_rpt_GRP_CURRENT_FACILITY_COLLATERAL_DETAILS]", parameters);
            return dt;
        }
        //fixed
        private DataTable GetTotalExistingCustomerCollateral(string counterpartyID)
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@counterparty_id", counterpartyID);
            parameters.Add("@itemid", _caseID);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[param].[SpBBL_RPT_EXISTING_CUSTOMER_COLLATERAL_TOTAL]", parameters);
            return dt;
        }
        //fixed
        private DataTable GetTotalExistingCustomerGurantee(string counterpartyID)
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@counterparty_id", counterpartyID);
            parameters.Add("@itemid", _caseID);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[param].[SpBBL_RPT_EXISTING_CUSTOMER_GURANTEE_TOTAL]", parameters);
            return dt;
        }
        //fixed
        private DataTable GetTotalExistingCustomerGroup()
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@itemid", _caseID);
            parameters.Add("@userid", _userID);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[param].[SpBBL_RPT_EXISTING_ALL_CUSTOMER_TOTAL_GROUP]", parameters);
            return dt;
        }
        //fixed
        private DataTable GetExistingAllCustomerCollateralTotalGroup()
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@itemid", _caseID);
            parameters.Add("@userid", _userID);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[param].[SpBBL_RPT_EXISTING_ALL_CUSTOMER_COLLATERAL_TOTAL_GROUP]", parameters);
            return dt;
        }
        //fixed
        private DataTable GetExistingAllCustomerGuaranteeTotalGroup()
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@itemid", _caseID);
            parameters.Add("@userid", _userID);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[param].[SpBBL_RPT_EXISTING_ALL_CUSTOMER_GURANTEE_TOTAL_GROUP]", parameters);
            return dt;
        }
        //fixed
        private DataTable GetCurrentApprovedFacilityGroup()
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@itemid", _caseID);
            parameters.Add("@userid", _userID);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[param].[SpBBL_RPT_CAF_CURRENT_APPROVED_FACILITY_GROUP]", parameters);
            return dt;
        }
        //fixed
        private DataTable GetCurrentApprovedFacilityTotalGroup()
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@itemid", _caseID);
            parameters.Add("@userid", _userID);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[param].[SpBBL_RPT_CAF_CURRENT_APPROVED_FACILITY_TOTAL_GROUP]", parameters);
            return dt;
        }

        private DataTable GetCurrentMortgage(string linkageId, string collateralId)
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@Linkage_id", linkageId);
            parameters.Add("@Collateral_id", collateralId);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[param].[SpBBL_rpt_CURRENT_MORTGAGE_TABLE]", parameters);
            return dt;
        }
        //fixed
        private DataTable GetTotalCurrentMortgage(string linkageId, string collateralId)
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@Linkage_id", linkageId);
            parameters.Add("@Collateral_id", collateralId);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[param].[SpBBL_RPT_CURRENT_MORTGAGE_TABLE_TOTAL]", parameters);
            return dt;
        }
        // Group End//

        private int RequestComponent(ExcelWorksheet ws, DataRow drRequest, int requestComponentIndex)
        {
            int requestAmountIndex;
            int interestIndex;
            int termIndex;

            requestAmountIndex = requestComponentIndex;
            IEnumerable<string> RequestAmountStrings = Utility.NewLineSplitText(drRequest["EXISTING_LIMIT"].ToString(), 12);
            foreach (string str in RequestAmountStrings)
            {
                if (!string.IsNullOrEmpty(str))
                {
                    ws.Cells[$"C{requestAmountIndex}"].Value = str;
                    requestAmountIndex++;
                }
            }

            interestIndex = requestComponentIndex;
            IEnumerable<string> interestOrFeeRateStrings = Utility.NewLineSplitText(drRequest["INTEREST RATE / FEE RATE"].ToString(), 12);
            foreach (string str in interestOrFeeRateStrings)
            {
                ws.Cells[$"H{interestIndex}"].Value = Regex.Replace(str, @"[ \t\u00A0]+", " ");//str;
                ws.Cells[$"H{interestIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                interestIndex++;
            }

            termIndex = requestComponentIndex;
            IEnumerable<string> termStrings = Utility.NewLineSplitText(drRequest["TERM"].ToString(), 10);
            foreach (string str in termStrings)
            {
                if (!string.IsNullOrEmpty(str))
                {
                    ws.Cells[$"I{termIndex}"].Value = str;
                    ws.Cells[$"I{termIndex}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                    termIndex++;
                }
            }

            requestComponentIndex = Math.Max(requestComponentIndex, Math.Max(interestIndex, Math.Max(requestAmountIndex, termIndex)));
            return requestComponentIndex;
        }

    }
}


