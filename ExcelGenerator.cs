using CAFGenerator.Model;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;

namespace CAFGenerator.Classes
{
    public class ExcelGenerator
    {
        private readonly string _caseID;
        private readonly string _userID;
        private readonly CreditApplicationType _creditApplicationType;
        private readonly string _unitCaption;
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();

        public ExcelGenerator(string caseID, string userID)
        {
            _caseID = caseID;
            _userID = userID;
            _creditApplicationType = GetCAType();
            _unitCaption = GetUnitCaption();
            HttpContext.Current.Session["non_split_words"] = GetNonSplitTextList();
        }

        public bool GenerateCAF(string TemplateFileName, string TemplatePath, string ExportFileName, string ExportPath, bool checkIsPricingBu)
        {
            string TrimedOaUserName = GetUserShortname().Trim().ToUpper();
            string getOaUserName = TrimedOaUserName.Split('\\').Last();

            Logger.Info($"Start - Case ID: {_caseID} User ID: {_userID}");
            if (!Directory.Exists(ExportPath))
            {
                Directory.CreateDirectory(ExportPath);
            }
            File.Copy(TemplatePath + TemplateFileName, ExportPath + ExportFileName, true);
            FileInfo fi = new FileInfo(ExportPath + ExportFileName);
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {

                Stopwatch stopWatch = new Stopwatch();
                stopWatch.Start();
                if (checkIsPricingBu == true)
                {
                    //Header
                    try
                    {
                        _ = new HeaderGenerator(_caseID, _userID, _creditApplicationType, excelPackage);
                    }
                    catch (Exception ex)
                    {
                        Logger.Error(ex, $"Error generating Header - Case ID: {_caseID} User ID: {_userID}");
                    }
                    stopWatch.Stop();
                    TimeSpan ts = stopWatch.Elapsed;
                    string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                    Logger.Debug($"Header Duration: {elapsedTime} - Case ID: {_caseID} User ID: {_userID}");

                    //09 Income and Deposit
                    stopWatch.Restart();
                    try
                    {
                        _ = new IncomeAndDepositSheetGenerator(_caseID, _userID, _creditApplicationType, checkIsPricingBu, excelPackage);
                    }
                    catch (Exception ex)
                    {
                        Logger.Error(ex, $"Error Income and Deposit - Case ID: {_caseID} User ID: {_userID} Username: {getOaUserName}");
                    }

                    finally
                    {
                        string[] templatesToDelete = { "สรุปข้อมูลลูกค้า", "คำขอ", "วงเงินภาระหนี้ปัจจุบัน", "วงเงินสินเชื่อระหว่างอนุมัติ", "สรุปวงเงินภาระหนี้หลังอนุมัติ", "สรุปข้อมูลทางการเงิน",
                        "Financial Highlights", "ปริมาณธุรกิจ", "Term Sheet", "Information Sheet", "FS_Template", "IaD_Template", "TS_Template", "InfoSheet_Template"};
                        foreach (string templateName in templatesToDelete)
                        {
                            ExcelWorksheet wsTemplate = excelPackage.Workbook.Worksheets[templateName];
                            if (wsTemplate != null)
                            {
                                excelPackage.Workbook.Worksheets.Delete(wsTemplate);
                            }
                        }
                    }
                    stopWatch.Stop();
                    ts = stopWatch.Elapsed;
                    elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                    Logger.Debug($"Income and Deposit Duration: {elapsedTime} - Case ID: {_caseID} User ID: {_userID} Username: {getOaUserName}");
                }
                else
                {
                    //Header
                    try
                    {
                        _ = new HeaderGenerator(_caseID, _userID, _creditApplicationType, excelPackage);
                    }
                    catch (Exception ex)
                    {
                        Logger.Error(ex, $"Error generating Header - Case ID: {_caseID} User ID: {_userID}");
                    }
                    stopWatch.Stop();
                    TimeSpan ts = stopWatch.Elapsed;
                    string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                    Logger.Debug($"Header Duration: {elapsedTime} - Case ID: {_caseID} User ID: {_userID} Username: {getOaUserName}");

                    //01 CustomerInfo
                    stopWatch.Restart();
                    try
                    {
                        _ = new CustomerInfoSheetGenerator(_caseID, _userID, _creditApplicationType, excelPackage);
                    }
                    catch (Exception ex)
                    {
                        Logger.Error(ex, $"Error generating Customer Info - Case ID: {_caseID} User ID: {_userID}");
                    }
                    stopWatch.Stop();
                    ts = stopWatch.Elapsed;
                    elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                    Logger.Debug($"Customer Info Duration: {elapsedTime} - Case ID: {_caseID} User ID: {_userID} Username: {getOaUserName}");

                    //02 Request
                    stopWatch.Restart();
                    //try
                    //{
                    //    _ = new RequestSheetGenerator(_caseID, _userID, _creditApplicationType, _unitCaption, excelPackage);
                    //}
                    //catch (Exception ex)
                    //{
                    //    Logger.Error(ex, $"Error generating Request - Case ID: {_caseID} User ID: {_userID} ");
                    //}
                    stopWatch.Stop();
                    ts = stopWatch.Elapsed;
                    elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                    Logger.Debug($"Request Duration: {elapsedTime} - Case ID: {_caseID} User ID: {_userID} Username: {getOaUserName}");

                    //03 Existing
                    stopWatch.Restart();
                    try
                    {
                        _ = new CurrentSheetGenerator(_caseID, _userID, _creditApplicationType, _unitCaption, excelPackage);
                    }
                    catch (Exception ex)
                    {
                        Logger.Error(ex, $"Error generating Existing - Case ID: {_caseID} User ID: {_userID}");
                    }
                    stopWatch.Stop();
                    ts = stopWatch.Elapsed;
                    elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                    Logger.Debug($"Existing Duration: {elapsedTime} - Case ID: {_caseID} User ID: {_userID} Username: {getOaUserName}");

                    //04 WIP
                    stopWatch.Restart();
                    //try
                    //{
                    //    _ = new WIPSheetGenerator(_caseID, _userID, _creditApplicationType, _unitCaption, excelPackage);
                    //}
                    //catch (Exception ex)
                    //{
                    //    Logger.Error(ex, $"Error generating Request - Case ID: {_caseID} User ID: {_userID}");
                    //}
                    stopWatch.Stop();
                    ts = stopWatch.Elapsed;
                    elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                    Logger.Debug($"WIP Duration: {elapsedTime} - Case ID: {_caseID} User ID: {_userID} Username: {getOaUserName}");


                    //05 Approved
                    stopWatch.Restart();
                    try
                    {
                        _ = new ApprovedSheetGenerator(_caseID, _userID, _creditApplicationType, _unitCaption, excelPackage);
                    }
                    catch (Exception ex)
                    {
                        Logger.Error(ex, $"Error generating Approved - Case ID: {_caseID} User ID: {_userID}");
                    }
                    stopWatch.Stop();
                    ts = stopWatch.Elapsed;
                    elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                    Logger.Debug($"Approved Duration: {elapsedTime} - Case ID: {_caseID} User ID: {_userID} Username: {getOaUserName}");

                    //06 Financial Summary
                    stopWatch.Restart();
                    try
                    {
                        _ = new FinancialSummarySheetGenerator(_caseID, _userID, _creditApplicationType, _unitCaption, excelPackage);
                    }
                    catch (Exception ex)
                    {
                        Logger.Error(ex, $"Error generating Financial Summary - Case ID: {_caseID} User ID: {_userID}");
                    }
                    finally
                    {
                        ExcelWorksheet wsTemplate = excelPackage.Workbook.Worksheets["FS_Template"];
                        excelPackage.Workbook.Worksheets.Delete(wsTemplate);
                    }
                    stopWatch.Stop();
                    ts = stopWatch.Elapsed;
                    elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                    Logger.Debug($"Financial Summary Duration: {elapsedTime} - Case ID: {_caseID} User ID: {_userID} Username: {getOaUserName}");

                    //07 Financial Highlights
                    stopWatch.Restart();
                    //try
                    //{
                    //    _ = new FinancialHighlightSheetGenerator(_caseID, _userID, _creditApplicationType, _unitCaption, excelPackage);
                    //}
                    //catch (Exception ex)
                    //{
                    //    Logger.Error(ex, $"Error generating Financial Highlights - Case ID: {_caseID} User ID: {_userID}");
                    //}
                    stopWatch.Stop();
                    ts = stopWatch.Elapsed;
                    elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                    Logger.Debug($"Financial Highlights Duration: {elapsedTime} - Case ID: {_caseID} User ID: {_userID} Username: {getOaUserName}");

                    //08 BusinessVolume
                    stopWatch.Restart();
                    try
                    {
                        _ = new BusinessVolumeSheetGenerator(_caseID, _userID, _creditApplicationType, _unitCaption, excelPackage);
                    }
                    catch (Exception ex)
                    {
                        Logger.Error(ex, $"Error generating Business Volume - Case ID: {_caseID} User ID: {_userID}");
                    }
                    stopWatch.Stop();
                    ts = stopWatch.Elapsed;
                    elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                    Logger.Debug($"Business Volume Duration: {elapsedTime} - Case ID: {_caseID} User ID: {_userID} Username: {getOaUserName}");

                    //09 Income and Deposit
                    stopWatch.Restart();
                    try
                    {
                        _ = new IncomeAndDepositSheetGenerator(_caseID, _userID, _creditApplicationType, checkIsPricingBu, excelPackage);
                    }
                    catch (Exception ex)
                    {
                        Logger.Error(ex, $"Error Income and Deposit - Case ID: {_caseID} User ID: {_userID}");
                    }

                    finally
                    {
                        ExcelWorksheet wsTemplate = excelPackage.Workbook.Worksheets["IaD_Template"];
                        excelPackage.Workbook.Worksheets.Delete(wsTemplate);
                    }
                    stopWatch.Stop();
                    ts = stopWatch.Elapsed;
                    elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                    Logger.Debug($"Income and Deposit Duration: {elapsedTime} - Case ID: {_caseID} User ID: {_userID} Username: {getOaUserName}");


                    //10 Term Sheet
                    stopWatch.Restart();
                    try
                    {
                        _ = new TermsheetSheetGenerator(_caseID, _userID, _creditApplicationType, excelPackage);
                    }
                    catch (Exception ex)
                    {
                        Logger.Error(ex, $"Error Term Sheet - Case ID: {_caseID} User ID: {_userID}");
                    }
                    finally
                    {
                        ExcelWorksheet wsTemplate = excelPackage.Workbook.Worksheets["TS_Template"];
                        excelPackage.Workbook.Worksheets.Delete(wsTemplate);
                    }
                    stopWatch.Stop();
                    ts = stopWatch.Elapsed;
                    elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                    Logger.Debug($"Term Sheet Duration: {elapsedTime} - Case ID: {_caseID} User ID: {_userID} Username: {getOaUserName}");

                    //11 Information Sheet
                    stopWatch.Restart();
                    try
                    {
                        _ = new InformationSheetGenerator(_caseID, _userID, excelPackage);
                    }
                    catch (Exception ex)
                    {
                        Logger.Error(ex, $"Error Information - Case ID: {_caseID} User ID: {_userID}");
                    }
                    finally
                    {
                        ExcelWorksheet wsTemplate = excelPackage.Workbook.Worksheets["InfoSheet_Template"];
                        excelPackage.Workbook.Worksheets.Delete(wsTemplate);
                    }
                    stopWatch.Stop();
                    ts = stopWatch.Elapsed;
                    elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                    Logger.Debug($"Information Duration: {elapsedTime} - Case ID: {_caseID} User ID: {_userID} Username: {getOaUserName}");
                }
                excelPackage.Save();
                Logger.Info($"End - Case ID: {_caseID} User ID: {_userID}");
            }

            return true;
        }



        private CreditApplicationType GetCAType()
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@transaction_id", _caseID);
            DataTable dt = SqlServerDBUtil.getDataTableFromQuery("SELECT t.transaction_id FROM [transaction_hierarchy] th JOIN [transaction] t ON th.child_transaction_id = t.transaction_id AND th.link_type_id = 175 WHERE th.ddate IS NULL AND t.ddate IS NULL AND th.transaction_id = @transaction_id", parameters);
            if (dt != null && dt.Rows.Count > 1)
            {
                return CreditApplicationType.Group;
            }
            else
            {
                return CreditApplicationType.SingleCustomer;
            }
        }

        private string GetUnitCaption()
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@itemid", _caseID);
            DataTable dt = SqlServerDBUtil.getDataTableFromQuery("SELECT SDT.NAME_1 AS UNIT FROM [TRANSACTION] T LEFT JOIN STATIC_DATA_TABLE SDT ON T.RENEW_FREQUENCY = SDT.SHORTNAME AND TABLE_NAME = 'AMOUNT_UNIT' WHERE TRANSACTION_ID = @itemid", parameters);
            if (dt != null && dt.Rows.Count > 0)
            {
                return dt.Rows[0]["UNIT"].ToString();
            }
            else
            {
                return "";
            }
        }
        private static List<string> GetNonSplitTextList()
        {
            List<string> nonSplitTexts = new List<string>();
            DataTable dt = SqlServerDBUtil.getDataTableFromQuery("SELECT non_split_word FROM [param].[T_BOS_CAF_NON_SPLIT_WORD]", null);
            if (dt != null && dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    nonSplitTexts.Add(row["non_split_word"].ToString());
                }
            }
            return nonSplitTexts;
        }

        private string GetUserShortname()
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@userID", _userID);
            DataTable dt = SqlServerDBUtil.getDataTableFromQuery("select user_shortname from software_user where user_id = @userID", parameters);
            if (dt != null && dt.Rows.Count > 0)
            {
                return dt.Rows[0]["user_shortname"].ToString();
            }
            else
            {
                return "";
            }
        }

    }
}