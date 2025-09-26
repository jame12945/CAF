using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;

namespace CAFGenerator.Classes
{
    public class InformationSheetGenerator
    {
        private readonly string _caseId;
        private readonly string _userId;
        private readonly ExcelPackage _excelPackage;

        public InformationSheetGenerator(string caseId, string userId, ExcelPackage excelPackage)
        {
            _caseId = caseId;
            _userId = userId;
            _excelPackage = excelPackage;
            GenerateSheet();
        }

        private void GenerateSheet()
        {
            ExcelWorksheet ws = _excelPackage.Workbook.Worksheets["Information Sheet"];
            ExcelWorksheet wsTemplate = _excelPackage.Workbook.Worksheets["InfoSheet_Template"];
            DataTable dtCustList = GetCafCustListData();
            DataTable dtCafHeader = GetCafHeader();

            if (dtCustList != null && dtCustList.Rows.Count > 0)
            {
                int lineIndex = 5;
                int currentIndex = 0;
                int outsideCondIndex = 0;
                int countNum = 1;
                int minFollowIndex = 0;
                int newRoleIndex = 0;
                int finalIndex = 0;

                foreach (DataRow row in dtCustList.Rows)
                    if (dtCustList.Rows.IndexOf(row) == 0)
                    {
                        Debug.WriteLine("CountNumber 1 st Time <==>");

                        // var sourceRange = wsTemplate.Cells["A5:M71"];//ใช้ความสามารถลากคลุมแบบ excel
                        var sourceRange = wsTemplate.Cells["A5:M21"];
                        sourceRange.Copy(ws.Cells[$"A{lineIndex}:M{lineIndex}"]);

                        //Debug.WriteLine($"Copying from {sourceRange.Address} to A{lineIndex}:M{lineIndex}");

                        // sourceRange.Copy(ws.Cells[$"A{lineIndex}:M{lineIndex}"]);//return 1 


                        var informationValueColD = ws.Cells[$"D{lineIndex}"].Value?.ToString() ?? string.Empty;

                        var customerValueColE = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                           ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["CUSTOMER_NAME"].ToString()
                            : string.Empty;

                        var customerValue = ws.Cells[$"D{lineIndex}"].RichText;
                        customerValue.Clear();
                        customerValue.Add(informationValueColD);
                        customerValue.Add(customerValueColE).Bold = false;

                        var rmPrefixValue = ws.Cells[$"H{lineIndex}"].Value?.ToString() ?? string.Empty;
                        var rmValue = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                            ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["RM_NO"].ToString()
                            : string.Empty;

                        var rmTotal = ws.Cells[$"H{lineIndex}"].RichText;
                        rmTotal.Clear();
                        rmTotal.Add(rmPrefixValue);
                        rmTotal.Add($" {rmValue}").Bold = false;

                        lineIndex += 2;
                        var extendCreditFacilityValue = GetCafInformationCreditFacilityType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                            ? GetCafInformationCreditFacilityType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["EXTEND_CREDIT_FACILITY"].ToString()
                            : string.Empty;

                        ws.Cells[$"D{lineIndex}"].Value = extendCreditFacilityValue;

                        var noExtendCreditFacilityValue = GetCafInformationCreditFacilityType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                            ? GetCafInformationCreditFacilityType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["NO_EXTEND_CREDIT_FACILITY"].ToString()
                            : string.Empty;

                        ws.Cells[$"F{lineIndex}"].Value = noExtendCreditFacilityValue;

                        var reduceCreditFacilityValue = GetCafInformationCreditFacilityType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                          ? GetCafInformationCreditFacilityType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["REDUCE_CREDIT_FACILITY"].ToString()
                          : string.Empty;

                        ws.Cells[$"H{lineIndex}"].Value = reduceCreditFacilityValue;

                        var cancleValue = GetCafInformationCreditFacilityType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                         ? GetCafInformationCreditFacilityType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["CANCEL"].ToString()
                         : string.Empty;

                        ws.Cells[$"K{lineIndex}"].Value = cancleValue;

                        lineIndex += 3; // (อย่าลืมเรื่มรวมเอาไปลบในตอนสุดท้ายจาก Index ปัจจุบัน ให้เหลือเว้นได้ 1 ช่องระหว่างลูกค้า )
                        var tsr001 = GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                            ? GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TSR_001"].ToString()
                            : string.Empty;
                        ws.Cells[$"D{lineIndex}"].Value = tsr001;

                        var tsr002 = GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                            ? GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TSR_002"].ToString()
                            : string.Empty;
                        ws.Cells[$"F{lineIndex}"].Value = tsr002;

                        var tsr003 = GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                           ? GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TSR_003"].ToString()
                           : string.Empty;
                        ws.Cells[$"H{lineIndex}"].Value = tsr003;

                        var tsr004 = GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                           ? GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TSR_004"].ToString()
                           : string.Empty;
                        ws.Cells[$"K{lineIndex}"].Value = tsr004;

                        lineIndex += 2;
                        var tsr007 = GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                          ? GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TSR_007"].ToString()
                          : string.Empty;
                        ws.Cells[$"D{lineIndex}"].Value = tsr007;

                        var tsr008 = GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                            ? GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TSR_008"].ToString()
                            : string.Empty;
                        ws.Cells[$"F{lineIndex}"].Value = tsr008;

                        var tsr009 = GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                           ? GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TSR_009"].ToString()
                           : string.Empty;
                        ws.Cells[$"H{lineIndex}"].Value = tsr009;

                        lineIndex += 2;
                        var tsr005 = GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                        ? GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TSR_005"].ToString()
                        : string.Empty;
                        ws.Cells[$"D{lineIndex}"].Value = tsr005;

                        lineIndex += 2;
                        var tsr006 = GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                        ? GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TSR_006"].ToString()
                        : string.Empty;
                        ws.Cells[$"D{lineIndex}"].Value = tsr006;
                        lineIndex += 2;
                        ws.Cells[$"D{lineIndex}"].Value = tsr006;

                        lineIndex += 3;
                        var typeOfCreditReview = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                            ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TYPE_OF_CREDIT_REVIEW_PLAN"].ToString()
                            : string.Empty;
                        ws.Cells[$"D{lineIndex}"].Value = typeOfCreditReview;

                        var typeOfCreditReviewAdHoc = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                            ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TYPE_OF_CREDIT_REVIEW_AD_HOC"].ToString()
                            : string.Empty;
                        ws.Cells[$"F{lineIndex}"].Value = typeOfCreditReviewAdHoc;
                        currentIndex = lineIndex + 1; //currentIndex = 22

                        if (dtCafHeader != null && dtCafHeader.Rows.Count > 0 && dtCafHeader.Rows[0]["BUSINESS_UNIT_SHORTNAME"].ToString() == "SAM")
                        {
                            var sourceRange2 = wsTemplate.Cells["A22:M25"];
                            sourceRange2.Copy(ws.Cells[$"A{currentIndex}:M{currentIndex}"]);

                            lineIndex += 4; // lineIndex ปัจจุบัน = 21
                            var phaseStatus = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                    ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["PHASE_STATUS_SAM"].ToString()
                                                    : string.Empty;

                            ws.Cells[$"D{lineIndex}"].Value = phaseStatus;

                            var phaseStatusLegal = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["PHASE_STATUS_LEGAL"].ToString()
                                                   : string.Empty;
                            ws.Cells[$"F{lineIndex}"].Value = phaseStatusLegal;
                            currentIndex = lineIndex + 1; // เขียนแบบนี้ไม่ได้ทำให้ lineIndex เปลี่ยนแปลงเพราะต้องตั้งค่าใหม่อยู่ดี
                            var sourceRange3 = wsTemplate.Cells["A26:M30"];
                            sourceRange3.Copy(ws.Cells[$"A{currentIndex}:M{currentIndex}"]);
                            //Debug.WriteLine("Currently lineIndex =>" + lineIndex);//lineIndex => 25

                            lineIndex += 4;
                            var bblConnectedParty = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                           ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["BBL_CONNECTED_PARTY"].ToString()
                           : string.Empty;
                            ws.Cells[$"B{lineIndex}"].Value = bblConnectedParty;
                            //อีกเงื่อนไขหนึ่ง (เงื่อนไขที่ 2 )
                            currentIndex = lineIndex + 2;
                            var sourceRange4 = wsTemplate.Cells["A31:M37"];
                            sourceRange4.Copy(ws.Cells[$"A{currentIndex}:M{currentIndex}"]);

                            lineIndex += 3;
                            var transactionApprovedNo = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                          ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TRANSACTION_APPROVED_NO"].ToString()
                          : string.Empty;
                            ws.Cells[$"H{lineIndex}"].Value = transactionApprovedNo;

                            var transactionApprovedYes = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                          ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TRANSACTION_APPROVED_YES"].ToString()
                          : string.Empty;
                            ws.Cells[$"K{lineIndex}"].Value = transactionApprovedYes;

                            lineIndex += 3;
                            var specialConditionAnsNo = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                          ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["NO_SPECIAL_CONDITION_NO"].ToString()
                          : string.Empty;
                            ws.Cells[$"H{lineIndex}"].Value = specialConditionAnsNo;

                            var specialConditionAnsYes = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                          ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["NO_SPECIAL_CONDITION_YES"].ToString()
                          : string.Empty;
                            ws.Cells[$"K{lineIndex}"].Value = specialConditionAnsYes;
                            lineIndex += 3;
                        }
                        else
                        {
                            var sourceRange3 = wsTemplate.Cells["A26:M31"];
                            sourceRange3.Copy(ws.Cells[$"A{currentIndex}:M{currentIndex}"]);
                            lineIndex += 4;

                            var bblConnectedParty = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                    ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["BBL_CONNECTED_PARTY"].ToString()
                                                    : string.Empty;
                            ws.Cells[$"B{lineIndex}"].Value = bblConnectedParty;
                            lineIndex += 3;
                        }

                        outsideCondIndex = lineIndex;
                        var sourceRange5 = wsTemplate.Cells["A38:M66"];
                        sourceRange5.Copy(ws.Cells[$"A{outsideCondIndex}:M{outsideCondIndex}"]);


                        var getChidTranSactionID = row["CHILD_TRANSACTION_ID"] == DBNull.Value ? string.Empty : row["CHILD_TRANSACTION_ID"].ToString();
                        var sllLevel = string.Empty;
                        if (!string.IsNullOrEmpty(getChidTranSactionID))
                        {
                            var cafInfo = GetCafInformationData(getChidTranSactionID);
                            if (cafInfo.Rows.Count > 0)
                            {
                                sllLevel = cafInfo.Rows[0]["SLL_LEVEL"].ToString();
                            }
                        }

                        if (!string.IsNullOrEmpty(sllLevel) && (sllLevel.StartsWith("=") || sllLevel.StartsWith("+") || sllLevel.StartsWith("-")
                            || sllLevel.StartsWith("@") || sllLevel.StartsWith("\"")))
                        {
                            sllLevel = "'" + sllLevel;
                        }

                        var normalRichText = ws.Cells[$"C{lineIndex}"].RichText.Add(sllLevel);
                        normalRichText.FontName = "Arial Unicode MS";
                        normalRichText.Size = 14;
                        normalRichText.Bold = false;


                        lineIndex += 2;
                        var sllCustomerCorporateGroupNo = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                           ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["SLL_CUSTOMER_CORPORATE_GROUP_NO"].ToString()
                                                           : string.Empty;
                        var sllCustomerCorporateGroupYes = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                           ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["SLL_CUSTOMER_CORPORATE_GROUP_YES"].ToString()
                                                           : string.Empty;
                        ws.Cells[$"F{lineIndex}"].Value = sllCustomerCorporateGroupNo;
                        ws.Cells[$"H{lineIndex}"].Value = sllCustomerCorporateGroupYes;

                        lineIndex += 2;

                        var commentGroupManager = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                           ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["COMMENT_GROUP_MANAGER"].ToString()
                                                           : string.Empty;
                        ws.Cells[$"B{lineIndex}"].Value = commentGroupManager;

                        lineIndex += 3;
                        //Debug.WriteLine("Currently lineIndex =>" + lineIndex); //35
                        var sllCustomerProjectNo = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["SLL_CUSTOMER_PROJECT_NO"].ToString()
                                                   : string.Empty;
                        var sllCustomerProjectYes = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["SLL_CUSTOMER_PROJECT_YES"].ToString()
                                                   : string.Empty;
                        var projectValue = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["PROJECT"].ToString()
                                                   : string.Empty;
                        ws.Cells[$"F{lineIndex}"].Value = sllCustomerProjectNo;
                        ws.Cells[$"H{lineIndex}"].Value = sllCustomerProjectYes;
                        ws.Cells[$"K{lineIndex}"].Value = projectValue;

                        lineIndex += 2;
                        var commentProjectManager = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["COMMENT_PROJECT_MANAGER"].ToString()
                                                   : string.Empty;
                        ws.Cells[$"B{lineIndex}"].Value = commentProjectManager;

                        lineIndex += 3;
                        // Debug.WriteLine("Currently lineIndex =>" + lineIndex); //
                        var anyViolationOnGusNo = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["ANY_VIOLATION_ON_GUS_NO"].ToString()
                                                   : string.Empty;

                        var anyViolationOnGusYes = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["ANY_VIOLATION_ON_GUS_YES"].ToString()
                                                   : string.Empty;
                        var approvalLevelGus = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["APPROVAL_LEVEL_GUS"].ToString()
                                                   : string.Empty;

                        ws.Cells[$"F{lineIndex}"].Value = anyViolationOnGusNo;
                        ws.Cells[$"H{lineIndex}"].Value = anyViolationOnGusYes;
                        ws.Cells[$"K{lineIndex}"].Value = approvalLevelGus;

                        lineIndex += 2;
                        var anyViolationOnIusNo = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["ANY_VIOLATION_ON_GUS_NO"].ToString()
                                                   : string.Empty;

                        var anyViolationOnIusYes = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["ANY_VIOLATION_ON_IUS_YES"].ToString()
                                                   : string.Empty;
                        var approvalLevelIus = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["APPROVAL_LEVEL_IUS"].ToString()
                                                   : string.Empty;

                        ws.Cells[$"F{lineIndex}"].Value = anyViolationOnIusNo;
                        ws.Cells[$"H{lineIndex}"].Value = anyViolationOnIusYes;
                        ws.Cells[$"K{lineIndex}"].Value = approvalLevelIus;

                        lineIndex += 2;
                        var anyViolationOnOtherRegulationNo = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["ANY_VIOLATION_ON_OTHER_REGULATION_NO"].ToString()
                                                   : string.Empty;

                        var anyViolationOnOtherRegulationYes = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["ANY_VIOLATION_ON_OTHER_REGULATION_YES"].ToString()
                                                   : string.Empty;

                        ws.Cells[$"F{lineIndex}"].Value = anyViolationOnOtherRegulationNo;
                        ws.Cells[$"H{lineIndex}"].Value = anyViolationOnOtherRegulationYes;

                        lineIndex += 2;
                        var anySignificantAccountAbnormalitiesNo = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["ANY_SIGNIFICANT_ACCOUNT_ABNORMALITIES_NO"].ToString()
                                                   : string.Empty;

                        var anySignificantAccountAbnormalitiesYes = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["ANY_SIGNIFICANT_ACCOUNT_ABNORMALITIES_YES"].ToString()
                                                   : string.Empty;

                        ws.Cells[$"F{lineIndex}"].Value = anySignificantAccountAbnormalitiesNo;
                        ws.Cells[$"H{lineIndex}"].Value = anySignificantAccountAbnormalitiesYes;

                        lineIndex += 2;
                        var customerAgreedToSignConsentOfCreditBureauYes = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["CUSTOMER_AGREED_TO_SIGN_CONSENT_OF_CREDIT_BUREAU_YES"].ToString()
                                                   : string.Empty;
                        var cbConsentNo = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                            ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["CB_CONSENT_NO"].ToString()
                                            : string.Empty;

                        var dataSearch = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                         ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["DATE_SEARCH"].ToString()
                                         : string.Empty;

                        var searchPrefix = ws.Cells[$"K{lineIndex}"].Value?.ToString() ?? string.Empty;
                        var dataSearchValue = ws.Cells[$"K{lineIndex}"].RichText;

                        ws.Cells[$"F{lineIndex}"].Value = customerAgreedToSignConsentOfCreditBureauYes;
                        ws.Cells[$"J{lineIndex}"].Value = cbConsentNo;

                        dataSearchValue.Clear();
                        dataSearchValue.Add(searchPrefix);
                        dataSearchValue.Add(" " + dataSearch);

                        lineIndex += 2;
                        var purposeRequestNew = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                 ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["PURPOSE_REQUEST_NEW"].ToString()
                                                 : string.Empty;

                        var purposeRequestCreditReview = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["PURPOSE_REQUEST_CREDIT_REVIEW"].ToString()
                                                   : string.Empty;

                        ws.Cells[$"F{lineIndex}"].Value = purposeRequestNew;
                        ws.Cells[$"H{lineIndex}"].Value = purposeRequestCreditReview;

                        lineIndex += 2;
                        var customerAgreedToSignConsentOfCreditBureauNo = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                 ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["CUSTOMER_AGREED_TO_SIGN_CONSENT_OF_CREDIT_BUREAU_NO"].ToString()
                                                 : string.Empty;

                        var reqNoOfExemptRequest = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                 ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["REQ_NO_OF_EXCEMPT_REQUEST"].ToString()
                                                 : string.Empty;
                        ws.Cells[$"F{lineIndex}"].Value = customerAgreedToSignConsentOfCreditBureauNo;
                        ws.Cells[$"L{lineIndex}"].Value = reqNoOfExemptRequest;

                        lineIndex += 2;
                        var kycChecklistNo = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                 ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["KYC_CHECKLIST_NO"].ToString()
                                                 : string.Empty;

                        var kycChecklistYes = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["KYC_CHECKLIST_YES"].ToString()
                                                   : string.Empty;
                        var ratingCode = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["RATING_CODE"].ToString()
                                                   : string.Empty;

                        ws.Cells[$"F{lineIndex}"].Value = kycChecklistNo;
                        ws.Cells[$"H{lineIndex}"].Value = kycChecklistYes;
                        ws.Cells[$"K{lineIndex}"].Value = ratingCode;

                        //try todo table การติดตามดูแล
                        currentIndex = lineIndex + 3;
                        var sourceRange6 = wsTemplate.Cells["A67:M68"];
                        sourceRange6.Copy(ws.Cells[$"A{currentIndex}:M{currentIndex}"]);

                        lineIndex += 5;
                        DataTable dtInformationSheetFollower = GetCafInformationFollowUp(GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["CASE_CUSTOMER"].ToString());
                        if (dtInformationSheetFollower != null && dtInformationSheetFollower.Rows.Count > 0)
                        {
                            newRoleIndex = lineIndex;

                            foreach (DataRow row2 in dtInformationSheetFollower.Rows)
                            {
                                IEnumerable<string> issuesToFollowStrings = Utility.NewLineSplitText(row2["ISSUES_TO_FOLLOW"].ToString(), 20);
                                foreach (string str in issuesToFollowStrings)
                                {
                                    if (!string.IsNullOrEmpty(str))
                                    {
                                        ws.Cells[$"B{newRoleIndex}"].Value = str;
                                        ws.Cells[$"B{newRoleIndex}:E{newRoleIndex}"].Merge = true;
                                        ws.Cells[$"B{newRoleIndex}"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        ws.Cells[$"E{newRoleIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                    }
                                }
                                minFollowIndex = newRoleIndex;
                                IEnumerable<string> followUpMethodStrings = Utility.NewLineSplitText(row2["FOLLOW_UP_METHOD"].ToString(), 30);
                                foreach (string str in followUpMethodStrings)
                                {
                                    if (!string.IsNullOrEmpty(str))
                                    {
                                        ws.Cells[$"F{newRoleIndex}"].Value = str;
                                        ws.Cells[$"F{newRoleIndex}:H{newRoleIndex}"].Merge = true;
                                        ws.Cells[$"F{newRoleIndex}"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        ws.Cells[$"H{newRoleIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        ws.Cells[$"J{newRoleIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        newRoleIndex++;
                                    }
                                }
                                int followedUpFrequencyIndex = minFollowIndex;
                                IEnumerable<string> followedUpFrequencyStrings = Utility.NewLineSplitText(row2["FOLLOW_UP_FREQUENCY"].ToString(), 30);
                                foreach (string str in followedUpFrequencyStrings)
                                {
                                    if (!string.IsNullOrEmpty(str))
                                    {
                                        ws.Cells[$"I{followedUpFrequencyIndex}"].Value = str;
                                        ws.Cells[$"I{followedUpFrequencyIndex}:J{followedUpFrequencyIndex}"].Merge = true;
                                        //ws.Cells[$"I{followedUpFrequencyIndex}"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        // ws.Cells[$"J{followedUpFrequencyIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        followedUpFrequencyIndex++;
                                    }
                                }
                                int followedByIndex = minFollowIndex;
                                IEnumerable<string> followedUpByStrings = Utility.NewLineSplitText(row2["FOLLOWED_UP_BY"].ToString(), 30);
                                foreach (string str in followedUpByStrings)
                                {
                                    if (!string.IsNullOrEmpty(str))
                                    {
                                        ws.Cells[$"K{followedByIndex}"].Value = str;
                                        ws.Cells[$"K{followedByIndex}:L{followedByIndex}"].Merge = true;
                                        // ws.Cells[$"K{followedByIndex}"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        followedByIndex++;
                                    }
                                }
                                ws.Cells[$"B{newRoleIndex - 1}:L{newRoleIndex - 1}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            }
                            ws.Cells[$"B{lineIndex}:L{newRoleIndex - 1}"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            ws.Cells[$"M{lineIndex}:M{newRoleIndex - 1}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            lineIndex = Math.Max(lineIndex, newRoleIndex);

                            //lineIndex++;
                        }
                        else
                        {
                            ws.Cells[$"A{lineIndex}"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            ws.Cells[$"E{lineIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            ws.Cells[$"H{lineIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            ws.Cells[$"J{lineIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            ws.Cells[$"L{lineIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            ws.Cells[$"B{lineIndex}:L{lineIndex}"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            ws.Cells[$"M{lineIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            lineIndex++;
                        }
                        //  lineIndex++;
                        Debug.WriteLine("Currently lineIndex After Table =>" + lineIndex);
                        ws.Cells[$"A{lineIndex}:M{lineIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        ws.Cells[$"A{lineIndex}"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        ws.Cells[$"M{lineIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        lineIndex += 2;
                        //ws.Cells[$"A{lineIndex}:M{lineIndex}"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    }
                    else
                    {
                        Debug.WriteLine("CountNumber 2 sc Time <==>");
                        var sourceRange = wsTemplate.Cells["A5:M21"];
                        sourceRange.Copy(ws.Cells[$"A{lineIndex}:M{lineIndex}"]);

                        //Debug.WriteLine($"Copying from {sourceRange.Address} to A{lineIndex}:M{lineIndex}");

                        // sourceRange.Copy(ws.Cells[$"A{lineIndex}:M{lineIndex}"]);//return 1 


                        var informationValueColD = ws.Cells[$"D{lineIndex}"].Value?.ToString() ?? string.Empty;

                        var customerValueColE = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                           ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["CUSTOMER_NAME"].ToString()
                            : string.Empty;

                        var customerValue = ws.Cells[$"D{lineIndex}"].RichText;
                        customerValue.Clear();
                        customerValue.Add(informationValueColD);
                        customerValue.Add(customerValueColE).Bold = false;

                        var rmPrefixValue = ws.Cells[$"H{lineIndex}"].Value?.ToString() ?? string.Empty;
                        var rmValue = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                            ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["RM_NO"].ToString()
                            : string.Empty;

                        var rmTotal = ws.Cells[$"H{lineIndex}"].RichText;
                        rmTotal.Clear();
                        rmTotal.Add(rmPrefixValue);
                        rmTotal.Add($" {rmValue}").Bold = false;

                        lineIndex += 2;
                        var extendCreditFacilityValue = GetCafInformationCreditFacilityType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                            ? GetCafInformationCreditFacilityType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["EXTEND_CREDIT_FACILITY"].ToString()
                            : string.Empty;

                        ws.Cells[$"D{lineIndex}"].Value = extendCreditFacilityValue;

                        var noExtendCreditFacilityValue = GetCafInformationCreditFacilityType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                            ? GetCafInformationCreditFacilityType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["NO_EXTEND_CREDIT_FACILITY"].ToString()
                            : string.Empty;

                        ws.Cells[$"F{lineIndex}"].Value = noExtendCreditFacilityValue;

                        var reduceCreditFacilityValue = GetCafInformationCreditFacilityType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                          ? GetCafInformationCreditFacilityType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["REDUCE_CREDIT_FACILITY"].ToString()
                          : string.Empty;

                        ws.Cells[$"H{lineIndex}"].Value = reduceCreditFacilityValue;

                        var cancleValue = GetCafInformationCreditFacilityType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                         ? GetCafInformationCreditFacilityType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["CANCEL"].ToString()
                         : string.Empty;

                        ws.Cells[$"K{lineIndex}"].Value = cancleValue;

                        lineIndex += 3; // (อย่าลืมเรื่มรวมเอาไปลบในตอนสุดท้ายจาก Index ปัจจุบัน ให้เหลือเว้นได้ 1 ช่องระหว่างลูกค้า )
                        var tsr001 = GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                            ? GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TSR_001"].ToString()
                            : string.Empty;
                        ws.Cells[$"D{lineIndex}"].Value = tsr001;

                        var tsr002 = GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                            ? GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TSR_002"].ToString()
                            : string.Empty;
                        ws.Cells[$"F{lineIndex}"].Value = tsr002;

                        var tsr003 = GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                           ? GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TSR_003"].ToString()
                           : string.Empty;
                        ws.Cells[$"H{lineIndex}"].Value = tsr003;

                        var tsr004 = GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                           ? GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TSR_004"].ToString()
                           : string.Empty;
                        ws.Cells[$"K{lineIndex}"].Value = tsr004;

                        lineIndex += 2;
                        var tsr007 = GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                          ? GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TSR_007"].ToString()
                          : string.Empty;
                        ws.Cells[$"D{lineIndex}"].Value = tsr007;

                        var tsr008 = GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                            ? GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TSR_008"].ToString()
                            : string.Empty;
                        ws.Cells[$"F{lineIndex}"].Value = tsr008;

                        var tsr009 = GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                           ? GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TSR_009"].ToString()
                           : string.Empty;
                        ws.Cells[$"H{lineIndex}"].Value = tsr009;

                        lineIndex += 2;
                        var tsr005 = GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                        ? GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TSR_005"].ToString()
                        : string.Empty;
                        ws.Cells[$"D{lineIndex}"].Value = tsr005;

                        lineIndex += 2;
                        var tsr006 = GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                        ? GetCafInformationCreditRequestType(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TSR_006"].ToString()
                        : string.Empty;
                        ws.Cells[$"D{lineIndex}"].Value = tsr006;
                        lineIndex += 2;
                        ws.Cells[$"D{lineIndex}"].Value = tsr006;

                        lineIndex += 3;
                        var typeOfCreditReview = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                            ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TYPE_OF_CREDIT_REVIEW_PLAN"].ToString()
                            : string.Empty;
                        ws.Cells[$"D{lineIndex}"].Value = typeOfCreditReview;

                        var typeOfCreditReviewAdHoc = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                            ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TYPE_OF_CREDIT_REVIEW_AD_HOC"].ToString()
                            : string.Empty;
                        ws.Cells[$"F{lineIndex}"].Value = typeOfCreditReviewAdHoc;
                        currentIndex = lineIndex + 1; //currentIndex = 22
                        if (dtCafHeader != null && dtCafHeader.Rows.Count > 0 && dtCafHeader.Rows[0]["BUSINESS_UNIT_SHORTNAME"].ToString() == "SAM")
                        {
                            var sourceRange2 = wsTemplate.Cells["A22:M25"];
                            sourceRange2.Copy(ws.Cells[$"A{currentIndex}:M{currentIndex}"]);

                            lineIndex += 4; // lineIndex ปัจจุบัน = 21
                            var phaseStatus = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                    ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["PHASE_STATUS_SAM"].ToString()
                                                    : string.Empty;

                            ws.Cells[$"D{lineIndex}"].Value = phaseStatus;

                            var phaseStatusLegal = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["PHASE_STATUS_LEGAL"].ToString()
                                                   : string.Empty;
                            ws.Cells[$"F{lineIndex}"].Value = phaseStatusLegal;
                            currentIndex = lineIndex + 1; // เขียนแบบนี้ไม่ได้ทำให้ lineIndex เปลี่ยนแปลงเพราะต้องตั้งค่าใหม่อยู่ดี
                            var sourceRange3 = wsTemplate.Cells["A26:M30"];
                            sourceRange3.Copy(ws.Cells[$"A{currentIndex}:M{currentIndex}"]);
                            //Debug.WriteLine("Currently lineIndex =>" + lineIndex);//lineIndex => 25

                            lineIndex += 4;
                            var bblConnectedParty = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                           ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["BBL_CONNECTED_PARTY"].ToString()
                           : string.Empty;
                            ws.Cells[$"B{lineIndex}"].Value = bblConnectedParty;
                            //อีกเงื่อนไขหนึ่ง (เงื่อนไขที่ 2 )
                            currentIndex = lineIndex + 2;
                            var sourceRange4 = wsTemplate.Cells["A31:M37"];
                            sourceRange4.Copy(ws.Cells[$"A{currentIndex}:M{currentIndex}"]);

                            lineIndex += 3;
                            var transactionApprovedNo = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                          ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TRANSACTION_APPROVED_NO"].ToString()
                          : string.Empty;
                            ws.Cells[$"H{lineIndex}"].Value = transactionApprovedNo;

                            var transactionApprovedYes = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                          ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["TRANSACTION_APPROVED_YES"].ToString()
                          : string.Empty;
                            ws.Cells[$"K{lineIndex}"].Value = transactionApprovedYes;

                            lineIndex += 3;
                            var specialConditionAnsNo = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                          ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["NO_SPECIAL_CONDITION_NO"].ToString()
                          : string.Empty;
                            ws.Cells[$"H{lineIndex}"].Value = specialConditionAnsNo;

                            var specialConditionAnsYes = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                          ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["NO_SPECIAL_CONDITION_YES"].ToString()
                          : string.Empty;
                            ws.Cells[$"K{lineIndex}"].Value = specialConditionAnsYes;
                            lineIndex += 3;
                        }
                        else
                        {
                            var sourceRange3 = wsTemplate.Cells["A26:M31"];
                            sourceRange3.Copy(ws.Cells[$"A{currentIndex}:M{currentIndex}"]);
                            lineIndex += 4;

                            var bblConnectedParty = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                    ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["BBL_CONNECTED_PARTY"].ToString()
                                                    : string.Empty;
                            ws.Cells[$"B{lineIndex}"].Value = bblConnectedParty;
                            lineIndex += 3;
                        }

                        outsideCondIndex = lineIndex;
                        var sourceRange5 = wsTemplate.Cells["A38:M66"];
                        sourceRange5.Copy(ws.Cells[$"A{outsideCondIndex}:M{outsideCondIndex}"]);


                        lineIndex += 2;
                        var sllCustomerCorporateGroupNo = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                           ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["SLL_CUSTOMER_CORPORATE_GROUP_NO"].ToString()
                                                           : string.Empty;
                        var sllCustomerCorporateGroupYes = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                           ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["SLL_CUSTOMER_CORPORATE_GROUP_YES"].ToString()
                                                           : string.Empty;
                        ws.Cells[$"F{lineIndex}"].Value = sllCustomerCorporateGroupNo;
                        ws.Cells[$"H{lineIndex}"].Value = sllCustomerCorporateGroupYes;

                        lineIndex += 2;

                        var commentGroupManager = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                           ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["COMMENT_GROUP_MANAGER"].ToString()
                                                           : string.Empty;
                        ws.Cells[$"B{lineIndex}"].Value = commentGroupManager;

                        lineIndex += 3;
                        //Debug.WriteLine("Currently lineIndex =>" + lineIndex); //35
                        var sllCustomerProjectNo = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["SLL_CUSTOMER_PROJECT_NO"].ToString()
                                                   : string.Empty;
                        var sllCustomerProjectYes = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["SLL_CUSTOMER_PROJECT_YES"].ToString()
                                                   : string.Empty;
                        var projectValue = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["PROJECT"].ToString()
                                                   : string.Empty;
                        ws.Cells[$"F{lineIndex}"].Value = sllCustomerProjectNo;
                        ws.Cells[$"H{lineIndex}"].Value = sllCustomerProjectYes;
                        ws.Cells[$"K{lineIndex}"].Value = projectValue;

                        lineIndex += 2;
                        var commentProjectManager = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["COMMENT_PROJECT_MANAGER"].ToString()
                                                   : string.Empty;
                        ws.Cells[$"B{lineIndex}"].Value = commentProjectManager;

                        lineIndex += 3;
                        // Debug.WriteLine("Currently lineIndex =>" + lineIndex); //
                        var anyViolationOnGusNo = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["ANY_VIOLATION_ON_GUS_NO"].ToString()
                                                   : string.Empty;

                        var anyViolationOnGusYes = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["ANY_VIOLATION_ON_GUS_YES"].ToString()
                                                   : string.Empty;
                        var approvalLevelGus = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["APPROVAL_LEVEL_GUS"].ToString()
                                                   : string.Empty;

                        ws.Cells[$"F{lineIndex}"].Value = anyViolationOnGusNo;
                        ws.Cells[$"H{lineIndex}"].Value = anyViolationOnGusYes;
                        ws.Cells[$"K{lineIndex}"].Value = approvalLevelGus;

                        lineIndex += 2;
                        var anyViolationOnIusNo = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["ANY_VIOLATION_ON_GUS_NO"].ToString()
                                                   : string.Empty;

                        var anyViolationOnIusYes = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["ANY_VIOLATION_ON_IUS_YES"].ToString()
                                                   : string.Empty;
                        var approvalLevelIus = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["APPROVAL_LEVEL_IUS"].ToString()
                                                   : string.Empty;

                        ws.Cells[$"F{lineIndex}"].Value = anyViolationOnIusNo;
                        ws.Cells[$"H{lineIndex}"].Value = anyViolationOnIusYes;
                        ws.Cells[$"K{lineIndex}"].Value = approvalLevelIus;

                        lineIndex += 2;
                        var anyViolationOnOtherRegulationNo = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["ANY_VIOLATION_ON_OTHER_REGULATION_NO"].ToString()
                                                   : string.Empty;

                        var anyViolationOnOtherRegulationYes = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["ANY_VIOLATION_ON_OTHER_REGULATION_YES"].ToString()
                                                   : string.Empty;

                        ws.Cells[$"F{lineIndex}"].Value = anyViolationOnOtherRegulationNo;
                        ws.Cells[$"H{lineIndex}"].Value = anyViolationOnOtherRegulationYes;

                        lineIndex += 2;
                        var anySignificantAccountAbnormalitiesNo = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["ANY_SIGNIFICANT_ACCOUNT_ABNORMALITIES_NO"].ToString()
                                                   : string.Empty;

                        var anySignificantAccountAbnormalitiesYes = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["ANY_SIGNIFICANT_ACCOUNT_ABNORMALITIES_YES"].ToString()
                                                   : string.Empty;

                        ws.Cells[$"F{lineIndex}"].Value = anySignificantAccountAbnormalitiesNo;
                        ws.Cells[$"H{lineIndex}"].Value = anySignificantAccountAbnormalitiesYes;

                        lineIndex += 2;
                        var customerAgreedToSignConsentOfCreditBureauYes = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["CUSTOMER_AGREED_TO_SIGN_CONSENT_OF_CREDIT_BUREAU_YES"].ToString()
                                                   : string.Empty;
                        var cbConsentNo = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                            ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["CB_CONSENT_NO"].ToString()
                                            : string.Empty;

                        var dataSearch = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                         ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["DATE_SEARCH"].ToString()
                                         : string.Empty;

                        var searchPrefix = ws.Cells[$"K{lineIndex}"].Value?.ToString() ?? string.Empty;
                        var dataSearchValue = ws.Cells[$"K{lineIndex}"].RichText;

                        ws.Cells[$"F{lineIndex}"].Value = customerAgreedToSignConsentOfCreditBureauYes;
                        ws.Cells[$"J{lineIndex}"].Value = cbConsentNo;

                        dataSearchValue.Clear();
                        dataSearchValue.Add(searchPrefix);
                        dataSearchValue.Add(" " + dataSearch);

                        lineIndex += 2;
                        var purposeRequestNew = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                 ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["PURPOSE_REQUEST_NEW"].ToString()
                                                 : string.Empty;

                        var purposeRequestCreditReview = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["PURPOSE_REQUEST_CREDIT_REVIEW"].ToString()
                                                   : string.Empty;

                        ws.Cells[$"F{lineIndex}"].Value = purposeRequestNew;
                        ws.Cells[$"H{lineIndex}"].Value = purposeRequestCreditReview;

                        lineIndex += 2;
                        var customerAgreedToSignConsentOfCreditBureauNo = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                 ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["CUSTOMER_AGREED_TO_SIGN_CONSENT_OF_CREDIT_BUREAU_NO"].ToString()
                                                 : string.Empty;

                        var reqNoOfExemptRequest = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                 ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["REQ_NO_OF_EXCEMPT_REQUEST"].ToString()
                                                 : string.Empty;
                        ws.Cells[$"F{lineIndex}"].Value = customerAgreedToSignConsentOfCreditBureauNo;
                        ws.Cells[$"L{lineIndex}"].Value = reqNoOfExemptRequest;

                        lineIndex += 2;

                        var kycChecklistNo = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                 ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["KYC_CHECKLIST_NO"].ToString()
                                                 : string.Empty;

                        var kycChecklistYes = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["KYC_CHECKLIST_YES"].ToString()
                                                   : string.Empty;
                        var ratingCode = GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows.Count > 0
                                                   ? GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["RATING_CODE"].ToString()
                                                   : string.Empty;

                        ws.Cells[$"F{lineIndex}"].Value = kycChecklistNo;
                        ws.Cells[$"H{lineIndex}"].Value = kycChecklistYes;
                        ws.Cells[$"K{lineIndex}"].Value = ratingCode;
                        //try todo table การติดตามดูแล
                        currentIndex = lineIndex + 3;
                        var sourceRange6 = wsTemplate.Cells["A67:M68"];
                        sourceRange6.Copy(ws.Cells[$"A{currentIndex}:M{currentIndex}"]);

                        lineIndex += 5;

                        DataTable dtInformationSheetFollower = GetCafInformationFollowUp(GetCafInformationData(row["CHILD_TRANSACTION_ID"].ToString()).Rows[0]["CASE_CUSTOMER"].ToString());
                        if (dtInformationSheetFollower != null && dtInformationSheetFollower.Rows.Count > 0)
                        {

                            newRoleIndex = lineIndex;
                            foreach (DataRow row2 in dtInformationSheetFollower.Rows)
                            {
                                IEnumerable<string> issuesToFollowStrings = Utility.NewLineSplitText(row2["ISSUES_TO_FOLLOW"].ToString(), 20);
                                foreach (string str in issuesToFollowStrings)
                                {
                                    if (!string.IsNullOrEmpty(str))
                                    {
                                        ws.Cells[$"B{newRoleIndex}"].Value = str;
                                        ws.Cells[$"B{newRoleIndex}:E{newRoleIndex}"].Merge = true;
                                        ws.Cells[$"B{newRoleIndex}"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        ws.Cells[$"E{newRoleIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                    }
                                }
                                minFollowIndex = newRoleIndex;
                                IEnumerable<string> followUpMethodStrings = Utility.NewLineSplitText(row2["FOLLOW_UP_METHOD"].ToString(), 30);
                                foreach (string str in followUpMethodStrings)
                                {
                                    if (!string.IsNullOrEmpty(str))
                                    {
                                        ws.Cells[$"F{newRoleIndex}"].Value = str;
                                        ws.Cells[$"F{newRoleIndex}:H{newRoleIndex}"].Merge = true;
                                        ws.Cells[$"F{newRoleIndex}"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        ws.Cells[$"H{newRoleIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        ws.Cells[$"J{newRoleIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        newRoleIndex++;
                                    }
                                }
                                int followedUpFrequencyIndex = minFollowIndex;
                                IEnumerable<string> followedUpFrequencyStrings = Utility.NewLineSplitText(row2["FOLLOW_UP_FREQUENCY"].ToString(), 30);
                                foreach (string str in followedUpFrequencyStrings)
                                {
                                    if (!string.IsNullOrEmpty(str))
                                    {
                                        ws.Cells[$"I{followedUpFrequencyIndex}"].Value = str;
                                        ws.Cells[$"I{followedUpFrequencyIndex}:J{followedUpFrequencyIndex}"].Merge = true;
                                        //ws.Cells[$"I{followedUpFrequencyIndex}"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        // ws.Cells[$"J{followedUpFrequencyIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        followedUpFrequencyIndex++;
                                    }
                                }
                                int followedByIndex = minFollowIndex;
                                IEnumerable<string> followedUpByStrings = Utility.NewLineSplitText(row2["FOLLOWED_UP_BY"].ToString(), 30);
                                foreach (string str in followedUpByStrings)
                                {
                                    if (!string.IsNullOrEmpty(str))
                                    {
                                        ws.Cells[$"K{followedByIndex}"].Value = str;
                                        ws.Cells[$"K{followedByIndex}:L{followedByIndex}"].Merge = true;
                                        // ws.Cells[$"K{followedByIndex}"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        followedByIndex++;
                                    }
                                }




                                ws.Cells[$"B{newRoleIndex - 1}:L{newRoleIndex - 1}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            }
                            ws.Cells[$"B{lineIndex}:L{newRoleIndex - 1}"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            ws.Cells[$"M{lineIndex}:M{newRoleIndex - 1}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            lineIndex = Math.Max(lineIndex, newRoleIndex);
                            // Debug.WriteLine("Currently lineIndex =>" + lineIndex); //
                            //lineIndex++;
                        }
                        else
                        {
                            ws.Cells[$"A{lineIndex}"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            ws.Cells[$"E{lineIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            ws.Cells[$"H{lineIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            ws.Cells[$"J{lineIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            ws.Cells[$"L{lineIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            ws.Cells[$"B{lineIndex}:L{lineIndex}"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            ws.Cells[$"M{lineIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            lineIndex++;
                        }
                        // lineIndex++;
                        Debug.WriteLine("Currently lineIndex After Table =>" + lineIndex);
                        ws.Cells[$"A{lineIndex}"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        ws.Cells[$"A{lineIndex}:M{lineIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        ws.Cells[$"M{lineIndex}"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        //ws.Cells[$"B{lineIndex}:L{Math.Max(newRoleIndex - 1, newRoleIndex - 1)}"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                        //Index แบ่งช่วงลูกค้า  ต้องไล่ Index ควรดูตอนสุดท้ายจาก lineIndex ตัวสุดท้ายแล้วบวกเข้าไป 1 เพื่อขึ้นในส่วนลูกค้าถัดไป
                        lineIndex += 2;
                    }
                for (int i = 5; i <= lineIndex; i++)
                {
                    ws.Row(i).Height = 22;
                }

            }
        }

        //Single & Group Start Function
        private DataTable GetCafHeader()
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@itemid", _caseId);
            parameters.Add("@userid", _userId);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[dbo].[SpBBL_rpt_CAF_Header]", parameters);
            return dt;
        }
        private DataTable GetCafCustListData()
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@itemid", _caseId);
            parameters.Add("@userid", _userId);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[param].[spBBL_rpt_GET_CAF_CUST_LIST_OTHER]", parameters);
            return dt;
        }

        private DataTable GetCafInformationData(string childTransactionId)
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@child_transaction_id", childTransactionId);
            parameters.Add("@itemid", _caseId);
            parameters.Add("@userid", _userId);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[param].[SpBBL_RPT_CAF_INFO_SHEET]", parameters);
            return dt;
        }

        private DataTable GetCafInformationCreditFacilityType(string childTransactionId)
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@itemid", _caseId);
            parameters.Add("@CHILD_TRANSACTION_ID", childTransactionId);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[param].[SpBBL_RPT_CAF_INFO_SHEET_CREDIT_FACILITY_TYPE]", parameters);
            return dt;
        }
        private DataTable GetCafInformationCreditRequestType(string childTransactionId)
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@itemid", _caseId);
            parameters.Add("@CHILD_TRANSACTION_ID", childTransactionId);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[param].[SpBBL_RPT_CAF_INFO_SHEET_REQUEST_TYPE]", parameters);
            return dt;
        }
        private DataTable GetCafInformationFollowUp(string caseCustomer)
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters.Add("@case_customer", caseCustomer);
            DataTable dt = SqlServerDBUtil.getDataTableFromStoredProcedure("[param].[spBBL_rpt_INFORMATION_SHEET_FOLLOW_UP]", parameters);
            return dt;
        }

        //Single & Group End Function

    }
}
