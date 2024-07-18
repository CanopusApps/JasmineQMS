using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Collections;
using TEPL.QMS.Common;
using TEPL.QMS.Common.Constants;
using System.Data;
using TEPL.QMS.Common.Models;
using TEPL.QMS.Workflow.Business;
using System.Runtime.InteropServices;
using System.Configuration;
using System.Text.RegularExpressions;
using TEPL.QMS.Workflow.Models;
using TEPL.QMS.DAL.Database.Component;
using System.Web;

namespace TEPL.QMS.BLL.Component
{
    public class DocumentUpload
    {
        DAL.Database.Component.DocumentOperations docOperObj = new DAL.Database.Component.DocumentOperations();
        QMSAdminDAL objAdminDAL = new QMSAdminDAL();

        public object TEPLQMS { get; private set; }
        public Object[] GenerateDocumentNo(DraftDocument objDoc)
        {
            Object[] ArrayOfObjects = new Object[3];
            try
            {
                string docNumberString = string.Empty;
                Boolean isMissingApprovals = false;
                objDoc.DocumentLevel = GetDocumentLevel(objDoc.DocumentCategoryCode);
                WorkflowActions objWFA = new WorkflowActions();
                List<WFApprovers> objApprovers = objWFA.GetWorkflowApprovers(QMSConstants.WorkflowID, objDoc.ProjectTypeID, objDoc.ProjectID, objDoc.DocumentLevel, objDoc.SectionID);
                //objWFA.GetWorkflowApprovers(new Guid("464dbb1a-0af6-4d32-bc59-9c7c43305432"), new Guid("ef875c62-76fa-490a-b33e-f15f045f0cca"), new Guid("55ca6d0c-265f-4f59-a883-99311d55b852"), "LEVEL 4", new Guid("60D0B83A-2593-478C-9F88-D8B47CF11360"));
                //objWFA.GetWorkflowApprovers(new Guid("464dbb1a-0af6-4d32-bc59-9c7c43305432"), new Guid("7101e9e2-66d7-4b61-af5a-888ac765deb3"), new Guid("d85997d2-686c-44cf-9395-2de3fb7eb42e"), "LEVEL 4", new Guid("60D0B83A-2593-478C-9F88-D8B47CF11360"));
                foreach (WFApprovers wfApp in objApprovers)
                {
                    if (string.IsNullOrEmpty(wfApp.ApprovalUser))
                    { isMissingApprovals = true; break; }
                }
                if (!isMissingApprovals)
                {
                    objDoc = docOperObj.GenerateDocumentNo(objDoc);
                    docNumberString = objDoc.DocumentNo + "#" + objDoc.DocumentID + "#" + objDoc.WFExecutionID;
                }
                else
                {
                    //Approvers not configured
                }
                ArrayOfObjects[0] = isMissingApprovals;
                ArrayOfObjects[1] = docNumberString;
                ArrayOfObjects[2] = objApprovers;
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return ArrayOfObjects;
        }
        public DraftDocument SubmitDocument(DraftDocument objDoc)
        {
            try
            {
                WorkflowActions objWF = new WorkflowActions();
                DocumentResponse objRes = new DocumentResponse();
                //docOperObj.UploadDocument(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.EditableFilePath, objDoc.EditableDocumentName, objDoc.DraftVersion, objDoc.EditableByteArray);
                //docOperObj.UploadDocument(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.ReadableFilePath, objDoc.ReadableDocumentName, objDoc.DraftVersion, objDoc.ReadableByteArray);
                docOperObj.UploadWithOutEncryptedDocument(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.EditableFilePath, objDoc.EditableDocumentName, objDoc.DraftVersion, objDoc.EditableByteArray);
                docOperObj.UploadWithOutEncryptedDocument(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.ReadableFilePath, objDoc.ReadableDocumentName, objDoc.DraftVersion, objDoc.ReadableByteArray);
                
                docOperObj.DocumentUpdate(objDoc);
                Stage objSt = objWF.GetWorkflowStage(QMSConstants.WorkflowID, objDoc.CurrentStageID);
                if (objSt.IsDocumentLevelRequired)
                    objDoc.DocumentLevel = GetDocumentLevel(objDoc.DocumentCategoryCode);
                else
                    objDoc.DocumentLevel = "";
                DocumentApprover objApprover = GetDocumentApprover(objDoc.ProjectTypeID, objDoc.ProjectID, objSt.NextStageID, objDoc.DocumentLevel, objDoc.SectionID);
                objWF.ExecuteAction(objDoc.WFExecutionID, objDoc.CurrentStageID, objDoc.UploadedUserID, objDoc.Action, objDoc.Comments, objDoc.UploadedUserID);
                objWF.CreateAction(objDoc.WFExecutionID, objSt.NextStageID, objApprover.ApprovalUser, "", objDoc.UploadedUserID);
                bool blAppLink = objDoc.ProjectTypeCode == "MP" ? true : false;
                //Send email - Doc Name, Doc Num, DOc ID, receipt email, stage, uploaded by ,
                string message = objDoc.UploadedUserName + " has uploaded a document having Document Number - <b>" + objDoc.DocumentNo + "</b>, and it is waiting for your review. Please check and take an appropriate action as applicable.";
                PrepareandSendMail(objApprover.ApprovalUserEmail, objDoc, objDoc.DocumentNo + " - QMS Reviewer - Document ready for Review", message, "ApproveRequest", blAppLink);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
            return objDoc;
        }
        public DraftDocument SubmitExistingUpload(DraftDocument objDoc)
        {
            try
            {
                WorkflowActions objWF = new WorkflowActions();
                DocumentResponse objRes = new DocumentResponse();
                docOperObj.UploadWithOutEncryptedDocument(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.EditableFilePath, objDoc.EditableDocumentName, objDoc.DraftVersion, objDoc.EditableByteArray);
                docOperObj.UploadWithOutEncryptedDocument(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.ReadableFilePath, objDoc.ReadableDocumentName, objDoc.DraftVersion, objDoc.ReadableByteArray);
                objDoc.DraftVersion = objDoc.DraftVersion + 0.001m;
                docOperObj.DocumentUpdate(objDoc);
                objWF.WorkflowInitate(QMSConstants.WorkflowID, objDoc.DocumentID, objDoc.CurrentStageID, objDoc.UploadedUserID, objDoc.UploadedUserID);
                objDoc = GetDocumentDetailsByID("User", objDoc.UploadedUserID, objDoc.DocumentID);
                //objDoc = GetPublishedDocumentDetailsByID(objDoc.UploadedUserID, objDoc.DocumentID, false);
                objDoc.Action = "Submitted";
                Stage objSt = objWF.GetWorkflowStage(QMSConstants.WorkflowID, objDoc.CurrentStageID);
                if (objSt.IsDocumentLevelRequired)
                    objDoc.DocumentLevel = GetDocumentLevel(objDoc.DocumentCategoryCode);
                DocumentApprover objApprover = GetDocumentApprover(objDoc.ProjectTypeID, objDoc.ProjectID, objSt.NextStageID, objDoc.DocumentLevel, objDoc.SectionID);
                objWF.ExecuteAction(objDoc.WFExecutionID, objDoc.CurrentStageID, objDoc.UploadedUserID, objDoc.Action, objDoc.Comments, objDoc.UploadedUserID);
                objWF.CreateAction(objDoc.WFExecutionID, objSt.NextStageID, objApprover.ApprovalUser, "", objDoc.UploadedUserID);

                bool blAppLink = objDoc.ProjectTypeCode == "MP" ? true : false;
                //Send email - Doc Name, Doc Num, DOc ID, receipt email, stage, uploaded by ,
                string message = objDoc.UploadedUserName + " has uploaded a document having Document Number - <b>" + objDoc.DocumentNo + "</b>, and it is waiting for your review. Please check and take an appropriate action as applicable.";
                PrepareandSendMail(objApprover.ApprovalUserEmail, objDoc, objDoc.DocumentNo + " - QMS Reviewer - Document ready for Review", message, "ApproveRequest", blAppLink);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
            return objDoc;
        }
        public void ReSubmitDocument(DraftDocument objDoc, bool isDocumentUploaded)
        {
            try
            {
                WorkflowActions objWF = new WorkflowActions();
                DocumentResponse objRes = new DocumentResponse();
                if (isDocumentUploaded)
                {
                    //docOperObj.UploadDocument(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.EditableFilePath, objDoc.EditableDocumentName, objDoc.DraftVersion, objDoc.EditableByteArray);
                    //docOperObj.UploadDocument(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.ReadableFilePath, objDoc.ReadableDocumentName, objDoc.DraftVersion, objDoc.ReadableByteArray);
                    docOperObj.UploadWithOutEncryptedDocument(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.EditableFilePath, objDoc.EditableDocumentName, objDoc.DraftVersion, objDoc.EditableByteArray);
                    docOperObj.UploadWithOutEncryptedDocument(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.ReadableFilePath, objDoc.ReadableDocumentName, objDoc.DraftVersion, objDoc.ReadableByteArray);
                    objDoc.DraftVersion = objDoc.DraftVersion + 0.001m;
                }
                docOperObj.DocumentDescriptionUpdate(objDoc);
                Stage objSt = objWF.GetWorkflowStage(QMSConstants.WorkflowID, objDoc.CurrentStageID);
                if (objSt.IsDocumentLevelRequired)
                    objDoc.DocumentLevel = GetDocumentLevel(objDoc.DocumentCategoryCode);
                DocumentApprover objApprover = GetDocumentApprover(objDoc.ProjectTypeID, objDoc.ProjectID, objSt.NextStageID, objDoc.DocumentLevel, objDoc.SectionID);
                objWF.ExecuteAction(objDoc.WFExecutionID, objDoc.CurrentStageID, objDoc.UploadedUserID, objDoc.Action, objDoc.Comments, objDoc.UploadedUserID);
                objWF.CreateAction(objDoc.WFExecutionID, objSt.NextStageID, objApprover.ApprovalUser, "", objDoc.UploadedUserID);

                bool blAppLink = objDoc.ProjectTypeCode == "MP" ? true : false;
                //Send email - Doc Name, Doc Num, DOc ID, receipt email, stage, uploaded by ,
                string message = objDoc.UploadedUserName + "  has re-uploaded the document having Document Number - <b>" + objDoc.DocumentNo.ToString() + "</b>, and waiting for your review. Please check and take an appropriate action as applicable.";
                PrepareandSendMail(objApprover.ApprovalUserEmail, objDoc, objDoc.DocumentNo + " - QMS Reviewer - Document ready for Re-review", message, "ApproveRequest", blAppLink);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
        }
        public string ApproveDocument(DraftDocument objDoc, bool isDocumentUploaded, bool IsMultiApproversChanged)
        {
            string strResult = string.Empty;
            try
            {
                WorkflowActions objWF = new WorkflowActions();
                if (isDocumentUploaded)
                {
                    docOperObj.UploadWithOutEncryptedDocument(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.EditableFilePath, objDoc.EditableDocumentName, objDoc.DraftVersion, objDoc.EditableByteArray);
                    docOperObj.UploadWithOutEncryptedDocument(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.ReadableFilePath, objDoc.ReadableDocumentName, objDoc.DraftVersion, objDoc.ReadableByteArray);
                    objDoc.DraftVersion = objDoc.DraftVersion + 0.001m;
                }
                docOperObj.DocumentDescriptionUpdate(objDoc);

                string strResponse = objWF.ExecuteAction(objDoc.WFExecutionID, objDoc.CurrentStageID, objDoc.ActionedID, objDoc.Action, objDoc.ActionComments, objDoc.ActionedID);
                List<Response> objRes = BindModels.ConvertJSON<Response>(strResponse);
                if (objRes[0].Status == "A")
                {
                    Stage objSt = objWF.GetWorkflowStage(QMSConstants.WorkflowID, objDoc.CurrentStageID);
                    //GetApproverDetails(objDoc);
                    if (objSt.NextStageID != Guid.Empty)
                    {
                        if (objSt.IsDocumentLevelRequired)
                            objDoc.DocumentLevel = GetDocumentLevel(objDoc.DocumentCategoryCode);
                        DocumentApprover objApprover = GetDocumentApprover(objDoc.ProjectTypeID, objDoc.ProjectID, objSt.NextStageID, objDoc.DocumentLevel, objDoc.SectionID);
                        string strMultiApprovers = string.Empty;

                        //if(objSt.CurrentStage == "QMS Reviewer")
                        //{
                            
                        //}
                        //else
                        //{

                        //}

                        if (objDoc.IsMultipleApprovers)
                        {
                            List<User> objUser = null;
                            if (objSt.NextStage == "Document Reviewer")
                            {
                                objUser = objDoc.MultipleApproversDisplay.DocumentReviewers;

                            }
                            else if (objSt.NextStage == "Document Approver")
                            {
                                objUser = objDoc.MultipleApproversDisplay.DocumentApprovers;

                            }
                            if (objUser != null)
                            {
                                foreach (User ua in objUser)
                                {
                                    strMultiApprovers += ua.ID + ",";
                                    if (!objApprover.ApprovalUserEmail.Contains(ua.EmailID))
                                        objApprover.ApprovalUserEmail = objApprover.ApprovalUserEmail.Trim(';') + ';' + ua.EmailID;
                                }
                            }
                            strMultiApprovers = strMultiApprovers.Trim(',');
                        }
                        if (IsMultiApproversChanged)
                            docOperObj.UpdateMultipleApprovers(objDoc);


                        objWF.CreateAction(objDoc.WFExecutionID, objSt.NextStageID, objApprover.ApprovalUser, strMultiApprovers, objDoc.ActionedID);
                        if (!string.IsNullOrEmpty(objApprover.ApprovalUser))
                        {
                            bool blAppLink = objDoc.ProjectTypeCode == "MP" ? true : false;
                            string message = objSt.CurrentStage + " has approved the document having the Document Number - <b>" + objDoc.DocumentNo.ToString() + "</b> in " + objDoc.CurrentStage + " Stage. You are identified as " + objSt.NextStage + ", hence please check and take an appropriate action.";
                            PrepareandSendMail(objApprover.ApprovalUserEmail, objDoc, objDoc.DocumentNo + " - " + objSt.NextStage + " - Document is ready for Approval", message, "ApproveRequest", blAppLink);
                        }
                        else
                        {
                            LoggerBlock.WriteLog(objSt.NextStageID + " is not configured for Document No:" + objDoc.DocumentNo);
                        }
                    }
                    strResult = "success";
                }
                else if (objRes[0].Status == "I")
                {
                    strResult = objRes[0].Message;
                }
                else
                {
                    strResult = "success";
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
            return strResult;
        }
        public void RejectDocument(DraftDocument objDoc, bool isDocumentUploaded)
        {
            try
            {
                WorkflowActions objWF = new WorkflowActions();
                DocumentResponse objRes = new DocumentResponse();
                if (isDocumentUploaded)
                {
                    docOperObj.UploadWithOutEncryptedDocument(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.EditableFilePath, objDoc.EditableDocumentName, objDoc.DraftVersion, objDoc.EditableByteArray);
                    docOperObj.UploadWithOutEncryptedDocument(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.ReadableFilePath, objDoc.ReadableDocumentName, objDoc.DraftVersion, objDoc.ReadableByteArray);
                    objDoc.DraftVersion = objDoc.DraftVersion + 0.001m;
                }
                docOperObj.DocumentDescriptionUpdate(objDoc);
                Stage objSt = objWF.GetWorkflowStage(QMSConstants.WorkflowID, objDoc.CurrentStageID);
                objWF.ExecuteAction(objDoc.WFExecutionID, objDoc.CurrentStageID, objDoc.ActionedID, objDoc.Action, objDoc.ActionComments, objDoc.ActionedID);
                objWF.CreateAction(objDoc.WFExecutionID, objSt.PreviousStageID, objDoc.UploadedUserID.ToString(), "", objDoc.ActionedID);
                QMSAdmin objAdmin = new QMSAdmin();
                LoginUser objUser = new LoginUser();
                objUser = objAdmin.GetUserDetails(objDoc.UploadedUserID);
                bool blAppLink = objDoc.ProjectTypeCode == "MP" ? true : false;
                string message = objDoc.ActionByName + " has rejected the document uploaded by you with Document Number - <b>" + objDoc.DocumentNo + "</b>. Please modify the document and re-submit again.";
                PrepareandSendMail(objUser.EmailID, objDoc, objDoc.DocumentNo + " - " + objSt.PreviousStage + " - Document is Rejected by " + objDoc.CurrentStage, message, "Request", blAppLink);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
        }
        public DraftDocument DocumentPublish(DraftDocument objDoc, bool isDocumentUploaded)
        {
            try
            {
                WorkflowActions objWF = new WorkflowActions();
                docOperObj.UploadWithOutEncryptedDocument(QMSConstants.StoragePath, QMSConstants.PublishedFolder, objDoc.EditableFilePath, objDoc.EditableDocumentName, objDoc.EditVersion, objDoc.EditableByteArray);
                docOperObj.UploadWithOutEncryptedDocument(QMSConstants.StoragePath, QMSConstants.PublishedFolder, objDoc.ReadableFilePath, objDoc.ReadableDocumentName, objDoc.EditVersion, objDoc.ReadableByteArray);
                if (isDocumentUploaded)
                {
                    docOperObj.UploadWithOutEncryptedDocument(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.EditableFilePath, objDoc.EditableDocumentName, objDoc.DraftVersion, objDoc.EditableByteArray);
                    docOperObj.UploadWithOutEncryptedDocument(QMSConstants.StoragePath, QMSConstants.DraftFolder, objDoc.ReadableFilePath, objDoc.ReadableDocumentName, objDoc.DraftVersion, objDoc.ReadableByteArray);
                    objDoc.DraftVersion = objDoc.DraftVersion + 0.001m;
                }
                docOperObj.DocumentDescriptionUpdate(objDoc);
                objWF.ExecuteAction(objDoc.WFExecutionID, objDoc.CurrentStageID, objDoc.ActionedID, objDoc.Action, objDoc.ActionComments, objDoc.ActionedID);
                objDoc = docOperObj.DocumentPublish(objDoc);

                QMSAdmin objAdmin = new QMSAdmin();
                LoginUser objUser = new LoginUser();
                objUser = objAdmin.GetUserDetails(objDoc.UploadedUserID);
                bool blAppLink = objDoc.ProjectTypeCode == "MP" ? true : false;
                string message = objDoc.ActionByName + " has published the document uploaded by you with Document Number - <b>" + objDoc.DocumentNo + "</b>. Please verify the document.";
                PrepareandSendMail(objUser.EmailID, objDoc, objDoc.DocumentNo + " - Document is Published by " + objDoc.CurrentStage, message, "Request", blAppLink);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
            return objDoc;
        }

        public string ReplaceDocument(DraftDocument objDoc, bool isDocumentUploaded)
        {
            string result = "";
            try
            {
                if (isDocumentUploaded)
                {
                    docOperObj.UploadWithOutEncryptedDocument(QMSConstants.StoragePath, QMSConstants.PublishedFolder, objDoc.EditableFilePath, objDoc.EditableDocumentName, objDoc.EditVersion, objDoc.EditableByteArray);
                    docOperObj.UploadWithOutEncryptedDocument(QMSConstants.StoragePath, QMSConstants.PublishedFolder, objDoc.ReadableFilePath, objDoc.ReadableDocumentName, objDoc.EditVersion, objDoc.ReadableByteArray);
                }
                objDoc.EditVersion = objDoc.EditVersion + 0.001m;
                docOperObj.DocumentUpdatePublised(objDoc);
                result = "success";
            }
            catch (Exception ex)
            {
                result = "fail";
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
            return result;
        }

        private string GetEmailIDs(string emails)
        {
            string emailIDs = string.Empty;
            /*
            string[] strEmails = emails.Split(';');
            foreach (string str in strEmails)
            {
                User obj = new User();
                obj = CommonMethods.GetUserADValues(@System.Web.HttpContext.Current.Request.LogonUserIdentity.Name);
                obj.CreatedBy = System.Web.HttpContext.Current.Session[QMSConstants.LoggedInUserID].ToString();
                QMSAdmin objAdmin = new QMSAdmin();
                objAdmin.GetUserDetails(obj);
                emailIDs += str + "@tataelectronics.co.in,";
            }

            return emailIDs.Trim(',');*/
            return emails;
        }
        private void PrepareandSendMail(string toemail, DraftDocument docObj, string subject, string mainMessage, string pageName, bool blAppLink)
        {
            try
            {
                string emailIDs = GetEmailIDs(toemail);

                StringBuilder body = new StringBuilder();
                body.Append("Dear User,");
                body.Append("<br/><br/>");
                body.Append(mainMessage);
                body.Append("<br/><br/>");
                if (blAppLink)
                    body.Append("Link: <a style='text-decoration:underline' target='_blank' href='" + QMSConstants.WebSiteURL + pageName + "?ID= " + docObj.DocumentID + "'>" + "Click here" + "</a>");

                string strMailTemplate = GetApprovalMailTempate();
                strMailTemplate = strMailTemplate.Replace("@@MailBody@@", body.ToString());
                CommonMethods.SendMail(emailIDs, subject, strMailTemplate);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
        }
        private string GetApprovalMailTempate()
        {
            string strMailtemplate = string.Empty;
            QMSAdminDAL objADM = new QMSAdminDAL();
            strMailtemplate = objADM.GetConfigValue("ApprovalMailTemplate");
            return strMailtemplate;
        }
        private string GetDocumentLevel(string CategoryCode)
        {
            string strDocLevel = string.Empty;
            strDocLevel = docOperObj.GetDocumentLevel(CategoryCode);
            return strDocLevel;
        }
        private DocumentApprover GetDocumentApprover(Guid ProjectTypeID, Guid ProjectID, Guid NextStageID, string strDocLevel, Guid SectionID)
        {
            WorkflowActions objWF = new WorkflowActions();

            List<DocumentApprover> objApprovers = objWF.GetWorkflowApprover(ProjectTypeID, ProjectID, NextStageID, strDocLevel, SectionID);

            return objApprovers[0];
        }
        public List<DraftDocument> GetRequestedDocuments(Guid CreatedID)
        {
            List<DraftDocument> objDocuments = null;
            try
            {
                objDocuments = new List<DraftDocument>();
                DataTable dt = docOperObj.GetRequestedDocuments(CreatedID);

                for (int z = 0; z < dt.Rows.Count; z++)
                {
                    DraftDocument obj = new DraftDocument();
                    obj.SeqNo = z;
                    obj.DocumentID = new Guid(dt.Rows[z]["DocumentID"].ToString());
                    obj.DocumentNo = dt.Rows[z]["DocumentNo"].ToString();
                    obj.DepartmentName = dt.Rows[z]["DepartmentName"].ToString();
                    obj.SectionName = dt.Rows[z]["SectionName"].ToString();
                    obj.DocumentCategoryName = dt.Rows[z]["DocumentCategoryName"].ToString();
                    obj.ProjectName = dt.Rows[z]["ProjectName"].ToString();
                    obj.CreatedDate = DateTime.Parse(dt.Rows[z]["CreatedDate"].ToString());
                    obj.CurrentStage = dt.Rows[z]["CurrentStage"].ToString();

                    objDocuments.Add(obj);
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return objDocuments;
        }
        public List<DraftDocument> GetApprovalPendingDocuments(Guid ActionedID)
        {
            List<DraftDocument> objDocuments = null;
            try
            {
                objDocuments = new List<DraftDocument>();
                DataTable dt = docOperObj.GetApprovalPendingDocuments(ActionedID);

                for (int z = 0; z < dt.Rows.Count; z++)
                {
                    DraftDocument obj = new DraftDocument();
                    obj.SeqNo = z;
                    obj.DocumentID = new Guid(dt.Rows[z]["DocumentID"].ToString());
                    obj.DocumentNo = dt.Rows[z]["DocumentNo"].ToString();
                    obj.DepartmentName = dt.Rows[z]["DepartmentName"].ToString();
                    obj.SectionName = dt.Rows[z]["SectionName"].ToString();
                    obj.DocumentCategoryName = dt.Rows[z]["DocumentCategoryName"].ToString();
                    obj.ProjectName = dt.Rows[z]["ProjectName"].ToString();
                    obj.UploadedUserName = dt.Rows[z]["UploadedUserName"].ToString();
                    obj.CreatedDate = DateTime.Parse(dt.Rows[z]["CreatedDate"].ToString());
                    obj.CurrentStage = dt.Rows[z]["CurrentStage"].ToString();

                    objDocuments.Add(obj);
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return objDocuments;
        }
        public Object[] GetDocumentDetailsByNo(string DocumentNo)
        {
            //DraftDocument objDocuments = null;
            //try
            //{
            //    string strReturn = docOperObj.GetDocumentDetailsByNo(DocumentNo);
            //    List<DraftDocument> objDraft = BindModels.ConvertJSON<DraftDocument>(strReturn);
            //    objDocuments = objDraft[0];
            //}
            //catch (Exception ex)
            //{
            //    LoggerBlock.WriteTraceLog(ex);
            //}
            //return objDocuments;

            Object[] ArrayOfObjects = new Object[3];
            try
            {
                string docNumberString = string.Empty;
                Boolean isMissingApprovals = false;
                DraftDocument objDocuments = null;
                string strReturn = docOperObj.GetDocumentDetailsByNo(DocumentNo);
                List<DraftDocument> objDraft = BindModels.ConvertJSON<DraftDocument>(strReturn);
                objDocuments = objDraft[0];
                string DocumentLevel = GetDocumentLevel(objDocuments.DocumentCategoryCode);
                WorkflowActions objWFA = new WorkflowActions();
                List<WFApprovers> objApprovers = objWFA.GetWorkflowApprovers(QMSConstants.WorkflowID, objDocuments.ProjectTypeID, objDocuments.ProjectID, DocumentLevel, objDocuments.SectionID);
                foreach (WFApprovers wfApp in objApprovers)
                {
                    if (string.IsNullOrEmpty(wfApp.ApprovalUser))
                    { isMissingApprovals = true; break; }
                }
                ArrayOfObjects[0] = isMissingApprovals;
                ArrayOfObjects[1] = objDocuments;
                ArrayOfObjects[2] = objApprovers;
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return ArrayOfObjects;
        }
        public DraftDocument GetDocumentDetailsByNoForRequest(string DocumentNo)
        {
            DraftDocument objDocuments = null;
            try
            {
                string strReturn = docOperObj.GetDocumentDetailsByNo(DocumentNo);
                List<DraftDocument> objDraft = BindModels.ConvertJSON<DraftDocument>(strReturn);
                objDocuments = objDraft[0];
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return objDocuments;

        }
        public DraftDocument GetDocumentDetailsByID(string role, Guid loggedInUserID, Guid DocumentID)
        {
            DraftDocument objDocuments = null;
            try
            {
                string strReturn = docOperObj.GetDocumentDetailsByID(role, loggedInUserID, DocumentID);
                List<DraftDocument> objDraft = BindModels.ConvertJSON<DraftDocument>(strReturn);
                objDocuments = objDraft[0];
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return objDocuments;
        }
        public DraftDocument GetPublishedDocumentDetailsByID(Guid loggedInUserID, Guid DocumentID, bool isHistory)
        {
            DraftDocument objDocuments = null;
            try
            {
                string strReturn = "";
                if (isHistory)
                    strReturn = docOperObj.GetPublishedDocumentHistoryDetailsByID(loggedInUserID, DocumentID);
                else
                    strReturn = docOperObj.GetPublishedDocumentDetailsByID(loggedInUserID, DocumentID);
                List<DraftDocument> objDraft = BindModels.ConvertJSON<DraftDocument>(strReturn);
                objDocuments = objDraft[0];
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return objDocuments;
        }
        public List<DraftDocument> GetPublishedDocumentHistoryByID(Guid loggedInUserID, Guid DocumentID)
        {
            List<DraftDocument> objDraft = null;
            try
            {
                string strReturn = docOperObj.GetPublishedDocumentHistoryByID(loggedInUserID, DocumentID);
                objDraft = BindModels.ConvertJSON<DraftDocument>(strReturn);
                //objDocuments = objDraft[1];
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return objDraft;
        }
        public DraftDocument GetPublishedDocumentHistoryDetailsByID(Guid loggedInUserID, Guid DocumentID)
        {
            DraftDocument objDocuments = null;
            try
            {
                string strReturn = docOperObj.GetPublishedDocumentHistoryDetailsByID(loggedInUserID, DocumentID);
                List<DraftDocument> objDraft = BindModels.ConvertJSON<DraftDocument>(strReturn);
                objDocuments = objDraft[1];
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return objDocuments;
        }
        public List<WorkflowStages> GetApproverDetails(DraftDocument objDoc)
        {
            WorkflowActions objWF = new WorkflowActions();
            DataTable dtStages = objWF.GetWorkflowStages(QMSConstants.WorkflowID);
            DataTable dtApproverMatrix = new DataTable();// objWF.GetWorkflowApprovers(new Guid(),objDoc.ProjectID, objDoc.CurrentStageID, "", objDoc.SectionID);
            //Document Category, Document Reviewer, Document Approver, QMS Approver, Document Controller.

            List<WorkflowStages> wfStages = new List<WorkflowStages>();

            foreach (DataRow dr in dtStages.Rows)
            {
                WorkflowStages obj = new WorkflowStages();
                //f
                string expression = "ApprovalStage='" + dr["NextStage"] + "' and Level='" + objDoc.DocumentLevel + "' and Department='" + objDoc.DepartmentName + "' and Section='" + objDoc.SectionName + "' and Active =1";
                DataRow[] drApproverMatrix = dtApproverMatrix.Select(expression);
                obj.WFActionedID = (Guid)drApproverMatrix[0]["ApprovalUser"];
                //get display names
            }
            return wfStages;
        }
        public byte[] DownloadDocument(string DocURL, [Optional] int DocVersion)
        {
            byte[] fileContent = null;
            try
            {
                fileContent = docOperObj.DownloadDocument(DocURL, DocVersion);
                //fileContent = libOp.DownloadDocument(DocURL, DocVersion);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return fileContent;
        }
        public List<DraftDocument> GetPublishedDocuments(string DepartmentCode, string SectionCode, string ProjectCode, string DocumentCategoryCode, string DocumentDescription, Guid UserID, bool IsProjectActive)
        {
            List<DraftDocument> objDocList = null;
            try
            {
                DataTable dt;
                dt = objAdminDAL.GetPublishedDocuments(DepartmentCode, SectionCode, ProjectCode, DocumentCategoryCode, DocumentDescription, UserID);
                objDocList = new List<DraftDocument>();
                objDocList = GetDocuments(dt, IsProjectActive);
                GetApprovalMailTempate();
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return objDocList;
        }
        public byte[] GetExportPublishedDocuments(string DepartmentCode, string SectionCode, string ProjectCode, string DocumentCategoryCode, string DocumentDescription, Guid UserID, bool IsProjectActive)
        {
            byte[] fileContent = null;
            try
            {
                string strExcelPath = QMSConstants.TempFolder + Guid.NewGuid() + ".xlsx";
                DataTable dt = objAdminDAL.GetPublishedDocuments(DepartmentCode, SectionCode, ProjectCode, DocumentCategoryCode, DocumentDescription, UserID);
                string[] selectedColumns = new[] { "DepartmentName", "SectionName", "ProjectName", "DocumentCategoryName", "DocumentNo", "DocumentDescription", "RevisionReason", "Version", "PublishedOn" };
                DataTable dt2 = dt.AsEnumerable().Where(row => row.Field<bool>("ProjectActive") == IsProjectActive).CopyToDataTable();
                DataTable dt1 = new DataView(dt2).ToTable(false, selectedColumns);
                dt1.TableName = "PublishedDocuments";
                ExcelOperations.ExportDataSet(dt1, strExcelPath);
                fileContent = File.ReadAllBytes(strExcelPath);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return fileContent;
        }

        public byte[] GetExportPendingDocuments(string DepartmentCode, string SectionCode, string ProjectCode, string DocumentCategoryCode, string DocumentDescription, Guid UserID, bool IsProjectActive)
        {
            byte[] fileContent = null;
            try
            {
                string strExcelPath = QMSConstants.TempFolder + Guid.NewGuid() + ".xlsx";
                DataTable dt = objAdminDAL.GetDraftDocuments(DepartmentCode, SectionCode, ProjectCode, DocumentCategoryCode, DocumentDescription, UserID);
                string[] selectedColumns = new[] { "DepartmentName", "SectionName", "ProjectName", "DocumentCategoryName", "DocumentNo", "DocumentDescription", "RevisionReason", "CurrentStage", "PendingDays", "ActionByName" };
                DataTable dt2 = dt.AsEnumerable().Where(row => row.Field<bool>("ProjectActive") == IsProjectActive).CopyToDataTable();
                DataTable dt1 = new DataView(dt2).ToTable(false, selectedColumns);
                dt1.TableName = "PendingDocuments";
                ExcelOperations.ExportDataSet(dt1, strExcelPath);
                fileContent = File.ReadAllBytes(strExcelPath);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return fileContent;
        }
        public List<DraftDocument> GetDraftDocuments(string DepartmentCode, string SectionCode, string ProjectCode, string DocumentCategoryCode, string DocumentDescription, Guid UserID, bool IsProjectActive)
        {
            List<DraftDocument> objDocList = null;
            try
            {
                DataTable dt;
                dt = objAdminDAL.GetDraftDocuments(DepartmentCode, SectionCode, ProjectCode, DocumentCategoryCode, DocumentDescription, UserID);
                objDocList = new List<DraftDocument>();
                objDocList = GetDocuments(dt, IsProjectActive);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return objDocList;
        }
        public List<DraftDocument> GetDocuments(DataTable dt, bool IsProjectActive)
        {
            List<DraftDocument> objDocList = new List<DraftDocument>();
            foreach (DataRow dr in dt.Rows)
            {
                if (Convert.ToBoolean(dr["ProjectActive"].ToString()) == IsProjectActive)
                {
                    DraftDocument objDocument = CommonMethods.CreateItemFromRow<DraftDocument>(dr); // CommonMethods.GetDocumentObject(dr);
                    objDocList.Add(objDocument);
                }
            }
            return objDocList;
        }

        public DraftDocument DocumentDirectUpload(DraftDocument objDoc)
        {
            try
            {
                WorkflowActions objWF = new WorkflowActions();
                DocumentResponse objRes = new DocumentResponse();

                docOperObj.UploadWithOutEncryptedDocument(QMSConstants.StoragePath, QMSConstants.PublishedFolder, objDoc.EditableFilePath, objDoc.EditableDocumentName, objDoc.EditVersion, objDoc.EditableByteArray);
                docOperObj.UploadWithOutEncryptedDocument(QMSConstants.StoragePath, QMSConstants.PublishedFolder, objDoc.ReadableFilePath, objDoc.ReadableDocumentName, objDoc.EditVersion, objDoc.ReadableByteArray);

                docOperObj.DocumentUpdate(objDoc);

                Stage objSt = objWF.GetWorkflowStage(QMSConstants.WorkflowID, objDoc.CurrentStageID);
                if (objSt.IsDocumentLevelRequired)
                    objDoc.DocumentLevel = GetDocumentLevel(objDoc.DocumentCategoryCode);
                else
                    objDoc.DocumentLevel = "";
                DocumentApprover objApprover = GetDocumentApprover(objDoc.ProjectTypeID, objDoc.ProjectID, objSt.NextStageID, objDoc.DocumentLevel, objDoc.SectionID);
                objWF.ExecuteAction(objDoc.WFExecutionID, objDoc.CurrentStageID, objDoc.UploadedUserID, objDoc.Action, objDoc.Comments, objDoc.UploadedUserID);
                objWF.CreateAction(objDoc.WFExecutionID, objSt.NextStageID, objApprover.ApprovalUser, "", objDoc.UploadedUserID);
                bool blAppLink = objDoc.ProjectTypeCode == "MP" ? true : false;
                //Send email - Doc Name, Doc Num, DOc ID, receipt email, stage, uploaded by ,
                string message = objDoc.UploadedUserName + " has uploaded a document having Document Number - <b>" + objDoc.DocumentNo + "</b>, and it is waiting for your review. Please check and take an appropriate action as applicable.";
                PrepareandSendMail(objApprover.ApprovalUserEmail, objDoc, objDoc.DocumentNo + " - QMS Reviewer - Document ready for Review", message, "ApproveRequest", blAppLink);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
            return objDoc;
        }
    }
}
