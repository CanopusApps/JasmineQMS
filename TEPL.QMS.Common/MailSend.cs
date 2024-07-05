using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using TEPL.QMS.Common.Constants;
using TEPL.QMS.Common.DAL;

namespace TEPL.QMS.Common
{
    public static class MailSend
    {
        public static void SendApprovalMail(string strTemplateName, string strDisplayName, string toemail, string ccEmail, string bccEmail, string subject, string fileName, string mainMessage)
        {
            StringBuilder strMailBody = new StringBuilder();
            strMailBody.Append("Dear " + strDisplayName + ",");
            strMailBody.Append("<br/><br/>");
            strMailBody.Append(mainMessage);
            strMailBody.Append("<br/><br/>");
            string strMailTemplate = ConfigDAL.GetConfigValue(strTemplateName);
            strMailTemplate = strMailTemplate.Replace("@@MailBody@@", strMailBody.ToString());
            SendMail(toemail, ccEmail, bccEmail, subject, fileName, strMailTemplate);
        }
        public static void PrepareandSendMail(string strTemplateName, string strDisplayName, string toemail, string ccEmail, string bccEmail, string subject, string fileName, string mainMessage)
        {
            try
            {
                StringBuilder body = new StringBuilder();
                body.Append("Dear " + strDisplayName + ",");
                body.Append("<br/><br/>");
                body.Append(mainMessage);
                body.Append("<br/><br/>");
                string strMailTemplate = ConfigDAL.GetConfigValue(strTemplateName);
                strMailTemplate = strMailTemplate.Replace("@@MailBody@@", body.ToString());
                SendMail(toemail, ccEmail, bccEmail, subject, fileName, strMailTemplate);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
        }
        public static void SendMail(string toEmail, string ccEmail, string bccEmail, string subject, string fileName, string messageBody)
        {
            try
            {
                int port = Convert.ToInt32(QMSConstants.SMTPPORT);
                string host = QMSConstants.SMTPHOST;// "smtp.office365.com";
                string username = QMSConstants.EmailUserName; //"rajesh.m@canopusgbs.com";
                string password = QMSConstants.EmailPassword; //"sfdsfs";
                string mailFrom = QMSConstants.EmailFrom; //"rajesh.m@canopusgbs.com";             

                using (SmtpClient client = new SmtpClient())
                {
                    MailAddress from = new MailAddress(mailFrom);
                    MailMessage message = new MailMessage
                    {
                        From = from
                    };
                    if (!string.IsNullOrEmpty(toEmail))
                    {
                        string[] Approvers = toEmail.Split(';');
                        {
                            foreach (string Approver in Approvers)
                            {
                                if (!string.IsNullOrEmpty(Approver))
                                    message.To.Add(new MailAddress(Approver,""));
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(bccEmail))
                    {
                        string[] Approvers = bccEmail.Split(';');
                        {
                            foreach (string Approver in Approvers)
                            {
                                 if (!string.IsNullOrEmpty(Approver))
                                    message.Bcc.Add(new MailAddress(Approver, ""));
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(ccEmail))
                    {
                        string[] Approvers = ccEmail.Split(';');
                        {
                            foreach (string Approver in Approvers)
                            {
                                if (!string.IsNullOrEmpty(Approver))
                                    message.CC.Add(new MailAddress(Approver, ""));
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(fileName))
                        message.Attachments.Add(new Attachment(fileName));
                    message.Subject = subject;
                    message.Body = messageBody;
                    message.IsBodyHtml = true;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    client.Host = host;
                    client.Port = port;
                    client.EnableSsl = false;
                    client.Credentials = new NetworkCredential
                    {
                        UserName = username,
                        Password = password
                    };
                    client.Send(message);
                }
                LoggerBlock.WriteLog("Mail sent sucessfully to :" + toEmail+",Cc:"+ccEmail+",Bcc"+bccEmail+",Subject:"+subject);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteLog("Error in mail sending to:" + toEmail + ",Cc:" + ccEmail + ",Bcc" + bccEmail + ",Subject:" + subject);
                LoggerBlock.WriteTraceLog(ex);
            }
        }
        public static void GetEmailAddress(string strEmails)
        {
            string[] Approvers = strEmails.Split(';');
            foreach (string Approver in Approvers)
            {
            }

        }
    }
}
