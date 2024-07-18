using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Net.Mail;
using System.Net;
using System.Configuration;
using System.Collections;
using TEPL.QMS.Common.Models;
using TEPL.QMS.Common.Constants;
using System.Collections.Specialized;
using System.Data;
using System.Reflection;
using Newtonsoft.Json;
using System.ComponentModel;
using System.DirectoryServices.AccountManagement;
using System.Security.Cryptography;
using System.IO;
using System.Data.SqlClient;
using System.DirectoryServices;

namespace TEPL.QMS.Common
{
    public static class CommonMethods
    {
        public static string CombineUrl(string baseUrl, params string[] pathSegments)
        {
            // Combine the base URL and path segments using StringBuilder
            StringBuilder urlBuilder = new StringBuilder(baseUrl.Trim('/') + "/");
            foreach (var segment in pathSegments)
            {
                if (!string.IsNullOrEmpty(segment))
                    urlBuilder.Append(segment).Append("/");
            }

            // The combined URL
            return urlBuilder.ToString().Trim('/');
        }
        public static string CombinePath(string baseUrl, params string[] pathSegments)
        {
            // Combine the base URL and path segments using StringBuilder
            StringBuilder urlBuilder = new StringBuilder(baseUrl.Trim('/') + "\\");
            foreach (var segment in pathSegments)
            {
                if (!string.IsNullOrEmpty(segment))
                    urlBuilder.Append(segment).Append("\\");
            }

            // The combined URL
            return urlBuilder.ToString().Trim('\\');
        }
        public static LoginUser GetUserADValues(string LoginID)
        {
            LoginUser obj = new LoginUser();
            try
            {
                string username = "dms.canopus";
                string propertyName = "displayName";

                string displayName = ActiveDirectoryHelper.GetUserProperty(username, propertyName);

                var context = new PrincipalContext(ContextType.Domain, "tataelectronics.co.in");
                //var principal = UserPrincipal.FindByIdentity(context, LoginID);
                UserPrincipal principal = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);
                obj.DisplayName = principal.DisplayName;
                obj.EmailID = principal.EmailAddress;
                obj.LoginID = principal.SamAccountName; //For Local
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return obj;
        }
        public static LoginUser GetUserADValuesDirectoryEntry(string LoginID)
        {
            LoginUser obj = new LoginUser();
            try
            {
                try
                {
                    DirectoryEntry userEntry = ActiveDirectoryHelper.GetUserDirectoryEntry(LoginID);
                    obj.EmailID = (string)userEntry.Properties["mail"][0];
                    obj.DisplayName = (string)userEntry.Properties["displayname"][0];
                    obj.LoginID = (string)userEntry.Properties["SAMAccountName"][0];
                    try
                    {
                        obj.DepartmentName = (string)userEntry.Properties["department"][0];
                    }
                    catch
                    {
                        obj.DepartmentName = "";
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Error authenticating user. " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return obj;
        }

        public static string ValidateUserUsingDirectoryEntry(string LoginID, string Password)
        {
            string loginID = "";
            try
            {
                try
                {
                    DirectoryEntry userEntry = ActiveDirectoryHelper.ValidateUserUsingDirectoryEntry(LoginID, Password);
                    loginID = (string)userEntry.Properties["SAMAccountName"][0];
                    
                }
                catch (Exception ex)
                {
                    throw new Exception("Error in ValidateUserUsingDirectoryEntry  " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return loginID;
        }

        public static LoginUser GetUserADValuesDirectoryEntryUsingPrinContext(string LoginID)
        {
            LoginUser obj = new LoginUser();
            try
            {
                try
                {
                    User userEntry = ActiveDirectoryHelper.GetUserDirectoryEntryUsingPrinContext(LoginID);
                    obj.EmailID = userEntry.EmailID;
                    obj.DisplayName = userEntry.DisplayName;
                    obj.LoginID = userEntry.LoginID;
                    obj.DepartmentName = userEntry.DepartmentName;
                }
                catch (Exception ex)
                {
                    throw new Exception("Error GetUserADValuesDirectoryEntryUsingPrinContext  " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return obj;
        }

        public static List<User> GetUsersForADUsingName(string Name)
        {
            List<User> listUsers = new List<User>();
            try
            {
                SearchResultCollection userFromAD = ActiveDirectoryHelper.GetUsersFromDirectory(Name);
                if (userFromAD != null)
                {
                    foreach (SearchResult child in userFromAD)
                    {
                        User obj = new User();
                        if (child != null)
                        {
                            obj.EmailID = GetProperty(child, "mail"); //(string)child.Properties["mail"][0];
                            obj.DisplayName = GetProperty(child, "displayname");
                            obj.LoginID = GetProperty(child, "SAMAccountName");
                            //try
                            //{
                            //    obj.DepartmentName = GetProperty(child, "department");
                            //}
                            //catch
                            //{
                            //    obj.DepartmentName = "";
                            //}
                            listUsers.Add(obj);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return listUsers;
        }

        public static List<User> GetUsersFromADUsingDirectorySearch(string Name)
        {
            List<User> listUsers = new List<User>();
            try
            {
                SearchResultCollection userFromAD = ActiveDirectoryHelper.GetUsersFromADUsingDirectorySearch(Name);
                if (userFromAD != null)
                {
                    foreach (SearchResult child in userFromAD)
                    {
                        User obj = new User();
                        if (child != null)
                        {
                            obj.EmailID = GetProperty(child, "mail"); //(string)child.Properties["mail"][0];
                            obj.DisplayName = GetProperty(child, "displayname");
                            obj.LoginID = GetProperty(child, "SAMAccountName");                            
                            listUsers.Add(obj);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return listUsers;
        }

        public static List<User> GetUsersFromADUsingPrincipalContext(string Name)
        {
            List<User> listUsers = new List<User>();
            try
            {
                listUsers = ActiveDirectoryHelper.GetUsersFromADUsingPrincipalContext(Name);
                //if (usersFromAD != null)
                //{
                //    foreach (Principal child in usersFromAD)
                //    {
                //        User obj = new User();
                //        if (child != null)
                //        {
                //            //obj.EmailID = GetProperty(child, "mail"); //(string)child.Properties["mail"][0];
                //            //obj.DisplayName = GetProperty(child, "displayname");
                //            //obj.LoginID = GetProperty(child, "SAMAccountName");
                //            listUsers.Add(obj);
                //        }
                //    }
                //}
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteLog("Error in GetUsersFromADUsingPrincipalContext in Common method class. Error Message is " + ex.Message.ToString());
                throw ex;
            }
            return listUsers;
        }

        public static string GetProperty(SearchResult searchResult, string PropertyName)
        {
            if (searchResult.Properties.Contains(PropertyName))
            {
                return searchResult.Properties[PropertyName][0].ToString();
            }
            else
            {
                return string.Empty;
            }
        }

        public static Hashtable getMetaData(DraftDocument objDoc)
        {
            Hashtable metaData = new Hashtable(){{SharePointConstants.DocumentID,objDoc.DocumentID.ToString()},
                                                {SharePointConstants.DocumentNo,objDoc.DocumentNo.ToString()},
                                                {SharePointConstants.DocumentDescription,objDoc.DocumentDescription},
                                                {SharePointConstants.RevisionReason,objDoc.RevisionReason},
                                                {SharePointConstants.DepartmentName,objDoc.DepartmentName},
                                                {SharePointConstants.SectionName,objDoc.SectionName},
                                                {SharePointConstants.ProjectName,objDoc.ProjectName},
                                                {SharePointConstants.DocumentCategoryName,objDoc.DocumentCategoryName}};
            return metaData;
        }
        public static DraftDocument GetDocumentObject(NameValueCollection RequestForm)
        {
            Dictionary<string, object> dict2 = new Dictionary<string, object>();
            foreach (string str in RequestForm.AllKeys)
            {
                if (RequestForm[str] != null)
                {
                    if (str == "MultipleApprovers")
                        dict2.Add(str, JsonConvert.DeserializeObject(RequestForm[str]));
                    else if (str == "MultipleApproversDisplay")
                        dict2.Add(str, JsonConvert.DeserializeObject(RequestForm[str]));
                    else dict2.Add(str, RequestForm[str]);

                }
            }
            string json = JsonConvert.SerializeObject(dict2);
            DraftDocument objDoc = JsonConvert.DeserializeObject<DraftDocument>(json);
            return objDoc;
        }
        public static LoginUser GetUserObject(string strJSON)
        {
            LoginUser[] objDoc = JsonConvert.DeserializeObject<LoginUser[]>(strJSON);
            return objDoc[0];
        }
        public static List<User> GetUsers(string strJSON)
        {
            List<User> objDoc = JsonConvert.DeserializeObject<List<User>>(strJSON);
            return objDoc;
        }

        public static Object GetObject(NameValueCollection RequestForm)
        {
            var formDictionary = RequestForm.AllKeys
                     .Where(p => RequestForm[p] != "null")
                     .ToDictionary(p => p, p => RequestForm[p]);
            string json = JsonConvert.SerializeObject(formDictionary);
            Object objDoc = JsonConvert.DeserializeObject<Object>(json);
            return objDoc;
        }
        public static NameValueCollection JSONDeserialize(this NameValueCollection _nvc, string _serializedString)
        {
            var deserializedobject = JsonConvert.DeserializeObject<Dictionary<string, string[]>>(_serializedString);
            foreach (var strCol in deserializedobject.Values)
                foreach (var str in strCol)
                    _nvc.Add(deserializedobject.FirstOrDefault(x => x.Value.Contains(str)).Key, str);
            return _nvc;
        }
        public static NameValueCollection ToNameValueCollection<T>(this T dynamicObject)
        {
            var nameValueCollection = new NameValueCollection();
            foreach (PropertyDescriptor propertyDescriptor in TypeDescriptor.GetProperties(dynamicObject))
            {
                string value = propertyDescriptor.GetValue(dynamicObject).ToString();
                nameValueCollection.Add(propertyDescriptor.Name, value);
            }
            return nameValueCollection;
        }

        public static T CreateItemFromRow<T>(DataRow row) where T : new()
        {
            // create a new object
            T item = new T();

            // set the item
            SetItemFromRow(item, row);

            // return 
            return item;
        }
        public static void SetItemFromRow<T>(T item, DataRow row) where T : new()
        {
            // go through each column
            foreach (DataColumn c in row.Table.Columns)
            {
                // find the property for the column
                PropertyInfo p = item.GetType().GetProperty(c.ColumnName);

                // if exists, set the value
                if (p != null && row[c] != DBNull.Value)
                {
                    p.SetValue(item, row[c], null);
                }
            }
        }

        public static DraftDocument GetDocumentObject(DataRow dr)
        {
            DraftDocument objDocument = new DraftDocument();
            try
            {

                if (dr.Table.Columns.Contains("DocumentPublishID") && dr["DocumentPublishID"] != DBNull.Value)
                    objDocument.DocumentPublishID = new Guid(dr["DocumentPublishID"].ToString());
                if (dr.Table.Columns.Contains("DocumentPublishID") && dr["DocumentID"] != DBNull.Value)
                    objDocument.DocumentID = new Guid(dr["DocumentID"].ToString());
                if (dr.Table.Columns.Contains("DocumentPublishID") && dr["DocumentNo"] != DBNull.Value)
                    objDocument.DocumentNo = dr["DocumentNo"].ToString();
                if (dr.Table.Columns.Contains("DocumentPublishID") && dr["EditableDocumentName"] != DBNull.Value)
                    objDocument.EditableDocumentName = dr["EditableDocumentName"].ToString();
                if (dr["EditableFilePath"] != DBNull.Value)
                    objDocument.EditableFilePath = dr["EditableFilePath"].ToString();
                if (dr["ReadableDocumentName"] != DBNull.Value)
                    objDocument.ReadableDocumentName = dr["ReadableDocumentName"].ToString();
                if (dr["ReadableFilePath"] != DBNull.Value)
                    objDocument.ReadableFilePath = dr["ReadableFilePath"].ToString();
                if (dr["DocumentDescription"] != DBNull.Value)
                    objDocument.DocumentDescription = dr["DocumentDescription"].ToString();
                if (dr["DepartmentName"] != DBNull.Value)
                    objDocument.DepartmentName = dr["DepartmentName"].ToString();
                if (dr["DepartmentCode"] != DBNull.Value)
                    objDocument.DepartmentCode = dr["DepartmentCode"].ToString();
                if (dr["SectionName"] != DBNull.Value)
                    objDocument.SectionName = dr["SectionName"].ToString();
                if (dr["SectionCode"] != DBNull.Value)
                    objDocument.SectionCode = dr["SectionCode"].ToString();
                if (dr["ProjectName"] != DBNull.Value)
                    objDocument.ProjectName = dr["ProjectName"].ToString();
                if (dr["ProjectCode"] != DBNull.Value)
                    objDocument.ProjectCode = dr["ProjectCode"].ToString();
                if (dr["DocumentCategoryName"] != DBNull.Value)
                    objDocument.DocumentCategoryName = dr["DocumentCategoryName"].ToString();
                if (dr["DocumentCategoryCode"] != DBNull.Value)
                    objDocument.DocumentCategoryCode = dr["DocumentCategoryCode"].ToString();
                if (dr["Comments"] != DBNull.Value)
                    objDocument.Comments = dr["Comments"].ToString();
                if (dr["DraftVersion"] != DBNull.Value)
                    objDocument.DraftVersion = Convert.ToDecimal(dr["DraftVersion"].ToString());
                if (dr["UploadedUserName"] != DBNull.Value)
                    objDocument.UploadedUserName = dr["UploadedUserName"].ToString();
                if (dr["CreatedDate"] != DBNull.Value)
                    objDocument.CreatedDate = DateTime.Parse(dr["CreatedDate"].ToString());
                if (dr["CurrentStage"] != DBNull.Value)
                    objDocument.CurrentStage = dr["CurrentStage"].ToString();
                if (dr["WFExecutionID"] != DBNull.Value)
                    objDocument.WFExecutionID = (Guid)dr["WFExecutionID"];
                if (dr["UploadedUserID"] != DBNull.Value)
                    objDocument.UploadedUserID = (Guid)dr["UploadedUserID"];
                if (dr["WorkflowID"] != DBNull.Value)
                    objDocument.WorkflowID = new Guid(dr["WorkflowID"].ToString());
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return objDocument;
        }

        public static string SendApprovalMail()
        {
            string strMailBody = string.Empty;
            return strMailBody;
        }

        public static void SendMail(string toEmail, string subject, string messageBody)
        {
            try
            {
                int port = Convert.ToInt32(QMSConstants.SMTPPORT);
                string host = QMSConstants.SMTPHOST;// "smtp.office365.com";
                string username = QMSConstants.EmailUserName; //"rajesh.m@canopusgbs.com";
                string password = QMSConstants.EmailPassword; //"2@Chakalakonda";
                string mailFrom = QMSConstants.EmailFrom; //"rajesh.m@canopusgbs.com";        

                using (SmtpClient client = new SmtpClient())
                {
                    MailAddress from = new MailAddress(mailFrom);
                    MailMessage message = new MailMessage
                    {
                        From = from
                    };
                    message.To.Add(toEmail);
                    message.Bcc.Add("rajesh.m-ext@tataelectronics.co.in");
                    //message.Bcc.Add("siva.d-ext@tataelectronics.co.in");
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
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
        }

        public static string EncryptFile(byte[] byteArray)
        {
            try
            {
                string password = QMSConstants.EncryptionKey; // Your Key Here
                UnicodeEncoding UE = new UnicodeEncoding();
                byte[] key = UE.GetBytes(password);

                string cryptFile = QMSConstants.LogPath + "\\" + Guid.NewGuid() + ".docx";
                FileStream fsCrypt = new FileStream(cryptFile, FileMode.Create);

                RijndaelManaged RMCrypto = new RijndaelManaged();

                CryptoStream cs = new CryptoStream(fsCrypt,
                    RMCrypto.CreateEncryptor(key, key),
                    CryptoStreamMode.Write);

                foreach (byte bt in byteArray)
                {
                    cs.WriteByte(bt);
                }

                cs.Close();
                fsCrypt.Close();
                return cryptFile;
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
        }
        public static void DecryptFile(string cryptFile)
        {
            try
            {
                string password = QMSConstants.EncryptionKey; // Your Key Here

                UnicodeEncoding UE = new UnicodeEncoding();
                byte[] key = UE.GetBytes(password);

                string decryptFile = QMSConstants.LogPath + "\\" + Guid.NewGuid() + ".docx";

                FileStream fsCrypt = new FileStream(cryptFile, FileMode.Open);

                RijndaelManaged RMCrypto = new RijndaelManaged();

                CryptoStream cs = new CryptoStream(fsCrypt,
                    RMCrypto.CreateDecryptor(key, key),
                    CryptoStreamMode.Read);

                byte[] byteArray = new byte[fsCrypt.Length];

                int data, i = 0;
                while ((data = cs.ReadByte()) != -1)
                {
                    byteArray[i] = (byte)data;
                    i++;
                }
                cs.Close();
                fsCrypt.Close();
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
        }
        public static string GetConfigValue(string strParamKey)
        {
            string strValue = string.Empty;
            try
            {
                using (SqlConnection con = new SqlConnection(QMSConstants.DBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(QMSConstants.spGetConfigValue, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@ParamKey", SqlDbType.NVarChar, 100).Value = strParamKey;

                        var ParamValue = cmd.Parameters.Add("@ParamValue", SqlDbType.NVarChar, -1);
                        ParamValue.Direction = ParameterDirection.Output;
                        con.Open();
                        cmd.ExecuteNonQuery();
                        strValue = (string)ParamValue.Value;
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return strValue;
        }
    }
}
