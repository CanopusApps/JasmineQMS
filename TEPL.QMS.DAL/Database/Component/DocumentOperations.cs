using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using TEPL.QMS.Common;
using TEPL.QMS.Common.Constants;
using TEPL.QMS.Common.Models;
using TEPL.QMS.DAL.Database.Interface;
namespace TEPL.QMS.DAL.Database.Component
{
    public class DocumentOperations : IDocumentOperations
    {
        public DraftDocument GenerateDocumentNo(DraftDocument objDoc)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(QMSConstants.DBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(QMSConstants.spGenerateDocumentNo, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@WorkflowID", SqlDbType.UniqueIdentifier).Value = objDoc.WorkflowID;
                        cmd.Parameters.Add("@WorkflowStageID", SqlDbType.UniqueIdentifier).Value = objDoc.CurrentStageID;
                        cmd.Parameters.Add("@CompanyCode", SqlDbType.NVarChar, 10).Value = objDoc.CompanyCode;
                        cmd.Parameters.Add("@DepartmentCode", SqlDbType.NVarChar, 10).Value = objDoc.DepartmentCode;
                        cmd.Parameters.Add("@SectionCode", SqlDbType.NVarChar, 10).Value = objDoc.SectionCode;
                        cmd.Parameters.Add("@ProjectCode", SqlDbType.NVarChar, 10).Value = objDoc.ProjectCode;
                        cmd.Parameters.Add("@DocumentCategoryCode", SqlDbType.NVarChar, 10).Value = objDoc.DocumentCategoryCode;
                        cmd.Parameters.Add("@CreatedID", SqlDbType.UniqueIdentifier).Value = objDoc.UploadedUserID;

                        var DocumentNo = cmd.Parameters.Add("@DocumentNo", SqlDbType.NVarChar, 50);
                        DocumentNo.Direction = ParameterDirection.Output;
                        var DocumentID = cmd.Parameters.Add("@DocumentID", SqlDbType.UniqueIdentifier);
                        DocumentID.Direction = ParameterDirection.Output;
                        var WFExecutionID = cmd.Parameters.Add("@WFExecutionID", SqlDbType.UniqueIdentifier);
                        WFExecutionID.Direction = ParameterDirection.Output;
                        con.Open();
                        cmd.ExecuteNonQuery();
                        objDoc.DocumentNo = (string)DocumentNo.Value;
                        objDoc.DocumentID = (Guid)DocumentID.Value;
                        objDoc.WFExecutionID = (Guid)WFExecutionID.Value;
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return objDoc;
        }
        public string UpdateMultipleApprovers(DraftDocument objDoc)
        {
            string strReturn = string.Empty;
            try
            {
                //objDoc.RevisionReason = "";
                using (SqlConnection con = new SqlConnection(QMSConstants.DBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(QMSConstants.spUpdateMultipleApprovers, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@DocumentID", SqlDbType.UniqueIdentifier).Value = objDoc.DocumentID;
                        cmd.Parameters.Add("@flgMultipleApprovers", SqlDbType.Bit).Value = objDoc.IsMultipleApprovers;
                        if (objDoc.IsMultipleApprovers)
                            cmd.Parameters.Add("@MultipleApprovers", SqlDbType.NVarChar, -1).Value = JsonConvert.SerializeObject(objDoc.MultipleApprovers);
                        else
                            cmd.Parameters.Add("@MultipleApprovers", SqlDbType.NVarChar, -1).Value = DBNull.Value;
                        con.Open();
                        cmd.ExecuteNonQuery();
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return strReturn;
        }
        public DraftDocument DocumentUpdate(DraftDocument objDoc)
        {
            try
            {
                //objDoc.RevisionReason = "";
                using (SqlConnection con = new SqlConnection(QMSConstants.DBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(QMSConstants.spDocumentUpdate, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@DocumentID", SqlDbType.UniqueIdentifier).Value = objDoc.DocumentID;
                        cmd.Parameters.Add("@DocumentPublishID", SqlDbType.UniqueIdentifier).Value = objDoc.DocumentPublishID;
                        cmd.Parameters.Add("@EditableDocumentName", SqlDbType.NVarChar, 50).Value = objDoc.EditableDocumentName;
                        cmd.Parameters.Add("@EditableFilePath", SqlDbType.NVarChar, 500).Value = objDoc.EditableFilePath;
                        cmd.Parameters.Add("@ReadableDocumentName", SqlDbType.NVarChar, 50).Value = objDoc.ReadableDocumentName;
                        cmd.Parameters.Add("@ReadableFilePath", SqlDbType.NVarChar, 500).Value = objDoc.ReadableFilePath;
                        cmd.Parameters.Add("@DocumentDescription", SqlDbType.NVarChar, 500).Value = objDoc.DocumentDescription;
                        cmd.Parameters.Add("@RevisionReason", SqlDbType.NVarChar, 500).Value = objDoc.RevisionReason;
                        SqlParameter param = cmd.Parameters.Add("@DraftVersion", SqlDbType.Decimal);
                        param.Precision = 18;
                        param.Scale = 3;
                        param.Value = objDoc.DraftVersion;
                        SqlParameter paramEdit = cmd.Parameters.Add("@EditVersion", SqlDbType.Decimal);
                        paramEdit.Precision = 18;
                        paramEdit.Scale = 3;
                        paramEdit.Value = objDoc.EditVersion;
                        cmd.Parameters.Add("@Comments", SqlDbType.NVarChar, -1).Value = objDoc.Comments;
                        cmd.Parameters.Add("@CreatedID", SqlDbType.UniqueIdentifier).Value = objDoc.UploadedUserID;
                        cmd.Parameters.Add("@LastModifiedID", SqlDbType.UniqueIdentifier).Value = objDoc.UploadedUserID;

                        con.Open();
                        cmd.ExecuteNonQuery();
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return objDoc;
        }
        public DraftDocument DocumentUpdatePublised(DraftDocument objDoc)
        {
            try
            {
                //objDoc.RevisionReason = "";
                using (SqlConnection con = new SqlConnection(QMSConstants.DBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(QMSConstants.spDocumentUpdatePublished, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@DocumentID", SqlDbType.UniqueIdentifier).Value = objDoc.DocumentID;
                        cmd.Parameters.Add("@EditableDocumentName", SqlDbType.NVarChar, 50).Value = objDoc.EditableDocumentName;
                        cmd.Parameters.Add("@ReadableDocumentName", SqlDbType.NVarChar, 50).Value = objDoc.ReadableDocumentName;
                        cmd.Parameters.Add("@DocumentDescription", SqlDbType.NVarChar, 500).Value = objDoc.DocumentDescription;
                        cmd.Parameters.Add("@RevisionReason", SqlDbType.NVarChar, 500).Value = objDoc.RevisionReason;
                        SqlParameter EditVersion = cmd.Parameters.Add("@EditVersion", SqlDbType.Decimal);
                        EditVersion.Precision = 18;
                        EditVersion.Scale = 3;
                        EditVersion.Value = objDoc.EditVersion;
                        cmd.Parameters.Add("@LastModifiedID", SqlDbType.UniqueIdentifier).Value = objDoc.UploadedUserID;

                        con.Open();
                        cmd.ExecuteNonQuery();
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return objDoc;
        }
        public DraftDocument DocumentDescriptionUpdate(DraftDocument objDoc)
        {
            try
            {
                //objDoc.RevisionReason = "";
                using (SqlConnection con = new SqlConnection(QMSConstants.DBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(QMSConstants.spDocumentDescriptionUpdate, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@DocumentID", SqlDbType.UniqueIdentifier).Value = objDoc.DocumentID;
                        cmd.Parameters.Add("@DocumentDescription", SqlDbType.NVarChar, 500).Value = objDoc.DocumentDescription;
                        cmd.Parameters.Add("@RevisionReason", SqlDbType.NVarChar, 500).Value = objDoc.RevisionReason;
                        SqlParameter param = cmd.Parameters.Add("@DraftVersion", SqlDbType.Decimal);
                        param.Precision = 18;
                        param.Scale = 3;
                        param.Value = objDoc.DraftVersion;
                        cmd.Parameters.Add("@LastModifiedID", SqlDbType.UniqueIdentifier).Value = objDoc.UploadedUserID;

                        con.Open();
                        cmd.ExecuteNonQuery();
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return objDoc;
        }
        public DraftDocument DocumentPublish(DraftDocument objDoc)
        {
            try
            {
                //objDoc.RevisionReason = "";
                using (SqlConnection con = new SqlConnection(QMSConstants.DBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(QMSConstants.spDocumentPublish, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@DocumentID", SqlDbType.UniqueIdentifier).Value = objDoc.DocumentID;
                        cmd.Parameters.Add("@DocumentDescription", SqlDbType.NVarChar, 500).Value = objDoc.DocumentDescription;
                        cmd.Parameters.Add("@EditableDocumentName", SqlDbType.NVarChar, 50).Value = objDoc.EditableDocumentName;
                        cmd.Parameters.Add("@ReadableDocumentName", SqlDbType.NVarChar, 50).Value = objDoc.ReadableDocumentName;
                        cmd.Parameters.Add("@RevisionReason", SqlDbType.NVarChar, 500).Value = objDoc.RevisionReason;
                        SqlParameter param = cmd.Parameters.Add("@OriginalVersion", SqlDbType.Decimal);
                        param.Precision = 18;
                        param.Scale = 3;
                        param.Value = objDoc.OriginalVersion;

                        con.Open();
                        cmd.ExecuteNonQuery();
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return objDoc;
        }
        public DraftDocument DocumentVersionUpdate(DraftDocument objDoc)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(QMSConstants.DBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(QMSConstants.spDocumentVersionUpdate, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@DocumentID", SqlDbType.UniqueIdentifier).Value = objDoc.DocumentID;
                        cmd.Parameters.Add("@EditableDocumentName", SqlDbType.NVarChar, 50).Value = objDoc.EditableDocumentName;
                        SqlParameter param = cmd.Parameters.Add("@DraftVersion", SqlDbType.Decimal);
                        param.Precision = 18;
                        param.Scale = 3;
                        param.Value = objDoc.DraftVersion;
                        cmd.Parameters.Add("@LastModifiedID", SqlDbType.UniqueIdentifier).Value = objDoc.ActionedID;

                        con.Open();
                        cmd.ExecuteNonQuery();
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return objDoc;
        }
        public DataTable GetRequestedDocuments(Guid CreatedID)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(QMSConstants.DBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(QMSConstants.spGetRequestedDocuments, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@CreatedID", SqlDbType.UniqueIdentifier).Value = CreatedID;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            sda.Fill(dt);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return dt;
        }
        public string GetDocumentDetailsByID(string role, Guid loggedInUserID, Guid DocumentID)
        {
            string strReturn = string.Empty;
            try
            {
                using (SqlConnection con = new SqlConnection(QMSConstants.DBCon))
                {
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand(QMSConstants.spGetDocumentDetailsByID, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@DocumentID", SqlDbType.UniqueIdentifier).Value = DocumentID;
                        cmd.Parameters.Add("@Role", SqlDbType.NVarChar, 10).Value = role;
                        cmd.Parameters.Add("@LoginUserID", SqlDbType.UniqueIdentifier).Value = loggedInUserID;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            sda.Fill(dt);
                            strReturn = dt.Rows[0][0].ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return strReturn;
        }
        public string GetDocumentDetailsByNo(string DocumentNo)
        {
            string strReturn = string.Empty;
            try
            {
                using (SqlConnection con = new SqlConnection(QMSConstants.DBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(QMSConstants.spGetDocumentDetailsByNo, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@DocumentNo", SqlDbType.NVarChar, 50).Value = DocumentNo;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            sda.Fill(dt);
                            strReturn = dt.Rows[0][0].ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return strReturn;
        }
        public string GetPublishedDocumentDetailsByID(Guid UserID, Guid DocumentID)
        {
            string strReturn = string.Empty;
            try
            {
                using (SqlConnection con = new SqlConnection(QMSConstants.DBCon))
                {
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand(QMSConstants.spGetPublishedDocumentDetailsByID, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@DocumentID", SqlDbType.UniqueIdentifier).Value = DocumentID;
                        cmd.Parameters.Add("@UserID", SqlDbType.UniqueIdentifier).Value = UserID;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            sda.Fill(dt);
                            strReturn = dt.Rows[0][0].ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return strReturn;
        }
        public string GetPublishedDocumentHistoryByID(Guid UserID, Guid DocumentID)
        {
            string strReturn = string.Empty;
            try
            {
                using (SqlConnection con = new SqlConnection(QMSConstants.DBCon))
                {
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand(QMSConstants.spGetPublishedDocumentHistoryByID, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@DocumentID", SqlDbType.UniqueIdentifier).Value = DocumentID;
                        cmd.Parameters.Add("@UserID", SqlDbType.UniqueIdentifier).Value = UserID;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            sda.Fill(dt);
                            strReturn = dt.Rows[0][0].ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return strReturn;
        }
        public string GetPublishedDocumentHistoryDetailsByID(Guid UserID, Guid DocumentID)
        {
            string strReturn = string.Empty;
            try
            {
                using (SqlConnection con = new SqlConnection(QMSConstants.DBCon))
                {
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand(QMSConstants.spGetPublishedDocumentHistoryDetailsByID, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@DocumentID", SqlDbType.UniqueIdentifier).Value = DocumentID;
                        cmd.Parameters.Add("@UserID", SqlDbType.UniqueIdentifier).Value = UserID;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            sda.Fill(dt);
                            strReturn = dt.Rows[0][0].ToString();
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return strReturn;
        }
        public string GetDocumentLevel(string CategoryCode)
        {
            string strDocLevel = string.Empty;
            try
            {
                using (SqlConnection con = new SqlConnection(QMSConstants.DBCon))
                {
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand(QMSConstants.spGetDocumentLevel, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@CategoryCode", SqlDbType.NVarChar, 10).Value = CategoryCode;
                        var DocumentLevel = cmd.Parameters.Add("@DocumentLevel", SqlDbType.NVarChar, 10);
                        DocumentLevel.Direction = ParameterDirection.Output;
                        cmd.ExecuteNonQuery();
                        strDocLevel = (string)DocumentLevel.Value;
                    }
                    con.Close();
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return strDocLevel;
        }
        public DataTable GetApprovalPendingDocuments([Optional] Guid ActionedID)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(QMSConstants.DBCon))
                {
                    using (SqlCommand cmd = new SqlCommand(QMSConstants.spGetApprovalPendingDocuments, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@ActionedID", SqlDbType.UniqueIdentifier).Value = ActionedID;
                        cmd.Parameters.Add("@siteName", SqlDbType.NVarChar, 20).Value = QMSConstants.SiteName;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            sda.Fill(dt);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
            }
            return dt;
        }
        public string UploadDocument(string StoragePath, string baseFolder, string folderPath, string filename, Decimal DocVerions, byte[] byteArray)
        {
            try
            {
                string filePath = CommonMethods.CombineUrl(StoragePath, baseFolder, folderPath);
                if (File.Exists(filePath + "\\" + filename))
                {
                    BackUpFile(StoragePath, "History", baseFolder, folderPath, filename, DocVerions);
                }
                else if (!Directory.Exists(filePath))
                {
                    CreateFolders(filePath);
                }
                string cryptFile = CommonMethods.CombineUrl(QMSConstants.StoragePath, baseFolder, folderPath, filename);
                cryptFile = UploadEncryptedDocument(cryptFile, byteArray);//UploadAESEncryptedDocument(cryptFile, byteArray);
                return cryptFile;
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
        }
        public string UploadExternalDocument(string StoragePath, string baseFolder, string folderPath, string filename, string DocVerions, byte[] byteArray)
        {
            try
            {
                string filePath = CommonMethods.CombineUrl(StoragePath, baseFolder, folderPath);
                if (File.Exists(filePath + "\\" + filename))
                {
                    //filename = (new Guid()).ToString() + "_" + filename;
                    BackUpFileExtDocument(StoragePath, "History", baseFolder, folderPath, filename, DocVerions);
                }
                else if (!Directory.Exists(filePath))
                {
                    CreateFolders(filePath);
                }
                string cryptFile = CommonMethods.CombineUrl(QMSConstants.StoragePath, baseFolder, folderPath, filename);
                cryptFile = UploadEncryptedDocument(cryptFile, byteArray);//UploadAESEncryptedDocument(cryptFile, byteArray);
                return cryptFile;
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
        }
        public void BackUpFile(string StoragePath, string BackUpFolder, string baseFolder, string folderPath, string srcFilename, Decimal DocVerions)
        {
            try
            {
                string srcFilePath = CommonMethods.CombineUrl(StoragePath, baseFolder, folderPath);
                string desFilePath = CommonMethods.CombineUrl(QMSConstants.BackUpPath, BackUpFolder, baseFolder, folderPath);
                string extn = Path.GetExtension(srcFilename);
                string Version = string.Empty;
                Version = Math.Ceiling(DocVerions).ToString();
                string desFileName = srcFilename.Replace(extn, "_V" + Version + extn);
                if (!Directory.Exists(desFilePath))
                    CreateFolders(desFilePath);
                else
                {
                    int i = 0;
                    string tepmFileName = desFileName;
                    while (File.Exists(desFilePath + "/" + tepmFileName))
                    {
                        tepmFileName = desFileName.Replace(extn, "_V" + i + extn);
                    }
                    if (i != 0)
                        desFileName = desFileName.Replace(extn, "_V" + i + extn);
                }
                File.Move(srcFilePath + "/" + srcFilename, desFilePath + "/" + desFileName);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
        }
        public void BackUpFileExtDocument(string StoragePath, string BackUpFolder, string baseFolder, string folderPath, string srcFilename, string DocVerions)
        {
            string srcFilePath = CommonMethods.CombineUrl(StoragePath, baseFolder, folderPath);
            string desFilePath = CommonMethods.CombineUrl(QMSConstants.BackUpPath, BackUpFolder, baseFolder, folderPath);
            string extn = Path.GetExtension(srcFilename);
            string Version = string.Empty;
            Version = DocVerions.ToString();
            string desFileName = srcFilename.Replace(extn, "_V" + Version + extn);
            if (!Directory.Exists(desFilePath))
                CreateFolders(desFilePath);
            else
            {
                int i = 0;
                string tepmFileName = desFileName;
                while (File.Exists(desFilePath + "/" + tepmFileName))
                {
                    tepmFileName = desFileName.Replace(extn, "_V" + i + extn);
                }
                if (i != 0)
                    desFileName = desFileName.Replace(extn, "_V" + i + extn);
            }
            File.Move(srcFilePath + "/" + srcFilename, desFilePath + "/" + desFileName);
        }
        public void DeleteFile(string StoragePath, string baseFolder, string folderPath, string srcFilename, Decimal DocVerions)
        {
            try
            {
                string srcFilePath = CommonMethods.CombineUrl(StoragePath, baseFolder, folderPath);
                string desFilePath = CommonMethods.CombineUrl(QMSConstants.BackUpPath, "DeletedFiles", baseFolder, folderPath);
                string extn = Path.GetExtension(srcFilename);
                string Version = string.Empty;
                Version =Math.Ceiling(DocVerions).ToString();
                string desFileName = srcFilename.Replace(extn, "_D" + Version + extn);
                string tepmFileName = desFileName;
                if (!Directory.Exists(desFilePath))
                    CreateFolders(desFilePath);
                else
                {
                    int i = 0;
                    while (File.Exists(desFilePath + "\\" + tepmFileName))
                    {
                        tepmFileName = desFileName.Replace(extn, "_D" + i + extn);
                    }
                    if (i != 0)
                        desFileName = desFileName.Replace(extn, "_D" + i + extn);
                }
                File.Move(srcFilePath + "\\" + srcFilename, desFilePath + "\\" + desFileName);
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
        }

        public byte[] DownloadDocument(string DocURL, int DocVersion)
        {
            try
            {
                byte[] byteArray = DownloadDrecrytedDocument(DocURL, DocVersion);// DownloadAESDrecrytedDocument(DocURL, DocVersion);
                return byteArray;
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
        }

        public string UploadWithOutEncryptedDocument(string StoragePath, string baseFolder, string folderPath, string filename, Decimal DocVerions, byte[] byteArray)
        {
            try
            {
                string filePath = CommonMethods.CombineUrl(StoragePath, baseFolder, folderPath);
                if (File.Exists(filePath + "\\" + filename))
                {
                    BackUpFile(StoragePath, "History", baseFolder, folderPath, filename, DocVerions);
                }
                else if (!Directory.Exists(filePath))
                {
                    CreateFolders(filePath);
                }
                string cryptFile = CommonMethods.CombineUrl(QMSConstants.StoragePath, baseFolder, folderPath, filename);
                FileStream fsCrypt = new FileStream(cryptFile, FileMode.Create);

                try
                {
                    foreach (byte bt in byteArray)
                    {
                        fsCrypt.WriteByte(bt);
                    }
                }
                catch (Exception ex)
                {
                    LoggerBlock.WriteTraceLog(ex);
                    File.Delete(cryptFile);
                    throw ex;
                }
                finally
                {
                    fsCrypt.Close();
                }
                return cryptFile;
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
        }

        public byte[] DownloadWithOutDecryptedDocument(string DocURL, int DocVersion)
        {
            try
            {
                FileStream fsCrypt = new FileStream(DocURL, FileMode.Open);

                byte[] byteArray = new byte[fsCrypt.Length];
                try
                {
                    int data, i = 0;
                    while ((data = fsCrypt.ReadByte()) != -1)
                    {
                        byteArray[i] = (byte)data;
                        i++;
                    }
                }
                catch (Exception ex)
                {
                    LoggerBlock.WriteTraceLog(ex);
                    throw ex;
                }
                finally
                {
                    fsCrypt.Close();
                }
                return byteArray;
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
        }

        public string UploadEncryptedDocument(string inputFile, byte[] byteArray)
        {
            try
            {
                // Derive a key using PBKDF2 with 32 bytes output size (equivalent to 256 bits)
                byte[] key = DeriveKey(QMSConstants.EncryptionKey, QMSConstants.EncryptionKey, 32);
                byte[] iv = DeriveKey(QMSConstants.EncryptionKey, QMSConstants.EncryptionKey, 16);
                using (Aes aesAlg = Aes.Create())
                {
                    aesAlg.Key = key;
                    aesAlg.IV = iv;
                    aesAlg.Mode = CipherMode.CBC;
                    ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);

                    using (MemoryStream fsInput = new MemoryStream(byteArray))
                    {
                        using (FileStream fsOutput = new FileStream(inputFile, FileMode.Create, FileAccess.Write))
                        {
                            using (CryptoStream csEncrypt = new CryptoStream(fsOutput, encryptor, CryptoStreamMode.Write))
                            {
                                fsInput.CopyTo(csEncrypt);
                            }
                        }
                    }
                }
                return inputFile;
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
        }

        public byte[] DownloadDrecrytedDocument(string DocURL, int DocVersion)
        {
            try
            {
                // Derive a key using PBKDF2 with 32 bytes output size (equivalent to 256 bits)
                byte[] key = DeriveKey(QMSConstants.EncryptionKey, QMSConstants.EncryptionKey, 32);
                byte[] iv = DeriveKey(QMSConstants.EncryptionKey, QMSConstants.EncryptionKey, 16);
                using (Aes aesAlg = Aes.Create())
                {
                    aesAlg.Key = key;
                    aesAlg.IV = iv;
                    aesAlg.Mode = CipherMode.CBC;
                    ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);
                    Byte[] byteArray = null;

                    using (FileStream fsEncrypted = new FileStream(DocURL, FileMode.Open))
                    {
                        using (MemoryStream fsDecrypted = new MemoryStream())
                        {
                            using (ICryptoTransform decryptor = aesAlg.CreateDecryptor())
                            {
                                using (CryptoStream cs = new CryptoStream(fsEncrypted, decryptor, CryptoStreamMode.Read))
                                {
                                    byte[] buffer = new byte[4096];
                                    int bytesRead;
                                    while ((bytesRead = cs.Read(buffer, 0, buffer.Length)) > 0)
                                    {
                                        fsDecrypted.Write(buffer, 0, bytesRead);
                                    }
                                    byteArray = fsDecrypted.ToArray();
                                }
                            }
                        }
                    }
                    return byteArray;
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
        }

        static byte[] DeriveKey(string password, string salt, int outputSizeInBytes)
        {
            using (Rfc2898DeriveBytes pbkdf2 = new Rfc2898DeriveBytes(password, Encoding.UTF8.GetBytes(salt), 10000, HashAlgorithmName.SHA256))
            {
                return pbkdf2.GetBytes(outputSizeInBytes);
            }
        }
        public string UploadAESEncryptedDocument(string inputFile, byte[] byteArray)
        {
            try
            {
                string password = QMSConstants.EncryptionKey;
                byte[] salt = new byte[] { 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08 };
                RijndaelManaged aes = new RijndaelManaged();
                aes.KeySize = 256;
                aes.BlockSize = 128;
                var key = new Rfc2898DeriveBytes(password, salt, 1000);
                aes.Key = key.GetBytes(aes.KeySize / 8);
                aes.IV = key.GetBytes(aes.BlockSize / 8);
                aes.Mode = CipherMode.CBC;

                using (MemoryStream fsInput = new MemoryStream(byteArray))
                {
                    using (FileStream fsEncrypted = new FileStream(inputFile, FileMode.Create))
                    {
                        using (ICryptoTransform encryptor = aes.CreateEncryptor())
                        {
                            using (CryptoStream cs = new CryptoStream(fsEncrypted, encryptor, CryptoStreamMode.Write))
                            {
                                byte[] buffer = new byte[4096];
                                int bytesRead;
                                while ((bytesRead = fsInput.Read(buffer, 0, buffer.Length)) > 0)
                                {
                                    cs.Write(buffer, 0, bytesRead);
                                }
                            }
                        }
                    }
                }
                return inputFile;
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
        }
        public byte[] DownloadAESDrecrytedDocument(string DocURL, int DocVersion)
        {
            try
            {
                string password = QMSConstants.EncryptionKey;
                byte[] salt = new byte[] { 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08 };
                RijndaelManaged aes = new RijndaelManaged();
                aes.KeySize = 256;
                aes.BlockSize = 128;
                var key = new Rfc2898DeriveBytes(password, salt, 1000);
                aes.Key = key.GetBytes(aes.KeySize / 8);
                aes.IV = key.GetBytes(aes.BlockSize / 8);
                aes.Mode = CipherMode.CBC;
                Byte[] byteArray = null;
                using (FileStream fsEncrypted = new FileStream(DocURL, FileMode.Open))
                {
                    using (MemoryStream fsDecrypted = new MemoryStream())
                    {
                        using (ICryptoTransform decryptor = aes.CreateDecryptor())
                        {
                            using (CryptoStream cs = new CryptoStream(fsEncrypted, decryptor, CryptoStreamMode.Read))
                            {
                                byte[] buffer = new byte[4096];
                                int bytesRead;
                                while ((bytesRead = cs.Read(buffer, 0, buffer.Length)) > 0)
                                {
                                    fsDecrypted.Write(buffer, 0, bytesRead);
                                }
                                byteArray = fsDecrypted.ToArray();
                            }
                        }
                    }
                }
                return byteArray;
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteTraceLog(ex);
                throw ex;
            }
        }
        public void CreateFolders(string folderPath)
        {
            Directory.CreateDirectory(folderPath);
        }
    }
}
