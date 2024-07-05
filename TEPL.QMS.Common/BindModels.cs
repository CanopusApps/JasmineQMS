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

namespace TEPL.QMS.Common
{
    public static class BindModels
    {
        public static List<T> ConvertJSON<T>(string strJSON)
        {
            List<T> data = new List<T>();
            data=JsonConvert.DeserializeObject<List<T>>(strJSON);
            return data;
        }
        public static T ConvertJSONObject<T>(string strJSON)
        {
            T data;
            data = JsonConvert.DeserializeObject<T>(strJSON);
            return data;
        }
        public static List<T> ConvertDataTable<T>(DataTable dt)
        {
            List<T> data = new List<T>();
            foreach (DataRow row in dt.Rows)
            {
                T item = GetItem<T>(row);
                data.Add(item);
            }
            return data;
        }
        private static T GetItem<T>(DataRow dr)
        {
            Type temp = typeof(T);
            T obj = Activator.CreateInstance<T>();

            foreach (DataColumn column in dr.Table.Columns)
            {
                foreach (PropertyInfo pro in temp.GetProperties())
                {
                    if (pro.Name == column.ColumnName)
                        pro.SetValue(obj, dr[column.ColumnName], null);
                    else
                        continue;
                }
            }
            return obj;
        }
        public static DraftDocument GetDocumentObject(NameValueCollection RequestForm)
        {
            var formDictionary = RequestForm.AllKeys
                     .Where(p => RequestForm[p] != "null")
                     .ToDictionary(p => p, p => RequestForm[p]);
            string json = JsonConvert.SerializeObject(formDictionary);
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
        public static DraftDocument GetDocumentObject1(NameValueCollection RequestForm)
        {
            DraftDocument objDoc = new DraftDocument();
            if (!string.IsNullOrEmpty(RequestForm[SharePointConstants.DepartmentCode]))
                objDoc.DepartmentCode = RequestForm[SharePointConstants.DepartmentCode].ToString();
            if (!string.IsNullOrEmpty(RequestForm[SharePointConstants.DocumentDescription]))
                objDoc.DepartmentCode = RequestForm[SharePointConstants.DocumentDescription].ToString();
            if (!string.IsNullOrEmpty(RequestForm[SharePointConstants.DepartmentName]))
                objDoc.DepartmentName = RequestForm[SharePointConstants.DepartmentName].ToString();
            if (!string.IsNullOrEmpty(RequestForm[SharePointConstants.SectionCode].ToString()))
                objDoc.SectionCode = RequestForm[SharePointConstants.SectionCode].ToString();
            if (!string.IsNullOrEmpty(RequestForm[SharePointConstants.SectionName].ToString()))
                objDoc.SectionName = RequestForm[SharePointConstants.SectionName].ToString();
            if (!string.IsNullOrEmpty(RequestForm[SharePointConstants.ProjectCode].ToString()))
                objDoc.ProjectCode = RequestForm[SharePointConstants.ProjectCode].ToString();
            if (!string.IsNullOrEmpty(RequestForm[SharePointConstants.ProjectName].ToString()))
                objDoc.ProjectName = RequestForm[SharePointConstants.ProjectName].ToString();
            if (!string.IsNullOrEmpty(RequestForm[SharePointConstants.DocumentCategoryCode].ToString()))
                objDoc.DocumentCategoryCode = RequestForm[SharePointConstants.DocumentCategoryCode].ToString();
            if (!string.IsNullOrEmpty(RequestForm[SharePointConstants.DocumentCategoryName].ToString()))
                objDoc.DocumentCategoryName = RequestForm[SharePointConstants.DocumentCategoryName].ToString();
            if (!string.IsNullOrEmpty(RequestForm[SharePointConstants.DocumentCategoryName].ToString()))
                objDoc.DocumentCategoryName = RequestForm[SharePointConstants.DocumentCategoryName].ToString();
            if (!string.IsNullOrEmpty(RequestForm[SharePointConstants.DocumentNo].ToString()))
                objDoc.DocumentNo = RequestForm[SharePointConstants.DocumentNo].ToString();
            if (!string.IsNullOrEmpty(RequestForm[SharePointConstants.DocumentID].ToString()))
                objDoc.DocumentID = new Guid(RequestForm[SharePointConstants.DocumentID].ToString());
            if (!string.IsNullOrEmpty(RequestForm[SharePointConstants.PublishingDocumentID].ToString()))
                objDoc.DocumentPublishID = new Guid(RequestForm[SharePointConstants.PublishingDocumentID].ToString());
            if (!string.IsNullOrEmpty(RequestForm[SharePointConstants.WFExecutionID].ToString()))
                objDoc.WFExecutionID = new Guid(RequestForm[SharePointConstants.WFExecutionID].ToString());
            if (!string.IsNullOrEmpty(RequestForm[SharePointConstants.Comments]))
            {
                objDoc.Comments = RequestForm[SharePointConstants.Comments].ToString();
            }
            else
            {
                objDoc.Comments = "";
            }
            objDoc.CompanyCode = QMSConstants.CompanyCode;
            if (!string.IsNullOrEmpty(RequestForm[SharePointConstants.DraftVersion].ToString()))
                objDoc.DraftVersion = Convert.ToDecimal(RequestForm[SharePointConstants.DraftVersion].ToString());
            return objDoc;
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
                if (dr.Table.Columns.Contains("DocumentPublishID") && dr["ReadableDocumentName"] != DBNull.Value)
                    objDocument.ReadableDocumentName = dr["ReadableDocumentName"].ToString();
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
                if (dr["EditableFilePath"] != DBNull.Value)
                    objDocument.EditableFilePath = dr["EditableFilePath"].ToString();
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
                    objDocument.WFExecutionID = new Guid(dr["WFExecutionID"].ToString());
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
    }
}
