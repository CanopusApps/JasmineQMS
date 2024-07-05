using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace TEPL.QMS.BLL.Interface
{
    public interface IDocumentUpload
    {
        string UploadDocument(string destination, string docLib, string folderPath, string filename, byte[] documentBytes);
        string UploadDocument(string destination, string docLib, string folderPath, string filename, byte[] documentBytes, Hashtable Metadata);
        DataTable GetListData(string siteURL, string lstName, List<string> lsColumns);
        List<List<string>> GetListData(string siteURL, string lstName, string viewFeilds, string strCondition);
    }
}
