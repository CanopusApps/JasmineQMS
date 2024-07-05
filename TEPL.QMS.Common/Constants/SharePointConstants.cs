using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TEPL.QMS.Common.Constants
{
    public static class SharePointConstants
    {
        public static string siteURL = Convert.ToString(ConfigurationManager.AppSettings["siteURL"]);
        public static string tempDocLib = Convert.ToString(ConfigurationManager.AppSettings["tempDocLib"]);
        public static string editDocLib = Convert.ToString(ConfigurationManager.AppSettings["editDocLib"]);
        public static string publishDocLib = Convert.ToString(ConfigurationManager.AppSettings["publishDocLib"]);

        public static string mstListViewFields = @"<FieldRef Name='Title'></FieldRef><FieldRef Name='Code'></FieldRef><FieldRef Name='ID'></FieldRef>";
        //public static string mstListViewFields = @"<FieldRef Name='Title'></FieldRef><FieldRef Name='Code'></FieldRef><FieldRef Name='ID'></FieldRef>";
        public static string mstListCondition = @"<Where><Eq><FieldRef Name='Active' /><Value Type='Boolean'>1</Value></Eq></Where>";
        public static string mstListDCViewFeilds = "< FieldRef Name='FileLeafRef'/><FieldRef Name = 'FileRef' />< FieldRef Name='Title'/><FieldRef Name = 'LinkFilename' />< FieldRef Name='Code'/>";
        public static string DocumentID = "DocumentID";
        public static string DocumentNo = "DocumentNo";
        public static string DocumentDescription = "DocumentDescription";
        public static string RevisionReason = "RevisionReason";
        public static string DepartmentName = "DepartmentName";
        public static string DepartmentCode = "DepartmentCode";
        public static string SectionName = "SectionName";
        public static string SectionCode = "SectionCode";
        public static string ProjectName = "ProjectName";
        public static string ProjectCode = "ProjectCode";
        public static string DocumentCategoryName = "DocumentCategoryName";
        public static string DocumentCategoryCode = "DocumentCategoryCode";

        public static string PublishingDocumentID = "PublishingDocumentID";
        public static string WFExecutionID = "WFExecutionID";
        public static string Comments = "Comments";
        public static string DraftVersion = "DraftVersion";

        public static string departmentListName = "Departments";
        public static string sectionListName = "Sections";
        public static string projectListName = "Projects";
        public static string categoryListName = "DocumentCategory";
        public static string approvalMatrixListName = Convert.ToString(ConfigurationManager.AppSettings["ApprovalMatrix"]);
    }
}
