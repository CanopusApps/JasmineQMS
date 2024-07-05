using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TEPL.QMS.Common.Models;

namespace TEPL.QMS.DAL.Database.Interface
{
    public interface IDocumentOperations
    {
        DraftDocument GenerateDocumentNo(DraftDocument objDoc);
        string GetDocumentDetailsByID(string role, Guid loggedInUserID, Guid DocumentID);
        string GetDocumentDetailsByNo(string DocumentNo);
    }
}
