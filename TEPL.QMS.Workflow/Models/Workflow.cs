using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TEPL.QMS.Workflow.Models
{
    public class Workflow
    {
        string Name { set; get; }
        List<Stage> Stages { set; get; }
    }
    public class Stage
    {
        public Guid CurrentStageID { set; get; }
        public string CurrentStage { set; get; }
        public Guid PreviousStageID { set; get; }
        public string PreviousStage { set; get; }
        public Guid NextStageID { set; get; }
        public string NextStage { set; get; }
        public bool IsDocumentLevelRequired { get; set; }
        public bool IsDepartmentRequired { get; set; }
    }
}
