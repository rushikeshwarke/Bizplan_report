using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace BEBIZ.Upload.Sheets
{
    public class InternalCollaboration : BaseSheet
    {
        public InternalCollaboration(DataSet ds) : base(ds)
        {



        }

        protected override void Init()
        {
            RowNumber = 4;
            SheetName = "Internal Collaboration";
            TableName = "BizPlan_Internal_Collaboration";
            MandatoryColumns = new string[] { "Service Line" };
            DuplicateColumns = new string[] { "Service Line", "Collaboration with", "Dimension", "Focus Area", "Focus Sub-Area _(Optional)", "Owner Mail ID", "Remarks" }; 
            NumericColumns = new string[] { "TCV (MUSD)", "Revenue_(MUSD)","TCV (MUSD)1", "Revenue_(MUSD)1","TCV (MUSD)2", "Revenue_(MUSD)2",
            "TCV (MUSD)3","Revenue_(MUSD)3"};

            DictMapping.Add("Service Line", "SL");
            DictMapping.Add("Collaboration with", "Collaboration");
            DictMapping.Add("Dimension", "Dimension");
            DictMapping.Add("Focus Area", "Focus_Area");
            DictMapping.Add("Focus Sub-Area _(Optional)", "Focus_SubArea");
            DictMapping.Add("Owner Mail ID", "Owner_MailId");
            //DictMapping.Add("Current Status of the Internal Collaboration Plan", "Current_Status");
            //DictMapping.Add("If Plan is in progress, by when it will be ready? (dd-mmm-yyyy)", "Estimated_Date");
            DictMapping.Add("Remarks", "Remarks");
            DictMapping.Add("TCV (MUSD)", "FY1_TCV");
            DictMapping.Add("Revenue_(MUSD)", "FY1_Revenue");
            DictMapping.Add("TCV (MUSD)1", "FY2_TCV");
            DictMapping.Add("Revenue_(MUSD)1", "FY2_Revenue");
            DictMapping.Add("TCV (MUSD)2", "FY3_TCV");
            DictMapping.Add("Revenue_(MUSD)2", "FY3_Revenue");
            DictMapping.Add("TCV (MUSD)3", "FY4_TCV");
            DictMapping.Add("Revenue_(MUSD)3", "FY4_Revenue");
        }


    }
}