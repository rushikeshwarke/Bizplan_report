using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace BEBIZ.Upload.Sheets
{
    public class PartnerEcoSystemsAlliances : BaseSheet
    {
        public PartnerEcoSystemsAlliances(DataSet ds, string sl) : base(ds, sl) { }

        protected override void Init()
        {
            RowNumber = 9;
            SheetName = "Partner EcoSystems_Alliances";
            TableName = "BizPlan_Partner_EcoSystems_Alliances";
            MandatoryColumns = new string[] { "Service Line" };
            DuplicateColumns = new string[] { "Service Line", "Dimension", "Ecosystem", "Focus Area", "Focus Sub-Area _(Optional)", "Owner Mail ID", "Alliance Name", "Remarks" , "Geo" };
            NumericColumns = new string[] {  "TCV (MUSD)", "Revenue_(MUSD)",
                                            "TCV (MUSD)1", "Revenue_(MUSD)1", "TCV (MUSD)2", "Revenue_(MUSD)2" ,
                                            "TCV (MUSD)3","Revenue_(MUSD)3"};

            DictMapping.Add("Service Line", "SL");
            DictMapping.Add("Dimension", "Dimension");
            DictMapping.Add("Ecosystem", "Ecosystem");
            DictMapping.Add("Focus Area", "Focus_Area");
            DictMapping.Add("Focus Sub-Area _(Optional)", "Focus_SubArea");
            DictMapping.Add("Owner Mail ID", "Owner_MailId");
            DictMapping.Add("Alliance Name", "Alliance_Name");
            DictMapping.Add("Geo", "Geo");
            DictMapping.Add("Relationship Status", "Relation_Status");
            //DictMapping.Add("Current Status of the Partner EcoSystem / Alliances Plan", "Current_Status");
            //DictMapping.Add("If Plan is in progress, by when it will be ready? (dd-mmm-yyyy)", "Estimated_date");
            DictMapping.Add("Remarks", "Remarks");
            //DictMapping.Add("TCV (MUSD)", "FY1_TCV");
            //DictMapping.Add("Revenue_(MUSD)", "FY1_Revenue");
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