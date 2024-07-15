using BEBIZ.Upload.Sheets;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace BizPlanReport.Upload.Sheets
{
    public class Maximus_Margin_Intervention : BaseSheet
    {
        public Maximus_Margin_Intervention(DataSet ds, string sl) : base(ds, sl) { }

        protected override void Init()
        {
            RowNumber = 2;//what should be the row number?
            SheetName = "Maximus - Margin Intervention";
            TableName = "bizplan_Maximus_OM";
            MandatoryColumns = new string[] { "Unit", "Subunit", "Master Customer Code" };//SL changed to Service Line
            DuplicateColumns = new string[] { "Unit", "Subunit", "Master Customer Code" };//SL changed to Service Line
            NumericColumns = new string[] {
                "FY22",
                "Q1'23",
                "Q2'23",
                "Q3'23",
                "Q4'23",
                "FY23",
                "Q1'24",
                "Q2'24",
                "Q3'24",
                "Q4'24",
                "FY24",
                "Q1'241",
                "Q2'241",
                "Q3'241",
                "Q4'241",
                "FY241"



            };//changed

            DictMapping.Add("Unit", "Unit"); //changed
            DictMapping.Add("Subunit", "Subunit");
            DictMapping.Add("Master Customer Code", "Master_Customer_Code");

            DictMapping.Add("FY22", "FY22_OM");
            DictMapping.Add("Q1'23", "Q1'23_OM");
            DictMapping.Add("Q2'23", "Q2'23_OM");
            DictMapping.Add("Q3'23", "Q3'23_OM");
            DictMapping.Add("Q4'23", "Q4'23_OM");
            DictMapping.Add("FY23", "FY23_OM");

            DictMapping.Add("Q1'24", "Q1'24_OM");
            DictMapping.Add("Q2'24", "Q2'24_OM");
            DictMapping.Add("Q3'24", "Q3'24_OM");
            DictMapping.Add("Q4'24", "Q4'24_OM");
            DictMapping.Add("FY24", "FY24_OM");

            DictMapping.Add("Q1'241", "Q1'24_MI");
            DictMapping.Add("Q2'241", "Q2'24_MI");
            DictMapping.Add("Q3'241", "Q3'24_MI");
            DictMapping.Add("Q4'241", "Q4'24_MI");
            DictMapping.Add("FY241", "FY24_MI");

            DictMapping.Add("What are the margin intervention plan and FY24 QOQ target", "Margin Intervention Plan and FY24 QOQ Target");

        }

    }
}