using BEBIZ.Upload.Sheets;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace BizPlanReport.Upload.Sheets
{
    public class DCGProductPlan : BaseSheet
    {
        public DCGProductPlan(DataSet ds, string sl) : base(ds, sl) { }

        protected override void Init()
        {
            RowNumber = 2;//what should be the row number?
            SheetName = "DCG Product Plan";
            TableName = "bizplan_DCG_Product_Plan";
            MandatoryColumns = new string[] { "DCG Unit (for ADM only)", "DCG Product Name", "Master Customer Code", "Core/Digital/Emerging", "Digital Offering Name" };//SL changed to Service Line
            DuplicateColumns = new string[] { "DCG Unit (for ADM only)", "DCG Product Name", "Master Customer Code", "Core/Digital/Emerging", "Digital Offering Name" };//SL changed to Service Line
            NumericColumns = new string[] {
                "Q1 FY23",
                "Q2 FY23",
                "Q3 FY23",
                "Q4 FY23",
                "FY23",
                "Q1 FY24",
                "Q2 FY24",
                "Q3 FY24 (Plan)",
                "Q4 FY24 (Plan)",
                "FY24",
                "Q1FY25 (Plan)",
                "Q2FY25 (Plan)",
                "Q3FY25 (Plan)",
                "Q4FY25 (Plan)",
                "FY25 (Plan)",
                "FY26 (Plan)",
                "FY27 (Plan)"


            };//changed

            DictMapping.Add("DCG Unit (for ADM only)", "DCG Unit"); //changed
            DictMapping.Add("DCG Product Name", "DCG Product Name");
            DictMapping.Add("Master Customer Code", "Master Customer Code");
            DictMapping.Add("Core/Digital/Emerging", "Core/Digital/Emerging");
            DictMapping.Add("Digital Offering Name", "Digital Offering Name");

            DictMapping.Add("Q1 FY23", "Q1'23");
            DictMapping.Add("Q2 FY23", "Q2'23");
            DictMapping.Add("Q3 FY23", "Q3'23");
            DictMapping.Add("Q4 FY23", "Q4'23");
            DictMapping.Add("FY23", "FY23");
            DictMapping.Add("Q1 FY24", "Q1'24");
            DictMapping.Add("Q2 FY24", "Q2'24");
            DictMapping.Add("Q3 FY24 (Plan)", "Q3'24 (Plan)");                      //column mapping changed
            DictMapping.Add("Q4 FY24 (Plan)", "Q4'24 (Plan)");                      //column mapping changed
            DictMapping.Add("FY24", "FY24");                      //column mapping changed
            DictMapping.Add("Q1FY25 (Plan)", "Q1'25 (Plan)");
            DictMapping.Add("Q2FY25 (Plan)", "Q2'25 (Plan)");
            DictMapping.Add("Q3FY25 (Plan)", "Q3'25 (Plan)");
            DictMapping.Add("Q4FY25 (Plan)", "Q4'25 (Plan)");
            DictMapping.Add("FY25 (Plan)", "FY25 (Plan)");
            DictMapping.Add("FY26 (Plan)", "FY26 (Plan)");
            DictMapping.Add("FY27 (Plan)", "FY27 (Plan)");
           
        }
    }
}