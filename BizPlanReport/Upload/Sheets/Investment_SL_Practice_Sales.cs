using BEBIZ4.Common;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace BEBIZ.Upload.Sheets
{
    public class Investment_SL_Practice_Sales : BaseSheet
    {
        public Investment_SL_Practice_Sales(DataSet ds, string sl) : base(ds, sl) { }

        protected override void Init()
        {
            RowNumber = 2;
            SheetName = "Investment_SL_Practice_Sales";
            TableName = "BizPlan_SL_Practice_Sales_Investment_New";
            MandatoryColumns = new string[] {"Service Line" ,"Sales Region","Geo","Country"
,"Role Designation","Job Level" };//,"Offering Area"
            DuplicateColumns = new string[] {  "Service Line" ,"Sales Region","Geo","Country"
,"Role Designation","Job Level" };//,"Offering Area","Remarks"
            NumericColumns = new string[] {
  "FY'23 Head Count (HC)"
,"Current HC"
,"Current Practice Sales Open Indents",
"Q4FY'24 HC Plan",
"Q1FY'25 HC Plan",
"Q2FY'25 HC Plan",
"Q3FY'25 HC Plan",
"Q4FY'25 HC Plan",
"FY'25 HC Plan",
"FY'26 HC Plan",
"FY'27 HC Plan"



               // "FY23 Quarterly delta %","FY21 Quarterly delta %",
            //"FY22 Quarterly delta %"
            };
           // DictLookUpMapping.Add("Investment Category ",EnumLookUp.InvestmentCategory_PracticeSales);

            DictMapping.Add("Service Line", "SL");
            DictMapping.Add("Sales Region", "SalesRegion");
            DictMapping.Add("Geo", "Geo");
            DictMapping.Add("Country", "Country");
            DictMapping.Add("Role Designation", "RoleDesignation");
            DictMapping.Add("Job Level", "JobLevel");
            //DictMapping.Add("Offering Area", "Offering_Area");
            //DictMapping.Add("Remarks", "Remarks");
            DictMapping.Add("FY'23 Head Count (HC)", "PrevFY_HeadCount");
            DictMapping.Add("Current HC", "Current_HeadCount");
            DictMapping.Add("Current Practice Sales Open Indents", "Current_Practice_Sales_Open_Indents");
            DictMapping.Add("Q4FY'24 HC Plan", "FY1Q4_EST_INV");
            DictMapping.Add("Q1FY'25 HC Plan", "FY2Q1_EST_INV");
            DictMapping.Add("Q2FY'25 HC Plan", "FY2Q2_EST_INV");
            DictMapping.Add("Q3FY'25 HC Plan", "FY2Q3_EST_INV");
            DictMapping.Add("Q4FY'25 HC Plan", "FY2Q4_EST_INV");
            DictMapping.Add("FY'25 HC Plan", "FY2_EST_INV");
            DictMapping.Add("FY'26 HC Plan", "FY3Q4_EST_INV");
            DictMapping.Add("FY'27 HC Plan", "FY4Q4_EST_INV");



        }
    }

   
}