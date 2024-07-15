using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace BEBIZ.Upload.Sheets
{
    public class InfosysAssetsPlan:BaseSheet
    {
        public InfosysAssetsPlan(DataSet ds, string sl): base(ds, sl)
        {

        }


        protected override void Init()
        {
            RowNumber = 2;
            SheetName = "Infosys Assets Plan";
            TableName = "BizPlan_Infosys_Assets_Plan";

            MandatoryColumns = new string[] { "Service Line", "Sub unit", "Name of Infosys Asset/IP", "Is Asset published in Service Store (Yes/No)",  "Is the revenue currently reported under digital business? Yes/No", };//SL changed to Service Line
         DuplicateColumns = new string[] { "Service Line", "Sub unit", "Name of Infosys Asset/IP", "Master Customer Code", "Is the revenue currently reported under digital business? Yes/No" };//SL //changed to Service Line
           // DuplicateColumns = new string[] { "Service Line" };

            NumericColumns = new string[] {
                "Q1'23","Q2'23","Q3'23","Q4'23","FY23",
                "Q1'24","Q2'24","Q3'24 (Plan)","Q4'24 (Plan)","FY24 (Plan)",
                "Q1 FY25 (Plan)","Q2 FY25 (Plan)", "Q3 FY25 (Plan)","Q4 FY25 (Plan)",
                //"Q1 FY26 (Plan)","Q2 FY26 (Plan)", "Q3 FY26 (Plan)","Q4 FY26 (Plan)",

                "FY25 (Plan)","FY26 (Plan)","FY27 (Plan)"
                //"FY26 (Plan)","FY27 (Plan)"

            };

            //"FY22 Quarterly delta %" ,"FY23 Quarterly delta %" , "FY24 Quarterly delta %" };

            DictMapping.Add("Service Line", "SL");//changed
            DictMapping.Add("Sub unit", "Sub_unit");
            DictMapping.Add("Name of Infosys Asset/IP", "Name_of_Infosys_Asset/IP");
            DictMapping.Add("Is Asset published in Service Store (Yes/No)", "Is_Asset_published_in_Service_Store");
            DictMapping.Add("Master Customer Code", "Master_Customer_Code");
            DictMapping.Add("Is the revenue currently reported under digital business? Yes/No", "Is_the_revenue_currently_reported_under_digital_business");//added
            DictMapping.Add("Is yes, please mention the Digital Services Name", "Is_yes_please_mention_the_Digital_Services_Name");//added



            DictMapping.Add("Q1'23", "Q1_FY23(Plan)");
            DictMapping.Add("Q2'23", "Q2_FY23(Plan)");
            DictMapping.Add("Q3'23", "Q3_FY23(Plan)");
            DictMapping.Add("Q4'23", "Q4_FY23(Plan)");

            DictMapping.Add("Q1'24", "Q1_FY24(Plan)");
            DictMapping.Add("Q2'24", "Q2_FY24(Plan)");
            DictMapping.Add("Q3'24 (Plan)", "Q3_FY24(Plan)");
            DictMapping.Add("Q4'24 (Plan)", "Q4_FY24(Plan)");


            DictMapping.Add("Q1 FY25 (Plan)", "Q1_FY25(Plan)");
            DictMapping.Add("Q2 FY25 (Plan)", "Q2_FY25(Plan)");
            DictMapping.Add("Q3 FY25 (Plan)", "Q3_FY25(Plan)");
            DictMapping.Add("Q4 FY25 (Plan)", "Q4_FY25(Plan)");


            DictMapping.Add("FY26 (Plan)", "Q4_F26(Plan)");

            DictMapping.Add("FY27 (Plan)", "Q4_F27(Plan)");



        }





    }
}