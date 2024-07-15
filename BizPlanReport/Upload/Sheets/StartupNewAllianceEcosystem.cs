using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace BEBIZ.Upload.Sheets
{
    public class StartupNewAllianceEcosystem:BaseSheet
    {
        public StartupNewAllianceEcosystem(DataSet ds, string sl)
           : base(ds, sl)
        {

        }


        protected override void Init()
        {
            RowNumber = 2;
            SheetName = "Startup&New Alliance Ecosystem";
            TableName = "BizPlan_Startup_New_Alliance_Ecosystem";
            
            MandatoryColumns = new string[] { "Service Line", "Sub unit", "Name of the Startup/Alliance", "Master Customer Code", "Service Area" };//SL changed to Service Line
            DuplicateColumns = new string[] { "Service Line", "Sub unit", "Name of the Startup/Alliance", "Master Customer Code", "Service Area", "TCV of Won Deal (KUSD)" };//SL //changed to Service Line
            

            NumericColumns = new string[] {
                "Q1'23","Q2'23","Q3'23","Q4'23","FY23",
                "Q1'24","Q2'24","Q3'24 (Plan)","Q4'24 (Plan)","FY24 (Plan)",
                "Q1 FY25 (Plan)","Q2 FY25 (Plan)","Q3 FY25 (Plan)","Q4 FY25 (Plan)","FY25 (Plan)",
                "FY26 (Plan)","FY27 (Plan)"

            };

            //"FY22 Quarterly delta %" ,"FY23 Quarterly delta %" , "FY24 Quarterly delta %" };

            DictMapping.Add("Service Line", "SL");//changed
            DictMapping.Add("Sub unit", "Sub_unit");
            DictMapping.Add("Name of the Startup/Alliance", "Name_of_the_Startup/Alliance");
            DictMapping.Add("Master Customer Code", "Master_Customer_Code");
            DictMapping.Add("Service Area", "Service_Area");
            //DictMapping.Add("Any Deal Closed (Won/Open)", "Any_Deal_Closed");//added
            DictMapping.Add("TCV of Won Deal (KUSD)", "TCV_of_the_Deal_KUSD");//added
            
        

            //DictMapping.Add("Q1'21", "Q1'21");
            //DictMapping.Add("Q2'21", "Q2'21");
            //DictMapping.Add("Q3'21", "Q3'21");
            //DictMapping.Add("Q4'21", "Q4'21");
          

            //DictMapping.Add("Q1'22", "Q1'22");
            //DictMapping.Add("Q2'22", "Q2'22");
            //DictMapping.Add("Q3'22", "Q3'22");
            //DictMapping.Add("Q4'22", "Q4'22");

           // DictMapping.Add("FY22", "Q4'22");



            DictMapping.Add("Q1'23", "Q1_FY23(Plan)");
            DictMapping.Add("Q2'23", "Q2_FY23(Plan)");
            DictMapping.Add("Q3'23", "Q3_FY23(Plan)");
            DictMapping.Add("Q4'23", "Q4_FY23(Plan)");

            DictMapping.Add("FY23", "FY23");

            DictMapping.Add("Q1'24", "Q1_FY24(Plan)");
            DictMapping.Add("Q2'24", "Q2_FY24(Plan)");
            DictMapping.Add("Q3'24 (Plan)", "Q3_FY24(Plan)");
            DictMapping.Add("Q4'24 (Plan)", "Q4_FY24(Plan)");

            DictMapping.Add("FY24 (Plan)", "FY24 (Plan)");

            DictMapping.Add("Q1 FY25 (Plan)", "Q1_FY25(Plan)");
            DictMapping.Add("Q2 FY25 (Plan)", "Q2_FY25(Plan)");
            DictMapping.Add("Q3 FY25 (Plan)", "Q3_FY25(Plan)");
            DictMapping.Add("Q4 FY25 (Plan)", "Q4_FY25(Plan)");

            DictMapping.Add("FY25 (Plan)", "FY25 (Plan)");

            DictMapping.Add("FY26 (Plan)", "FY26 (Plan)");  
            DictMapping.Add("FY27 (Plan)", "FY27 (Plan)");

        }



    }
}