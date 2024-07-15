using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace BEBIZ.Upload.Sheets
{
    public class SLVolume : BaseSheet
    {

        public SLVolume(DataSet ds, string sl) : base(ds, sl) { }

        protected override void Init()
        {
            RowNumber = 2;//what should be the row number?
            SheetName = "SL_Vol";
            TableName = "BizPlan_SL_Volume";
            MandatoryColumns = new string[] { "Service Line", "Services Type" };//SL changed to Service Line
            DuplicateColumns = new string[] { "Service Line", "Services Type" };//SL changed to Service Line
            NumericColumns = new string[] {
                "FY24 TD EST RPP (USD)",
                //"FY25 TD EST RPP (USD)",

            //"Q1'24 EST RPP","Q2'24 EST RPP","Q3'24 EST RPP","Q4'24 EST RPP",
            "Q1'25 EST RPP","Q2'25 EST RPP","Q3'25 EST RPP","Q4'25 EST RPP",
            //"Q1'26 EST RPP","Q2'26 EST RPP","Q3'26 EST RPP","Q4'26 EST RPP",

               "FY'25 EST RPP", "FY'26 EST RPP", "FY'27 EST RPP"
                //"FY'27 EST RPP", "FY'28 EST RPP"
            };//changed

            DictMapping.Add("Service Line", "SL"); //changed
            DictMapping.Add("Services Type", "Services_Type");


            //DictMapping.Add("Q1'22 EST RPP (USD)", "FY1Q1_EST_RPP_USD");                      //column mapping changed
            //DictMapping.Add("Q2'22 EST RPP (USD)", "FY1Q2_EST_RPP_USD");                      //column mapping changed
            //DictMapping.Add("Q3'22 EST RPP (USD)", "FY1Q3_EST_RPP_USD");                      //column mapping changed
            //DictMapping.Add("Q4'22 EST RPP (USD)", "FY1Q4_EST_RPP_USD");                      //column mapping changed
            //DictMapping.Add("Remarks for RPP", "FY1_Notes_for_RPP_changes");            //column mapping changed

            //DictMapping.Add("Q1'23 EST RPP", "FY1Q1_EST_RPP_USD");                      //column mapping changed
            //DictMapping.Add("Q2'23 EST RPP", "FY1Q2_EST_RPP_USD");                      //column mapping changed
            //DictMapping.Add("Q3'23 EST RPP", "FY1Q3_EST_RPP_USD");                      //column mapping changed
            //DictMapping.Add("Q4'23 EST RPP", "FY1Q4_EST_RPP_USD");                      //column mapping changed
            //DictMapping.Add("Remarks for RPP", "FY1_Notes_for_RPP_changes");            //column mapping changed


            //DictMapping.Add("Q1'24 EST RPP", "FY1Q1_EST_RPP_USD");                      //column mapping changed
            //DictMapping.Add("Q2'24 EST RPP", "FY1Q2_EST_RPP_USD");                      //column mapping changed
            //DictMapping.Add("Q3'24 EST RPP", "FY1Q3_EST_RPP_USD");                      //column mapping changed
            //DictMapping.Add("Q4'24 EST RPP", "FY1Q4_EST_RPP_USD");                      //column mapping changed
            //DictMapping.Add("Remarks for RPP", "FY1_Notes_for_RPP_changes");            //column mapping changed

            //  DictMapping.Add("FY'24 EST RPP", "FY2Q4_EST_RPP_USD");


            DictMapping.Add("FY24 TD EST RPP (USD)", "FY1Q4_EST_RPP_USD");
            //DictMapping.Add("FY25 TD EST RPP (USD)", "FY1Q4_EST_RPP_USD");


            DictMapping.Add("Q1'25 EST RPP", "FY2Q1_EST_RPP_USD");
            DictMapping.Add("Q2'25 EST RPP", "FY2Q2_EST_RPP_USD");
            DictMapping.Add("Q3'25 EST RPP", "FY2Q3_EST_RPP_USD");
            DictMapping.Add("Q4'25 EST RPP", "FY2Q4_EST_RPP_USD");
            DictMapping.Add("Remarks for RPP", "FY2_Notes_for_RPP_changes");


            //DictMapping.Add("Q1'26 EST RPP", "FY2Q1_EST_RPP_USD");
            //DictMapping.Add("Q2'26 EST RPP", "FY2Q2_EST_RPP_USD");
            //DictMapping.Add("Q3'26 EST RPP", "FY2Q3_EST_RPP_USD");
            //DictMapping.Add("Q4'26 EST RPP", "FY2Q4_EST_RPP_USD");
            //DictMapping.Add("Remarks for RPP", "FY2_Notes_for_RPP_changes");


            //            DictMapping.Add("Q1'25 EST RPP", "FY3Q1_EST_RPP_USD");
            //            DictMapping.Add("Q2'25 EST RPP", "FY3Q2_EST_RPP_USD");
            //            DictMapping.Add("Q3'25 EST RPP", "FY3Q3_EST_RPP_USD");
            //            DictMapping.Add("Q4'25 EST RPP", "FY3Q4_EST_RPP_USD");
            //            DictMapping.Add("Remarks for RPP3", "FY3_Notes_for_RPP_changes");

            DictMapping.Add("FY'26 EST RPP", "FY3Q4_EST_RPP_USD");
            //DictMapping.Add("FY'27 EST RPP", "FY3Q4_EST_RPP_USD");


            //            DictMapping.Add("Q1'26 EST RPP", "FY4Q1_EST_RPP_USD");
            //            DictMapping.Add("Q2'26 EST RPP", "FY4Q2_EST_RPP_USD");
            //            DictMapping.Add("Q3'26 EST RPP", "FY4Q3_EST_RPP_USD");
            //            DictMapping.Add("Q4'26 EST RPP", "FY4Q4_EST_RPP_USD");
            //            DictMapping.Add("Remarks for RPP4", "FY4_Notes_for_RPP_changes");

            DictMapping.Add("FY'27 EST RPP", "FY4Q4_EST_RPP_USD");
            //DictMapping.Add("FY'28 EST RPP", "FY4Q4_EST_RPP_USD");


        }


    }
}