using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace BEBIZ.Upload.Sheets
{
    public class InOrganicSLRev : BaseSheet
    {
        public InOrganicSLRev(DataSet ds, string sl ) : base(ds,sl)
        {



        }

        protected override void Init()
        {
            RowNumber = 2;//what should be the row number?
            SheetName = "In-Organic_Rev (M&A)";//changed
            TableName = "BizPlan_InOrganic_SL_Revenue";
            MandatoryColumns = new string[] { "Service Line" };//changed. And no Services Type present.
            //DuplicateColumns = new string[] { "Service Line", "Dimension", "Focus Area" };//changed. And no Services Type present.
            DuplicateColumns = new string[] { "Service Line", "Focus Area" };//changed. And no Services Type present.
            NumericColumns = new string[] {
            
             "Q1'24 EST REV in KUSD", "Q2'24 EST REV in KUSD", "Q3'24 EST REV in KUSD", "Q4'24 EST REV in KUSD",
              
            //"Q1'25 EST REV in KUSD", "Q2'25 EST REV in KUSD", "Q3'25 EST REV in KUSD", "Q4'25 EST REV in KUSD",
            "FY25 EST REV in KUSD",
                "FY26 EST REV in KUSD"
            };



            DictMapping.Add("Service Line", "SL");//changed 


            DictMapping.Add("Master Customer Code", "MCC");
            DictMapping.Add("Vertical", "Vertical");
            DictMapping.Add("Region", "Region");



            // DictMapping.Add("Dimension", "Dimension");//added
            DictMapping.Add("Focus Area", "Focus_Area");//added
            DictMapping.Add("Focus Sub-Area _(Optional)", "Focus_Sub_Area");//added
            DictMapping.Add("M&A Pipeline _(Name of the possible candidate)", "M&A_Pipeline");//added
            DictMapping.Add("Shortlisted? (Y/N)", "Shortlisted");//added
            DictMapping.Add("Remarks", "Remarks");


         


            DictMapping.Add("Q1'24 EST REV in KUSD", "FY1Q1_EST_REV_KUSD");//added
            DictMapping.Add("Q2'24 EST REV in KUSD", "FY1Q2_EST_REV_KUSD");//added
            DictMapping.Add("Q3'24 EST REV in KUSD", "FY1Q3_EST_REV_KUSD");//added
            DictMapping.Add("Q4'24 EST REV in KUSD", "FY1Q4_EST_REV_KUSD");//added
          // DictMapping.Add("FY24 REV in KUSD", "FY2Q4_EST_REV_KUSD");
            //                                                               //COLUMN MAPPING CHANGED

            //DictMapping.Add("Q1'25 EST REV in KUSD", "FY3Q1_EST_REV_KUSD");
            //DictMapping.Add("Q2'25 EST REV in KUSD", "FY3Q2_EST_REV_KUSD");
            //DictMapping.Add("Q3'25 EST REV in KUSD", "FY3Q3_EST_REV_KUSD");
            //DictMapping.Add("Q4'25 EST REV in KUSD", "FY3Q4_EST_REV_KUSD");
            DictMapping.Add("FY25 EST REV in KUSD", "FY2Q4_EST_REV_KUSD");

            //DictMapping.Add("Q1'26 EST REV in KUSD", "FY3Q1_EST_REV_KUSD");
            //DictMapping.Add("Q2'26 EST REV in KUSD", "FY3Q2_EST_REV_KUSD");
            //DictMapping.Add("Q3'26 EST REV in KUSD", "FY3Q3_EST_REV_KUSD");
            //DictMapping.Add("Q4'26 EST REV in KUSD", "FY3Q4_EST_REV_KUSD");
            DictMapping.Add("FY26 EST REV in KUSD", "FY3Q4_EST_REV_KUSD");

        }
    }


}