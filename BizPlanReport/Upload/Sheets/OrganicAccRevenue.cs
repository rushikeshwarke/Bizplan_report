using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace BEBIZ.Upload.Sheets
{
    public class OrganicAccRevenue : BaseSheet
    {
        public OrganicAccRevenue(DataSet ds, string sl)
            : base(ds, sl)
        {

        }


        protected override void Init()
        {
            RowNumber = 2;
            SheetName = "Acc_Rev";
            TableName = "BizPlan_Organic_Acc_Revenue";
            //if (SL.ToLower() == "orc" || SL.ToLower() == "sap" || SL.ToLower() == "ecas")
            //{
            MandatoryColumns = new string[] { "Service Line", "Services Type", "Master Customer Code", "Native Currency", "SDM / DH", "Cloud Infra Pass-through/SaaS Passthrough/Sevices" };//SL changed to Service Line
            DuplicateColumns = new string[] { "Service Line", "Services Type", "Master Customer Code", "Native Currency", "SDM / DH", "RegionGroupLatest", "Alliance Partner", "For Others, pls specify", "Cloud Infra Pass-through/SaaS Passthrough/Sevices" };//SL //changed to Service Line
            //}
            //else
            //{
            //    // ECAS and EAIS
            //    MandatoryColumns = new string[] { "Service Line", "Services Type", "Master Customer Code", "Native Currency" };// for ecas and eais
            //    DuplicateColumns = new string[] { "Service Line", "Services Type", "Master Customer Code", "Native Currency", "RegionGroupLatest", "SDM / DH", "For Others, pls specify", "Remarks (Only for \"Emerging Services\" )" , "Cloud Infra Pass-through/SaaS Passthrough/Sevices" }; //// for ecas and eais
            //}

            NumericColumns = new string[] {
                //"Q1'17 ACT REV in KNC","Q2'17 ACT REV in KNC","Q3'17 ACT REV in KNC","Q4'17 ACT REV in KNC" ,
                //"Q1'18 ACT REV in KNC", "Q2'18 ACT REV in KNC","Q3'18 ACT REV in KNC","Q4'18 ACT REV in KNC" ,
               // "Q1'19 ACT REV in KNC", "Q2'19 ACT REV in KNC", "Q3'19 ACT REV in KNC", "Q4'19 ACT REV in KNC",
              //  "Q1'20 ACT REV in KNC", "Q2'20 ACT REV in KNC", "Q3'20 ACT REV in KNC", "Q4'20 ACT REV in KNC","FY20 ACT REV in KNC",
                //"Q1'21 ACT REV in KNC", "Q2'21 ACT REV in KNC", "Q3'21 ACT REV in KNC", "Q4'21 ACT REV in KNC","FY21 ACT REV in KNC",
                //"Q1'22 ACT REV in KNC", "Q2'22 ACT REV in KNC", "Q3'22 EST REV in KNC", "Q4'22 EST REV in KNC","FY22 ACT REV in KNC",
                //"Q1'23 EST REV in KNC","Q2'23 EST REV in KNC",
                //"Q3'23 EST REV in KNC","Q4'23 EST REV in KNC",
                //"Q1'24 EST REV in KNC","Q2'24 EST REV in KNC",
                "Q3'24 EST REV in KNC","Q4'24 EST REV in KNC","FY24 EST REV in KNC", 
                "Q1'25 EST REV in KNC", "Q2'25 EST REV in KNC", "Q3'25 EST REV in KNC","Q4'25 EST REV in KNC",
                "FY25 EST REV in KNC","FY26 EST REV in KNC", "FY27 EST REV in KNC"
            };

            DictLookUpMapping.Add("Cloud Infra Pass-through/SaaS Passthrough/Sevices", BEBIZ4.Common.EnumLookUp.PassThrough);


            //"FY22 Quarterly delta %" ,"FY23 Quarterly delta %" , "FY24 Quarterly delta %" };

            DictMapping.Add("Service Line", "SL");//changed
            DictMapping.Add("Services Type", "Services_Type");
            DictMapping.Add("Master Customer Code", "MCC");
            //DictMapping.Add("Master Customer Name", "Master_Cust_Name");//added
            DictMapping.Add("Native Currency", "NC");
            DictMapping.Add("SDM / DH", "SDM_DH");
            DictMapping.Add("Vertical", "Vertical");//added
            DictMapping.Add("RegionGroupLatest", "RegionGroup");//added
            //DictMapping.Add("Remarks (Only for \"Emerging Services\" )", "Remarks");            
            DictMapping.Add("Alliance Partner", "Hyperscalers_Alliances");
            DictMapping.Add("For Others, pls specify", "For_Others_pls_specify");
            DictMapping.Add("Cloud Infra Pass-through/SaaS Passthrough/Sevices", "Cloud_Consumption_Passthrough");



            //DictMapping.Add("Q1'21 ACT REV in KNC", "FY3Q1_EST_REV_KNC");
            //DictMapping.Add("Q2'21 ACT REV in KNC", "FY3Q2_EST_REV_KNC");
            //DictMapping.Add("Q3'21 ACT REV in KNC", "FY3Q3_EST_REV_KNC");
            //DictMapping.Add("Q4'21 ACT REV in KNC", "FY3Q4_EST_REV_KNC");
            //DictMapping.Add("FY21 ACT REV in KNC", "FY22 ACT REV in KNC");

            //DictMapping.Add("Q1'22 ACT REV in KNC", "FY4Q1_EST_REV_KNC");
            //DictMapping.Add("Q2'22 ACT REV in KNC", "FY4Q2_EST_REV_KNC");
            //DictMapping.Add("Q3'22 EST REV in KNC", "FY4Q3_EST_REV_KNC");
            //DictMapping.Add("Q4'22 EST REV in KNC", "FY4Q4_EST_REV_KNC");
            //DictMapping.Add("FY22 ACT REV in KNC", "FY22 ACT REV in KNC");

            //DictMapping.Add("Q1'23 EST REV in KNC", "FY5Q1_EST_REV_KNC");
            //DictMapping.Add("Q2'23 EST REV in KNC", "FY5Q2_EST_REV_KNC");
            //DictMapping.Add("Q3'23 EST REV in KNC", "FY5Q3_EST_REV_KNC");
            //DictMapping.Add("Q4'23 EST REV in KNC", "FY5Q4_EST_REV_KNC");
            //DictMapping.Add("FY23 ACT REV in KNC", "FY23 ACT REV in KNC");

            //DictMapping.Add("Q1'24 EST REV in KNC", "FY6Q1_EST_REV_KNC");
            //DictMapping.Add("Q2'24 EST REV in KNC", "FY6Q2_EST_REV_KNC");
            DictMapping.Add("Q3'24 EST REV in KNC", "FY5Q3_EST_REV_KNC");
            DictMapping.Add("Q4'24 EST REV in KNC", "FY5Q4_EST_REV_KNC");
            DictMapping.Add("FY24 EST REV in KNC", "FY5_EST_REV_KNC");

            DictMapping.Add("Q1'25 EST REV in KNC", "FY6Q1_EST_REV_KNC");
            DictMapping.Add("Q2'25 EST REV in KNC", "FY6Q2_EST_REV_KNC");
            DictMapping.Add("Q3'25 EST REV in KNC", "FY6Q3_EST_REV_KNC");
            DictMapping.Add("Q4'25 EST REV in KNC", "FY6Q4_EST_REV_KNC");
            DictMapping.Add("FY25 EST REV in KNC", "FY6_EST_REV_KNC");

            // DictMapping.Add("Q1'26 EST REV in KNC", "FY8Q1_EST_REV_KNC");
            // DictMapping.Add("Q2'26 EST REV in KNC", "FY8Q2_EST_REV_KNC");
            // DictMapping.Add("Q3'26 EST REV in KNC", "FY8Q3_EST_REV_KNC");
            // DictMapping.Add("Q4'26 EST REV in KNC", "FY8Q4_EST_REV_KNC");
            DictMapping.Add("FY26 EST REV in KNC", "FY7Q4_EST_REV_KNC");
            DictMapping.Add("FY27 EST REV in KNC", "FY8Q4_EST_REV_KNC");



            //DictMapping.Add("FY22 Quarterly delta %", "FY4_Qtr_Delta_Perc");
            //DictMapping.Add("FY23 Quarterly delta %", "FY5_Qtr_Delta_Perc");
            //DictMapping.Add("FY23 Quarterly delta %", "FY6_Qtr_Delta_Perc");

        }

        public override List<string> GetMCC()
        {
            List<string> lst = base.GetMCC();
            string[] itmesToIgnore = new string[] { "TBD" };
            lst = lst.Where(k => !itmesToIgnore.Contains(k)).ToList();
            return lst;
        }


        protected override void ReplaceNullWithZeroInDT(DataTable dt)
        {
            base.ReplaceNullWithZeroInDT(dt);

            //for (int i = 0; i < dt.Rows.Count; i++)
            //{
            //    DataRow row = dt.Rows[i];

            //        if (row["Cloud Infra Pass-through/SaaS Passthrough/Sevices"].ToString() == "Passthrough-TBD")
            //        {
            //        row["Cloud Infra Pass-through/SaaS Passthrough/Sevices"] = "";
            //        }

            //}
            //dt.AcceptChanges();

        }
    }
}