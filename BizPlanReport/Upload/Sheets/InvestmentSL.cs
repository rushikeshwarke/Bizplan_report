using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace BEBIZ.Upload.Sheets
{
    public class SLInvestment : BaseSheet
    {
        public SLInvestment(DataSet ds, string sl) : base(ds, sl) { }

        protected override void Init()
        {
            RowNumber = 2;
            SheetName = "Investment_SL";
            TableName = "BizPlan_SL_Acc_Investment";
            if (SL.ToLower() == "sap" || SL.ToLower() == "orc" || SL.ToLower() == "ecas")
            {
                MandatoryColumns = new string[] { "Service Line", "Services Type", "SDM / DH", "Investment Category "  };//SL changed to Service Line
                DuplicateColumns = new string[] { "Service Line", "Services Type", "SDM / DH", "Investment Category ", "Remarks (MUST for investment category \"Others\")" , "Investment sub-category" };//SL changed to Service Line
            }
            else
            {

                MandatoryColumns = new string[] { "Service Line", "Services Type", "Investment Category "  };// for  ecais and eais
                DuplicateColumns = new string[] { "Service Line", "Services Type", "Investment Category ", "Remarks (MUST for investment category \"Others\")" , "Investment sub-category" };// for ecais and eais
            }


            NumericColumns = new string[] {


            //"Q1'24 EST INV", "Q2'24 EST INV", "Q3'24 EST INV", "Q4'24 EST INV", "FY24 EST INV",
            "Q1'25 EST INV", "Q2'25 EST INV", "Q3'25 EST INV", "Q4'25 EST INV", "FY25 EST INV",
            //"Q1'26 EST INV", "Q2'26 EST INV", "Q3'26 EST INV", "Q4'26 EST INV",
                "FY26 EST INV","FY27 EST INV"
            };

               // "FY21 Quarterly delta %", "FY22 Quarterly delta %", "FY23 Quarterly delta %" };
            DictLookUpMapping.Add("Investment Category ", BEBIZ4.Common.EnumLookUp.InvestmentCategory);

            //DictMapping.Add("SL_Acc_Flag", "SL_Acc_Flag");

            DictMapping.Add("Service Line", "SL");//SL 
            DictMapping.Add("Services Type", "Services_Type");
            DictMapping.Add("SDM / DH", "SDM_DH"); 
            DictMapping.Add("Investment Category ", "Investment_Category");
            DictMapping.Add("Investment sub-category", "Investment_sub_category");

            DictMapping.Add("Remarks (MUST for investment category \"Others\")", "Remarks");

            //DictMapping.Add("Q1'22 EST INV", "FY1Q1_EST_INV");
            //DictMapping.Add("Q2'22 EST INV", "FY1Q2_EST_INV");
            //DictMapping.Add("Q3'22 EST INV", "FY1Q3_EST_INV");
            //DictMapping.Add("Q4'22 EST INV", "FY1Q4_EST_INV");

            ////added these new columns and changed the mapping
            //DictMapping.Add("Q1'23 EST INV", "FY1Q1_EST_INV");
            //DictMapping.Add("Q2'23 EST INV", "FY1Q2_EST_INV");
            //DictMapping.Add("Q3'23 EST INV", "FY1Q3_EST_INV");
            //DictMapping.Add("Q4'23 EST INV", "FY1Q4_EST_INV");


            //   DictMapping.Add("Q1'24 EST INV", "FY1Q1_EST_INV");
            //   DictMapping.Add("Q2'24 EST INV", "FY1Q2_EST_INV");
            //   DictMapping.Add("Q3'24 EST INV", "FY1Q3_EST_INV");
            //   DictMapping.Add("Q4'24 EST INV", "FY1Q4_EST_INV");
            ////   DictMapping.Add("FY24 EST INV", "FY2Q4_EST_INV");


            DictMapping.Add("Q1'25 EST INV", "FY1Q1_EST_INV");
            DictMapping.Add("Q2'25 EST INV", "FY1Q2_EST_INV");
            DictMapping.Add("Q3'25 EST INV", "FY1Q3_EST_INV");
            DictMapping.Add("Q4'25 EST INV", "FY1Q4_EST_INV");

            DictMapping.Add("FY25 EST INV", "FY1_EST_INV");


            //DictMapping.Add("Q1'26 EST INV", "FY1Q1_EST_INV");
            //DictMapping.Add("Q2'26 EST INV", "FY1Q2_EST_INV");
            //DictMapping.Add("Q3'26 EST INV", "FY1Q3_EST_INV");
            //DictMapping.Add("Q4'26 EST INV", "FY1Q4_EST_INV");


            DictMapping.Add("FY26 EST INV", "FY2Q4_EST_INV");
            //DictMapping.Add("FY27 EST INV", "FY2Q4_EST_INV");


            DictMapping.Add("FY27 EST INV", "FY3Q4_EST_INV");
            //DictMapping.Add("FY28 EST INV", "FY3Q4_EST_INV");


        }
        protected override void FormatData()
        {
            base.FormatData();
            //Data.Columns.Add("SL_Acc_Flag");
            //Data.AsEnumerable().ToList().ForEach(row => row["SL_Acc_Flag"] = "SL_Invest");
            //whereCondition = "SL_Acc_Flag = 'SL_Invest'";
        }
    }
}