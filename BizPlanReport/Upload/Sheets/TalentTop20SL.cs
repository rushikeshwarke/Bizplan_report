using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace BEBIZ.Upload.Sheets
{
    public class TalentTop20SL : BaseSheet
    {
        public TalentTop20SL(DataSet ds)
            : base(ds)
        {
           


        }

        protected override void Init()
        {
            RowNumber = 4;
            SheetName = "Talent_Top20Accounts (Optional)";
            TableName = "BizPlan_Talent_Skill_Plan";
            MandatoryColumns = new string[] { "Service Line", "Master Customer Code" };//changes SL to Service Line
            DuplicateColumns = new string[] { "Service Line", "Skill", "Master Customer Code" };//changes SL to Service Line
            NumericColumns = new string[] { "Q2'20 EST _HC (as on _30-Sep-19)", "Q3'20 EST HC", "Q4'20 EST HC", "Q1'23 EST HC", "Q2'23 EST HC", "Q3'23 EST HC", "Q4'23 EST HC", "Q1'21 EST HC", "Q2'21 EST HC", "Q3'21 EST HC", "Q4'21 EST HC",
            "Q1'22 EST HC", "Q2'22 EST HC", "Q3'22 EST HC", "Q4'22 EST HC"};

            DictMapping.Add("SL_Top20_Flag", "SL_Top20_Flag");
            DictMapping.Add("Service Line", "SL");
            ////DictMapping.Add("Master Customer Code", "Master_Cust_Code");//added
            DictMapping.Add("Master Customer Code", "MCC");
           
            DictMapping.Add("Skill", "Skill");

            DictMapping.Add("Q2'20 EST _HC (as on _30-Sep-19)", "CurrFYQ2_EST_HC");//added



            //DictMapping.Add("Q1'19 EST HC", "FY1Q1_EST_HC");
            //DictMapping.Add("Q2'19 EST HC", "FY1Q2_EST_HC");
            DictMapping.Add("Q3'20 EST HC", "CurrFYQ3_EST_HC"); 
            DictMapping.Add("Q4'20 EST HC", "CurrFYQ4_EST_HC");
            DictMapping.Add("Q1'21 EST HC", "FY1Q1_EST_HC");
            DictMapping.Add("Q2'21 EST HC", "FY1Q2_EST_HC");
            DictMapping.Add("Q3'21 EST HC", "FY1Q3_EST_HC");
            DictMapping.Add("Q4'21 EST HC", "FY1Q4_EST_HC");
            DictMapping.Add("Q1'22 EST HC", "FY2Q1_EST_HC");
            DictMapping.Add("Q2'22 EST HC", "FY2Q2_EST_HC");
            DictMapping.Add("Q3'22 EST HC", "FY2Q3_EST_HC");
            DictMapping.Add("Q4'22 EST HC", "FY2Q4_EST_HC");

            //mapping changed
            DictMapping.Add("Q1'23 EST HC", "FY3Q1_EST_HC");
            DictMapping.Add("Q2'23 EST HC", "FY3Q2_EST_HC");
            DictMapping.Add("Q3'23 EST HC", "FY3Q3_EST_HC");
            DictMapping.Add("Q4'23 EST HC", "FY3Q4_EST_HC");

            //DictMapping.Add("Q3'18 EST HC (as on 31-Dec-17)", "CurrFYQ3_EST_HC_ason31Dec17");
           // DictMapping.Add("Q4'18 EST HC (as on 31-Mar-18)", "CurrFYQ4_EST_HC_ason31Mar18");
             
        }
        protected override void FormatData()
        {
            base.FormatData();
            Data.Columns.Add("SL_Top20_Flag");
            Data.AsEnumerable().ToList().ForEach(row => row["SL_Top20_Flag"] = "Top20");
            whereCondition = " SL_Top20_Flag = 'Top20' ";

        }
    }
}