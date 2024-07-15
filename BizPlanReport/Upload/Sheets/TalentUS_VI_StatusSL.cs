using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace BEBIZ.Upload.Sheets
{
    public class TalentUsStatus_SL : BaseSheet
    {
        public TalentUsStatus_SL(DataSet ds) : base(ds)
        {



        }

        protected override void Init()
        {
            RowNumber = 4;
            SheetName = "Talent US VI Status_SL";
            TableName = "BizPlan_Talent_Status_SL";
            MandatoryColumns = new string[] { "Service Line" };
            DuplicateColumns = new string[] { "Service Line", "Parameter" };
            NumericColumns = new string[] { "Q2'20", "Q3'20 ", "Q4'20 ", "Q1'23 ", "Q2'23 ", "Q3'23 ", "Q4'23 ",
            "Q1'21 ","Q2'21 ","Q3'21 ","Q4'21 ","Q1'22 ","Q2'22 ","Q3'22 ","Q4'22 " };

            DictMapping.Add("Service Line", "SL");
            DictMapping.Add("Parameter", "Parameter");
            DictMapping.Add("Q2'20", "FY1_Q2");
            DictMapping.Add("Q3'20 ", "FY1_Q3");
            DictMapping.Add("Q4'20 ", "FY1_Q4");
            DictMapping.Add("Q1'21 ", "FY2_Q1");
            DictMapping.Add("Q2'21 ", "FY2_Q2");
            DictMapping.Add("Q3'21 ", "FY2_Q3");
            DictMapping.Add("Q4'21 ", "FY2_Q4");
            DictMapping.Add("Q1'22 ", "FY3_Q1");
            DictMapping.Add("Q2'22 ", "FY3_Q2");
            DictMapping.Add("Q3'22 ", "FY3_Q3");
            DictMapping.Add("Q4'22 ", "FY3_Q4");
            DictMapping.Add("Q1'23 ", "FY4_Q1");
            DictMapping.Add("Q2'23 ", "FY4_Q2");
            DictMapping.Add("Q3'23 ", "FY4_Q3");
            DictMapping.Add("Q4'23 ", "FY4_Q4");
        }

    }
}