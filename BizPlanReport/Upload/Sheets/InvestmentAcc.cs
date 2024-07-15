using BEBIZ4.Common;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace BEBIZ.Upload.Sheets
{
    public class AccInvestment : BaseSheet
    {
        public AccInvestment(DataSet ds)
            : base(ds)
        {
          

        }
        protected override void Init()
        {
            RowNumber = 3;
            SheetName = "Investment_Acc(Optional)";
            TableName = "BizPlan_SL_Acc_Investment";
            MandatoryColumns = new string[] { "Service Line", "Services Type", "Master Customer Code", "SDM / DH", "Investment Category " };  //SL changed to Service Line
            DuplicateColumns = new string[] { "Service Line", "Services Type", "Master Customer Code", "SDM / DH", "Investment Category " };  //SL changed to Service Line
            NumericColumns = new string[] { "Q1'22 EST INV", "Q2'22 EST INV", "Q3'22 EST INV", "Q4'22 EST INV", "Q1'23 EST INV", "Q2'23 EST INV", "Q3'23 EST INV", "Q4'23 EST INV", "Q1'21 EST INV", "Q2'21 EST INV", "Q3'21 EST INV", "Q4'21 EST INV", "FY22 Quarterly delta %", "FY23 Quarterly delta %", "FY21 Quarterly delta %" };

            DictLookUpMapping.Add("Investment Category ", EnumLookUp.InvestmentCategory);

            DictMapping.Add("SL_Acc_Flag", "SL_Acc_Flag");
            DictMapping.Add("Service Line", "SL");
            DictMapping.Add("Services Type", "Services_Type");
            DictMapping.Add("Master Customer Code", "MCC");
            //DictMapping.Add("Master Customer Name", "Master_Cust_Name");//added
            DictMapping.Add("SDM / DH", "SDM_DH");
            
            DictMapping.Add("Investment Category ", "Investment_Category");

            //DictMapping.Add("Q1'19 EST INV", "FY1Q1_EST_INV");
            //DictMapping.Add("Q2'19 EST INV", "FY1Q2_EST_INV");
            //DictMapping.Add("Q3'19 EST INV", "FY1Q3_EST_INV");
            //DictMapping.Add("Q4'19 EST INV", "FY1Q4_EST_INV");
            //DictMapping.Add("Q1'20 EST INV", "FY1Q1_EST_INV");
            //DictMapping.Add("Q2'20 EST INV", "FY1Q2_EST_INV");
            //DictMapping.Add("Q3'20 EST INV", "FY1Q3_EST_INV");
            //DictMapping.Add("Q4'20 EST INV", "FY1Q4_EST_INV");


            DictMapping.Add("Q1'21 EST INV", "FY1Q1_EST_INV");
            DictMapping.Add("Q2'21 EST INV", "FY1Q2_EST_INV");
            DictMapping.Add("Q3'21 EST INV", "FY1Q3_EST_INV");
            DictMapping.Add("Q4'21 EST INV", "FY1Q4_EST_INV");

            //columns added and changed mapping

            DictMapping.Add("Q1'22 EST INV", "FY2Q1_EST_INV");
            DictMapping.Add("Q2'22 EST INV", "FY2Q2_EST_INV");
            DictMapping.Add("Q3'22 EST INV", "FY2Q3_EST_INV");
            DictMapping.Add("Q4'22 EST INV", "FY2Q4_EST_INV");

            DictMapping.Add("Q1'23 EST INV", "FY3Q1_EST_INV");
            DictMapping.Add("Q2'23 EST INV", "FY3Q2_EST_INV");
            DictMapping.Add("Q3'23 EST INV", "FY3Q3_EST_INV");
            DictMapping.Add("Q4'23 EST INV", "FY3Q4_EST_INV");



        
            DictMapping.Add("FY21 Quarterly delta %", "FY1_Quarterly_Delta_Perc");
            DictMapping.Add("FY22 Quarterly delta %", "FY2_Quarterly_Delta_Perc");
            DictMapping.Add("FY23 Quarterly delta %", "FY3_Quarterly_Delta_Perc");
        }
        protected override void FormatData()
        {
            base.FormatData();
            Data.Columns.Add("SL_Acc_Flag");
            Data.AsEnumerable().ToList().ForEach(row => row["SL_Acc_Flag"] = "Acc_Invest");
            whereCondition = "SL_Acc_Flag = 'Acc_Invest'";
        }
    }
}