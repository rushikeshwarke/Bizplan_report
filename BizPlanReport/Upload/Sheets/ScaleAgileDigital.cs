using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace BEBIZ.Upload.Sheets
{
    public class ScaleAgileDigital : BaseSheet
    {
        public ScaleAgileDigital(DataSet ds) : base(ds)
        {



        }

        protected override void Init()
        {
            RowNumber = 5;
            SheetName = "Scale Agile Digital";
            TableName = "BizPlan_Scale_Agile";
            MandatoryColumns = new string[] { "Service Line" };
            DuplicateColumns = new string[] { "Service Line", "Parameters" };
            NumericColumns = new string[] { "Target", "Actuals", "Target1", "Target2", "Target3" };

            DictMapping.Add("Service Line", "SL");
            DictMapping.Add("Parameters", "Parameters");
            DictMapping.Add("Target", "H1FY1_Target");
            DictMapping.Add("Actuals", "H1FY1_Actuals");
            DictMapping.Add("Target1", "H2FY1_Target");
            DictMapping.Add("Target2", "H1FY2_Target");
            DictMapping.Add("Target3", "H2FY2_Target");

        }


    }
}