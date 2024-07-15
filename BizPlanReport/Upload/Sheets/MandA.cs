using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace BEBIZ.Upload.Sheets
{
    public class MandA : BaseSheet
    {
        public MandA(DataSet ds)
            : base(ds)
        { }

        protected override void Init()
        {
            RowNumber = 2;
            SheetName = "M&A";
            TableName = "BizPlan_MandA";
            MandatoryColumns = new string[] { "SL", "Theme" };
            DuplicateColumns = new string[] { "SL", "Theme" };
            NumericColumns = new string[] { "Likely_IT_Spend"};

            DictMapping.Add("SL", "SL");
            DictMapping.Add("Theme", "Theme");
            DictMapping.Add("Likely IT spend MUSD", "Likely_IT_Spend");
            DictMapping.Add("Current internal situation/What stops us from achieving this int", "Pain_Points");
            DictMapping.Add("Why M&A?", "Why_MA");
            DictMapping.Add("Key attributes desired from the target", "Key_Attritute");
            DictMapping.Add("Expected Closure", "Expected_Closure"); 
        }
    }


}