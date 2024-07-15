using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace BEBIZ.Upload.Sheets
{
    public class RandD : BaseSheet
    {
        public RandD(DataSet ds, string sl) : base(ds, sl) { }

        protected override void Init()
        {
            RowNumber = 2;
            SheetName = "Investment_Solution_R&D";
            TableName = "BizPlan_Solution_RD_Investment";
            MandatoryColumns = new string[] { "Service Line" };
            DuplicateColumns = new string[] { "Service Line", "Dimension", "Key Solution/Platform", "Focus Area", "Focus Sub-Area _(Optional)" };
            NumericColumns = new string[] { "Effort Investment (Person Months)", "# of clients deployed", "Monetization (MUSD)" ,
            "Effort Investment (Person Months)", "# of clients deployed", "Monetization (MUSD)" ,
            "Effort Investment (Person Months)1", "# of clients deployed1", "Monetization (MUSD)1" ,
            "Effort Investment (Person Months)2", "# of clients deployed2", "Monetization (MUSD)2" 
            //"Effort Investment (Person Months)3", "# of clients deployed3", "Monetization (MUSD)3"
            };

            DictMapping.Add("Service Line", "SL");
            DictMapping.Add("Dimension", "Dimension");
            DictMapping.Add("Key Solution/Platform", "Solution_Platform");
            DictMapping.Add("Focus Area", "Focus_Area");
            DictMapping.Add("Focus Sub-Area _(Optional)", "Focus_SubArea");
            DictMapping.Add("Owner Mail ID", "OwnerMailID");
            DictMapping.Add("Sensed or Incubated?", "Sensed_Incubated");
          

            DictMapping.Add("Effort Investment (Person Months)", "FY1_Eff_Investment");
            DictMapping.Add("# of clients deployed", "FY1_Clients_Deployed");
            DictMapping.Add("Monetization (MUSD)", "FY1_Monetization");

            DictMapping.Add("Effort Investment (Person Months)1", "FY2_Eff_Investment");
            DictMapping.Add("# of clients deployed1", "FY2_Clients_Deployed");
            DictMapping.Add("Monetization (MUSD)1", "FY2_Monetization");

            DictMapping.Add("Effort Investment (Person Months)2", "FY3_Eff_Investment");
            DictMapping.Add("# of clients deployed2", "FY3_Clients_Deployed");
            DictMapping.Add("Monetization (MUSD)2", "FY3_Monetization");



        }

    }
}