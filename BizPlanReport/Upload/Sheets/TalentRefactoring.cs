using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace BEBIZ.Upload.Sheets
{
    public class TalentRefactoring : BaseSheet
    {
        public TalentRefactoring(DataSet ds) : base(ds)
        {
           

        }
        protected override void Init()
        {
            RowNumber = 2;
            
            SheetName = "Talent_Refactoring_SL";//changed the sheet name
            TableName = "BizPlan_Talent_Refactoring";
            MandatoryColumns = new string[] { "Service Line",  };//changes SL to Service Line and service type to be added or not?
            DuplicateColumns = new string[] { "Service Line", "Service Type" };//changes SL to Service Line and service type to be added or not?
            NumericColumns = new string[] {  "FY18 Opening HC","Talent Refactoring FY18","External Hire FY18","FY18 Closing HC",
                "Talent Refactoring FY19","External Hire FY19","FY19 Closing HC","Talent Refactoring FY20","External Hire FY20",
                "FY20 Closing HC","Talent Refactoring FY21","External Hire FY21","FY21 Closing HC","Talent Refactoring FY22",
                "External _Hire _FY22","FY22 Closing _HC",
            };

            DictMapping.Add("Service Line", "SL"); //changes SL to Service Line
            ////DictMapping.Add("New Service Offering", "NSO");
            DictMapping.Add("Service Type", "Service_Type");// added



            DictMapping.Add("FY18 Opening HC", "PrevFY_Opening_HC");
            DictMapping.Add("Talent Refactoring FY18", "PrevFY_Talent_Refactoring");
            DictMapping.Add("External Hire FY18", "PrevFY_External_Hire");
            DictMapping.Add("FY18 Closing HC", "PrevFY_Closing_HC");

            //DictMapping.Add("FY19 Opening HC", "CurrFY_Opening_HC");   
            DictMapping.Add("Talent Refactoring FY19", "CurrFY_Talent_Refactoring");
            DictMapping.Add("External Hire FY19", "CurrFY_External_Hire");
            DictMapping.Add("FY19 Closing HC", "CurrFY_Closing_HC");
            DictMapping.Add("Talent Refactoring FY20", "FY1_Talent_Refactoring");
            DictMapping.Add("External Hire FY20", "FY1_External_Hire");
            DictMapping.Add("FY20 Closing HC", "FY1_Closing_HC");
            DictMapping.Add("Talent Refactoring FY21", "FY2_Talent_Refactoring");
            DictMapping.Add("External Hire FY21", "FY2_External_Hire");
            DictMapping.Add("FY21 Closing HC", "FY2_Closing_HC");
            DictMapping.Add("Talent Refactoring FY22", "FY3_Talent_Refactoring");
            DictMapping.Add("External _Hire _FY22", "FY3_External_Hire");
            DictMapping.Add("FY22 Closing _HC", "FY3_Closing_HC");

            //added new columns to be added in DB
            DictMapping.Add("Talent Refactoring FY23", "FY4_Talent_Refactoring");
            DictMapping.Add("External _Hire _FY23", "FY4_External_Hire");
            DictMapping.Add("FY23 Closing _HC", "FY4_Closing_HC");
        }
    }
}