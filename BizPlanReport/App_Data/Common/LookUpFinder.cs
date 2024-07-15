using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace BEBIZ4.Common
{
    public class LookUpFinder
    {
        public static List<string> GetLookUpItems(EnumLookUp key) {
            Dictionary<EnumLookUp, string> dictMapping = new Dictionary<EnumLookUp, string>();
            dictMapping.Add(EnumLookUp.InvestmentCategory, "InvestmentCategoryAll.txt");
            dictMapping.Add(EnumLookUp.InvestmentCategory_PracticeSales, "InvestmentCategory_Practice_Sales.txt");
            string fileName = dictMapping[key];
            string path = HttpContext.Current.Server.MapPath("~/Resources/LookupData");
            path = Path.Combine(path, fileName);
            var listItems = File.ReadAllLines(path).ToList();
            listItems = listItems.Select(k => k.Trim()).ToList();
            return listItems; 
        }
    }

    public enum EnumLookUp
    {
        InvestmentCategory_PracticeSales,
        InvestmentCategory
    }
}