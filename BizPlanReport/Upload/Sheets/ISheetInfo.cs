using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BEBIZ.Upload.Sheets
{
    public interface ISheetInfo
    {
        string SheetName { get; set; }
        string SL { get; set; }
        List<ErrorEntity> Validate();
        void Save();
        


    }
}
