using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

 
    public class ErrorEntity
    {
        public int RowNo { get; set; }
        public ErrorType ErrType { get; set; }
        public string SheetName { get; set; }
        public string Message { get; set; }
    }

 public   enum ErrorType
    {
        NonNumeric, Mandatory, Duplicate, InvalidColumn, Others, InvalidData,InvalidLookUpData
    }
 