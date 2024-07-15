using BEBIZ.Upload.Sheets;
using ClosedXML.Excel;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Excel = Microsoft.Office.Interop.Excel;
using VBIDE = Microsoft.Vbe.Interop;


namespace BEBIZ.Upload
{
    public class SummaryDeckTableNames
    {
        public static string SummaryDeck2Actuals = "SummaryDeck2Actuals";
        public static string SummaryDeck2Region = "SummaryDeck2Region";
        public static string SummaryDeck2Vertical = "SummaryDeck2Vertical";
        public static string SummaryDeck2VerticalGrouping = "SummaryDeck2VerticalGrouping";
        public static string SummaryDeck2RevGrowth = "SummaryDeck2RevGrowth";
        public static string SummaryDeck2DigRevGrowth = "SummaryDeck2DigRevGrowth";
        public static string SummaryDeck2DigRev = "SummaryDeck2DigRev";
    }



    public partial class Upload : Page
    {
        public void FillExcelSheet_SummaryDeck2(DataTable dtQtrly, DataTable dtClientCat, DataTable dtDigital, Microsoft.Office.Interop.Excel.Worksheet sheet)
        {
            object[,] rawData = null;



            try
            {
                Microsoft.Office.Interop.Excel.Range c1;
                Microsoft.Office.Interop.Excel.Range c2;
                Microsoft.Office.Interop.Excel.Range range_excel;

                if (1 == 1)
                {
                    string[] dtQtrly_columnsToDelete = { "QTR_17_Q1", "QTR_17_Q2", "QTR_17_Q3", "QTR_17_Q4", "FY_17", "FY_18", "FY_19", "FY_20", "FY_21", "FY_22" , "FY_23"  , "Trend"};
                    string[] dtDigital_columnsToDelete = { "FY_18", "FY_19", "FY_20", "FY_21", "FY_22" };
                    //string[] rowsToDelete = { "Revenue Growth %'age", "BilledMonths", "PersonMonths","Digital %'age","T&M","RPP (31Mar2019) in KUSD" };
                    //string[] rowsToDelete_digital = {"T&M", "RPP (31Mar2019) in KUSD" };

                    dtQtrly_columnsToDelete.ToList().ForEach(col => dtQtrly.Columns.Remove(col));
                    dtDigital_columnsToDelete.ToList().ForEach(col => dtDigital.Columns.Remove(col));
                    //foreach (DataRow row in dtQtrly.Rows)
                    //{
                    //    string param = row["PARAMETER"].ToString().Trim();
                    //    if (rowsToDelete.Contains(param))
                    //        row.Delete();
                    //}
                    //foreach (DataRow row in dtDigital.Rows)
                    //{
                    //    string param_digital = row["PARAMETER"].ToString().Trim();
                    //    if (rowsToDelete_digital.Contains(param_digital))
                    //        row.Delete();
                    //}

                    dtQtrly.AcceptChanges();
                    dtDigital.AcceptChanges();

                    // Loading the data to an 2D array
                    rawData = new object[dtQtrly.Rows.Count, dtQtrly.Columns.Count];
                    //for (int col = 0; col < dtQtrly.Columns.Count; col++)
                    //    rawData[0, col] = dtQtrly.Columns[col].ColumnName;

                    for (int col = 0; col < dtQtrly.Columns.Count; col++)
                        for (int row = 0; row < dtQtrly.Rows.Count; row++)
                            rawData[row, col] = dtQtrly.Rows[row][col];



                    c1 = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[3, 2];
                    c2 = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[dtQtrly.Rows.Count + 2, dtQtrly.Columns.Count + 1];
                    range_excel = sheet.get_Range(c1, c2);

                    //Fill Array in Excel
                    range_excel.Value2 = rawData;
                    // range_excel.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    //range_excel.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid;
                    //range_excel.Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }

                if (1 == 2) // digtial
                {


                    // Loading the data to an 2D array

                    rawData = new object[dtDigital.Rows.Count, dtDigital.Columns.Count];
                    //for (int col = 0; col < dtDigital.Columns.Count; col++)
                    //    rawData[0, col] = dtDigital.Columns[col].ColumnName;

                    for (int col = 0; col < dtDigital.Columns.Count; col++)
                        for (int row = 0; row < dtDigital.Rows.Count; row++)
                            rawData[row, col] = dtDigital.Rows[row][col];



                    c1 = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[77, 2];
                    c2 = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[dtDigital.Rows.Count + 76, dtDigital.Columns.Count + 1];
                    range_excel = sheet.get_Range(c1, c2);

                    //Fill Array in Excel
                    range_excel.Value2 = rawData;
                    // range_excel.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    // range_excel.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid;
                    // range_excel.Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }


                if (1 == 1)
                {
                    //string[] rowsToDelete = { "Digital %", "Core%" };
                    //foreach (DataRow row in dtClientCat.Rows)
                    //{
                    //    string param = row["category"].ToString();
                    //    if (rowsToDelete.Contains(param))
                    //        row.Delete();
                    //}
                    string[] columnsToDelete = { };//"category", "Values" ,"sl"};
                    columnsToDelete.ToList().ForEach(col => dtClientCat.Columns.Remove(col));
                    dtClientCat.AcceptChanges();

                    // Loading the data to an 2D array
                    rawData = new object[dtClientCat.Rows.Count + 1, dtClientCat.Columns.Count];

                    for (int col = 0; col < dtClientCat.Columns.Count; col++)
                        for (int row = 0; row < dtClientCat.Rows.Count; row++)
                            rawData[row, col] = dtClientCat.Rows[row][col];

                    c1 = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[4, 29];
                    c2 = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[dtClientCat.Rows.Count + 4, dtClientCat.Columns.Count + 28];
                    range_excel = sheet.get_Range(c1, c2);

                    //Fill Array in Excel
                    range_excel.Value2 = rawData;
                    // range_excel.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    // range_excel.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid;
                    // range_excel.Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }




                ReleaseObject(range_excel);
                ReleaseObject(c2);
                ReleaseObject(c1);
                ReleaseObject(sheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public void FillExcelSheet_SummaryDeck1(DataTable[] _dts, Microsoft.Office.Interop.Excel.Worksheet sheet)
        {

            object[,] rawData = null;
            string[] ordering = { "SAP", "ORC", "EAIS", "ECAS", "EAS" };
           // string[] ordering = { "SAP", "ORC", "EAIS", "ECAS", "Fluido", "EAS" };
            string[] verticalGroupOrdering = { "CRL", "FSHIL", "COREMFG", "CMT", "SURE", "Others" };
            string[] verticalMSDOrdering = { "CRL", "FSHIL", "COREMFG", "CMT", "SURE" };

            int startRow = 0;
            int startCol = 0;
            int endRow = 0;
            int endCol = 0;

            // _dts = _dts.Where(k => k.TableName == SummaryDeckTableNames.SummaryDeck2Vertical).ToArray();
            List<DataTable> lstDts = new List<DataTable>();
            foreach (DataTable dt in _dts)  // Modify the collection to order the service line.
            {
                if (dt.TableName == SummaryDeckTableNames.SummaryDeck2Actuals || dt.TableName == SummaryDeckTableNames.SummaryDeck2DigRev)
                {

                    DataTable dtClone = dt.Clone();
                    dtClone.TableName = dt.TableName;
                    foreach (string sl in ordering)
                    {
                        DataRow rowTemp = dt.Rows.OfType<DataRow>().FirstOrDefault(k => k[0].ToString() == sl);
                        if (rowTemp != null)
                            dtClone.Rows.Add(rowTemp.ItemArray);
                    }
                    lstDts.Add(dtClone);
                }
                else if (dt.TableName == SummaryDeckTableNames.SummaryDeck2VerticalGrouping)
                {
                    DataTable dtClone = dt.Clone();
                    dtClone.TableName = dt.TableName;
                    foreach (string vertical in verticalGroupOrdering)
                    {
                        DataRow rowTemp = dt.Rows.OfType<DataRow>().FirstOrDefault(k => k["ver"].ToString() == vertical);
                        if (rowTemp != null)
                            dtClone.Rows.Add(rowTemp.ItemArray);
                    }
                    lstDts.Add(dtClone);
                }
                else if (dt.TableName == SummaryDeckTableNames.SummaryDeck2Vertical)
                {
                    DataTable dtClone = dt.Clone();
                    dtClone.TableName = dt.TableName;
                    var lstOthers = dt.Rows.OfType<DataRow>().Select(k => k["vertical"].ToString()).ToList().Except(verticalMSDOrdering);
                    List<string> lstAll = new List<string>(verticalMSDOrdering);
                    lstAll.AddRange(lstOthers);

                    foreach (string vertical in lstAll)
                    {
                        DataRow rowTemp = dt.Rows.OfType<DataRow>().FirstOrDefault(k => k["vertical"].ToString() == vertical);
                        if (rowTemp != null)
                            dtClone.Rows.Add(rowTemp.ItemArray);
                    }
                    lstDts.Add(dtClone);
                }
                else
                    lstDts.Add(dt);
            }


            try
            {
                Microsoft.Office.Interop.Excel.Range c1 = default(Microsoft.Office.Interop.Excel.Range);
                Microsoft.Office.Interop.Excel.Range c2 = default(Microsoft.Office.Interop.Excel.Range);
                Microsoft.Office.Interop.Excel.Range range_excel = default(Microsoft.Office.Interop.Excel.Range);

                foreach (DataTable dt in lstDts)
                {
                    if (dt.TableName == SummaryDeckTableNames.SummaryDeck2Actuals)
                    { startRow = 3; startCol = 1; endRow = 7; endCol = 7; }
                    if (dt.TableName == SummaryDeckTableNames.SummaryDeck2Region)
                    { startRow = 3; startCol = 11; endRow = dt.Rows.Count + 2; endCol = 12; }
                    if (dt.TableName == SummaryDeckTableNames.SummaryDeck2Vertical)
                    { startRow = 3; startCol = 15; endRow = dt.Rows.Count + 2; endCol = 16; }
                    if (dt.TableName == SummaryDeckTableNames.SummaryDeck2VerticalGrouping)
                    {
                        startRow = 3; startCol = 19; endRow = dt.Rows.Count + 2; endCol = 20;
                    }
                    if (dt.TableName == SummaryDeckTableNames.SummaryDeck2RevGrowth)
                    { startRow = 15; startCol = 1; endRow = 16; endCol = 10; }
                    if (dt.TableName == SummaryDeckTableNames.SummaryDeck2DigRevGrowth)
                    { startRow = 19; startCol = 1; endRow = 21; endCol = 12; }
                    if (dt.TableName == SummaryDeckTableNames.SummaryDeck2DigRev)
                    { startRow = 24; startCol = 1; endRow = 28; endCol = 8; }

                    rawData = new object[dt.Rows.Count, dt.Columns.Count];

                    for (int col = 0; col < dt.Columns.Count; col++)
                        for (int row = 0; row < dt.Rows.Count; row++)
                            rawData[row, col] = dt.Rows[row][col];

                    c1 = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[startRow, startCol];
                    c2 = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[endRow, endCol];
                    range_excel = sheet.get_Range(c1, c2);

                    //Fill Array in Excel
                    range_excel.Value2 = rawData;
                    // range_excel.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    // range_excel.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid;
                    //range_excel.Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                    //  range_excel.Borders.LineStyle = Excel.XlLineStyle.xlDot;
                }



                ReleaseObject(range_excel);
                ReleaseObject(c2);
                ReleaseObject(c1);
                ReleaseObject(sheet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public void FillInvestmentSales(DataTable dt, Microsoft.Office.Interop.Excel.Worksheet excel, string sl)
        {


            try
            {

                Microsoft.Office.Interop.Excel.Range c1;
                Microsoft.Office.Interop.Excel.Range c2;
                Microsoft.Office.Interop.Excel.Range range_excel;



                object[,] rawData = new object[dt.Rows.Count, dt.Columns.Count];
                for (int col = 0; col < dt.Columns.Count; col++)
                    for (int row = 0; row < dt.Rows.Count; row++)
                        rawData[row, col] = dt.Rows[row][col];

                // Filling the data 
                c1 = (Microsoft.Office.Interop.Excel.Range)excel.Cells[3, 1];
                c2 = (Microsoft.Office.Interop.Excel.Range)excel.Cells[dt.Rows.Count + 2, dt.Columns.Count];
                range_excel = excel.get_Range(c1, c2);

                range_excel.Value2 = rawData;
                range_excel.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                range_excel.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid;
                range_excel.Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);





                ReleaseObject(range_excel);
                ReleaseObject(c2);
                ReleaseObject(c1);
                ReleaseObject(excel);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }



    }


}