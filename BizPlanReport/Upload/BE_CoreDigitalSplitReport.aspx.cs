using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Globalization;
using System.Data;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
 
using Excel = Microsoft.Office.Interop.Excel;
using VBIDE = Microsoft.Vbe.Interop;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Data.SqlClient;
using System.Configuration;

namespace BECodeProd
{

    public partial class BE_CoreDigitalSplitReport : Page
    {

        static string connString = ConfigurationManager.ConnectionStrings["DBConnectionString"].ToString();
        public string fileName = "BEData.BETrendReport.cs";

    static string userid;
    protected void Page_Load(object sender, EventArgs e)
    {
        
        if (!IsPostBack)
        {
                userid = UserIdentity.GetCurrentUser();
            //string cmdtext1 = "EXEC dbo.[SPROC_GetServiceLine]";
            //DataSet dsn1 = new DataSet();
            //dsn1 = service.GetSUForuser(cmdtext1);

                List<string> lstSU = GetSUForuser(userid);

            if (lstSU.Count > 1)
            {
                ddlServiceLine.DataSource = lstSU.Select(k => k.ToString()).Distinct().ToList();
                ddlServiceLine.DataBind();
                ddlServiceLine.Items.Insert(0, "ALL");
            }
            else if (lstSU.Count == 1)
            {
                ddlServiceLine.DataSource = lstSU.Select(k => k.ToString()).Distinct().ToList();
                ddlServiceLine.DataBind();

            }

            string PrevQtr = DateUtility.GetQuarter("prev");
            string currentQtr = DateUtility.GetQuarter("current");
            string nextQtr = DateUtility.GetQuarter("next");
            string nextQtrPlus1 = DateUtility.GetQuarter("next1");

            ddlQuarter.Items.Insert(0, PrevQtr);
            ddlQuarter.Items.Insert(1, currentQtr);
            ddlQuarter.Items.Insert(2, nextQtr);
            ddlQuarter.Items.Insert(3, nextQtrPlus1);

            ddlQuarter.SelectedIndex = 1;

            btnreport.Visible = true;
            lblError.Visible = false;




        }
    }

    //code for btnreport

    public static bool TryKillProcessByMainWindowHwnd(int hWnd)
    {
        uint processID;
        GetWindowThreadProcessId((IntPtr)hWnd, out processID);
        if (processID == 0) return false;
        try
        {
            Process.GetProcessById((int)processID).Kill();
        }
        catch (ArgumentException)
        {
            return false;
        }
        catch (Exception ex)
        {
            return false;
        }
        return true;
    }

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);



    protected void btnreport_Click(object sender, EventArgs e)
    {
            string userID = UserIdentity.GetCurrentUser();

        DataTable dtBE = new DataTable();





        string SU = ddlServiceLine.SelectedItem.Text;

        string Qtr = ddlQuarter.Text.Remove(2);

        int tempyear = Convert.ToInt32(ddlQuarter.Text.Remove(0, 3)) + 2000 - 1;
        string FinYear = string.Format("{0}-{1}", tempyear, (tempyear - 2000 + 1));


        dtBE = GetBE_CoreDigitalSplit(FinYear, SU, Qtr);

        if (dtBE == null || dtBE.Rows.Count == 0)
        {
            lbl.Text = "";
            Session["key"] = null;
            //Page.ClientScript.RegisterStartupScript(this.GetType(), Guid.NewGuid().ToString(), "<script language=JavaScript>alert('No Data to download!');</script>");
            string message = "alert('No Data to download!')";
            ScriptManager.RegisterStartupScript((sender as Control), this.GetType(), "alert", message, true);
            return;
        }

        var currQtr = dtBE.Columns[5].ColumnName.Insert(2, "'");
        var nextQtr = dtBE.Columns[6].ColumnName.Insert(2, "'");

        dtBE.Columns[5].ColumnName = currQtr;
        dtBE.Columns[6].ColumnName = nextQtr;

        string folder = "ExcelOperations";
        var MyDir = new DirectoryInfo(Server.MapPath(folder));

        string Filename = "BE_Core_Digital_Split_" + currQtr + "_" + nextQtr + "_" + userID + "_" + DateTime.Now.ToString("ddMMMyyyy_HHmmss") + "IST.xlsx";

        if (MyDir.GetFiles().SingleOrDefault(k => k.Name == Filename) != null)
            System.IO.File.Delete(MyDir.FullName + "\\" + Filename);


        FileInfo file = new FileInfo(MyDir.FullName + "\\" + Filename);


        ExcelPackage pck = new ExcelPackage();
        ExcelWorksheet ws;
        ws = pck.Workbook.Worksheets.Add("Data");

        ws.Cells["A1"].LoadFromDataTable(dtBE, true);

        int rowcountSheet0 = dtBE.Rows.Count;
        int colcountSheet0 = dtBE.Columns.Count;

        ws.Cells[1, 1, 1, colcountSheet0].Style.Font.Bold = true;
        var fill = ws.Cells[1, 1, 1, colcountSheet0].Style.Fill;
        fill.PatternType = ExcelFillStyle.Solid;
        fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
        ws.Cells[1, 1, rowcountSheet0 + 1, colcountSheet0].AutoFitColumns();

        ws.Cells[1, 1, rowcountSheet0 + 1, colcountSheet0 + 1].Style.Font.Name = "calibri";
        ws.Cells[1, 1, rowcountSheet0 + 1, colcountSheet0 + 1].Style.Font.Size = 9;

        pck.SaveAs(file);
        pck.Dispose();
        ReleaseObject(pck);
        ReleaseObject(ws);

        DownloadFileBEReport_new(Filename);


    }





    void DownloadFileBEReport_new(string Filename)
    {
            string UserId = UserIdentity.GetCurrentUser();

        Excel.Application oExcel;
        Excel.Workbook oBook = default(Excel.Workbook);


        try
        {
            string folder = "ExcelOperations";
            var MyDir = new DirectoryInfo(Server.MapPath(folder));

            Object oMissing = System.Reflection.Missing.Value;

            oExcel = new Excel.Application();





            oBook = oExcel.Workbooks.
                Open(MyDir.FullName + "\\" + Filename + "", 0, false, 5, "", "", true,
                Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);



            //Adding permission to excel file//
            System.Threading.Thread.Sleep(1000);
            oBook.Activate();
            oBook.Permission.Enabled = true;
            oBook.Permission.RemoveAll();
            string strExpiryDate = DateTime.Now.AddDays(60).Date.ToString();
            DateTime dtTempDate = Convert.ToDateTime(strExpiryDate);
            DateTime dtExpireDate = new DateTime(dtTempDate.Year, dtTempDate.Month, dtTempDate.Day);
            UserPermission userper = oBook.Permission.Add("Everyone", MsoPermission.msoPermissionChange);
            userper.ExpirationDate = dtExpireDate;


            oExcel.DisplayAlerts = false;
            int hWnd1 = oExcel.Application.Hwnd;
            oBook.Save();
            oBook.Close(false);
            oExcel.Quit();
            if (oExcel != null)
            {
                TryKillProcessByMainWindowHwnd(hWnd1);
            }


            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oBook);
            oBook = null;






            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Session.Add("key", Filename);

            loading.Style.Add("visibility", "visible");
            lbl.Text = "Downloaded";
            up.Update();
            iframe.Attributes.Add("src", "Download.aspx");

            ClientScript.RegisterStartupScript(this.GetType(), "myStopFunction", "myStopFunction();", true);
            ClientScript.RegisterStartupScript(this.GetType(), "isvaliduploadClose", "isvaliduploadClose();", true);

        }
        catch (Exception ex)
        {
            throw ex;
        }

    }


    public void ReleaseObject(object o)
    {
        try
        {
            if (o != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
        }
        catch (Exception) { }
        finally { o = null; }
    }


        public List<string> GetSUForuser(string userid)
        {


            DataSet ds = new DataSet();

            SqlCommand objCommand;

            List<string> lstempCollection = new List<string>();

            try
            {

                objCommand = new SqlCommand();
                SqlConnection con = new SqlConnection(connString);

                SqlParameter objBE = new SqlParameter();
                objBE.ParameterName = "@userid";
                objBE.Direction = ParameterDirection.Input;
                objBE.SqlDbType = SqlDbType.VarChar;
                objBE.Value = userid;


                objCommand = new SqlCommand();
                objCommand.CommandText = "spBEGetSU_dummy";
                objCommand.CommandType = CommandType.StoredProcedure;
                objCommand.Connection = con;
                SqlParameterCollection objParamColl = objCommand.Parameters;
                objParamColl.Add(objBE);

                SqlDataAdapter da = new SqlDataAdapter(objCommand);
                da.Fill(ds);

                
               

                if (ds != null && ds.Tables != null && ds.Tables.Count > 0)
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        // empCollection = new DUPUCCMap();
                        string tmp = ds.Tables[0].Rows[i]["SU"].ToString();

                        lstempCollection.Add(tmp);
                    }

                }
            }
            catch (Exception ex)
            {

                
                throw ex;
            }
            finally
            {
                
            }


            return lstempCollection;
        }



        public DataTable GetBE_CoreDigitalSplit(string finYear, string SU, string Qtr)
        {
            SqlConnection con = new SqlConnection(connString);
            DataSet ds = new DataSet();
            SqlParameter objParm1, objParm2, objParm3;
            SqlCommand objCommand;
            SqlParameterCollection objParamColl;

            try
            {
                objParm1 = new SqlParameter();
                objParm1.ParameterName = "@SL";
                objParm1.Direction = ParameterDirection.Input;
                objParm1.SqlDbType = SqlDbType.NVarChar;
                objParm1.Value = SU;

                objParm2 = new SqlParameter();
                objParm2.ParameterName = "@CurrQtr";
                objParm2.Direction = ParameterDirection.Input;
                objParm2.SqlDbType = SqlDbType.VarChar;
                objParm2.Value = Qtr;

                objParm3 = new SqlParameter();
                objParm3.ParameterName = "@CurrQtrFY";
                objParm3.Direction = ParameterDirection.Input;
                objParm3.SqlDbType = SqlDbType.VarChar;
                objParm3.Value = finYear;


                objCommand = new SqlCommand();
                objCommand.CommandText = "sp_BE_Core_Digital_split_report";
                objCommand.CommandType = CommandType.StoredProcedure;
                objCommand.Connection = con;

                objParamColl = objCommand.Parameters;

                objParamColl.AddRange(new SqlParameter[] { objParm1, objParm2, objParm3 });
                SqlDataAdapter da = new SqlDataAdapter(objCommand);
                da.Fill(ds);


              
                if (ds != null && ds.Tables != null && ds.Tables.Count > 0)
                {
                    DataTable dt = new DataTable();
                    dt = ds.Tables[0];
                    return dt;
                }

            }

            catch (Exception ex)
            {

                
                throw ex;

            }

            finally
            {

                

            }



            return new DataTable();

        }



    }

    public class DateUtility
    {


        public static string GetQuarter(string strquarter)
        {
            // string strquarter = Session[Constants.QUARTER] + "";

            if (strquarter.ToLower() == "current")
            {

                DateTime todaydate = DateTime.Now;
                int year = todaydate.Year - 2000;
                int nextyear = year + 1;
                if (todaydate.Month == 1 || todaydate.Month == 2 || todaydate.Month == 3)
                    strquarter = "Q4'" + year;
                else if (todaydate.Month == 4 || todaydate.Month == 5 || todaydate.Month == 6)
                    strquarter = "Q1'" + nextyear;
                else if (todaydate.Month == 7 || todaydate.Month == 8 || todaydate.Month == 9)
                    strquarter = "Q2'" + nextyear;
                else
                    strquarter = "Q3'" + nextyear;
            }

            else if (strquarter.ToLower() == "next")
            {

                DateTime todaydate = DateTime.Now.AddMonths(3);
                int year = todaydate.Year - 2000;
                int nextyear = year + 1;
                if (todaydate.Month == 1 || todaydate.Month == 2 || todaydate.Month == 3)
                    strquarter = "Q4'" + year;
                else if (todaydate.Month == 4 || todaydate.Month == 5 || todaydate.Month == 6)
                    strquarter = "Q1'" + nextyear;
                else if (todaydate.Month == 7 || todaydate.Month == 8 || todaydate.Month == 9)
                    strquarter = "Q2'" + nextyear;
                else
                    strquarter = "Q3'" + nextyear;
            }

            else if (strquarter.ToLower() == "next1")
            {

                DateTime todaydate = DateTime.Now.AddMonths(6);
                int year = todaydate.Year - 2000;
                int nextyear = year + 1;
                if (todaydate.Month == 1 || todaydate.Month == 2 || todaydate.Month == 3)
                    strquarter = "Q4'" + year;
                else if (todaydate.Month == 4 || todaydate.Month == 5 || todaydate.Month == 6)
                    strquarter = "Q1'" + nextyear;
                else if (todaydate.Month == 7 || todaydate.Month == 8 || todaydate.Month == 9)
                    strquarter = "Q2'" + nextyear;
                else
                    strquarter = "Q3'" + nextyear;
            }

            else if (strquarter.ToLower() == "next2")
            {

                DateTime todaydate = DateTime.Now.AddMonths(9);
                int year = todaydate.Year - 2000;
                int nextyear = year + 1;
                if (todaydate.Month == 1 || todaydate.Month == 2 || todaydate.Month == 3)
                    strquarter = "Q4'" + year;
                else if (todaydate.Month == 4 || todaydate.Month == 5 || todaydate.Month == 6)
                    strquarter = "Q1'" + nextyear;
                else if (todaydate.Month == 7 || todaydate.Month == 8 || todaydate.Month == 9)
                    strquarter = "Q2'" + nextyear;
                else
                    strquarter = "Q3'" + nextyear;
            }

            else if (strquarter.ToLower() == "prev")
            {

                DateTime todaydate = DateTime.Now.AddMonths(-3);
                int year = todaydate.Year - 2000;
                int nextyear = year + 1;
                if (todaydate.Month == 1 || todaydate.Month == 2 || todaydate.Month == 3)
                    strquarter = "Q4'" + year;
                else if (todaydate.Month == 4 || todaydate.Month == 5 || todaydate.Month == 6)
                    strquarter = "Q1'" + nextyear;
                else if (todaydate.Month == 7 || todaydate.Month == 8 || todaydate.Month == 9)
                    strquarter = "Q2'" + nextyear;
                else
                    strquarter = "Q3'" + nextyear;
            }

            return strquarter;
        }


    }

}