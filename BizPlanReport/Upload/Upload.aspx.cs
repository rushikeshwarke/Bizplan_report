using BEBIZ.Upload.Sheets;
using BizPlanReport.Upload.Sheets;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Excel = Microsoft.Office.Interop.Excel;
using VBIDE = Microsoft.Vbe.Interop;
public class UserIdentity
{
    public static string GetCurrentUser()
    {
        string userid = HttpContext.Current.User.Identity.Name;
        string[] userids = userid.Split('\\');
        if (userids.Length == 2)
        {
            userid = userids[1];
        }

        return userid;
    }
}
namespace BEBIZ.Upload
{


    public partial class Upload : System.Web.UI.Page
    {
        static string connString = ConfigurationManager.ConnectionStrings["DBConnectionString"].ToString();
        protected void Page_Load(object sender, EventArgs e)
        {





            if (!IsPostBack)
            {
                SqlConnection con = new SqlConnection(connString);
                SqlCommand cmd = new SqlCommand() { CommandText = "SELECT 'ALL' as Vertical union select DISTINCT txtVertical from demclientcodeportfolio where  txtVertical not in ('TEST','') and isActive='Y'", CommandType = CommandType.Text };

                cmd.Connection = con;
                SqlDataAdapter sdr = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                sdr.Fill(ds);
                ddlVertical.DataSource = ds.Tables[0].DefaultView;
                ddlVertical.DataValueField = "Vertical";
                ddlVertical.DataTextField = "Vertical";
                ddlVertical.DataBind();

                hdnRole.Value = GetRole();

                string userID = UserIdentity.GetCurrentUser();
                List<string> lstSU = GetSUForuser(userID);

                lstSU = lstSU.Select(k => k.ToString()).Distinct().ToList();

                if (lstSU.Count == 4)

                    lstSU.Insert(0, "EAS");

                ddl_Internal.DataSource = lstSU;
                ddl_Internal.DataBind();


                ddl_Internal_SelectedIndexChanged(null, null);



                if (hdnRole.Value == "0")
                {
                    Response.Redirect("Error.aspx?Message=You are not authorised to access this portal, Please contact the Administrator.");

                }


            }
        }

        private string GetCurrentUser()
        {
            throw new NotImplementedException();
        }

        protected string GetRole()
        {
            List<string> lstSPOC = new List<string>();
            List<string> lstAdmin = new List<string>();

            //lstSPOC.Add("prema_haridas");
            //lstSPOC.Add("patel_jignesh");
            //lstSPOC.Add("srinivas_manjunath");

            string role = UserIdentity.GetCurrentUser().ToLower().Trim();
            SqlConnection con = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand() { CommandText = " select MailId from eas_dashboardaaccess where AccessType = 'admin' ; select MailId from eas_dashboardaaccess where Bizplan_Access = 'Y' ", CommandType = CommandType.Text };
            cmd.Connection = con;
            SqlDataAdapter sdr = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            sdr.Fill(ds);
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                lstAdmin.Add(row[0].ToString().ToLower().Trim());
            }

            foreach (DataRow row in ds.Tables[1].Rows)
            {
                lstSPOC.Add(row[0].ToString().ToLower().Trim());
            }



            if (lstSPOC.Contains(role))
                return "SPOC";
            if (lstAdmin.Contains(role))
                return "ADMIN";
            return "0";


        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            hidTAB.Value = "#home";
            if (fuUploader.HasFile)
                if (Path.GetExtension(fuUploader.FileName) == ".xlsx")
                {
                    string path = string.Concat(Server.MapPath("~/UploadOperations/" + fuUploader.FileName));
                    fuUploader.SaveAs(path);
                    DataSet ds = UploadExcel(path);
                    string sl = ddlSL.Text;
                    LoadData(ds, sl);
                   
                    ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "success", "alert('Data Saved Successfully!');", true);
                }
        }

        public void FinYearSplit(string sl )
        {
            DataTable dt = new DataTable();
            SqlConnection sqlconn = new SqlConnection(connString);
            SqlCommand sqlcmd1 = new SqlCommand("sp_BizPlan_FinYearSplit", sqlconn);
            sqlcmd1.Parameters.AddWithValue("@SL", sl);
            sqlcmd1.CommandType = CommandType.StoredProcedure;
            sqlconn.Open();
            sqlcmd1.ExecuteNonQuery();
            sqlconn.Close();
            
        }
        private void LoadData(DataSet ds, string sl)
        {
            // *****************************************************
            //ISheetInfo accInvest = new AccInvestment(ds);//orc
            //ISheetInfo talentTop20 = new TalentTop20SL(ds);//orc
            //ISheetInfo refactor = new TalentRefactoring(ds);
            //ISheetInfo talentUS = new TalentUsStatus_SL(ds);
            //ISheetInfo internalcollaboration = new InternalCollaboration(ds);
            //ISheetInfo Scale_AgileDigital = new ScaleAgileDigital(ds);
            // OLD**************************************************

             

            ISheetInfo organic = new OrganicAccRevenue(ds, sl);//orc
                                                               //   ISheetInfo inorganic = new InOrganicSLRev(ds, sl);//orc

            ISheetInfo slVolume = null;
            //if (sl.ToLower() != "orc")
            slVolume = new SLVolume(ds, sl);
            ISheetInfo DCG = new DCGProductPlan(ds, sl);
            ISheetInfo Maximus_OM = new Maximus_Margin_Intervention(ds, sl);
            ISheetInfo InfosysAssetPlan = new InfosysAssetsPlan(ds, sl);
            ISheetInfo slInvest = new SLInvestment(ds, sl);//orc
            ISheetInfo Investment_Practice_Sales = new Investment_SL_Practice_Sales(ds, sl);
            ISheetInfo StartupNewAllianceEcosystem = new StartupNewAllianceEcosystem(ds, sl);

            ////ISheetInfo talentSkillPlan = new TalentSkillPlan(ds, sl);//orc            
            //orc
            //  ISheetInfo Investment_RandD = new RandD(ds, sl);//orc


            //ISheetInfo PartnerEcoSys = new PartnerEcoSystemsAlliances(ds, sl);//orc

            //ISheetInfo[] allSheets = { organic, inorganic, slVolume, slInvest,
            //talentSkillPlan,  Investment_Practice_Sales,Investment_RandD ,PartnerEcoSys};



            List<ISheetInfo> lstallSheets = new List<ISheetInfo>();

            lstallSheets.Clear();

            //if (sl.ToLower() != "orc")

            lstallSheets.Add(organic);
            lstallSheets.Add(slVolume);
            lstallSheets.Add(DCG);
            lstallSheets.Add(Maximus_OM);
            lstallSheets.Add(InfosysAssetPlan);
            lstallSheets.Add(slInvest);
            lstallSheets.Add(Investment_Practice_Sales); // ok 
            lstallSheets.Add(StartupNewAllianceEcosystem); // ok 





            var allSheets = lstallSheets.ToArray();



            allSheets.ToList().ForEach(k => k.SL = sl);
            DataTable dtError;
            if (IsValidFileData(sl, ds, allSheets, out dtError))
            {
                try
                {
                    allSheets.ToList().ForEach(k => k.Save());
                    FinYearSplit(sl);
                    lblError.Text = "Saved Successfully";
                }
                catch (Exception ex)
                {

                    lblError.Text = ex.Message;
                    throw ex;
                }
        }
            else
            {
                lblError.Text = "One or more error occured while processing the request, Please check the error report ";
                ExportErrorReport(dtError);
    }



}

        bool IsValidFileData(string sl, DataSet ds, ISheetInfo[] sheets, out DataTable dtError0)
        {

            DataTable dtError = new DataTable();
            dtError.Columns.Add("Sheet Name");
            dtError.Columns.Add("Error Type");
            dtError.Columns.Add("Row No");
            dtError.Columns.Add("Error Message");
            foreach (ISheetInfo sheet in sheets)
            {
                sheet.SL = sl;
                var lst = sheet.Validate();

                foreach (var item in lst)
                {
                    dtError.Rows.Add(item.SheetName, item.ErrType.ToString(), item.RowNo, item.Message);
                }

                if (sheet.SheetName == "InfosysAssetsPlan")
                {
                    for (int i = 0; i < dtError.Rows.Count; i++)
                    {
                        DataRow dr = dtError.Rows[i];
                        if (dr["Error Type"].ToString() == "Duplicate")
                            dr.Delete();
                    }
                    dtError.AcceptChanges();
                }

            }
            dtError0 = dtError.Copy();
            return dtError.Rows.Count == 0;

        }


        protected DataSet UploadExcel(string path)
        {
            DataSet ds = new DataSet();

            string conString = string.Empty;
            //conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0;IMEX=1;\"";
            conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;IMEX=1;\"";

            conString = string.Format(conString, path);

            using (OleDbConnection excel_con = new OleDbConnection(conString))
            {

                excel_con.Open();

                DataTable worksheets = excel_con.GetSchema("Tables");
                string w = worksheets.Columns["TABLE_NAME"].ToString();
                List<string> lstsheetNames = new List<string>();
                Action<DataRow> actionToGetSheetName = (k) => { lstsheetNames.Add(k["TABLE_NAME"] + ""); };
                worksheets.Rows.OfType<DataRow>().ToList().ForEach(actionToGetSheetName);
                lstsheetNames = lstsheetNames.Where(k => k.Contains("$")).ToList();
                lstsheetNames = lstsheetNames.Select(k => k.Trim().Trim('\'')).ToList();
                lstsheetNames = lstsheetNames.Where(k => !k.Contains("FilterDatabase")).ToList();
                foreach (string sheetname in lstsheetNames)
                {
                    using (OleDbDataAdapter oda = new OleDbDataAdapter(string.Format("SELECT *  FROM [{0}]", sheetname), excel_con))
                    {
                        DataTable dt = new DataTable() { TableName = sheetname };
                        oda.Fill(dt);
                        ds.Tables.Add(dt);

                    }
                }
                excel_con.Close();
            }
            return ds;

        }


        protected void ExportErrorReport(DataTable dtError)
        {
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dtError, "Error Report");

                Response.Clear();
                Response.Buffer = true;
                Response.Charset = "";
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=ErrorReport.xlsx");
                using (MemoryStream MyMemoryStream = new MemoryStream())
                {
                    wb.SaveAs(MyMemoryStream);
                    MyMemoryStream.WriteTo(Response.OutputStream);
                    Response.Flush();
                    Response.End();
                }
            }
        }

        protected void btnDownload_Click(object sender, EventArgs e)
        {

        }
        //protected void btnDownload_Click(object sender, EventArgs e)
        //{
        //    SqlCommand cmdAccount = new SqlCommand() { CommandText = "sp_Fetch_BizPlan_SL_Account", CommandType = CommandType.StoredProcedure };
        //    cmdAccount.Parameters.AddWithValue("@SL", ddlSL.SelectedItem.Text);
        //    cmdAccount.Parameters.AddWithValue("@Model", ddlModel.SelectedItem.Text);
        //    if (ddlModel.SelectedItem.Text == "Moderate")
        //    {
        //        if (txtPercentageforModerate.Text == "")
        //        {
        //            cmdAccount.Parameters.AddWithValue("@ModeratePercentage", null);
        //        }
        //        else
        //        {
        //            cmdAccount.Parameters.AddWithValue("@ModeratePercentage", txtPercentageforModerate.Text);
        //        }

        //        if (txtNumberforModerate.Text == "")
        //        {
        //            cmdAccount.Parameters.AddWithValue("@num", 2);
        //        }
        //        else
        //        {
        //            cmdAccount.Parameters.AddWithValue("@num", txtNumberforModerate.Text);
        //        }
        //    }
        //    else
        //    {
        //        cmdAccount.Parameters.AddWithValue("@ModeratePercentage", null);
        //        cmdAccount.Parameters.AddWithValue("@num", 0);
        //    }

        //    SqlCommand cmdInvestment = new SqlCommand() { CommandText = "sp_Fetch_BizPlan_SL_Acc_Investment", CommandType = CommandType.StoredProcedure };
        //    cmdInvestment.Parameters.AddWithValue("@SL", ddlSL.SelectedItem.Text);
        //    cmdInvestment.Parameters.AddWithValue("@Model", ddlModel.SelectedItem.Text);
        //    if (ddlModel.SelectedItem.Text == "Moderate")
        //    {
        //        if (txtPercentageforModerate.Text == "")
        //        {
        //            cmdInvestment.Parameters.AddWithValue("@ModeratePercentage", null);
        //        }
        //        else
        //        {
        //            cmdInvestment.Parameters.AddWithValue("@ModeratePercentage", txtPercentageforModerate.Text);
        //        }

        //        if (txtNumberforModerate.Text == "")
        //        {
        //            cmdInvestment.Parameters.AddWithValue("@num", 2);
        //        }
        //        else
        //        {
        //            cmdInvestment.Parameters.AddWithValue("@num", txtNumberforModerate.Text);
        //        }
        //    }
        //    else
        //    {
        //        cmdInvestment.Parameters.AddWithValue("@ModeratePercentage", null);
        //        cmdInvestment.Parameters.AddWithValue("@num", 0);
        //    }


        //    DataSet dsAccount = GetDataSet(cmdAccount);
        //    DataTable dtAccount = dsAccount.Tables[0];
        //    DataSet dsInvestment = GetDataSet(cmdInvestment);
        //    DataTable dtInvestment = dsInvestment.Tables[0];

        //    string folder = "ExcelOperations";
        //    var myDir = new DirectoryInfo(Server.MapPath(folder));

        //    string filename = ddlSL.SelectedItem.Text + "_3 year Business Plan_" + DateTime.Now.ToString("ddMMMyyyy_HHmm") + "IST.xlsx";

        //    String downloadFile = myDir + "\\" + filename;

        //    if (myDir.GetFiles().SingleOrDefault(k => k.Name == filename) != null)
        //    {
        //        System.IO.File.Delete(downloadFile);
        //    }

        //    Microsoft.Office.Interop.Excel.Application oExcel = null;
        //    Microsoft.Office.Interop.Excel.Workbook oBook = default(Microsoft.Office.Interop.Excel.Workbook);
        //    Microsoft.Office.Interop.Excel.Sheets ws = null;

        //    Excel.Worksheet sheet_Account = null;
        //    Excel.Worksheet sheet_Investment = null;
        //    VBIDE.VBComponent oModule;
        //    String sCode;
        //    Object oMissing = System.Reflection.Missing.Value;

        //    //instance of excel
        //    oExcel = new Microsoft.Office.Interop.Excel.Application();
        //    string templatePath = Server.MapPath(folder) + "\\Template\\" + "BEBiz_Template.xlsx" + "";
        //    oBook = oExcel.Workbooks.
        //        Open(templatePath, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
        //    ws = oBook.Sheets;

        //    sheet_Account = ws.Item["Account"];
        //    sheet_Investment = ws.Item["Investments"];

        //    FillExcelSheet(dtAccount, sheet_Account);
        //    FillExcelSheet(dtInvestment, sheet_Investment);

        //    //oBook.Activate();
        //    //oBook.Permission.Enabled = true;
        //    //oBook.Permission.RemoveAll();
        //    //string strExpiryDate = DateTime.Now.AddDays(60).Date.ToString();
        //    //DateTime dtTempDate = Convert.ToDateTime(strExpiryDate);
        //    //DateTime dtExpireDate = new DateTime(dtTempDate.Year, dtTempDate.Month, dtTempDate.Day);
        //    //UserPermission userper = oBook.Permission.Add("Everyone", MsoPermission.msoPermissionChange);
        //    //userper.ExpirationDate = dtExpireDate;

        //    oExcel.DisplayAlerts = false;
        //    oBook.SaveAs(downloadFile);

        //    oBook.Close(false, templatePath, null);

        //    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(sheet_Account);
        //    sheet_Account = null;
        //    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(sheet_Investment);
        //    sheet_Investment = null;

        //    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ws);
        //    ws = null;
        //    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oBook);
        //    oBook = null;

        //    oExcel.Quit();
        //    oExcel = null;

        //    GC.Collect();
        //    GC.WaitForPendingFinalizers();
        //    GC.Collect();
        //    GC.WaitForPendingFinalizers();


        //    Session["key"] = filename;

        //    loading.Style.Add("visibility", "visible");
        //    lbl.Text = "Downloaded";
        //    up.Update();

        //    iframe.Attributes.Add("src", "Download.aspx");

        //    ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "myStopFunction", "myStopFunction()", true);
        //}
        protected void btn_ServiceLineBPlan_Input_Click(object sender, EventArgs e)
        {
            hidTAB.Value = "#menu3";
            string sl = ddlInput.SelectedItem.Text;

            //BEBIZ4.Upload.BizTableAdapters.sp_BizPlan_template_Acc_Rev_V1TableAdapter dd = new BEBIZ4.Upload.BizTableAdapters.sp_BizPlan_template_Acc_Rev_V1TableAdapter();
            SqlCommand cmdAcc_Rev = new SqlCommand() { CommandText = "sp_BizPlan_template_Acc_Rev_V2", CommandType = CommandType.StoredProcedure };
            cmdAcc_Rev.Parameters.AddWithValue("@SL", sl);

            DataTable dtAcc_Rev = GetDataTable(cmdAcc_Rev);

            //DataTable dtAcc_Rev = dd.GetData(sl);
            string filenameToBeSaved = "EAS_" + sl + "_FY21_27_Biz_Plan_Template_" + DateTime.Now.ToString("ddMMMyyyyHHmmsstt") + ".xlsx";
            string folder = "ExcelOperations";
            string templatePath = Server.MapPath(folder) + "\\Template\\InputTemplates\\EAS_" + sl + "_FY21_27_Biz_Plan_Template_v0.3.xlsx" + "";
            //var workbook = new XLWorkbook(templatePath);
            //IXLWorksheet Worksheet = workbook.Worksheet("Acc_Rev");
            // Worksheet.Row(5).Cell(1).InsertData(dtAcc_Rev.Rows);


            string saveFileName = Server.MapPath(folder) + "\\" + filenameToBeSaved;

            //workbook.SaveAs(saveFileName);


            Microsoft.Office.Interop.Excel.Application oExcel = null;
            Microsoft.Office.Interop.Excel.Workbook oBook = default(Microsoft.Office.Interop.Excel.Workbook);
            Microsoft.Office.Interop.Excel.Sheets ws = null;

            Excel.Worksheet sheet_Account = null;


            Object oMissing = System.Reflection.Missing.Value;

            //instance of excel
            oExcel = new Microsoft.Office.Interop.Excel.Application();
            oExcel.DisplayAlerts = false;
            oBook = oExcel.Workbooks.
                Open(templatePath, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            ws = oBook.Sheets;

            sheet_Account = ws.Item["Acc_Rev"];


            FillExcelSheet_Acc_Rev(dtAcc_Rev, sheet_Account);



            int hWnd1 = oExcel.Application.Hwnd;







            oBook.SaveAs(saveFileName);

            oBook.Close(false, templatePath, null);



            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(sheet_Account);
            sheet_Account = null;


            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ws);
            ws = null;
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oBook);
            oBook = null;

            oExcel.Quit();


            if (oExcel != null)
            {
                TryKillProcessByMainWindowHwnd(hWnd1);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();





            Session["key"] = filenameToBeSaved;

            //iframe.Attributes.Add("src", "Download.aspx");

            ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "download", "document.getElementById('iframe').src = 'Download.aspx';", true);

            ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "myStopFunction", "myStopFunction()", true);
        }

        protected void btn_ServiceLineBPlan_Download_Click(object sender, EventArgs e)
        {
            hidTAB.Value = "#menu1";
            string sl = ddl_SL_1.SelectedItem.Text;

            string spName = "";
            string isBEPortal = "Yes"; // While generating EAS file, to be taken care in the sp  sp_Fetch_BizPlan_SL_Account_V1_EAS
            if (sl == "EAS")
            {
                //spName = "sp_Fetch_BizPlan_SL_Account_V1_EAS_karthik";
                spName = "sp_Fetch_BizPlan_SL_Account_EAS_Online_FY26";
            }
            else
            {
                //spName = "sp_Fetch_BizPlan_SL_Account_V3";
                spName = "sp_Fetch_BizPlan_SL_Account_Online_FY26";
            }
            //SqlCommand cmdAccount = new SqlCommand() { CommandText = spName, CommandType = CommandType.StoredProcedure };
            //cmdAccount.Parameters.AddWithValue("@SL", sl);
            //cmdAccount.Parameters.AddWithValue("@isBEPortal", isBEPortal);

            SqlCommand cmdInvestment = new SqlCommand() { CommandText = "sp_Fetch_BizPlan_SL_Acc_Investment_Online", CommandType = CommandType.StoredProcedure };
            cmdInvestment.Parameters.AddWithValue("@SL", sl);

            SqlCommand cmdStartup = new SqlCommand() { CommandText = "sp_Fetch_BizPlan_Startup_New_Alliance_Ecosystem", CommandType = CommandType.StoredProcedure };
            cmdStartup.Parameters.AddWithValue("@SL", sl);

            SqlCommand cmdAssets = new SqlCommand() { CommandText = "sp_Fetch_BizPlan_Infosys_Assets_Plan", CommandType = CommandType.StoredProcedure };
            cmdAssets.Parameters.AddWithValue("@SL", sl);

            SqlCommand cmdPracticeSales = new SqlCommand() { CommandText = "SP_Bizplan_PracticeSales", CommandType = CommandType.StoredProcedure };
            cmdPracticeSales.Parameters.AddWithValue("@SL", sl);

            SqlCommand cmdDCG = new SqlCommand() { CommandText = "sp_Fetch_BizPlan_DCG_Product_Plan", CommandType = CommandType.StoredProcedure };
            cmdDCG.Parameters.AddWithValue("@SL", sl);

            SqlCommand cmdMaximus_Margin_Intervention = new SqlCommand() { CommandText = "sp_Fetch_BizPlan_Maximus_Margin_Intervention", CommandType = CommandType.StoredProcedure };
            cmdMaximus_Margin_Intervention.Parameters.AddWithValue("@SL", sl);

            SqlCommand cmdBPlan = new SqlCommand() { CommandText = spName, CommandType = CommandType.StoredProcedure };
            cmdBPlan.Parameters.AddWithValue("@SL", sl);
            cmdBPlan.Parameters.AddWithValue("@isBEPortal", isBEPortal);

            DataSet dsDCG = GetDataSet(cmdDCG);
            DataTable dtDCG = dsDCG.Tables[0];

            DataSet dsMaximus = GetDataSet(cmdMaximus_Margin_Intervention);
            DataTable dtMaximus = dsMaximus.Tables[0];


            DataSet dsInvestment = GetDataSet(cmdInvestment);
            DataTable dtInvestment = dsInvestment.Tables[0];

            DataSet dsAssets = GetDataSet(cmdAssets);
            DataTable dtAssets = dsAssets.Tables[0];

            DataSet dsStartup = GetDataSet(cmdStartup);
            DataTable dtStartup = dsStartup.Tables[0];

            DataSet dsBplan = GetDataSet(cmdBPlan);
            DataTable dtBPlan = dsBplan.Tables[0];

            DataSet dsPracticeSales = GetDataSet(cmdPracticeSales);
            DataTable dtPracticeSales = dsPracticeSales.Tables[0];


            string folder = "ExcelOperations";
            var myDir = new DirectoryInfo(Server.MapPath(folder));

            string filename = sl + "_Service_Line_BPlan_" + DateTime.Now.ToString("ddMMMyyyy_HHmm") + "IST.xlsx";

            String downloadFile = myDir + "\\" + filename;

            if (myDir.GetFiles().SingleOrDefault(k => k.Name == filename) != null)
            {
                System.IO.File.Delete(downloadFile);
            }

            Microsoft.Office.Interop.Excel.Application oExcel = null;
            Microsoft.Office.Interop.Excel.Workbook oBook = default(Microsoft.Office.Interop.Excel.Workbook);
            Microsoft.Office.Interop.Excel.Sheets ws = null;

            Excel.Worksheet sheet_DCG = null;
            Excel.Worksheet sheet_Maximus = null;
            Excel.Worksheet sheet_Investment = null;
            Excel.Worksheet sheet_Assets = null;
            Excel.Worksheet sheet_Startup = null;
            Excel.Worksheet sheet_BPlan = null;
            Excel.Worksheet sheet_PracticeSales = null;
            //added by ankita
            Excel.Worksheet sheet_ServiceLineSummary = null;
            VBIDE.VBComponent oModule;
            String sCode;
            Object oMissing = System.Reflection.Missing.Value;

            //instance of excel
            oExcel = new Microsoft.Office.Interop.Excel.Application();
            string templatePath = Server.MapPath(folder) + "\\Template\\" + "BizPlan_CDO_Template.xlsx" + "";
            oBook = oExcel.Workbooks.
                Open(templatePath, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            ws = oBook.Sheets;

            //sheet_Account = ws.Item["Account"];
            sheet_PracticeSales = ws.Item["Practice Sales"];
            sheet_Investment = ws.Item["Investment"];
            sheet_Assets = ws.Item["Infosys Assets Plan"];
            sheet_Startup = ws.Item["Startup&New Alliance Ecosystem"];
            sheet_BPlan = ws.Item["BPlan (Digital+Core)"];
            sheet_Maximus = ws.Item["Maximus Margin Intervention"];
            sheet_DCG = ws.Item["DCG Product Plan"];
            //Added by ankita
            //sheet_ServiceLineSummary = ws.Item["Service Line Summary"];

            FillExcelSheet(dtDCG, sheet_DCG);
            FillExcelSheet(dtMaximus, sheet_Maximus);
            FillInvestmentSales(dtPracticeSales, sheet_PracticeSales, sl);
            FillInvestmentSales(dtInvestment, sheet_Investment, sl);
            FillInvestmentSales(dtAssets, sheet_Assets, sl);
            FillInvestmentSales(dtStartup, sheet_Startup, sl);
            FillExcelSheet(dtBPlan, sheet_BPlan);
            //SqlCommand cmdUSRegion = new SqlCommand() { CommandText = "SP_Bizplan_CDO_USSubRegion_V1", CommandType = CommandType.StoredProcedure };
            //cmdUSRegion.Parameters.AddWithValue("@SL", sl);

            //DataSet dsRegion = GetDataSet(cmdUSRegion);
            //FillExcelSheetUSSubRegion(dsRegion, sheet_ServiceLineSummary);


            int hWnd1 = oExcel.Application.Hwnd;

            oExcel.DisplayAlerts = false;
            RefreshPivots(ws);




            oBook.SaveAs(downloadFile);

            oBook.Close(false, templatePath, null);



            //System.Runtime.InteropServices.Marshal.FinalReleaseComObject(sheet_Account);
            //sheet_Account = null;
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(sheet_Investment);
            sheet_Investment = null;

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ws);
            ws = null;
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oBook);
            oBook = null;

            oExcel.Quit();


            if (oExcel != null)
            {
                TryKillProcessByMainWindowHwnd(hWnd1);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();


            Session["key"] = filename;



            //iframe.Attributes.Add("src", "Download.aspx");

            ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "download", "document.getElementById('iframe').src = 'Download.aspx';", true);

            ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "myStopFunction", "myStopFunction()", true);
        }


        public void ApplyPermission(Microsoft.Office.Interop.Excel.Workbook oBook)
        {
            oBook.Activate();
            oBook.Permission.Enabled = true;
            oBook.Permission.RemoveAll();
            string strExpiryDate = DateTime.Now.AddDays(60).Date.ToString();
            DateTime dtTempDate = Convert.ToDateTime(strExpiryDate);
            DateTime dtExpireDate = new DateTime(dtTempDate.Year, dtTempDate.Month, dtTempDate.Day);
            UserPermission userper = oBook.Permission.Add("Everyone", MsoPermission.msoPermissionChange);
            userper.ExpirationDate = dtExpireDate;
        }
        public void FillExcelSheetUSSubRegion(DataSet ds, Microsoft.Office.Interop.Excel.Worksheet excel)
        {

            DataTable dtQtrRev = new DataTable();
            DataTable dtYearly_Rev = new DataTable();
            DataTable dtQonQGrowth = new DataTable();
            DataTable dtYonYGrowth = new DataTable();
            string[] selectedQtrRev = new[] {"USsubRegion","Q1 FY19 (A)","Q2 FY19 (A)","Q3 FY19 (A)",
            "Q4 FY19 (A)","Q1 FY20 (A)","Q2 FY20 (EST)","Q3 FY20 (EST)","Q4 FY20 (EST)",
            "Q1 FY21 (EST)","Q2 FY21 (EST)","Q3 FY21 (EST)","Q4 FY21 (EST)"};
            dtQtrRev = new DataView(ds.Tables[0]).ToTable(false, selectedQtrRev);

            string[] selectedYearly_Rev = new[] { "USsubRegion", "FY20 (EST)", "FY21 (Plan)", "FY22 (Plan)", "FY23 (Plan)" };
            dtYearly_Rev = new DataView(ds.Tables[0]).ToTable(false, selectedYearly_Rev);

            string[] selectedQonQGrowth = new[] { "USsubRegion", "Q2FY19 Vs Q1FY19 (%)", "Q3FY19 Vs Q2FY19 (%)", "Q4FY19 Vs Q3FY19 (%)",
            "Q1FY20 Vs Q4FY19 (%)","Q2FY20 Vs Q1FY20 (%)","Q3FY20 Vs Q2FY20 (%)","Q4FY20 Vs Q3FY20 (%)","Q1FY21 Vs Q4FY20 (%)",
            "Q2FY21 Vs Q1FY21 (%)","Q3FY21 Vs Q2FY21 (%)","Q4FY21 Vs Q3FY21 (%)"};
            dtQonQGrowth = new DataView(ds.Tables[0]).ToTable(false, selectedQonQGrowth);


            string[] selectedYonYGrowth = new[] { "USsubRegion", "FY19 Vs FY20 (%)", "FY21 Vs FY20 (%)", "FY22 Vs FY21 (%)", "FY23 Vs FY22 (%)" };
            dtYonYGrowth = new DataView(ds.Tables[0]).ToTable(false, selectedYonYGrowth);

            Microsoft.Office.Interop.Excel.Range cQtrRev1;
            Microsoft.Office.Interop.Excel.Range cQtrRev2;
            Microsoft.Office.Interop.Excel.Range range_QtrRev;
            Microsoft.Office.Interop.Excel.Range cYearly_Rev1;
            Microsoft.Office.Interop.Excel.Range cYearly_Rev2;
            Microsoft.Office.Interop.Excel.Range range_Yearly_Rev;
            Microsoft.Office.Interop.Excel.Range cQonQGrowth1;
            Microsoft.Office.Interop.Excel.Range cQonQGrowth2;
            Microsoft.Office.Interop.Excel.Range range_QonQGrowth;
            Microsoft.Office.Interop.Excel.Range cYonYGrowth1;
            Microsoft.Office.Interop.Excel.Range cYonYGrowth2;
            Microsoft.Office.Interop.Excel.Range range_YonYGrowth;




            cQtrRev1 = (Microsoft.Office.Interop.Excel.Range)excel.Cells[141, 2];
            cQtrRev2 = (Microsoft.Office.Interop.Excel.Range)excel.Cells[146, 14];
            range_QtrRev = excel.get_Range(cQtrRev1, cQtrRev2);
            range_QtrRev.Value2 = GetObjectArrayFromDT(dtQtrRev);


            cYearly_Rev1 = (Microsoft.Office.Interop.Excel.Range)excel.Cells[141, 18];
            cYearly_Rev2 = (Microsoft.Office.Interop.Excel.Range)excel.Cells[146, 22];
            range_Yearly_Rev = excel.get_Range(cYearly_Rev1, cYearly_Rev2);
            range_Yearly_Rev.Value2 = GetObjectArrayFromDT(dtYearly_Rev);


            cQonQGrowth1 = (Microsoft.Office.Interop.Excel.Range)excel.Cells[151, 2];
            cQonQGrowth2 = (Microsoft.Office.Interop.Excel.Range)excel.Cells[156, 13];
            range_QonQGrowth = excel.get_Range(cQonQGrowth1, cQonQGrowth2);
            range_QonQGrowth.Value2 = GetObjectArrayFromDT(dtQonQGrowth);

            cYonYGrowth1 = (Microsoft.Office.Interop.Excel.Range)excel.Cells[151, 18];
            cYonYGrowth2 = (Microsoft.Office.Interop.Excel.Range)excel.Cells[156, 22];
            range_YonYGrowth = excel.get_Range(cYonYGrowth1, cYonYGrowth2);
            range_YonYGrowth.Value2 = GetObjectArrayFromDT(dtYonYGrowth);


            ReleaseObject(cQtrRev1);
            ReleaseObject(cQtrRev2);
            ReleaseObject(range_QtrRev);
            ReleaseObject(cYearly_Rev1);
            ReleaseObject(cYearly_Rev2);
            ReleaseObject(range_Yearly_Rev);
            ReleaseObject(cQonQGrowth1);
            ReleaseObject(cQonQGrowth2);
            ReleaseObject(range_QonQGrowth);
            ReleaseObject(cYonYGrowth1);
            ReleaseObject(cYonYGrowth2);
            ReleaseObject(range_YonYGrowth);



        }

       

        private object[,] GetObjectArrayFromDT(DataTable dt)
        {
            object[,] rawData = new object[dt.Rows.Count, dt.Columns.Count];
            for (int col = 0; col < dt.Columns.Count; col++)
            {
                for (int row = 0; row < dt.Rows.Count; row++)
                {
                    // rawData[row + 1, col] = dt.Rows[row].ItemArray[col];
                    rawData[row, col] = dt.Rows[row][col];
                }
            }
            return rawData;
        }

        public void FillExcelSheet(DataTable dt, Microsoft.Office.Interop.Excel.Worksheet excel)
        {

            try
            {
                //// Copy the DataTable to an object array
                //object[,] rawData = new object[dt.Rows.Count, dt.Columns.Count];

                //// Copy the column names to the first row of the object array
                ////for (int col = 0; col < dt.Columns.Count; col++)
                ////{
                ////    rawData[0, col] = dt.Columns[col].ColumnName;
                ////}
                //// Copy the values to the object array

                //for (int col = 0; col < dt.Columns.Count; col++)
                //{
                //    for (int row = 0; row < dt.Rows.Count; row++)
                //    {
                //        // rawData[row + 1, col] = dt.Rows[row].ItemArray[col];
                //        rawData[row, col] = dt.Rows[row][col];
                //    }
                //}


                object[,] rawData = GetObjectArrayFromDT(dt);


                Microsoft.Office.Interop.Excel.Range c1;
                Microsoft.Office.Interop.Excel.Range c2;
                Microsoft.Office.Interop.Excel.Range range_excel;

                if (excel.Name == "DCG Product Plan" || excel.Name == "Biz Plan Revenues View 2" || excel.Name == "Biz Plan Revenues View3" || excel.Name == "Biz Plan Revenues View4" || excel.Name == "BPlan (Digital+Core)")
                {
                    c1 = (Microsoft.Office.Interop.Excel.Range)excel.Cells[3, 1];
                    c2 = (Microsoft.Office.Interop.Excel.Range)excel.Cells[dt.Rows.Count + 2, dt.Columns.Count];
                    range_excel = excel.get_Range(c1, c2);
                }
                else if (excel.Name == "Maximus Margin Intervention")
                {
                    c1 = (Microsoft.Office.Interop.Excel.Range)excel.Cells[17, 1];
                    c2 = (Microsoft.Office.Interop.Excel.Range)excel.Cells[dt.Rows.Count + +16, dt.Columns.Count];
                    range_excel = excel.get_Range(c1, c2);
                }
                else
                {
                    c1 = (Microsoft.Office.Interop.Excel.Range)excel.Cells[2, 1];
                    c2 = (Microsoft.Office.Interop.Excel.Range)excel.Cells[dt.Rows.Count + 1, dt.Columns.Count];
                    range_excel = excel.get_Range(c1, c2);
                }


                //Fill Array in Excel
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

       

        public void FillExcelSheet_Acc_Rev(DataTable dt, Microsoft.Office.Interop.Excel.Worksheet excel)
        {

            try
            {


                object[,] rawData = GetObjectArrayFromDT(dt);


                Microsoft.Office.Interop.Excel.Range c1;
                Microsoft.Office.Interop.Excel.Range c2;
                Microsoft.Office.Interop.Excel.Range range_excel;




                c1 = (Microsoft.Office.Interop.Excel.Range)excel.Cells[4, 2];
                c2 = (Microsoft.Office.Interop.Excel.Range)excel.Cells[dt.Rows.Count + 3, dt.Columns.Count + 1];
                range_excel = excel.get_Range(c1, c2);



                //Fill Array in Excel
                range_excel.Value2 = rawData;
                //range_excel.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                //range_excel.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid;
                //range_excel.Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);



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
        private string GetUserID()
        {
            string[] machineUsers = HttpContext.Current.User.Identity.Name.Split('\\');
            if (machineUsers.Length == 2)
                return machineUsers[1];
            return "";
        }
        public DataTable GetDataTable(SqlCommand cmd)
        {

            cmd.CommandTimeout = 160;
            SqlConnection con = new SqlConnection(connString);
            cmd.Connection = con;
            DataTable dt = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            try
            {
                da.Fill(dt);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return dt;
        }
        public DataSet GetDataSet(SqlCommand cmd)
        {

            cmd.CommandTimeout = int.MaxValue;
            SqlConnection con = new SqlConnection(connString);
            cmd.Connection = con;
            cmd.CommandTimeout = int.MaxValue;

            DataSet ds = new DataSet();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            try
            {
                da.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return ds;
        }
        protected void btnFinRefresh_Click(object sender, EventArgs e)
        {
            System.Threading.Thread.Sleep(5000);
        }


        protected void lnkSPOReport_Click(object sender, EventArgs e) // Internal Bizplan report
        {
            hidTAB.Value = "#menu2";

            string dataType = ddl_Type.SelectedValue;

            string Incl_Fluido_Simplus;

            //if (dataType == "Digital")
            //{
            //    cmdSPOReport = new SqlCommand() { CommandText = "sp_BizPlan_EAS_Report_digital_Online", CommandType = CommandType.StoredProcedure };
            //}
            //else
            //{
            //    cmdSPOReport = new SqlCommand() { CommandText = "sp_BizPlan_EAS_Report_Online_FY27_Q3_24", CommandType = CommandType.StoredProcedure };
            //}


            Incl_Fluido_Simplus = dataType == "Consolidated" ? "Yes" : "No";

            SqlCommand cmdSPOReport = new SqlCommand() { CommandText = "sp_Internal_BizPlan_EAS_Report_Online", CommandType = CommandType.StoredProcedure };
            cmdSPOReport.Parameters.AddWithValue("@Vertical", ddlVertical.SelectedItem.Text);
            cmdSPOReport.Parameters.AddWithValue("@PracticeLine", ddl_Internal.SelectedItem.Text);
            cmdSPOReport.Parameters.AddWithValue("@Incl_Fluido_Simplus", Incl_Fluido_Simplus);
            DataSet dsSPOReport = GetDataSet(cmdSPOReport);


            DataTable dtEAS_Biz_Plan_data = dsSPOReport.Tables[0];

            if (ddl_Internal.SelectedItem.Text != "EAS")
            {
                var columnName = File.ReadLines(Server.MapPath("~/Upload") + "\\EAS Column Name.txt");
                foreach (var lineRead in columnName)
                {
                    string text = lineRead;
                    dtEAS_Biz_Plan_data.Columns.Remove(text);
                }
                dtEAS_Biz_Plan_data.AcceptChanges();
            }


            string folder = "ExcelOperations";
            var myDir = new DirectoryInfo(Server.MapPath(folder));

            string filename;
            if (dataType == "Digital")
            {
                if (ddl_Internal.SelectedItem.Text == "EAS")
                {
                    filename = "BizPlan_EAS_data_digital_" + GetUserID() + "_" + DateTime.Now.ToString("ddMMMyyyy_HHmm") + "IST.xlsx";
                }
                else
                    filename = "BizPlan_" + ddl_Internal.SelectedItem.Text + "_digital_" + GetUserID() + "_" + DateTime.Now.ToString("ddMMMyyyy_HHmm") + "IST.xlsx";
            }
            else
            {
                if (ddl_Internal.SelectedItem.Text == "EAS")
                {
                    filename = "BizPlan_EAS_data_" + GetUserID() + "_" + DateTime.Now.ToString("ddMMMyyyy_HHmm") + "IST.xlsx";
                }
                else
                    filename = "BizPlan_" + ddl_Internal.SelectedItem.Text + "_data_" + GetUserID() + "_" + DateTime.Now.ToString("ddMMMyyyy_HHmm") + "IST.xlsx";
            }



            String downloadFile = myDir + "\\" + filename;

            if (myDir.GetFiles().SingleOrDefault(k => k.Name == filename) != null)
            {
                System.IO.File.Delete(downloadFile);
            }

            Microsoft.Office.Interop.Excel.Application oExcel = null;
            Microsoft.Office.Interop.Excel.Workbook oBook = default(Microsoft.Office.Interop.Excel.Workbook);
            Microsoft.Office.Interop.Excel.Sheets ws = null;

            Excel.Worksheet EAS_Biz_Plan_data = null;



            VBIDE.VBComponent oModule;
            String sCode;
            Object oMissing = System.Reflection.Missing.Value;

            //instance of excel
            oExcel = new Microsoft.Office.Interop.Excel.Application();
            // string templatePath = Server.MapPath(folder) + "\\Template\\" + "BizPlan_EAS_data.xlsx" + ""; // TODO: karthik
            string templatePath;
            if (dataType == "Digital")
            {
                if (ddl_Internal.SelectedItem.Text == "EAS")
                {
                    templatePath = Server.MapPath(folder) + "\\Template\\" + "BizPlan_EAS_data_digital_template.xlsx" + "";
                }
                else { templatePath = Server.MapPath(folder) + "\\Template\\" + "BizPlan_SL_data_template.xlsx" + ""; }
            }
            else
            {
                if (ddl_Internal.SelectedItem.Text == "EAS")
                {
                    templatePath = Server.MapPath(folder) + "\\Template\\" + "BizPlan_EAS_data_template.xlsx" + "";
                }
                else
                    templatePath = Server.MapPath(folder) + "\\Template\\" + "BizPlan_SL_data_template.xlsx" + "";
            }


            oBook = oExcel.Workbooks.
                Open(templatePath, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            ws = oBook.Sheets;

            EAS_Biz_Plan_data = ws.Item["EAS_Biz_Plan_data"];

            if (ddl_Internal.SelectedItem.Text == "EAS")
            {
                Excel.Worksheet BizPlanRevenuesView2 = null;
                Excel.Worksheet BizPlanRevenuesView3 = null;
                Excel.Worksheet BizPlanRevenuesView4 = null;
                Excel.Worksheet BizPlandata2 = null;
                Excel.Worksheet BizPlandata3 = null;
                //Excel.Worksheet BizPlandata3_Volume = null;
                //Excel.Worksheet BizPlandata3_Margin = null;
                //Excel.Worksheet BizSummaryDeck2 = null;//TODO:Karthik
                //Excel.Worksheet BizSummaryDeck1 = null;//TODO:Karthik



                DataTable dtBizPlanRevenuesView2 = dsSPOReport.Tables[1];
                DataTable dtBizPlanRevenuesView3 = dsSPOReport.Tables[2];
                DataTable dtBizPlanRevenuesView4 = dsSPOReport.Tables[3];
                DataTable dtBizPlanData_2 = dsSPOReport.Tables[4];
                DataTable dtBizPlanData_3 = dsSPOReport.Tables[5];//Revenue
                //DataTable dtBizPlanData_4 = dsSPOReport.Tables[6];//maring
                //DataTable dtBizPlanData_5 = dsSPOReport.Tables[7];//Volume
                // DataTable dtSummaryDeck2_QtrylyTrend = dsSPOReport.Tables[6];//TODO:Karthik
                // DataTable dtSummaryDeck2_ClientCat = dsSPOReport.Tables[7];//TODO:Karthik
                // DataTable dtSummaryDeck2_Digital = dsSPOReport.Tables[8];//TODO:Karthik

                //int index = 9;
                //DataTable dtSummaryDeck2Actuals = dsSPOReport.Tables[index + 0];
                //DataTable dtSummaryDeck2Region = dsSPOReport.Tables[index + 1];
                //DataTable dtSummaryDeck2Vertical = dsSPOReport.Tables[index + 2];
                //DataTable dtSummaryDeck2RevGrowth = dsSPOReport.Tables[index + 3];
                //DataTable dtSummaryDeck2DigRevGrowth = dsSPOReport.Tables[index + 4];
                //DataTable dtSummaryDeck2DigRev = dsSPOReport.Tables[index + 5];
                //DataTable dtSummaryDeck2VerticalGrouping = dsSPOReport.Tables[index + 6];

                //dtSummaryDeck2Actuals.TableName = SummaryDeckTableNames.SummaryDeck2Actuals;
                //dtSummaryDeck2Region.TableName = SummaryDeckTableNames.SummaryDeck2Region;
                //dtSummaryDeck2Vertical.TableName = SummaryDeckTableNames.SummaryDeck2Vertical;
                //dtSummaryDeck2RevGrowth.TableName = SummaryDeckTableNames.SummaryDeck2RevGrowth;
                //dtSummaryDeck2DigRevGrowth.TableName = SummaryDeckTableNames.SummaryDeck2DigRevGrowth;
                //dtSummaryDeck2DigRev.TableName = SummaryDeckTableNames.SummaryDeck2DigRev;
                //dtSummaryDeck2VerticalGrouping.TableName = SummaryDeckTableNames.SummaryDeck2VerticalGrouping;

                //DataTable[] dtSummaryDeck2Tables = { dtSummaryDeck2Actuals, dtSummaryDeck2DigRev, dtSummaryDeck2DigRevGrowth, dtSummaryDeck2Region, dtSummaryDeck2RevGrowth, dtSummaryDeck2Vertical, dtSummaryDeck2VerticalGrouping };

                BizPlanRevenuesView2 = ws.Item["Biz Plan Revenues View 2"];

                if (!(dataType == "Digital"))
                {
                    BizPlanRevenuesView3 = ws.Item["Biz Plan Revenues View3"];
                }

                BizPlanRevenuesView4 = ws.Item["Biz Plan Revenues View4"];
                BizPlandata2 = ws.Item["Data_2"];
                BizPlandata3 = ws.Item["Data_3"];
                //BizPlandata3_Volume = ws.Item["Data_3_Volume"];
                //BizPlandata3_Margin = ws.Item["Data_3_Margin"];
                //BizSummaryDeck2 = ws.Item["SummaryDeck2"];
                //BizSummaryDeck1 = ws.Item["SummaryDeck1"];

                FillExcelSheet(dtBizPlanRevenuesView2, BizPlanRevenuesView2);
                if (!(dataType == "Digital"))
                {
                    FillExcelSheet(dtBizPlanRevenuesView3, BizPlanRevenuesView3);
                }

                FillExcelSheet(dtBizPlanRevenuesView4, BizPlanRevenuesView4);
                FillExcelSheet(dtBizPlanData_2, BizPlandata2);
                FillExcelSheet(dtBizPlanData_3, BizPlandata3);
                //FillExcelSheet(dtBizPlanData_4, BizPlandata3_Margin);
                //FillExcelSheet(dtBizPlanData_5, BizPlandata3_Volume);

                //FillExcelSheet_SummaryDeck2(dtSummaryDeck2_QtrylyTrend, dtSummaryDeck2_ClientCat, dtSummaryDeck2_Digital, BizSummaryDeck2);//TODO:Karthik
                //FillExcelSheet_SummaryDeck1(dtSummaryDeck2Tables, BizSummaryDeck1);

                FillExcelSheet(dtEAS_Biz_Plan_data, EAS_Biz_Plan_data);

                oExcel.DisplayAlerts = false;

                RefreshPivots(ws);


                oBook.SaveAs(downloadFile);

                oBook.Close(false, templatePath, null);


                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(EAS_Biz_Plan_data);
                EAS_Biz_Plan_data = null;


                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(BizPlanRevenuesView2);
                BizPlanRevenuesView2 = null;
                if (!(dataType == "Digital"))
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(BizPlanRevenuesView3);
                    BizPlanRevenuesView3 = null;
                }
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(BizPlanRevenuesView4);
                BizPlanRevenuesView4 = null;
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(BizPlandata2);
                BizPlandata2 = null;

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(BizPlandata3);
                BizPlandata3 = null;

                //System.Runtime.InteropServices.Marshal.FinalReleaseComObject(BizPlandata3_Volume);
                //BizPlandata3_Volume = null;

            }
            else
            {
                FillExcelSheet(dtEAS_Biz_Plan_data, EAS_Biz_Plan_data);
                oExcel.DisplayAlerts = false;
                RefreshPivots(ws);
                oBook.SaveAs(downloadFile);
                oBook.Close(false, templatePath, null);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(EAS_Biz_Plan_data);
                EAS_Biz_Plan_data = null;
            }

            int hWnd1 = oExcel.Application.Hwnd;

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ws);
            ws = null;
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oBook);
            oBook = null;

            oExcel.Quit();
            if (oExcel != null)
            {
                TryKillProcessByMainWindowHwnd(hWnd1);
            }


            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();


            Session["key"] = filename;

            //loading.Style.Add("visibility", "visible");
            //lbl.Text = "Downloaded";
            //up.Update();

            //iframe.Attributes.Add("src", "Download.aspx");
            //ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "download", "document.getElementById('iframe').src = 'Download.aspx';", true);
            ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "abc", "download1();myStopFunction();", true);
            //ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "myStopFunction", "()", true);
        }




        public void RefreshPivots(Microsoft.Office.Interop.Excel.Sheets excelsheets)
        {

            foreach (Microsoft.Office.Interop.Excel.Worksheet pivotSheet in excelsheets)
            {
                Microsoft.Office.Interop.Excel.PivotTables pivotTables = (Microsoft.Office.Interop.Excel.PivotTables)pivotSheet.PivotTables();
                int pivotTablesCount = pivotTables.Count;
                if (pivotTablesCount > 0)
                {
                    for (int i = 1; i <= pivotTablesCount; i++)
                    {
                        Microsoft.Office.Interop.Excel.PivotTable pivotTable = pivotTables.Item(i);
                        pivotTable.RefreshTable();
                    }
                }
            }
        }

        public List<string> GetSUForuser(string userid)
        {


            List<string> lstempCollection = new List<string>();

            try
            {




                SqlParameter objBE = new SqlParameter();
                objBE.ParameterName = "@userid";
                objBE.Direction = ParameterDirection.Input;
                objBE.SqlDbType = SqlDbType.VarChar;
                objBE.Value = userid;

                SqlCommand objCommand = new SqlCommand() { CommandText = "spBEGetSU_dummy", CommandType = CommandType.StoredProcedure };
                SqlParameterCollection objParamColl = objCommand.Parameters;
                objParamColl.Add(objBE);

                DataSet ds = GetDataSet(objCommand);



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



            return lstempCollection;
        }

        protected void ddl_Internal_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ddl_Internal.Text != "EAS")
            {
                ddlVertical.SelectedIndex = 0;
                div_vertical.Visible = false;
            }
            else
                div_vertical.Visible = true;
        }


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




        protected void btnConsolidatedDownload_Click(object sender, EventArgs e)
        {
            string sl = ddlSLConsolidated.SelectedValue;

            SqlConnection con = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand() { CommandText = "sp_bizplan_Consolidate_InputFileData", CommandType = CommandType.StoredProcedure };
            cmd.Parameters.AddWithValue("@SL", sl); 
            cmd.Connection = con;
            SqlDataAdapter sdr = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            sdr.Fill(ds);
            List<string> listSheets = new List<string>();
            listSheets.Add("");
            string[] sheets = ds.Tables[0].Rows.OfType<DataRow>().Select(k => k[0] + "").ToArray();
            listSheets.AddRange(sheets);



            string fileName = sl + "_Consolidated_Input_" + DateTime.Now.ToString("ddMMMyyyy_hhmmtt");

            using (XLWorkbook wb = new XLWorkbook())
            {

                for (int i = 1; i < ds.Tables.Count; i++)
                {
                    string sheetName = listSheets[i];
                    wb.Worksheets.Add(ds.Tables[i], sheetName);
                }


                Response.Clear();
                Response.Buffer = true;
                Response.Charset = "";
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename="+ fileName + ".xlsx");
                using (MemoryStream MyMemoryStream = new MemoryStream())
                {
                    wb.SaveAs(MyMemoryStream);
                    MyMemoryStream.WriteTo(Response.OutputStream);
                    Response.Flush();
                    Response.End();
                }
            }




        }
    }
}


