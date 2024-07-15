using BEBIZ4.Common;
using DocumentFormat.OpenXml.Office2013.PowerPoint.Roaming;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;

namespace BEBIZ.Upload.Sheets
{
    public abstract class BaseSheet : ISheetInfo
    {
        public string SheetName { get; set; }
        public string SL { get; set; }
        public string[] allSheets { get; set; }
        protected string TableName { get; set; }
        protected Dictionary<string, string> DictMapping = new Dictionary<string, string>();
        protected Dictionary<string, EnumLookUp> DictLookUpMapping = new Dictionary<string, EnumLookUp>(); // column, enum
        protected string[] MandatoryColumns { get; set; }
        protected string[] DuplicateColumns { get; set; }
        protected string[] NumericColumns { get; set; }
        public DataTable Data { get; set; }
        protected string[] Columns { get; set; }
        protected int RowNumber { get; set; } //USED FOR DISPLAYING THE ERROR MESSAGE. pLEASE DONT GET CONFUSED.
        protected string whereCondition { get; set; }
        string DBConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["DBConnectionString"].ConnectionString;

        public BaseSheet(DataSet ds,string sl)
        {
            SL = sl;
            DictMapping.Clear();
            Init();
            Data = ds.Tables[SheetName + "$"];
            FormatData();
        }

        public BaseSheet(DataSet ds)
        {

            DictMapping.Clear();
            Init();
            Data = ds.Tables[SheetName + "$"];
            FormatData();
        }


        protected abstract void Init();

        protected virtual void FormatData()
        {

            //for (int i = Data.Rows.Count - 1; i >= 0; i--)
            //{
            //    var row = Data.Rows[i];
            //    if (MandatoryColumns.All(k => isCellEmpty(row[k])))
            //        Data.Rows.RemoveAt(i);
            //}
            //Data.AcceptChanges();
            //if (Data != null)
            {
                Data.RemoveEmptyCols();
                Data.RemoveEmptyRows();
            }
        }

        public List<ErrorEntity> Validate()
        {
            List<ErrorEntity> lst = new List<ErrorEntity>();
           // if (Data == null) return lst;
            
                if (Data.Rows.Count == 0)
                    return lst;
            

            lst.AddRange(ValidateColumns());
            lst.AddRange(ValidateMandatory());
            lst.AddRange(ValidateDuplicate());
            lst.AddRange(ValidateNumeric());
            lst.AddRange(ValidateServiceLine());
            lst.AddRange(ValidateMCC());
            lst.AddRange(LookUpValidation());

            return lst;

        }


       
        List<ErrorEntity> LookUpValidation()
        {
            List<ErrorEntity> lst = new List<ErrorEntity>();


            for (int i = 0; i < Data.Rows.Count; i++)
            {
                foreach (var item in DictLookUpMapping)
                {
                    string value = Data.Rows[i][item.Key].ToString().Trim();
                    var lookUp = LookUpFinder.GetLookUpItems(item.Value);
                    if (!lookUp.Contains(value))
                        lst.Add(new ErrorEntity() { RowNo = RowNumber + i, ErrType = ErrorType.InvalidLookUpData, SheetName = SheetName, Message = " [" + value + "] is not a valid LookUp Item" });
                }
            }
            return lst;

        }

        List<ErrorEntity> ValidateServiceLine()
        {
            List<ErrorEntity> lst = new List<ErrorEntity>();
            for (int i = 0; i < Data.Rows.Count; i++)
            {
                if(SheetName!= "Startup&New Alliance Ecosystem" && SheetName!= "Infosys Assets Plan" && SheetName != "DCG Product Plan" && SheetName != "Maximus - Margin Intervention") {
                    if (Data.Rows[i]["Service Line"].ToString().Trim().ToLower() != SL.ToLower().Trim())
                        lst.Add(new ErrorEntity() { RowNo = RowNumber + i, ErrType = ErrorType.InvalidData, SheetName = SheetName, Message = "Service Line should be [" + SL + "]" });
                }

            }
            return lst;

        }
        List<ErrorEntity> ValidateColumns()
        {
            

            List<ErrorEntity> lst = new List<ErrorEntity>();
            var xlColumns = Data.Columns.OfType<DataColumn>().Select(k => k.ColumnName).ToList();

            DataTable dt = new DataTable();
            dt.Columns.Add("ColumnNames", typeof(string));
            Debugger("xl columns");
            foreach (string col in xlColumns)
            {
                dt.Rows.Add(col);
                Debugger(col);
            }
            Debugger("DictMapping columns");
            foreach (var item in DictMapping.Keys)
            {
                Debugger(item);

                if (!xlColumns.Contains(item))
                    lst.Add(new ErrorEntity() { RowNo = -1, ErrType = ErrorType.InvalidColumn, SheetName = SheetName, Message = " columns [ " + item + " ] is missing" });
            }
            return lst;

        }

        public virtual List<string> GetMCC()
        {
            List<string> lstMCCList = GetMccList();
            return lstMCCList;
        }


        List<ErrorEntity> ValidateMCC()
        {
            List<ErrorEntity> lst = new List<ErrorEntity>();
            List<string> lstMCCList = GetMCC();
            string columnName = "Master Customer Code";
            var xlColumns = Data.Columns.OfType<DataColumn>().Select(k => k.ColumnName).ToList();

            if (xlColumns.Contains(columnName))
            {
                for (int i = 0; i < Data.Rows.Count; i++)
                {
                    string mcc = Data.Rows[i][columnName].ToString().ToLower();
                    if (!lstMCCList.Contains(mcc))
                        lst.Add(new ErrorEntity() { RowNo = RowNumber + i, ErrType = ErrorType.Others, SheetName = SheetName, Message = mcc + "  is not a valid Master Customer Code" });
                }
            }
            return lst;

        }

        private List<string> GetMccList()
        {
            DataTable dt = new DataTable();
            SqlConnection sqlconn = new SqlConnection(DBConnectionString);
            SqlCommand sqlcmd = new SqlCommand("sp_BizPlan_Validate_Get_MCCList", sqlconn);
            sqlcmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter da = new SqlDataAdapter(sqlcmd);
            da.Fill(dt);
            List<string> lst = dt.Rows.OfType<DataRow>().Select(k => (k[0] + "").Trim().ToLower()).ToList();
            return lst;
        }

        //List<ErrorEntity> ValidateMCC()
        //{
        //    List<ErrorEntity> lst = new List<ErrorEntity>();
        //    try
        //    {

        //        SqlConnection sqlconn = new SqlConnection(DBConnectionString);
        //        var xlColumns = Data.Columns.OfType<DataColumn>().Select(k => k.ColumnName).ToList();
        //        string item = "Master Customer Name";
        //        if (xlColumns.Contains(item))
        //        {
        //            string present;
        //            string mcc;
        //            for (int i = 0; i < Data.Rows.Count; i++)
        //            {
        //                mcc = Data.Rows[i]["Master Customer Name"].ToString();
        //                SqlCommand sqlcmd = new SqlCommand("BizPlan_Validate_MCC", sqlconn);
        //                sqlcmd.CommandType = CommandType.StoredProcedure;
        //                sqlcmd.Parameters.AddWithValue("@masterCustomerName", mcc);
        //                sqlcmd.Parameters.Add("@present", SqlDbType.VarChar, 100);
        //                sqlcmd.Parameters["@present"].Direction = ParameterDirection.Output;
        //                sqlconn.Open();
        //                sqlcmd.ExecuteNonQuery();
        //                present = sqlcmd.Parameters["@present"].Value.ToString();
        //                sqlconn.Close();
        //                if (present == "No")
        //                    lst.Add(new ErrorEntity() { RowNo = RowNumber + i, ErrType = ErrorType.Others, SheetName = SheetName, Message = mcc + " Master Customer Name is not present" });

        //            }
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine(e.Message);
        //    }

        //    return lst;



        //}

        List<ErrorEntity> ValidateNumeric()
        {
            string numericformat = "columns [ {0} ] contains a Non Numeric value";
            List<ErrorEntity> lst = new List<ErrorEntity>();
            for (int i = 0; i < Data.Rows.Count; i++)
            {
                var row = Data.Rows[i];
                List<string> lstCols = new List<string>();
                foreach (DataColumn col in Data.Columns)
                    if (NumericColumns.Contains(col.ColumnName))
                        if (!isValidNumeric(row[col.ColumnName]))
                            lstCols.Add(col.ColumnName);

                if (lstCols.Count > 0)
                    lst.Add(new ErrorEntity() { RowNo = RowNumber + i, Message = string.Format(numericformat, string.Join(",", lstCols)), SheetName = SheetName, ErrType = ErrorType.NonNumeric });
            }
            return lst;
        }




        List<ErrorEntity> ValidateMandatory()
        {

            Dictionary<int, string> dictMandatory = new Dictionary<int, string>();
            List<ErrorEntity> lst = new List<ErrorEntity>();

            DataTable dt = Data.Copy();
            int rowNo;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                List<string> lstMandatoryCols = new List<string>();
                rowNo = RowNumber + i;
                foreach (DataColumn col in dt.Columns)
                    if (MandatoryColumns.Contains(col.ColumnName))
                        if (string.IsNullOrWhiteSpace(dt.Rows[i][col].ToString().Trim()))
                            lstMandatoryCols.Add(col.ColumnName);

                if (lstMandatoryCols.Count > 0)
                    dictMandatory.Add(rowNo, string.Join(",", lstMandatoryCols));

            }

            string Mandatoryformat = "Columns [{0}] are mandatory.";
            foreach (var item in dictMandatory)
                lst.Add(new ErrorEntity() { RowNo = item.Key, ErrType = global::ErrorType.Mandatory, SheetName = SheetName, Message = string.Format(Mandatoryformat, item.Value) });

            return lst;
        }

        public void Debugger(string info)
        {
            string file = System.Configuration.ConfigurationManager.AppSettings["Debugger"];

            File.AppendAllText(file, info + Environment.NewLine);

        }

        List<ErrorEntity> ValidateDuplicate()
        {
            Debugger("278");

            List<ErrorEntity> lst = new List<ErrorEntity>();
            var dtNew = Data.Copy();

            dtNew.Columns.OfType<DataColumn>().Select(k => k.ColumnName).ToList().Except(DuplicateColumns).ToList().ForEach(k => dtNew.Columns.Remove(k));
            dtNew.AcceptChanges();
            Debugger("285");

            for (int i = 0; i < DuplicateColumns.Length; i++)
            {
                Debugger(DuplicateColumns[i]);
            }

            Debugger("a");
            dtNew.Columns.OfType<DataColumn>().Select(k => k.ColumnName).ToList().ForEach(k => Debugger(k));

            Debugger("b");


            for (int i = 0; i < DuplicateColumns.Length; i++)
            {
                try
                {
                    Debugger(DuplicateColumns[i]);
                    dtNew.Columns[DuplicateColumns[i]].SetOrdinal(i);
                }
                catch (Exception ex)
                {
                    Debugger("--");
                    Debugger(DuplicateColumns[i]);

                }
            }

            Debugger("289");

            List<string> lstAll = new List<string>();


            dtNew.AsEnumerable().ToList().ForEach(k => lstAll.Add(string.Join("", k.ItemArray)));
            var duplicateItems = lstAll.GroupBy(k => k).Where(k => k.Count() > 1).Select(k => k.Key).ToList();
            string Duplicateformat = "Row {0} is duplicated in {1}.";
            Debugger("297");
            foreach (string item in duplicateItems)
            {
                List<int> lstDuplicateRows = new List<int>();
                for (int i = 0; i < dtNew.Rows.Count; i++)
                {
                    if (item == string.Join("", dtNew.Rows[i].ItemArray))
                        lstDuplicateRows.Add(i + 1);
                }
                lstDuplicateRows = lstDuplicateRows.OrderBy(k => k).ToList();
                lstDuplicateRows = lstDuplicateRows.Select(k => k + RowNumber).ToList();

                lst.Add(new ErrorEntity()
                {
                    RowNo = lstDuplicateRows.First(),
                    Message = string.Format(Duplicateformat, lstDuplicateRows.First(), string.Join(",", lstDuplicateRows.Skip(1))),
                    SheetName = SheetName,
                    ErrType = ErrorType.Duplicate
                });

            }
            return lst;
        }

        protected virtual void ReplaceNullWithZeroInDT(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                foreach (string col in NumericColumns)
                {
                    object val = row[col];
                    if (val == null || val == DBNull.Value)
                        val = 0;
                    row[col] = val;
                }
            }
        }


        public virtual void Save()
        {
            try
            {
                var dt = Data.Copy();
                ReplaceNullWithZeroInDT(dt);


                SqlConnection con = new SqlConnection(DBConnectionString);
                if (string.IsNullOrWhiteSpace(whereCondition))
                    if(TableName== "BizPlan_Infosys_Assets_Plan" || TableName== "BizPlan_Startup_New_Alliance_Ecosystem")
                        whereCondition = string.Format(" Sub_unit = '{0}'", SL);
                    else if (TableName == "bizplan_DCG_Product_Plan")
                    {
                        whereCondition = string.Format(" [DCG Unit] = '{0}'", SL);
                    }
                    else if (TableName == "bizplan_Maximus_OM")
                    {
                        whereCondition = string.Format(" Subunit = '{0}'", SL);
                    }
                    else
                        whereCondition = string.Format(" SL = '{0}'", SL);
                else
                    if (TableName == "BizPlan_Infosys_Assets_Plan" || TableName == "BizPlan_Startup_New_Alliance_Ecosystem")
                        whereCondition += string.Format(" and  Sub_unit = '{0}'", SL);
                    else
                        whereCondition += string.Format(" and  SL = '{0}'", SL);

                //TODO: Ankita  take care..
                string queryToMoveDataToAudit = string.Format("insert into {0}_Audit select *, getdate() from {0} where {1} ;", TableName, whereCondition);
                string queryToTruncatetheTable = string.Format("Delete from {0} where {1};", TableName, whereCondition);
                string query = queryToMoveDataToAudit + queryToTruncatetheTable;
                SqlCommand cmdToMoveDataToAudit = new SqlCommand(query, con);
                cmdToMoveDataToAudit.CommandType = CommandType.Text;

                SqlBulkCopy bulkCopy = new SqlBulkCopy(con, SqlBulkCopyOptions.TableLock | SqlBulkCopyOptions.FireTriggers | SqlBulkCopyOptions.UseInternalTransaction, null);
                bulkCopy.DestinationTableName = TableName;
                bulkCopy.BulkCopyTimeout = 1700;
                Debugger("------------------");
                foreach (var item in DictMapping)
                {
                    bulkCopy.ColumnMappings.Add(item.Key, item.Value);
                    Debugger(item.Key + "~" + item.Value);
                }

                Debugger("rows count " + dt.Rows.Count.ToString());

                con.Open();
                cmdToMoveDataToAudit.ExecuteNonQuery();
                bulkCopy.WriteToServer(dt);
                con.Close();

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + "["+  SheetName+"]" ,ex);
            }

        }



        bool isValidNumeric(object value)
        {
            double d;
            string temp = (value + "").Trim();
            if (string.IsNullOrWhiteSpace(temp))
                return true;


            temp = temp.Replace("%", "");
            temp = temp.Replace("(", "");
            temp = temp.Replace(")", "");




            bool isValid = double.TryParse(temp, out d);



            return isValid;

        }
        bool isCellEmpty(object value)
        {
            string temp = (value + "").Trim();
            return string.IsNullOrWhiteSpace(temp);

        }
    }
}