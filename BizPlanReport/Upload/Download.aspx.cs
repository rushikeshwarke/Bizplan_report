using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;

namespace NSOBEReport
{
    public partial class Download : System.Web.UI.Page
    {

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session["key"] != null)
            {
                string fianlname = Session["key"].ToString();
                string folder = "ExcelOperations";                
                var MyDir = new DirectoryInfo(Server.MapPath(folder)); 
                string sPath = Server.MapPath(folder+"\\");
                Session["key"] = null;

                if (MyDir.GetFiles().SingleOrDefault(k => k.Name == fianlname) != null)
                {
                    Response.AppendHeader("Content-Disposition", "attachment; filename=" + fianlname);
                    Response.TransmitFile(sPath + fianlname);
                    Response.End();
                }
            }
        }
    }
}