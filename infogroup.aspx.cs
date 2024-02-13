using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class infogroup : System.Web.UI.Page
{
    MyDataModule myservice = new MyDataModule();
    protected void Page_Load(object sender, EventArgs e)
    {
        SqlParameter[] param = { new SqlParameter("@where", "") };
        DataTable table = new DataTable();
        table = myservice.GetRelatedResources("spinfogroup", param);
        ExportExcel(table);
    }

    public void ExportExcel(DataTable DtProfile)
    {
        using (XLWorkbook wb = new XLWorkbook())
        {
            wb.Worksheets.Add(DtProfile, "Profile");
            Response.Clear();
            Response.Buffer = true;
            Response.Charset = "";
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;filename=dataExport.xlsx");
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