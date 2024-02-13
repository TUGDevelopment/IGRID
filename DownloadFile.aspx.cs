using System;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Text;
using System.Linq;
using ClosedXML.Excel;

public partial class _Default : System.Web.UI.Page
{
	string constr = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
	MyDataModule cs = new MyDataModule();
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsCallback && !IsPostBack)
        {
 
				download();  
 

        }
    }
   void download()
    {
        string strQuery = "spGetTrackingReport_XML";
        SqlCommand cmd = new SqlCommand(strQuery);
        DataTable dt = spGetData(cmd);
        using (XLWorkbook wb = new XLWorkbook())
        {
            wb.Worksheets.Add(dt, "Customers");

            Response.Clear();
            Response.Buffer = true;
            Response.Charset = "";
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;filename=SqlExport.xlsx");
            using (MemoryStream MyMemoryStream = new MemoryStream())
            {
                wb.SaveAs(MyMemoryStream);
                MyMemoryStream.WriteTo(Response.OutputStream);
                Response.Flush();
                Response.End();
            }
        }
        //string _name = DateTime.Now.ToString("yyyyMMddHHmmss");
        //Response.ContentType = "application/vnd.ms-excel";
        //Response.AppendHeader("Content-Disposition", "attachment; filename=DataFile_"+ _name + ".xlsx");
        //Response.TransmitFile(Server.MapPath("~/ExcelFiles/DataFile.xlsx"));
        //Response.End();
    }
     
    public DataTable spGetData(SqlCommand cmd)
    {
        DataTable dt = new DataTable();
        SqlConnection con = new SqlConnection(constr);
        SqlDataAdapter sda = new SqlDataAdapter();
        cmd.CommandType = CommandType.StoredProcedure;
        cmd.Connection = con;
        try
        {
            con.Open();
            sda.SelectCommand = cmd;
            sda.Fill(dt);
            return dt;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        finally
        {
            con.Close();
            sda.Dispose();
            con.Dispose();
        }
    }
//    protected void ASPxGridView1_DataBinding(object sender, EventArgs e)
//    {
//        string strQuery = "select *" +
//             " from tblncp";
//        SqlCommand cmd = new SqlCommand(strQuery);
//        DataTable dt = cs.GetData(cmd);
//        ASPxGridView1.DataSource = dt;
//    }
    
}
