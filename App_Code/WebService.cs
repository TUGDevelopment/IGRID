using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.ServiceModel;
using System.ServiceModel.Configuration;
using System.Web;
using System.Web.Script.Services;
using System.Web.Services;
using System.Xml.Serialization;
using System.Text;
using SAP.Middleware.Connector;
using System.Xml;
using ClosedXML.Excel;
using ClosedXML_Test;
using System.Diagnostics;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.ExtendedProperties;
//using Excel = Microsoft.Office.Interop.Excel;
/// <summary>
/// Summary description for WebService
/// </summary>
[WebService(Namespace = "http://tempuri.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
// [System.Web.Script.Services.ScriptService]
public class WebService : System.Web.Services.WebService
{
    static readonly string rootFolder = HttpContext.Current.Server.MapPath(@"~/ExcelFiles/");
    public string strConn = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
    public string CurUserName = HttpContext.Current.User.Identity.Name.Replace(@"THAIUNION\", @"");
    MyDataModule cs = new MyDataModule();
    public WebService()
    {

        //Uncomment the following line if using designed components 
        //InitializeComponent(); 
    }

    static DataTable GetTable()
    {
        DataTable table = new DataTable("Test");//DataTable with name - works fine
        //table.Columns.Add("Id", typeof(int));
        //table.Columns.Add("Condition", typeof(string));
        //table.Columns.Add("DocumentNo", typeof(string));
        //table.Columns.Add("Material", typeof(string));
        //table.Columns.Add("MaterialGroup", typeof(string));
        //table.Columns.Add("Brand", typeof(string));
        //table.Columns.Add("Description", typeof(string));
        //table.Columns.Add("RequestType", typeof(string));
        //table.Columns.Add("FinalInfoGroup", typeof(string));
        //table.Columns.Add("CreateOn", typeof(DateTime));
        //table.Columns.Add("CreateBy", typeof(string));
        //table.Columns.Add("Assign", typeof(string));
        //table.Columns.Add("Title", typeof(string));
        //table.Columns.Add("ReferenceMaterial", typeof(string));
        //table.Columns.Add("Vendor", typeof(string));
        //table.Columns.Add("VendorDescription", typeof(string));

        table.Columns.Add("Id", typeof(int));
        table.Columns.Add("Condition", typeof(string));
        table.Columns.Add("RequestType", typeof(string));
        table.Columns.Add("DocumentNo", typeof(string));
        table.Columns.Add("DMS No./ Artwork", typeof(string));
        table.Columns.Add("Material No.", typeof(string));
        table.Columns.Add("Description", typeof(string));
        table.Columns.Add("Group", typeof(string));
        table.Columns.Add("Brand", typeof(string));
        table.Columns.Add("Assignee(PG Name)", typeof(string));
        table.Columns.Add("CreateOn", typeof(string));
        table.Columns.Add("ActiveBy(PA Name)", typeof(string));
        table.Columns.Add("FinalInfoGroup", typeof(string));
        table.Columns.Add("ReferenceMaterial", typeof(string));
        table.Columns.Add("Vendor", typeof(string)); 
        table.Columns.Add("Vendor description", typeof(string));

        // Here we add five DataRows.
        return table;
    }
 
    [WebMethod]
    public void massinfogroup()
    {
        //Save the uploaded Excel file.
        string[] files = Directory.GetFiles(rootFolder);
        foreach (string _file in files)
        {
            //string filePath = HttpContext.Current.Server.MapPath(@"~/ExcelFiles/VK11_20211208142528.xlsx");
            //Open the Excel file using ClosedXML.
            using (XLWorkbook workBook = new XLWorkbook(_file))
            {
                //Read the first Sheet from Excel file.
                IXLWorksheet workSheet = workBook.Worksheet(3);

                //Create a new DataTable.
                DataTable dt = new DataTable();

                //Loop through the Worksheet rows.
                bool firstRow = true;
                foreach (IXLRow row in workSheet.Rows())
                {
                    //Use the first row to add columns to DataTable.
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Columns.Add(cell.Value.ToString());
                        }
                        firstRow = false;
                    }
                    else
                    {
                         bool checkstatus = false;
                        if ((row.Cell(18).ValueCached == "Yes" && row.Cell(19).ValueCached == "Yes" && row.Cell(2).GetValue<string>().ToLower().Trim().Contains("known so"))
                            || (row.Cell(17).ValueCached == "Yes" && row.Cell(18).ValueCached == "Yes" && row.Cell(19).ValueCached == "Yes")){
                            checkstatus = true;
                            GetsaveInfoGrouprpa(row.Cell(1).GetValue<string>().ToLower().Trim(), "0", CurUserName, "True");
                            row.Cell(20).SetValue("Yes").SetDataType(XLCellValues.Text);
                            string _subject = string.Format(@"SEC PKG Info already saved PKG Material no.: {0} / {1}<br /><br />E-Mail Material Info already saved",
                                    row.Cell(6).GetValue<string>().Trim(), row.Cell(7).GetValue<string>().Trim());
                            string material_query = @" select abc =STUFF(((SELECT DISTINCT  ';' + (select top 1  b.Email from ulogin b where b.[user_name]=f.ActiveBy)
											 from TransApprove f where MatDoc='" + row.Cell(1).GetValue<string>().ToLower().Trim() + "'  and fn in ('PA','PG','PA_Approve','PG_Approve') FOR XML PATH(''))), 1, 1, '')";
                            var table = cs.builditems(material_query);
                            foreach (DataRow dr in table.Rows)
                                cs.sendemail(@dr["abc"].ToString(), "", "<br/>Comment : ", _subject, "");
                        }
                        //dt.Rows.Add();
                        //int i = 0;
                        //foreach (IXLCell cell in row.Cells())
                        //{
                        //    dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                        //    i++;
                        //}
                    }
                }
                //foreach (DataRow _r in dt.Rows)
                //{
                //    Context.Response.Write("\n" + _r["Id"] + _r["Status Delete"] + _r["Status Assign"] + _r["Status Source List"]);
                //    
                //}

                workBook.Save();
                workBook.Dispose();
                //using (var stream = new MemoryStream())
                //{
                //    workBook.SaveAs(stream);
                //    var content = stream.ToArray();
                //}
                //File.Delete(_file);
                //workBook.SaveAs(_file);
            }

            string result = Path.GetFileNameWithoutExtension(_file);
            string _body = string.Format("Dear All, <br/>web service update file {0} in Grid Complete.", result);
            //cs.sendemail(@"Voravut.Somboornpong@thaiunion.com", "", _body, "RPA proces _"+ result, _file);
            var pathnew = string.Format("{0}{1}.xlsx", HttpContext.Current.Server.MapPath(@"~/FileTest/"), result);
            //string  path2 = @"c:\temp2\MySample.txt";
            //File.Move(_file, pathnew);
            var process = System.Diagnostics.Process.GetProcessesByName("Excel");
            foreach (var p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch { }
                }
            }
            if (Session != null) { Session.Clear(); }
            Context.Response.Write("success");
        }
    }
    public void GetsaveInfoGrouprpa(string Id, string InfoGroup, string user, string Check_PChanged)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spsaveInfoGroup";
            cmd.Parameters.AddWithValue("@Id", Id);
            cmd.Parameters.AddWithValue("@InfoGroup", InfoGroup);
            cmd.Parameters.AddWithValue("@user", user);
            cmd.Parameters.AddWithValue("@Check_PChanged", Check_PChanged);
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
    }
    public bool FileIsLocked(string strFullFileName)
    {
        bool blnReturn = false;
        System.IO.FileStream fs;
        try
        {
            fs = System.IO.File.Open(strFullFileName, FileMode.OpenOrCreate, FileAccess.Read,  FileShare.None);
            fs.Close();
        }
        catch (System.IO.IOException ex)
        {
            blnReturn = true;
        }
        return blnReturn;
    }
    public void GetsaveInfoGroup2(string Id, string InfoGroup, string user, string Check_PChanged)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spsaveInfoGroup";
            cmd.Parameters.AddWithValue("@Id", Id);
            cmd.Parameters.AddWithValue("@InfoGroup", InfoGroup);
            cmd.Parameters.AddWithValue("@user", user);
            cmd.Parameters.AddWithValue("@Check_PChanged", Check_PChanged);
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
    }
    [WebMethod]
    public void ExportDataSetToExcel()
    {
        try
        {
            using (SqlConnection con = new SqlConnection(strConn))
            {
                SqlCommand cmd = new SqlCommand();
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "spinfogrouprpa";
                //cmd.Parameters.AddWithValue("@user", user);
                cmd.Parameters.AddWithValue("@where", "");
                cmd.Connection = con;
                con.Open();
                DataTable dtx = GetTable();
                DataTable dt = new DataTable();
                SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
                oAdapter.Fill(dt);
                con.Close();
                foreach (DataRow row in dt.Rows)
                {
                    dtx.Rows.Add(
                    row["Id"],
                    row["Condition"],
                    row["RequestType"],
                    row["DocumentNo"],
                    row["DMSNo"],
                    row["Material"],
                    row["Description"],
                    row["MaterialGroup"],
                    row["Brand"],
                    row["Assign"],
                    row["CreateOn"],
                    row["CreateBy"],
                    row["FinalInfoGroup"],
                    row["ReferenceMaterial"],
                    row["Vendor"],
                    row["Vendor description"]);
                }
                string _d = DateTime.Now.ToString("yyyyMMdd HHmmss");
                //D:\RPA_Data\07 MKT\Packaging\TU_002_AssignMaterial\Input
                //string rootFolder = HttpContext.Current.Server.MapPath(@"~/ExcelFiles/");
                string Pathfilename = string.Format("{0}Export Igrid_{1}.xlsx", rootFolder, _d);
                string[] files = Directory.GetFiles(rootFolder);
                foreach (string _file in files){
                    File.Delete(_file);
                }
                var workbook = new XLWorkbook();
                using (XLWorkbook wb = new XLWorkbook())
                {
                    wb.Worksheets.Add(dtx);
                    var worksheet = wb.Worksheets.Add(_d);
                    wb.SaveAs(@Pathfilename);
                }
            }
            Context.Response.Write("success");
        }
        catch (Exception e)
        {
            Context.Response.Write(e.Message);
            // Action after the exception is caught  
        }
    }
 
    [WebMethod]
    public string iGridMigration(string Keys)
    {
        //master_artwork();
        //header
        ServiceReference.IGRID_OUTBOUND_MODEL iGrid_Model = new ServiceReference.IGRID_OUTBOUND_MODEL();
        //myh.OUTBOUND_HEADERS = new ServiceReference.IGRID_OUTBOUND_HEADER_MODEL;

        ServiceReference.IGRID_OUTBOUND_HEADER_MODEL result = new ServiceReference.IGRID_OUTBOUND_HEADER_MODEL();
        List<ServiceReference.IGRID_OUTBOUND_HEADER_MODEL> iGrid_Header_List = new List<ServiceReference.IGRID_OUTBOUND_HEADER_MODEL>();

        ServiceReference.MM73_OUTBOUND_MATERIAL_NUMBER matNumber = new ServiceReference.MM73_OUTBOUND_MATERIAL_NUMBER();
        ServiceReference.MM73_OUTBOUND_MATERIAL_NUMBERResponse resp = new ServiceReference.MM73_OUTBOUND_MATERIAL_NUMBERResponse();
        try
        {
            SqlParameter[] param = { new SqlParameter("@Keys", Keys) };
            var _table = cs.GetRelatedResources("spGetDocumentNoiGrid", param);
            string _ArtworkNumber = "", _Date = "", _Time = "", _Material = "", _PAUserName = "", _Subject = "Material {0} Created in SAP and send to Artwork Complete";

            foreach (DataRow dr in _table.Rows)
            {
                if (dr["StatusApp"].ToString() == "5")
                    _Subject = "Cancel form iGrid and send info to Artwork Complete";
                _ArtworkNumber = string.Format("{0}", dr["DMSNo"]);
                _Date = String.Format("{0:yyyyMMdd}", dr["CreateOn"]);
                _Time = String.Format("{0:HH:mm:ss}", dr["CreateOn"]); //"10:22:03";
                _Material = string.Format("{0}", dr["Material"]);
                _PAUserName = string.Format("{0}", dr["CreateBy"]);
                DataTable _dt = cs.builditems(@"select isnull(url,'')url,isnull(ReferenceMaterial,'')ReferenceMaterial from TransArtworkURL where Matdoc="
                + string.Format("{0}", dr["Id"]));
                if (_dt.Rows.Count > 0)
                {
                    DataRow r = _dt.Rows[0];
                    result.ArtworkURL = string.Format("{0}", r["url"]);//"http://artwork.thaiunion.com/content/aw-file.pdf";
                    result.ReferenceMaterial = string.Format("{0}", r["ReferenceMaterial"]);
                }
                else
                {
                    result.ArtworkURL = "";
                    result.ReferenceMaterial = "";
                }
                result.ArtworkNumber = _ArtworkNumber;
                result.Date = _Date;
                result.Time = _Time; //"10:22:03";
                result.RecordType = "I";
                result.MaterialNumber = string.Format("{0}", dr["Material"]);
                result.MaterialDescription = string.Format("{0}", dr["Description"]); //"CTN3 - 60960,LUCKY";
                result.MaterialCreatedDate = String.Format("{0:yyyyMMdd}", dr["ModifyOn"]);
                result.Status = dr["Status"].ToString();
                result.PAUserName = string.Format("{0}", dr["CreateBy"]);
                result.PGUserName = string.Format("{0}", dr["Assignee"]);
                //            result.Plant = string.Format("{0}", dr["Plant"].ToString().Replace(';',','));
                result.Plant = string.Format("{0}", dr["Plant"].ToString());
                result.PrintingStyleofPrimary = string.Format("{0}", dr["PrintingStyleofPrimary"]);
                result.PrintingStyleofSecondary = string.Format("{0}", dr["PrintingStyleofSecondary"]);

                //string CustomerDesign = string.Format("{0}", dr["CustomerDesign"]);
                //string[] words = CustomerDesign.Split('|');
                result.CustomersDesign = cs.splittext(dr["CustomerDesign"].ToString(), 0);
                result.CustomersDesignDetail = cs.splittext(dr["CustomerDesign"].ToString(), 1);

                result.CustomersSpec = cs.splittext(dr["CustomerSpec"].ToString(), 0);
                result.CustomersSpecDetail = cs.splittext(dr["CustomerSpec"].ToString(), 1);
                result.CustomersSize = cs.splittext(dr["CustomerSize"].ToString(), 0);
                result.CustomersSizeDetail = cs.splittext(dr["CustomerSize"].ToString(), 1);
                result.CustomerNominatesVendor = cs.splittext(dr["CustomerVendor"].ToString(), 0);
                result.CustomerNominatesVendorDetail = cs.splittext(dr["CustomerVendor"].ToString(), 1);
                result.CustomerNominatesColorPantone = cs.splittext(dr["CustomerColor"].ToString(), 0);
                result.CustomerNominatesColorPantoneDetail = cs.splittext(dr["CustomerColor"].ToString(), 1);
                result.CustomersBarcodeScanable = cs.splittext(dr["CustomerScanable"].ToString(), 0);
                result.CustomersBarcodeScanableDetail = cs.splittext(dr["CustomerScanable"].ToString(), 1);
                result.CustomersBarcodeSpec = cs.splittext(dr["CustomerBarcodeSpec"].ToString(), 0);
                result.CustomersBarcodeSpecDetail = cs.splittext(dr["CustomerBarcodeSpec"].ToString(), 1);
                result.FirstInfoGroup = string.Format("{0}", dr["FirstInfoGroup"]);
                result.SONumber = string.Format("{0}", dr["SO"]);
                result.SOitem = "";
                result.SOPlant = string.Format("{0}", dr["SOPlant"]);
                result.PICMKT = string.Format("{0}", dr["PICMkt"]);
                result.Destination = string.Format("{0}", dr["Destination"]);
                result.RemarkNoteofPA = string.Format("{0}", dr["Remark"]);
                result.FinalInfoGroup = string.Format("{0}", dr["FinalInfoGroup"]);
                result.RemarkNoteofPG = "";
                result.CompleteInfoGroup = "";
                result.ProductionExpirydatesystem = "";
                result.Seriousnessofcolorprinting = "";
                result.CustIngreNutritionAnalysis = "";
                result.ShadeLimit = "";
                result.PackageQuantity = "";
                result.WastePercent = "";
                iGrid_Header_List.Add(result);
                iGrid_Model.OUTBOUND_HEADERS = iGrid_Header_List.ToArray();
            }
            List<ServiceReference.IGRID_OUTBOUND_ITEM_MODEL> iGrid_Item_List = new List<ServiceReference.IGRID_OUTBOUND_ITEM_MODEL>();
            DataTable dt = new DataTable();
            using (SqlConnection con = new SqlConnection(strConn))
            {
                SqlCommand cmd = new SqlCommand();
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "spInboundArtwork_Interface";
                cmd.Parameters.AddWithValue("@Keys", string.Format("{0}", Keys.ToString()));
                cmd.Connection = con;
                con.Open();
                SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
                oAdapter.Fill(dt);
                con.Close();
                List<InboundArtwork> _itemsArtwork = new List<InboundArtwork>();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    var detail = new ServiceReference.IGRID_OUTBOUND_ITEM_MODEL();
                    DataRow dr = dt.Rows[i];

                    detail.ArtworkNumber = string.Format("{0}", _ArtworkNumber);
                    detail.Date = _Date;
                    detail.Time = _Time;
                    detail.Characteristic = dr["cols"].ToString();
                    //detail.Description = dr["Description"].ToString();
                    //detail.Value = dr["value"].ToString();
                    string[] splitHeader = dr["value"].ToString().Split(';');
                    if (splitHeader != null && splitHeader.Length > 1)
                        foreach (string word in splitHeader)
                        {
                            detail = new ServiceReference.IGRID_OUTBOUND_ITEM_MODEL();
                            detail.ArtworkNumber = string.Format("{0}", _ArtworkNumber);
                            detail.Date = _Date;
                            detail.Time = _Time;
                            detail.Characteristic = dr["cols"].ToString();
                            detail.Value = word.ToString();
                            detail.Description = detail.Value.ToString();
                            iGrid_Item_List.Add(detail);
                        }
                    else
                    {
                        detail.Description = dr["Description"].ToString();
                        detail.Value = dr["value"].ToString();
                        iGrid_Item_List.Add(detail);
                    }
                }
            }
            iGrid_Model.OUTBOUND_ITEMS = iGrid_Item_List.ToArray();
            ServiceReference.MM73Client client = new ServiceReference.MM73Client();
            matNumber.param = iGrid_Model;
            //resp = client.MATERIAL_NUMBER(matNumber);
            string Start = DateTime.Now.ToString();
            resp = client.MATERIAL_NUMBER(matNumber);
            string dtEnd = DateTime.Now.ToString();
            //Context.Response.Write(JsonConvert.SerializeObject(resp));
            return resp.MM73_OUTBOUND_MATERIAL_NUMBERResult.msg;
        }
        catch (Exception e)
        {
            return e.Message;
            // Action after the exception is caught  
        }
    }
    [WebMethod()]
    public void GetData(string data)
    {
        string datapath = "~/FileTest/" + data + ".json";
        using (StreamReader sr = new StreamReader(Server.MapPath(datapath)))
        {
            string json = sr.ReadToEnd();
            dynamic dynJson = JsonConvert.DeserializeObject(json);
            foreach (var item in dynJson)
            {
                //string json2 = convertthai(ro.Keyword);
                string value = item.tmpstr.ToString();
                using (SqlConnection oConn = new SqlConnection(strConn))
                {
                    oConn.Open();
                    //sName = "PrimarySize Where Description like '%307%'";
                    string strQuery = "select * from " + value;
                    DataTable dt = new DataTable();
                    SqlDataAdapter oAdapter = new SqlDataAdapter(strQuery, oConn);
                    // Fill the dataset.
                    oAdapter.Fill(dt);
                    oConn.Close();
                    Context.Response.Write(JsonConvert.SerializeObject(dt));
                }
            }
        }
    }
    [WebMethod]
    public void GetPackStyle4Digits(string Product_Code)
    {
        RfcDestination destination = new SAPConnection().SVIPRD();
        IRfcFunction function = destination.Repository.CreateFunction("RFC_READ_TABLE");
        function.SetValue("ROWCOUNT", "10");//99999999
        function.SetValue("DELIMITER", "|");

        function.SetValue("QUERY_TABLE", "AUSP");
        IRfcTable fieldsTable = function.GetTable("FIELDS");
        fieldsTable.Append();
        IRfcTable optsTable = function.GetTable("OPTIONS");
        optsTable.Append();
        fieldsTable.SetValue("FIELDNAME", "OBJEK");
        fieldsTable.Append();
        fieldsTable.SetValue("FIELDNAME", "ATWRT");
        optsTable.SetValue("TEXT", "OBJEK EQ '" + Product_Code + "' and");
        optsTable.Append();
        optsTable.SetValue("TEXT", "ATINN = '0000000142'");
        //IRfcStructure struck = function.GetStructure("ORDER_OBJECTS");
        //struck.SetValue(0, "X");
        //struck.SetValue(1, "X");
        //struck.SetValue(2, "X");
        //struck.SetValue(3, "X");
        //struck.SetValue(4, "X");
        //struck.SetValue(5, "X");
        //struck.SetValue(6, "X");
        function.Invoke(destination);
        IRfcTable OrderHeader = function.GetTable("DATA");
        DataTable dt = new SAPConnection().GetDataTableFromRFCTable(OrderHeader);
        foreach (DataRow r in dt.Rows)
        {
            string[] arr = r["wa"].ToString().Split('|');
            string result = string.Format("{0}", arr[1]);
            Context.Response.Write(result);
        }
    }
    [WebMethod]
    public void sapconnect()
    {
        string data = "ZSDT006;TVM1T";
        // Split string on spaces (this will separate all the words).
        string[] words = data.Split(';');
        foreach (string word in words)
        {
            //  RfcDestination destination = new SAPCONNECT().SVIDEV();
            RfcDestination destination = new SAPConnection().SVIPRD();
            IRfcFunction function = destination.Repository.CreateFunction("RFC_READ_TABLE");
            function.SetValue("ROWCOUNT", "99999999");
            function.SetValue("DELIMITER", ";");

            function.SetValue("QUERY_TABLE", word);
            IRfcTable fieldsTable = function.GetTable("FIELDS");
            fieldsTable.Append();
            IRfcTable optsTable = function.GetTable("OPTIONS");
            optsTable.Append();
            switch (word)
            {
                case "ZSDT006":
                    fieldsTable.SetValue("FIELDNAME", "ZBAND_ID");
                    fieldsTable.Append();
                    fieldsTable.SetValue("FIELDNAME", "ZBAND_DESC");
                    optsTable.SetValue("TEXT", "ZDEL_FLAG NE 'X'");
                    break;
                case "TVM1T":
                    fieldsTable.SetValue("FIELDNAME", "MVGR1");
                    fieldsTable.Append();
                    fieldsTable.SetValue("FIELDNAME", "BEZEI");
                    optsTable.SetValue("TEXT", "SPRAS = 'EN'");
                    break;
            }
            //IRfcStructure struck = function.GetStructure("ORDER_OBJECTS");
            //struck.SetValue(0, "X");
            //struck.SetValue(1, "X");
            //struck.SetValue(2, "X");
            //struck.SetValue(3, "X");
            //struck.SetValue(4, "X");
            //struck.SetValue(5, "X");
            //struck.SetValue(6, "X");
            function.Invoke(destination);
            IRfcTable OrderHeader = function.GetTable("DATA");
            DataTable dt = new SAPConnection().GetDataTableFromRFCTable(OrderHeader);
            int i = dt.Rows.Count;
            foreach (DataRow r in dt.Rows)
            {

                string[] arr = r["wa"].ToString().Split(';');
                string newS = string.Format("{0}", arr[0]);
                if (word == "TVM1T")
                {
                    string s = arr[0].ToString();
                    StringBuilder sb = new StringBuilder(s);
                    sb[0] = 'Z';
                    newS = sb.ToString();
                }
                string qry = "insert into TransBrand_ImportSAP values (@ZBAND_ID,@ZBAND_DESC,@ZDEL_FLAG); exec spupdatefromsap;";
                if (dt.Rows.IndexOf(r) == 0)
                    qry = "truncate table TransBrand_ImportSAP; " + qry;
                using (SqlConnection CN = new SqlConnection(strConn))
                {
                    SqlCommand SqlCom = new SqlCommand(qry, CN);
                    //We are passing Original File Path and file byte data as sql parameters.
                    SqlCom.Parameters.Add(new SqlParameter("@ZBAND_ID", newS.ToString()));
                    SqlCom.Parameters.Add(new SqlParameter("@ZBAND_DESC", arr[1].ToString()));
                    SqlCom.Parameters.Add(new SqlParameter("@ZDEL_FLAG", ""));
                    CN.Open();
                    SqlCom.ExecuteNonQuery();
                    CN.Close();
                }
            }
            //ZPKG_SEC_BRAND'
            //SqlParameter[] param = { new SqlParameter("@xxx", "WINSHUTADM".ToString()) };
            //var table = cs.GetRelatedResources("spupdatefromsap", param);
        }
    }
    public class ServiceClientFactory<TChannel> : ClientBase<TChannel> where TChannel : class
    {
        public TChannel Create(string url)
        {
            this.Endpoint.Address = new EndpointAddress(new Uri(url));
            return this.Channel;
        }
    }
    [WebMethod]
    public string convertthai(string text)
    {
        try
        {
            dynamic dynJson = JsonConvert.DeserializeObject(text);
            foreach (var item in dynJson)
            {
                text = item.tmpstr.ToString();
            }    //otherstuff        

        }
        catch (Exception ex)
        {
            //error loging stuff
        }
        return text;
    }
    [WebMethod]
    public void GetDocumenturl(string Id)
    {
        DataTable resp = new DataTable();
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetDocumenturl";
            cmd.Parameters.AddWithValue("@DocumentNo", string.Format("{0}", Id.ToString()));
            cmd.Connection = con;
            con.Open();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(resp);
            con.Close();
        }
        Context.Response.Write(JsonConvert.SerializeObject(resp));
    }
    [WebMethod]
    public void SelectMaster(string user)
    {
        DataTable resp = new DataTable();
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spSelectMaster";
            cmd.Parameters.AddWithValue("@user", string.Format("{0}", user.ToString()));
            cmd.Connection = con;
            con.Open();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(resp);
            con.Close();
        }
        Context.Response.Write(JsonConvert.SerializeObject(resp));
    }
    public void ArtworkURL(string document, string url, string ReferenceMaterial)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spinsetArtworkURL";
            cmd.Parameters.AddWithValue("@document", string.Format("{0}", document));
            cmd.Parameters.AddWithValue("@url", string.Format("{0}", url));
            cmd.Parameters.AddWithValue("@ReferenceMaterial", string.Format("{0}", ReferenceMaterial));
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
    }
    [WebMethod]
    public void testsendmaster(string SubChanged_Id)
    {
        string strSQL = " select Id,Changed_Charname,Description from TransMaster where Changed_id ='" + SubChanged_Id + "'";
        DataTable dt = cs.builditems(strSQL);
        foreach (DataRow dr in dt.Rows)
        {
            string _Id = dr["Id"].ToString();
            string _Description = dr["Description"].ToString();
            string[] value = { dr["Changed_Charname"].ToString(), _Description, _Id };
            master_artwork(value);
        }
    }
    public void master_artwork(string[] name)
    {
        //loop details
        myService.CHARACTERISTICS list = new myService.CHARACTERISTICS();
        List<myService.CHARACTERISTIC> iGrid_CHARACTERISTICS = new List<myService.CHARACTERISTIC>();
        myService.CHARACTERISTIC item = new myService.CHARACTERISTIC();
        item.NAME = name[0].ToString();
        item.DESCRIPTION = name[1].ToString();
        if (name[0].ToString() == "ZPKG_SEC_BRAND")
        {
            item.VALUE = name[2].ToString();
        }
        else
            item.VALUE = name[1].ToString();

        iGrid_CHARACTERISTICS.Add(item);
        list.CHARACTERISTICS1 = iGrid_CHARACTERISTICS.ToArray();
        myService.MM72_OUTBOUND_MATERIAL_CHARACTERISTIC matNumber = new myService.MM72_OUTBOUND_MATERIAL_CHARACTERISTIC();
        myService.MM72_OUTBOUND_MATERIAL_CHARACTERISTICResponse resp = new myService.MM72_OUTBOUND_MATERIAL_CHARACTERISTICResponse();
        myService.MM72Client client = new myService.MM72Client();

        matNumber.param = list;
        resp = client.MATERIAL_CHARACTERISTIC(matNumber);
        //++++++++++++++++++++++++++++++++

        string datapath = "~/FileTest/master" + name[0].ToString() + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml";
        using (FileStream fs = new FileStream(Server.MapPath(datapath), FileMode.Create))
        {
            new XmlSerializer(typeof(myService.CHARACTERISTICS)).Serialize(fs, list);
        }
        cs.sendemail("Pornpimon.Bouban@thaiunion.com;Walaipan.Supoltawanitch@thaiunion.com;Nongrat.Jantarasuwan@thaiunion.com", "",
            string.Format("Name : {0}<br/>Status: {1},<br/>msg: {2}", name[0].ToString(), resp.MM72_OUTBOUND_MATERIAL_CHARACTERISTICResult.status, resp.MM72_OUTBOUND_MATERIAL_CHARACTERISTICResult.msg),
            string.Format("master {0} Created in SAP Complete", name[1].ToString()), Server.MapPath(datapath));
    }
    [WebMethod]
    public string OutboundArtwork_Xml(string Keys)
    {
        //master_artwork();
        //header
        ServiceReference.IGRID_OUTBOUND_MODEL iGrid_Model = new ServiceReference.IGRID_OUTBOUND_MODEL();
        //myh.OUTBOUND_HEADERS = new ServiceReference.IGRID_OUTBOUND_HEADER_MODEL;

        ServiceReference.IGRID_OUTBOUND_HEADER_MODEL result = new ServiceReference.IGRID_OUTBOUND_HEADER_MODEL();
        List<ServiceReference.IGRID_OUTBOUND_HEADER_MODEL> iGrid_Header_List = new List<ServiceReference.IGRID_OUTBOUND_HEADER_MODEL>();

        ServiceReference.MM73_OUTBOUND_MATERIAL_NUMBER matNumber = new ServiceReference.MM73_OUTBOUND_MATERIAL_NUMBER();
        ServiceReference.MM73_OUTBOUND_MATERIAL_NUMBERResponse resp = new ServiceReference.MM73_OUTBOUND_MATERIAL_NUMBERResponse();
        var _table = cs.builditems("select *,case when statusapp=4 then 'Completed' when statusapp=5 then 'Canceled' end as 'Status' from SapMaterial Where DocumentNo='" + Keys + "'");
        string _ArtworkNumber = "", _Date = "", _Time = "", _Material = "", _PAUserName = "", _Subject = "Material {0} Created in SAP and send to Artwork Complete";

        foreach (DataRow dr in _table.Rows)
        {
            if (dr["StatusApp"].ToString() == "0") return "Inprocess";
            if (dr["StatusApp"].ToString() == "5")
                _Subject = "Cancel form iGrid and send info to Artwork Complete";
            _ArtworkNumber = string.Format("{0}", dr["DMSNo"]);
            _Date = String.Format("{0:yyyyMMdd}", dr["CreateOn"]);
            _Time = String.Format("{0:HH:mm:ss}", dr["CreateOn"]); //"10:22:03";
            _Material = string.Format("{0}", dr["Material"]);
            _PAUserName = string.Format("{0}", dr["CreateBy"]);
            DataTable _dt = cs.builditems(@"select isnull(url,'')url,isnull(ReferenceMaterial,'')ReferenceMaterial from TransArtworkURL where Matdoc="
            + string.Format("{0}", dr["Id"]));
            if (_dt.Rows.Count > 0)
            {
                DataRow r = _dt.Rows[0];
                result.ArtworkURL = string.Format("{0}", r["url"]);//"http://artwork.thaiunion.com/content/aw-file.pdf";
                result.ReferenceMaterial = string.Format("{0}", r["ReferenceMaterial"]);
            }
            else
            {
                result.ArtworkURL = "";
                result.ReferenceMaterial = "";
            }
            result.ArtworkNumber = _ArtworkNumber;
            result.Date = _Date;
            result.Time = _Time; //"10:22:03";
            result.RecordType = "I";
            result.MaterialNumber = dr["StatusApp"].ToString() == "5" ? "" : string.Format("{0}", dr["Material"]);
            result.MaterialDescription = string.Format("{0}", dr["Description"]); //"CTN3 - 60960,LUCKY";
            result.MaterialCreatedDate = String.Format("{0:yyyyMMdd}", dr["ModifyOn"]);
            result.Status = dr["Status"].ToString();
            result.PAUserName = string.Format("{0}", dr["CreateBy"]);
            result.PGUserName = string.Format("{0}", dr["Assignee"]);
            //            result.Plant = string.Format("{0}", dr["Plant"].ToString().Replace(';',','));
            result.Plant = string.Format("{0}", dr["Plant"].ToString());
            result.PrintingStyleofPrimary = string.Format("{0}", dr["PrintingStyleofPrimary"]);
            result.PrintingStyleofSecondary = string.Format("{0}", dr["PrintingStyleofSecondary"]);

            //string CustomerDesign = string.Format("{0}", dr["CustomerDesign"]);
            //string[] words = CustomerDesign.Split('|');
            result.CustomersDesign = cs.splittext(dr["CustomerDesign"].ToString(), 0);
            result.CustomersDesignDetail = cs.splittext(dr["CustomerDesign"].ToString(), 1);

            result.CustomersSpec = cs.splittext(dr["CustomerSpec"].ToString(), 0);
            result.CustomersSpecDetail = cs.splittext(dr["CustomerSpec"].ToString(), 1);
            result.CustomersSize = cs.splittext(dr["CustomerSize"].ToString(), 0);
            result.CustomersSizeDetail = cs.splittext(dr["CustomerSize"].ToString(), 1);
            result.CustomerNominatesVendor = cs.splittext(dr["CustomerVendor"].ToString(), 0);
            result.CustomerNominatesVendorDetail = cs.splittext(dr["CustomerVendor"].ToString(), 1);
            result.CustomerNominatesColorPantone = cs.splittext(dr["CustomerColor"].ToString(), 0);
            result.CustomerNominatesColorPantoneDetail = cs.splittext(dr["CustomerColor"].ToString(), 1);
            result.CustomersBarcodeScanable = cs.splittext(dr["CustomerScanable"].ToString(), 0);
            result.CustomersBarcodeScanableDetail = cs.splittext(dr["CustomerScanable"].ToString(), 1);
            result.CustomersBarcodeSpec = cs.splittext(dr["CustomerBarcodeSpec"].ToString(), 0);
            result.CustomersBarcodeSpecDetail = cs.splittext(dr["CustomerBarcodeSpec"].ToString(), 1);
            result.FirstInfoGroup = string.Format("{0}", dr["FirstInfoGroup"]);
            result.SONumber = string.Format("{0}", dr["SO"]);
            result.SOitem = "";
            result.SOPlant = string.Format("{0}", dr["SOPlant"]);
            result.PICMKT = string.Format("{0}", dr["PICMkt"]);
            result.Destination = string.Format("{0}", dr["Destination"]);
            result.RemarkNoteofPA = string.Format("{0}", dr["Remark"]);
            result.FinalInfoGroup = string.Format("{0}", dr["FinalInfoGroup"]);
            result.RemarkNoteofPG = "";
            result.CompleteInfoGroup = "";
            result.ProductionExpirydatesystem = "";
            result.Seriousnessofcolorprinting = "";
            result.CustIngreNutritionAnalysis = "";
            result.ShadeLimit = "";
            result.PackageQuantity = "";
            result.WastePercent = "";
            iGrid_Header_List.Add(result);
            iGrid_Model.OUTBOUND_HEADERS = iGrid_Header_List.ToArray();
        }
        List<ServiceReference.IGRID_OUTBOUND_ITEM_MODEL> iGrid_Item_List = new List<ServiceReference.IGRID_OUTBOUND_ITEM_MODEL>();
        DataTable dt = new DataTable();
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spInboundArtwork";
            cmd.Parameters.AddWithValue("@Keys", string.Format("{0}", Keys.ToString()));
            cmd.Connection = con;
            con.Open();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            List<InboundArtwork> _itemsArtwork = new List<InboundArtwork>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                var detail = new ServiceReference.IGRID_OUTBOUND_ITEM_MODEL();
                DataRow dr = dt.Rows[i];

                detail.ArtworkNumber = string.Format("{0}", _ArtworkNumber);
                detail.Date = _Date;
                detail.Time = _Time;
                detail.Characteristic = dr["cols"].ToString();
                //detail.Description = dr["Description"].ToString();
                //detail.Value = dr["value"].ToString();
                string[] splitHeader = dr["value"].ToString().Split(';');
                if (splitHeader != null && splitHeader.Length > 1)
                    foreach (string word in splitHeader)
                    {
                        detail = new ServiceReference.IGRID_OUTBOUND_ITEM_MODEL();
                        detail.ArtworkNumber = string.Format("{0}", _ArtworkNumber);
                        detail.Date = _Date;
                        detail.Time = _Time;
                        detail.Characteristic = dr["cols"].ToString();
                        detail.Value = word.ToString();
                        detail.Description = detail.Value.ToString();
                        iGrid_Item_List.Add(detail);
                    }
                else
                {
                    detail.Description = dr["Description"].ToString();
                    detail.Value = dr["value"].ToString();
                    iGrid_Item_List.Add(detail);
                }
            }
        }
        iGrid_Model.OUTBOUND_ITEMS = iGrid_Item_List.ToArray();
        ServiceReference.MM73Client client = new ServiceReference.MM73Client();
        matNumber.param = iGrid_Model;
        //resp = client.MATERIAL_NUMBER(matNumber);
        //Context.Response.Write(JsonConvert.SerializeObject(resp));
        string datapath = "~/FileTest/XML_iGrid_Model" + Keys.ToString() + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml";
        using (FileStream fs = new FileStream(Server.MapPath(datapath), FileMode.Create))
        {
            new XmlSerializer(typeof(ServiceReference.IGRID_OUTBOUND_MODEL)).Serialize(fs, iGrid_Model);
        }
        return "Message Sent Succesfully";
    }
    [WebMethod]
    public string OutboundArtwork(string Keys)
    {
        //master_artwork();
        //header
        ServiceReference.IGRID_OUTBOUND_MODEL iGrid_Model = new ServiceReference.IGRID_OUTBOUND_MODEL();
        //myh.OUTBOUND_HEADERS = new ServiceReference.IGRID_OUTBOUND_HEADER_MODEL;

        ServiceReference.IGRID_OUTBOUND_HEADER_MODEL result = new ServiceReference.IGRID_OUTBOUND_HEADER_MODEL();
        List<ServiceReference.IGRID_OUTBOUND_HEADER_MODEL> iGrid_Header_List = new List<ServiceReference.IGRID_OUTBOUND_HEADER_MODEL>();

        ServiceReference.MM73_OUTBOUND_MATERIAL_NUMBER matNumber = new ServiceReference.MM73_OUTBOUND_MATERIAL_NUMBER();
        ServiceReference.MM73_OUTBOUND_MATERIAL_NUMBERResponse resp = new ServiceReference.MM73_OUTBOUND_MATERIAL_NUMBERResponse();
        try
        {
            var _table = cs.builditems("select *,case when statusapp=4 then 'Completed' when statusapp=5 then 'Canceled' end as 'Status' from SapMaterial Where DocumentNo='" + Keys + "'");
            string _ArtworkNumber = "", _Date = "", _Time = "", _Material = "", _PAUserName = "", _Subject = "Material {0} Created in SAP and send to Artwork Complete";

            foreach (DataRow dr in _table.Rows)
            {
                if (dr["StatusApp"].ToString() == "0") return "Inprocess";
                if (dr["StatusApp"].ToString() == "5")
                    _Subject = "Cancel form iGrid and send info to Artwork Complete";
                _ArtworkNumber = string.Format("{0}", dr["DMSNo"]);
                _Date = String.Format("{0:yyyyMMdd}", dr["CreateOn"]);
                _Time = String.Format("{0:HH:mm:ss}", dr["CreateOn"]); //"10:22:03";
                _Material = string.Format("{0}", dr["Material"]);
                _PAUserName = string.Format("{0}", dr["CreateBy"]);
                DataTable _dt = cs.builditems(@"select isnull(url,'')url,isnull(ReferenceMaterial,'')ReferenceMaterial from TransArtworkURL where Matdoc="
                + string.Format("{0}", dr["Id"]));
                if (_dt.Rows.Count > 0)
                {
                    DataRow r = _dt.Rows[0];
                    result.ArtworkURL = string.Format("{0}", r["url"]);//"http://artwork.thaiunion.com/content/aw-file.pdf";
                    result.ReferenceMaterial = string.Format("{0}", r["ReferenceMaterial"]);
                }
                else
                {
                    result.ArtworkURL = "";
                    result.ReferenceMaterial = "";
                }
                result.ArtworkNumber = _ArtworkNumber;
                result.Date = _Date;
                result.Time = _Time; //"10:22:03";
                result.RecordType = "I";
                result.MaterialNumber = dr["StatusApp"].ToString() == "5" ? "" : string.Format("{0}", dr["Material"]);
                result.MaterialDescription = string.Format("{0}", dr["Description"]); //"CTN3 - 60960,LUCKY";
                result.MaterialCreatedDate = String.Format("{0:yyyyMMdd}", dr["ModifyOn"]);
                result.Status = dr["Status"].ToString();
                result.PAUserName = string.Format("{0}", dr["CreateBy"]);
                result.PGUserName = string.Format("{0}", dr["Assignee"]);
                //            result.Plant = string.Format("{0}", dr["Plant"].ToString().Replace(';',','));
                result.Plant = string.Format("{0}", dr["Plant"].ToString());
                result.PrintingStyleofPrimary = string.Format("{0}", dr["PrintingStyleofPrimary"]);
                result.PrintingStyleofSecondary = string.Format("{0}", dr["PrintingStyleofSecondary"]);

                //string CustomerDesign = string.Format("{0}", dr["CustomerDesign"]);
                //string[] words = CustomerDesign.Split('|');
                result.CustomersDesign = cs.splittext(dr["CustomerDesign"].ToString(), 0);
                result.CustomersDesignDetail = cs.splittext(dr["CustomerDesign"].ToString(), 1);

                result.CustomersSpec = cs.splittext(dr["CustomerSpec"].ToString(), 0);
                result.CustomersSpecDetail = cs.splittext(dr["CustomerSpec"].ToString(), 1);
                result.CustomersSize = cs.splittext(dr["CustomerSize"].ToString(), 0);
                result.CustomersSizeDetail = cs.splittext(dr["CustomerSize"].ToString(), 1);
                result.CustomerNominatesVendor = cs.splittext(dr["CustomerVendor"].ToString(), 0);
                result.CustomerNominatesVendorDetail = cs.splittext(dr["CustomerVendor"].ToString(), 1);
                result.CustomerNominatesColorPantone = cs.splittext(dr["CustomerColor"].ToString(), 0);
                result.CustomerNominatesColorPantoneDetail = cs.splittext(dr["CustomerColor"].ToString(), 1);
                result.CustomersBarcodeScanable = cs.splittext(dr["CustomerScanable"].ToString(), 0);
                result.CustomersBarcodeScanableDetail = cs.splittext(dr["CustomerScanable"].ToString(), 1);
                result.CustomersBarcodeSpec = cs.splittext(dr["CustomerBarcodeSpec"].ToString(), 0);
                result.CustomersBarcodeSpecDetail = cs.splittext(dr["CustomerBarcodeSpec"].ToString(), 1);
                result.FirstInfoGroup = string.Format("{0}", dr["FirstInfoGroup"]);
                result.SONumber = string.Format("{0}", dr["SO"]);
                result.SOitem = "";
                result.SOPlant = string.Format("{0}", dr["SOPlant"]);
                result.PICMKT = string.Format("{0}", dr["PICMkt"]);
                result.Destination = string.Format("{0}", dr["Destination"]);
                result.RemarkNoteofPA = string.Format("{0}", dr["Remark"]);
                result.FinalInfoGroup = string.Format("{0}", dr["FinalInfoGroup"]);
                result.RemarkNoteofPG = "";
                result.CompleteInfoGroup = "";
                result.ProductionExpirydatesystem = "";
                result.Seriousnessofcolorprinting = "";
                result.CustIngreNutritionAnalysis = "";
                result.ShadeLimit = "";
                result.PackageQuantity = "";
                result.WastePercent = "";
                iGrid_Header_List.Add(result);
                iGrid_Model.OUTBOUND_HEADERS = iGrid_Header_List.ToArray();
            }
            List<ServiceReference.IGRID_OUTBOUND_ITEM_MODEL> iGrid_Item_List = new List<ServiceReference.IGRID_OUTBOUND_ITEM_MODEL>();
            DataTable dt = new DataTable();
            using (SqlConnection con = new SqlConnection(strConn))
            {
                SqlCommand cmd = new SqlCommand();
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "spInboundArtwork";
                cmd.Parameters.AddWithValue("@Keys", string.Format("{0}", Keys.ToString()));
                cmd.Connection = con;
                con.Open();
                SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
                oAdapter.Fill(dt);
                con.Close();
                List<InboundArtwork> _itemsArtwork = new List<InboundArtwork>();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    var detail = new ServiceReference.IGRID_OUTBOUND_ITEM_MODEL();
                    DataRow dr = dt.Rows[i];

                    detail.ArtworkNumber = string.Format("{0}", _ArtworkNumber);
                    detail.Date = _Date;
                    detail.Time = _Time;
                    detail.Characteristic = dr["cols"].ToString();
                    //detail.Description = dr["Description"].ToString();
                    //detail.Value = dr["value"].ToString();
                    string[] splitHeader = dr["value"].ToString().Split(';');
                    if (splitHeader != null && splitHeader.Length > 1)
                        foreach (string word in splitHeader)
                        {
                            detail = new ServiceReference.IGRID_OUTBOUND_ITEM_MODEL();
                            detail.ArtworkNumber = string.Format("{0}", _ArtworkNumber);
                            detail.Date = _Date;
                            detail.Time = _Time;
                            detail.Characteristic = dr["cols"].ToString();
                            detail.Value = word.ToString();
                            detail.Description = detail.Value.ToString();
                            iGrid_Item_List.Add(detail);
                        }
                    else
                    {
                        detail.Description = dr["Description"].ToString();
                        detail.Value = dr["value"].ToString();
                        iGrid_Item_List.Add(detail);
                    }
                }
            }
            iGrid_Model.OUTBOUND_ITEMS = iGrid_Item_List.ToArray();
            ServiceReference.MM73Client client = new ServiceReference.MM73Client();
            matNumber.param = iGrid_Model;
            //resp = client.MATERIAL_NUMBER(matNumber);
            string Start = DateTime.Now.ToString();
            resp = client.MATERIAL_NUMBER(matNumber);
            string dtEnd = DateTime.Now.ToString();
            //Context.Response.Write(JsonConvert.SerializeObject(resp));
            string datapath = "~/FileTest/iGrid_Model" + Keys.ToString() + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml";
            using (FileStream fs = new FileStream(Server.MapPath(datapath), FileMode.Create))
            {
                new XmlSerializer(typeof(ServiceReference.IGRID_OUTBOUND_MODEL)).Serialize(fs, iGrid_Model);
            }
            SqlParameter[] param = { new SqlParameter("@keys", string.Format("{0}", Keys)) };
            cs.GetExecuteNonQuery("spupdateOutbound", param);
            cs.sendemail(@cs.Getuser(_PAUserName, "email"), "Nongrat.Jantarasuwan@thaiunion.com;Pornpimon.Bouban@thaiunion.com;Walaipan.Supoltawanitch@thaiunion.com;Voravut.Somboornpong@thaiunion.com",
                string.Format("Artwork Number {3} <br/> Workflow IGrid {0} <br/> Status: {1},<br/>msg: {2} <br/>Start Time: {4}<br/>End Time: {5}", Keys.ToString(), resp.MM73_OUTBOUND_MATERIAL_NUMBERResult.status, resp.MM73_OUTBOUND_MATERIAL_NUMBERResult.msg, _ArtworkNumber, Start, dtEnd),
                string.Format(_Subject, _Material.ToString()), Server.MapPath(datapath));
            return resp.MM73_OUTBOUND_MATERIAL_NUMBERResult.msg;
        }
        catch (Exception e)
        {
            cs.sendemail("pornnicha.thanarak@thaiunion.com;Nongrat.Jantarasuwan@thaiunion.com;Pornpimon.Bouban@thaiunion.com;Walaipan.Supoltawanitch@thaiunion.com;Voravut.Somboornpong@thaiunion.com;Suprawee.Pakagawin@thaiunion.com;Thitapa.Pojang@thaiunion.com", "",
            string.Format("{0}", e.Message), string.Format("iGrid can't send to Artwork , status fail , iGrid No : {0}", Keys.Substring(0, 16)), "");
            return e.Message;
            // Action after the exception is caught  
        }
    }
    public string Insert(string value)
    {
        // Insert string at index 6.
        //string adjusted = "";
        int div = value.Length / 30; int a = 0;
        if (value.Length > 30)
            for (int i = 1; i <= div; i++)
            {
                value = value.Insert((30 * i) + a, ";");
                a++;
            }
        return value;
    }
    [WebMethod]
    public string inputArtworkNumber(ArtworkObject _artworkObject, List<InboundArtwork> _itemsArtwork)
    {
        myService.SERVICE_RESULT_MODEL Results = new myService.SERVICE_RESULT_MODEL();
        CreateDocument ro = new CreateDocument();
        string _Subject = "", _Body = "", Keys = "0";
        var _total = cs.ReadItems("select count(*)total from SapMaterial Where statusapp=0 and dmsno='" + _artworkObject.ArtworkNumber + "'");
        if (Convert.ToInt32(_total.ToString()) > 0)
        {
            Keys = string.Format("{0}", cs.ReadItems(@"select top 1 DocumentNo from SapMaterial Where statusapp=0 and dmsno='" + _artworkObject.ArtworkNumber + "'"));
            _Subject = "Artwork number duplicate value";
            _Body = string.Format("artwork number : {0} <br/>CreateBy : {1}", _artworkObject.ArtworkNumber, cs.Getuser(_artworkObject.PAUserName, "fn"));
            goto jumptoexit;
        }
        string datapath = "~/FileTest/dArtwork" + Keys.ToString() + ".xml";
        using (FileStream fs = new FileStream(Server.MapPath(datapath), FileMode.Create))
        {
            new XmlSerializer(typeof(List<InboundArtwork>)).Serialize(fs, _itemsArtwork);
        }
        ro.CreateBy = CurUserName;
        ro.Code = _artworkObject.MaterialNumber == "" ? _artworkObject.ReferenceMaterial : _artworkObject.MaterialNumber;
        if (_artworkObject.RecordType == "U" && _artworkObject.MaterialNumber != "")
            ro.Condition = "7";
        else if (_artworkObject.RecordType == "I")
        {
            if (_artworkObject.ReferenceMaterial == "")
            {
                ro.Condition = "1";
            }
            else
            {
                ro.Condition = "4";
            }
        }
        if (string.IsNullOrEmpty(ro.Condition))
        {
            _Subject = "data invalid";
            goto jumptoexit;
        }
        //WebService s = new WebService();
        DateTime myDate = DateTime.ParseExact(_artworkObject.Date + " " + _artworkObject.Time + ",531", "yyyyMMdd HH:mm:ss,fff",
                                              System.Globalization.CultureInfo.InvariantCulture);
        SqlParameter[] param = { new SqlParameter("@Code", ro.Code),
                new SqlParameter("@Condition",string.Format("{0}",ro.Condition)),
                //new SqlParameter("@CreateOn",string.Format("{0}",myDate)),
                new SqlParameter("@CreateBy",string.Format("{0}",_artworkObject.PAUserName))};
        var table = cs.GetRelatedResources("spCreateDocument", param);
        //s.saveCreateroot(ro);
        //dynamic dynJson = JsonConvert.DeserializeObject(s.ToString());
        foreach (DataRow value in table.Rows)
        {
            Keys = string.Format("{0}", value["DocumentNo"]);
            //			string datapath = "~/FileTest/dArtwork" + Keys.ToString() + ".xml";
            //            using (FileStream fs = new FileStream(Server.MapPath(datapath), FileMode.Create))
            //            {
            //                new XmlSerializer(typeof(List<InboundArtwork>)).Serialize(fs, _itemsArtwork);
            //            }
            string _group = "";
            ArtworkURL(value["ID"].ToString(), _artworkObject.ArtworkURL, _artworkObject.ReferenceMaterial);
            //assignee PG input
            Assign _assign = new Assign();
            _assign.Id = value["ID"].ToString();
            _assign.Assignee = _artworkObject.PGUserName.ToString();
            artworkAssign(_assign);
            foreach (var p in _itemsArtwork)
                if (p.Characteristic == "ZPKG_SEC_GROUP")
                    _group = p.Value;
            att(_artworkObject, Keys);
            string _LidType = "", _ContainerType = "", _ChangePoint = "";
            string json = JsonConvert.SerializeObject(_itemsArtwork);
            DataTable pDt = JsonConvert.DeserializeObject<DataTable>(json);
            sapmaterial items = new sapmaterial();
            DataTable destination = new DataTable(pDt.TableName);
            destination = pDt.Clone();
            pDt.DefaultView.Sort = "Characteristic desc";
            pDt = pDt.DefaultView.ToTable(true);
            //DataRow[] foundRows = pDt.Select("Characteristic ASC");
            foreach (DataRow r in pDt.Rows)
            {
                if (r["Characteristic"].ToString() == "ZPKG_SEC_CONTAINER_TYPE")
                    _ContainerType = string.Format("{0}", r["Value"]);
                if (r["Characteristic"].ToString() == "ZPKG_SEC_LID_TYPE")
                    _LidType = string.Format("{0}", r["Value"]);
                if (r["Characteristic"].ToString() == "ZPKG_SEC_CHANGE_POINT")
                    _ChangePoint = string.Format("{0}", r["Value"]);
                DataRow dr = destination.Select("Characteristic='" + r["Characteristic"].ToString() + "'").FirstOrDefault();
                if (dr != null)
                {
                    if (string.IsNullOrEmpty(dr["Value"].ToString()))
                        dr["Value"] = r["Value"]; //changes the Product_name
                    else
                        dr["Value"] += string.Format(";{0}", r["Value"]);
                }
                else
                    destination.ImportRow(r);
                //List<string> list = new List<string>();
                //return String.Join(";", list.ToArray());
            }
            string[] userlevel = { "PA", "PG" };
            foreach (string data in userlevel)
            {
                using (SqlConnection con = new SqlConnection(strConn))
                {
                    SqlCommand cmd = new SqlCommand();
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spInsertMultipleRows";
                    cmd.Parameters.AddWithValue(@"@Description", string.Format("{0}", value["Description"].ToString()));
                    cmd.Parameters.AddWithValue("@Brand", "");
                    cmd.Parameters.AddWithValue("@Primarysize", "");
                    cmd.Parameters.AddWithValue("@Version", value["Version"].ToString());
                    cmd.Parameters.AddWithValue("@ChangePoint", _ChangePoint.ToString());
                    cmd.Parameters.AddWithValue("@MaterialGroup", "");
                    cmd.Parameters.AddWithValue("@CreateBy", _artworkObject.PAUserName);
                    cmd.Parameters.AddWithValue("@RequestNo", value["ID"].ToString());
                    cmd.Parameters.AddWithValue("@userlevel", data);
                    cmd.Parameters.AddWithValue("@PackingStyle", "");
                    cmd.Parameters.AddWithValue("@Packing", "");
                    cmd.Parameters.AddWithValue("@StyleofPrinting", "");
                    cmd.Parameters.AddWithValue("@ContainerType", "");
                    cmd.Parameters.AddWithValue("@LidType", "");
                    cmd.Parameters.AddWithValue("@TotalColour", "");
                    cmd.Parameters.AddWithValue("@StatusApp", string.Format("{0}", 0));
                    cmd.Parameters.AddWithValue("@ProductCode", "");
                    cmd.Parameters.AddWithValue("@FAOZone", "");
                    //cmd.Parameters.AddWithValue("@Plant", string.Format("{0}", _artworkObject.Plant.Replace(',', ';')));
                    cmd.Parameters.AddWithValue("@Plant", string.Format("{0}", _artworkObject.Plant.ToString()));
                    cmd.Parameters.AddWithValue("@Processcolour", "");
                    cmd.Parameters.AddWithValue("@PlantRegisteredNo", "");
                    cmd.Parameters.AddWithValue("@CompanyNameAddress", "");
                    cmd.Parameters.AddWithValue("@PMScolour", "");
                    cmd.Parameters.AddWithValue("@Symbol", "");
                    cmd.Parameters.AddWithValue("@CatchingArea", "");
                    cmd.Parameters.AddWithValue("@CatchingPeriodDate", "");
                    cmd.Parameters.AddWithValue("@Grandof", "");
                    cmd.Parameters.AddWithValue("@Flute", "");
                    cmd.Parameters.AddWithValue("@Vendor", "");
                    cmd.Parameters.AddWithValue("@Dimension", "");
                    cmd.Parameters.AddWithValue("@RSC", "");
                    cmd.Parameters.AddWithValue("@Accessories", "");
                    cmd.Parameters.AddWithValue("@PrintingStyleofPrimary", string.Format("{0}", _artworkObject.PrintingStyleofPrimary));
                    cmd.Parameters.AddWithValue("@PrintingStyleofSecondary", string.Format("{0}", _artworkObject.PrintingStyleofSecondary));
                    cmd.Parameters.AddWithValue("@CustomerDesign", _artworkObject.CustomersDesign + "|" + _artworkObject.CustomersDesignDetail);
                    cmd.Parameters.AddWithValue("@CustomerSpec", _artworkObject.CustomersSpec + "|" + _artworkObject.CustomersSpecDetail);
                    cmd.Parameters.AddWithValue("@CustomerSize", _artworkObject.CustomersSize + "|" + _artworkObject.CustomersSizeDetail);
                    cmd.Parameters.AddWithValue("@CustomerVendor", _artworkObject.CustomerNominatesVendor + "|" + _artworkObject.CustomerNominatesVendorDetail);
                    cmd.Parameters.AddWithValue("@CustomerColor", _artworkObject.CustomerNominatesColorPantone + "|" + _artworkObject.CustomerNominatesColorPantoneDetail);
                    cmd.Parameters.AddWithValue("@CustomerScanable", _artworkObject.CustomersBarcodeScanable + "|" + _artworkObject.CustomersBarcodeScanableDetail);
                    cmd.Parameters.AddWithValue("@CustomerBarcodeSpec", _artworkObject.CustomersBarcodeSpec + "|" + _artworkObject.CustomersBarcodeSpecDetail);
                    cmd.Parameters.AddWithValue("@FirstInfoGroup", string.Format("{0}", _artworkObject.FirstInfoGroup));
                    cmd.Parameters.AddWithValue("@SO", string.Format("{0}", _artworkObject.SONumber));
                    cmd.Parameters.AddWithValue("@PICMkt", string.Format("{0}", _artworkObject.PICMKT));
                    cmd.Parameters.AddWithValue("@SOPlant", string.Format("{0}", _artworkObject.SOPlant));
                    cmd.Parameters.AddWithValue("@Destination", string.Format("{0}", _artworkObject.Destination));
                    cmd.Parameters.AddWithValue("@Remark", string.Format("{0}", _artworkObject.RemarkNoteofPA));
                    cmd.Parameters.AddWithValue("@GrossWeight", "");
                    cmd.Parameters.AddWithValue("@FinalInfoGroup", string.Format("{0}", _artworkObject.FinalInfoGroup));
                    cmd.Parameters.AddWithValue("@Note", string.Format("{0}", _artworkObject.RemarkNoteofPG));
                    cmd.Parameters.AddWithValue("@SheetSize", "");
                    cmd.Parameters.AddWithValue("@Typeof", "");
                    cmd.Parameters.AddWithValue("@TypeofCarton2", "");
                    cmd.Parameters.AddWithValue("@DMSNo", _artworkObject.ArtworkNumber);

                    cmd.Parameters.AddWithValue("@TypeofPrimary", "");
                    cmd.Parameters.AddWithValue("@PrintingSystem", "");
                    cmd.Parameters.AddWithValue("@Direction", "");
                    cmd.Parameters.AddWithValue("@RollSheet", "");
                    cmd.Parameters.AddWithValue("@RequestType", "");
                    cmd.Parameters.AddWithValue("@PlantAddress", "");

                    cmd.Parameters.AddWithValue("@Fixed_Desc", "");
                    cmd.Parameters.AddWithValue("@Inactive", "");
                    cmd.Parameters.AddWithValue("@Catching_Method", "");
                    cmd.Parameters.AddWithValue("@Scientific_Name", "");
                    cmd.Parameters.AddWithValue("@Specie", "");
                    cmd.Parameters.AddWithValue("@SustainMaterial", string.Format("{0}", _artworkObject.SustainMaterial));
                    cmd.Parameters.AddWithValue("@SustainPlastic", string.Format("{0}", _artworkObject.SustainPlastic));
                    cmd.Parameters.AddWithValue("@SustainReuseable", string.Format("{0}", _artworkObject.SustainReuseable));
                    cmd.Parameters.AddWithValue("@SustainRecyclable", string.Format("{0}", _artworkObject.SustainRecyclable));
                    cmd.Parameters.AddWithValue("@SustainComposatable", string.Format("{0}", _artworkObject.SustainComposatable));
                    cmd.Parameters.AddWithValue("@SustainCertification", string.Format("{0}", _artworkObject.SustainCertification));
                    cmd.Parameters.AddWithValue("@SustainCertSourcing", string.Format("{0}", _artworkObject.SustainCertSourcing));
                    cmd.Parameters.AddWithValue("@SustainOther", string.Format("{0}", _artworkObject.SustainOther));
                    cmd.Parameters.AddWithValue("@SusSecondaryPKGWeight", string.Format("{0}", _artworkObject.SusSecondaryPKGWeight));
                    cmd.Parameters.AddWithValue("@SusRecycledContent", string.Format("{0}", _artworkObject.SusRecycledContent));

                    //    cmd.Parameters.AddWithValue("@ArtworkNumber", string.Format("{0}", _artworkObject.ArtworkNumber));
                    //    cmd.Parameters.AddWithValue("@Date", string.Format("{0}", _artworkObject.Date));
                    //    cmd.Parameters.AddWithValue("@Time", string.Format("{0}", _artworkObject.Time));
                    //    cmd.Parameters.AddWithValue("@RecordType", string.Format("{0}", _artworkObject.RecordType));
                    //    cmd.Parameters.AddWithValue("@MaterialNumber", string.Format("{0}", _artworkObject.MaterialNumber));
                    //    cmd.Parameters.AddWithValue("@MaterialDescription", string.Format("{0}", _artworkObject.MaterialDescription));
                    //    cmd.Parameters.AddWithValue("@MaterialCreatedDate", string.Format("{0}", _artworkObject.MaterialCreatedDate));
                    //    cmd.Parameters.AddWithValue("@ArtworkURL", string.Format("{0}", _artworkObject.ArtworkURL));
                    //    cmd.Parameters.AddWithValue("@Status", string.Format("{0}", _artworkObject.Status));
                    //    cmd.Parameters.AddWithValue("@PAUserName", string.Format("{0}", _artworkObject.PAUserName));
                    //    cmd.Parameters.AddWithValue("@PGUserName", string.Format("{0}", _artworkObject.PGUserName));
                    //	  cmd.Parameters.AddWithValue("@ReferenceMaterial", string.Format("{0}", _artworkObject.ReferenceMaterial));
                    //    cmd.Parameters.AddWithValue("@Plant", string.Format("{0}", _artworkObject.Plant));
                    //    cmd.Parameters.AddWithValue("@PrintingStyleofPrimary", string.Format("{0}", _artworkObject.PrintingStyleofPrimary));
                    //    cmd.Parameters.AddWithValue("@PrintingStyleofSecondary", string.Format("{0}", _artworkObject.PrintingStyleofSecondary));
                    //    cmd.Parameters.AddWithValue("@CustomersDesign", string.Format("{0}", _artworkObject.CustomersDesign));
                    //    cmd.Parameters.AddWithValue("@CustomersDesignDetail", string.Format("{0}", _artworkObject.CustomersDesignDetail));
                    //    cmd.Parameters.AddWithValue("@CustomersSpec", string.Format("{0}", _artworkObject.CustomersSpec));
                    //    cmd.Parameters.AddWithValue("@CustomersSpecDetail", string.Format("{0}", _artworkObject.CustomersSpecDetail));
                    //    cmd.Parameters.AddWithValue("@CustomersSize", string.Format("{0}", _artworkObject.CustomersSize));
                    //    cmd.Parameters.AddWithValue("@CustomersSizeDetail", string.Format("{0}", _artworkObject.CustomersSizeDetail));
                    //    cmd.Parameters.AddWithValue("@CustomerNominatesVendor", string.Format("{0}", _artworkObject.CustomerNominatesVendor));
                    //    cmd.Parameters.AddWithValue("@CustomerNominatesVendorDetail", string.Format("{0}", _artworkObject.CustomerNominatesVendorDetail));
                    //    cmd.Parameters.AddWithValue("@CustomerNominatesColorPantone", string.Format("{0}", _artworkObject.CustomerNominatesColorPantone));
                    //    cmd.Parameters.AddWithValue("@CustomerNominatesColorPantoneDetail", string.Format("{0}", _artworkObject.CustomerNominatesColorPantoneDetail));
                    //    cmd.Parameters.AddWithValue("@CustomersBarcodeScanable", string.Format("{0}", _artworkObject.CustomersBarcodeScanable));
                    //    cmd.Parameters.AddWithValue("@CustomersBarcodeScanableDetail", string.Format("{0}", _artworkObject.CustomersBarcodeScanableDetail));
                    //    cmd.Parameters.AddWithValue("@CustomersBarcodeSpec", string.Format("{0}", _artworkObject.CustomersBarcodeSpec));
                    //    cmd.Parameters.AddWithValue("@CustomersBarcodeSpecDetail", string.Format("{0}", _artworkObject.CustomersBarcodeSpecDetail));
                    //    cmd.Parameters.AddWithValue("@FirstInfoGroup", string.Format("{0}", _artworkObject.FirstInfoGroup));
                    //    cmd.Parameters.AddWithValue("@SONumber", string.Format("{0}", _artworkObject.SONumber));
                    //    cmd.Parameters.AddWithValue("@SOitem", string.Format("{0}", _artworkObject.SOitem));
                    //    cmd.Parameters.AddWithValue("@SOPlant", string.Format("{0}", _artworkObject.SOPlant));
                    //    cmd.Parameters.AddWithValue("@PICMKT", string.Format("{0}", _artworkObject.PICMKT));
                    //    cmd.Parameters.AddWithValue("@Destination", string.Format("{0}", _artworkObject.Destination));
                    //    cmd.Parameters.AddWithValue("@RemarkNoteofPA", string.Format("{0}", _artworkObject.RemarkNoteofPA));
                    //    cmd.Parameters.AddWithValue("@FinalInfoGroup", string.Format("{0}", _artworkObject.FinalInfoGroup));
                    //    cmd.Parameters.AddWithValue("@RemarkNoteofPG", string.Format("{0}", _artworkObject.RemarkNoteofPG));
                    //    cmd.Parameters.AddWithValue("@CompleteInfoGroup", string.Format("{0}", _artworkObject.CompleteInfoGroup));
                    //    cmd.Parameters.AddWithValue("@ProductionExpirydatesystem", string.Format("{0}", _artworkObject.ProductionExpirydatesystem));
                    //    cmd.Parameters.AddWithValue("@Seriousnessofcolorprinting", string.Format("{0}", _artworkObject.Seriousnessofcolorprinting));
                    //    cmd.Parameters.AddWithValue("@CustIngreNutritionAnalysis", string.Format("{0}", _artworkObject.CustIngreNutritionAnalysis));
                    //    cmd.Parameters.AddWithValue("@ShadeLimit", string.Format("{0}", _artworkObject.ShadeLimit));
                    //    cmd.Parameters.AddWithValue("@PackageQuantity", string.Format("{0}", _artworkObject.PackageQuantity));
                    //    cmd.Parameters.AddWithValue("@WastePercent", string.Format("{0}", _artworkObject.WastePercent));
                    cmd.Connection = con;
                    con.Open();
                    cmd.ExecuteNonQuery();
                    con.Close();
                }
            }
            foreach (DataRow item in destination.Rows)
            {
                string charac = item["Characteristic"].ToString();
                if (charac.ToString().Contains("ZPKG_SEC_PRIMARY_SIZE"))
                {
                    item["Value"] = cs.ReadItems(@"SELECT top 1 code from MasPrimarySize a where isnull(Inactive,'')<>'X' and UPPER(a.Description)=N'" + item["Value"].ToString().ToUpper()
                        + "' and UPPER(a.ContainerType)=N'" + _ContainerType.ToUpper() + "' and (a.DescriptionType) =N'" + _LidType.ToUpper() + "'");
                }
                if (charac.ToString().Contains("ZPKG_SEC_ACCESSORIES"))
                {
                    item["Value"] = Insert(item["Value"].ToString());
                }
                //Initialize SQL Server Connection
                SqlConnection cn = new SqlConnection(strConn);
                SqlCommand cmd = new SqlCommand("spUpdateSapMaterial", cn);
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.AddWithValue("@Description", string.Format("{0}", item["Description"]));
                //cmd.Parameters.AddWithValue("@ArtworkNumber", string.Format("{0}", _artworkObject.ArtworkNumber));
                //cmd.Parameters.AddWithValue("@Date", string.Format("{0}", item.Date));
                cmd.Parameters.AddWithValue("@Value", string.Format("{0}", item["Value"]));
                cmd.Parameters.AddWithValue("@Group", string.Format("{0}", _group));
                cmd.Parameters.AddWithValue("@Characteristic", string.Format("{0}", charac));
                cmd.Parameters.AddWithValue("@Keys", string.Format("{0}", value["ID"]));
                // Running the query.
                cn.Open();
                cmd.ExecuteNonQuery();
                cn.Close();
                //Set insert query
                //Dictionary<string, object> paras = new Dictionary<string, object>();
                //paras.Add("Title", string.Format("{0}", item.Characteristic));
                //paras.Add("MaterialType", _group.ToString());
                //using (SqlDataReader results = cs.executeProcedure("spReadCharacteristic", paras))
                //{
                //    while (results.Read())
                //    {
                //        //do something with the rows returned
                //        results["shortname"].ToString();
                //    }
                //} 
                //jumpto:
                //Context.Response.Write(item.Characteristic);
            }
            //foreach (var p in _itemsArtwork)
            //{
            //    if (p.Characteristic == "ZPKG_SEC_PRIMARY_SIZE")
            //    {
            //        using (SqlConnection cn = new SqlConnection(strConn))
            //        {
            //            using (SqlCommand cmd = new SqlCommand("spSecPrimarySize"))
            //            {
            //                cmd.Parameters.AddWithValue("@PrimarySize", string.Format("{0}", p.Value));
            //                cmd.Parameters.AddWithValue("@Keys", string.Format("{0}", value["ID"]));
            //            }
            //        }
            //    }
            //}
            //skip approve
            using (SqlConnection CN = new SqlConnection(strConn))
            {
                string qry = "spUpdateArtwork";
                SqlCommand SqlCom = new SqlCommand(qry, CN);
                SqlCom.CommandType = CommandType.StoredProcedure;
                SqlCom.Parameters.Add(new SqlParameter("@Keys", value["ID"].ToString()));
                CN.Open();
                SqlCom.ExecuteNonQuery();
                CN.Close();
            }

            _Subject = string.Format("system iGrid Request No.:{0}", value["DocumentNo"]);
            _Body = @"integration between iGrid and Artwork system (xECM)<br/> CreateBy : "
            + cs.Getuser(_artworkObject.PAUserName, "fullname") + " <br/> Material Group : " + _group;
        }
    jumptoexit:
        cs.sendemail(@"voravut.somboornpong@thaiunion.com", "", _Body, _Subject, "");
        return Keys.ToString();
        //Context.Response.Write(JsonConvert.SerializeObject(Results));   
    }
    void deletefile(string datapath)
    {
        string file = Server.MapPath(datapath);
        if (Directory.Exists(Path.GetDirectoryName(file)))
            File.Delete(file);
    }
    void att(ArtworkObject _artworkObject, string Keys)
    {
        string datapath = "~/FileTest/hArtwork" + Keys.ToString() + ".xml";
        using (FileStream fs = new FileStream(Server.MapPath(datapath), FileMode.Create))
        {
            new XmlSerializer(typeof(ArtworkObject)).Serialize(fs, _artworkObject);
        }
    }
    [WebMethod]
    public void uploadfile(attachment ro)
    {
        //string[] files = Directory.GetFiles(Server.MapPath(@"~/FileTest/" + Data));
        //foreach (string file in files)
        //{
        //Read File Bytes into a byte array
        //byte[] FileData = ReadFile(file);

        //Initialize SQL Server Connection
        using (SqlConnection CN = new SqlConnection(strConn))
        {
            string qry = "insert into tblFiles values (@Name,@ContentType,@Data,@MatDoc,@ActiveBy)";
            SqlCommand SqlCom = new SqlCommand(qry, CN);
            //We are passing Original File Path and file byte data as sql parameters.
            SqlCom.Parameters.Add(new SqlParameter("@Name", ro.Name));
            SqlCom.Parameters.Add(new SqlParameter("@ContentType", ro.ContentType));
            SqlCom.Parameters.Add(new SqlParameter("@Data", ro.Data));
            SqlCom.Parameters.Add(new SqlParameter("@MatDoc", ro.MatDoc));
            SqlCom.Parameters.Add(new SqlParameter("@ActiveBy", ro.ActiveBy));
            //Open connection and execute insert query.
            CN.Open();
            SqlCom.ExecuteNonQuery();
            CN.Close();
        }
        //string folderPath = Server.MapPath(@"~/FileTest/" + Data);
        //Directory.Delete(folderPath, true);
    }
    //Open file in to a filestream and read data in a byte array.
    byte[] ReadFile(string sPath)
    {
        //Initialize byte array with a null value initially.
        byte[] data = null;

        //Use FileInfo object to get file size.
        FileInfo fInfo = new FileInfo(sPath);
        long numBytes = fInfo.Length;

        //Open FileStream to read file
        FileStream fStream = new FileStream(sPath, FileMode.Open, FileAccess.Read);

        //Use BinaryReader to read file stream into byte array.
        BinaryReader br = new BinaryReader(fStream);

        //When you use BinaryReader, you need to supply number of bytes to read from file.
        //In this case we want to read entire file. So supplying total number of bytes.
        data = br.ReadBytes((int)numBytes);

        //Close BinaryReader
        br.Close();

        //Close FileStream
        fStream.Close();

        return data;
    }

    [WebMethod]
    public void SendEmailUpdateMaster(string _name)
    {
        //string datapath = "~/FileTest/" + _name;
        string _email = "";
        string _Id = "";
        string _Description = "";
        string _Body = "";
        string _Attached = "";
        string SubChanged_Id = cs.ReadItems(@"select cast(substring('"
        + _name.ToString() + "',2,len('" + _name.ToString() + "')-1) as nvarchar(max)) value");
        testsendmaster(SubChanged_Id);
        DataTable dt = new DataTable();
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spsendemail_upm";
            cmd.Parameters.AddWithValue("@Changed_Id", _name.ToString());
            cmd.Connection = con;
            con.Open();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
        }
        foreach (DataRow dr in dt.Rows)
        {
            _email = dr["Email"].ToString();
            _Id = dr["Id"].ToString();
            _Description = dr["Description"].ToString();
            _Body = dr["Body"].ToString();
            _Attached = dr["attached"].ToString();
        }
        //        MailMessage msg = new MailMessage();
        //        string[] words = _email.Split(';');
        //        foreach (string word in words)
        //        {
        //            msg.To.Add(new MailAddress(word));
        //            //Console.WriteLine(word);
        //        }
        //        //msg.To.Add(new MailAddress(_email));
        //        msg.From = new MailAddress("wshuttleadm@thaiunion.com");
        //        msg.Subject = "Maintained characteristic master data in SAP" + "[" + _Body.Substring(0, 6) + "]";
        //        //msg.Body = "Id  " + _Id.ToString() + "Description  " + _Description.ToString() + " Changed";
        //        //msg.Body = "Maintained characteristic master completed";
        //        msg.Body = _Body;
        //        //msg.Attachments.Add(new System.Net.Mail.Attachment(_Attached));
        //        msg.IsBodyHtml = true;
        //
        //        SmtpClient client = new SmtpClient();
        //        client.UseDefaultCredentials = false;
        //        client.Credentials = new System.Net.NetworkCredential("wshuttleadm@thaiunion.com", "WSP@ss2018");
        //        client.Port = 587; // You can use Port 25 if 587 is blocked (mine is!)
        //        client.Host = "smtp.office365.com";
        //        client.DeliveryMethod = SmtpDeliveryMethod.Network;
        //        client.EnableSsl = true;
        //        try
        //        {
        //            client.Send(msg);
        //            Context.Response.Write("Message Sent Succesfully");
        //        }
        //        catch (Exception ex)
        //        {
        //            Context.Response.Write(ex.ToString());
        //        }
        cs.sendemail(_email, "", _Body,
            "Maintained characteristic master data in SAP" + "[" + _Body.Substring(0, 6) + "]",
            _Attached);
    }

    [WebMethod]
    public void SendEmail(string _name)
    {
        //string datapath = "~/FileTest/" + _name;
        string _email = "";
        string _Material = "";
        string _Description = "";
        string _Body = "";
        string _Attached = "";

        DataTable dt = new DataTable();
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spsendemail";
            cmd.Parameters.AddWithValue("@Material", _name.ToString());
            cmd.Connection = con;
            con.Open();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
        }
        foreach (DataRow dr in dt.Rows)
        {
            _email = dr["Email"].ToString();
            _Material = dr["Material"].ToString();

            _Description = dr["Description"].ToString();
            _Body = dr["body"].ToString();
            _Attached = dr["attached"].ToString();
        }
        //        MailMessage msg = new MailMessage();
        //        string[] words = _email.Split(';');
        //        foreach (string word in words)
        //        {
        //            msg.To.Add(new MailAddress(word));
        //            //Console.WriteLine(word);
        //        }
        //        //msg.To.Add(new MailAddress(_email));
        //        msg.From = new MailAddress("wshuttleadm@thaiunion.com");
        //        msg.Subject = "System SEC PKG Template is created No. : " + _Material.ToString() + "/" + _Description.ToString() + "/" + "Create Material";
        //        //msg.Body = "Material  " + _Material.ToString() + " Created";
        //        msg.Body = _Body;
        //        msg.Attachments.Add(new System.Net.Mail.Attachment(_Attached));
        //        msg.IsBodyHtml = true;
        //
        //        SmtpClient client = new SmtpClient();
        //        client.UseDefaultCredentials = false;
        //        client.Credentials = new System.Net.NetworkCredential("wshuttleadm@thaiunion.com", "WSP@ss2018");
        //        client.Port = 587; // You can use Port 25 if 587 is blocked (mine is!)
        //        client.Host = "smtp.office365.com";
        //        client.DeliveryMethod = SmtpDeliveryMethod.Network;
        //        client.EnableSsl = true;
        //        try
        //        {
        //            client.Send(msg);
        //            Context.Response.Write("Message Sent Succesfully");
        //        }
        //        catch (Exception ex)
        //        {
        //            Context.Response.Write(ex.ToString());
        //        }
        cs.sendemail(_email, "", _Body,
            "System SEC PKG Template is created No. : " + _Material.ToString() + "/" + _Description.ToString() + "/" + "Create Material",
            _Attached);
        string _DocumentNo = cs.ReadItems("select DocumentNo from SapMaterial Where Material='" + _Material.ToString() + "' and StatusApp<>5");
        OutboundArtwork(_DocumentNo.ToString());
        //string senderID = "voravut.somb@gmail.com";
        //string senderPassword = "063446620";
        //string result = "Email Sent Successfully";

        //string body = " " + _name + " has sent an email from " + _email;
        //body += "Phone : " + _phone;
        //body += _description;
        //try
        //{
        //    MailMessage mail = new MailMessage();
        //    mail.To.Add("voravut.somboornpong@thaiunion.com");
        //    mail.From = new MailAddress(senderID);
        //    mail.Subject = "My Test Email!";
        //    mail.Body = body;
        //    mail.IsBodyHtml = true;
        //    SmtpClient smtp = new SmtpClient();
        //    smtp.Host = "smtp.gmail.com"; //Or Your SMTP Server Address
        //    smtp.Credentials = new System.Net.NetworkCredential(senderID, senderPassword);
        //    smtp.Port = 587;
        //    smtp.EnableSsl = true;
        //    smtp.Send(mail);
        //}
        //catch (Exception ex)
        //{
        //    result = "problem occurred";
        //    Context.Response.Write("Exception in sendEmail:" + ex.Message);
        //}
        //Context.Response.Write(result);
    }

    bool IsValidEmail(string email)
    {
        try
        {
            var addr = new System.Net.Mail.MailAddress(email);
            return addr.Address == email;
        }
        catch
        {
            return false;
        }
    }
    [WebMethod]
    public void jobalertemail(string data)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spalertemail";
            cmd.Connection = con;
            con.Open();
            //cmd.ExecuteNonQuery();
            var getValue = cmd.ExecuteScalar();
            if (getValue.ToString() == "0")
            {
                MailMessage msg = new MailMessage();
                msg.To.Add(new MailAddress("voravut.somboornpong@thaiunion.com"));
                List<string> li = new List<string>();
                li.Add("Nongrat.Jantarasuwan@thaiunion.com");
                li.Add("pornnicha.thanarak@thaiunion.com");
                li.Add("Walaipan.Supoltawanitch@thaiunion.com");
                li.Add("Pornpimon.Bouban@thaiunion.com");
                //li.Add("saihacksoft@gmail.com");  
                msg.CC.Add(string.Join<string>(",", li)); // Sending CC  
                msg.From = new MailAddress("wshuttleadm@thaiunion.com");
                msg.Subject = "[iGrid Support] Winshuttle down";
                msg.Body = "Dear All <br/>job Winshuttle fail.";
                msg.IsBodyHtml = true;
                SmtpClient client = new SmtpClient();
                client.UseDefaultCredentials = false;
                client.Credentials = new System.Net.NetworkCredential("wshuttleadm@thaiunion.com", "WSP@ss2018");
                client.Port = 587; // You can use Port 25 if 587 is blocked (mine is!)
                client.Host = "smtp.office365.com";
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.EnableSsl = true;
                client.Send(msg);
            }
            con.Close();
        }
        Context.Response.Write("success");
    }
    [WebMethod]
    public void GetUpdate(string sName)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spUpdateProcedure";
            cmd.Parameters.AddWithValue("@Material", sName);
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
            //alert email template
            SendEmail(sName);
        }
        Context.Response.Write("success");
    }
    [WebMethod]
    public DataSet GetImpactmat(string sName)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetImpactmat";
            cmd.Parameters.AddWithValue("@Active", sName);
            cmd.Connection = con;
            con.Open();
            DataSet oDataset = new DataSet();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(oDataset);
            con.Close();
            return oDataset;
        }
    }

    [WebMethod]
    public DataSet GetMasterData(string sName)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetMasterData";
            cmd.Parameters.AddWithValue("@Active", sName);
            cmd.Connection = con;
            con.Open();
            DataSet oDataset = new DataSet();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(oDataset);
            con.Close();
            return oDataset;
        }
    }
    [WebMethod]
    public DataSet GetQuery(string sName)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spQuery";
            cmd.Parameters.AddWithValue("@Material", sName);
            cmd.Connection = con;
            con.Open();
            DataSet oDataset = new DataSet();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(oDataset);
            con.Close();
            return oDataset;
        }
    }
    //[WebMethod]
    //public int SaveDocument(Byte[] filebyte)
    //{
    //    int adInteger; 
    //    using (SqlConnection con = new SqlConnection(strConn))
    //    {
    //        using (SqlCommand cmd = new SqlCommand("INSERT INTO masmaterial(docbinaryarray)  VALUES(@docbinaryarray);SELECT SCOPE_IDENTITY();",con))
    //        {
    //            cmd.CommandType = CommandType.Text;
    //            cmd.Parameters.AddWithValue("@docbinaryarray", filebyte);
    //            con.Open();
    //            adInteger = (int)cmd.ExecuteScalar();

    //            if (con.State == System.Data.ConnectionState.Open) con.Close();
    //            return adInteger;
    //        }
    //    }
    //}
    //    [WebMethod]
    //    public void Gettest(string name)
    //    {
    //        using (SqlConnection con = new SqlConnection(strConn))
    //        {
    //            SqlCommand cmd = new SqlCommand();
    //            cmd.CommandType = CommandType.StoredProcedure;
    //            cmd.CommandText = "sptest";
    //            cmd.Parameters.AddWithValue("@Param", name);
    //            cmd.Connection = con;
    //            con.Open();
    //            DataTable dt = new DataTable();
    //            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
    //            oAdapter.Fill(dt);
    //            con.Close();
    //            Context.Response.Write(JsonConvert.SerializeObject(dt));
    //        }
    //    }

    [WebMethod]
    public void UpdatePrimarySize(objPrimarySize ro)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spUpdatePrimarySize";
            cmd.Parameters.AddWithValue("@Item", ro.Item);
            cmd.Parameters.AddWithValue("@Id", ro.Id);
            cmd.Parameters.AddWithValue("@Description", ro.Description);
            cmd.Parameters.AddWithValue("@Can", ro.Can);
            cmd.Parameters.AddWithValue("@LidType", ro.LidType);
            cmd.Parameters.AddWithValue("@ContainerType", ro.ContainerType);
            cmd.Parameters.AddWithValue("@DescriptionType", ro.DescriptionType);
            cmd.Parameters.AddWithValue("@Changed_Action", ro.Changed_Action);
            cmd.Parameters.AddWithValue("@Changed_By", ro.Changed_By);
            cmd.Parameters.AddWithValue("@Active", ro.Active);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }

    //[WebMethod]
    //public void UpdateMaster(string json){
    //List<objTransMaster> _x = JsonConvert.DeserializeObject<List<objTransMaster>>(json);
    //objTransMaster ro = _x[0];
    //    using (SqlConnection con = new SqlConnection(strConn))
    //    {
    //        SqlCommand cmd = new SqlCommand();
    //        cmd.CommandType = CommandType.StoredProcedure;
    //        cmd.CommandText = "spUpdateMaster";
    //        cmd.Parameters.AddWithValue("@Changed_Tabname", ro.Changed_Tabname);
    //        cmd.Parameters.AddWithValue("@Changed_Charname", ro.Changed_Charname);
    //        cmd.Parameters.AddWithValue("@Old_Id", ro.Old_Id);
    //        cmd.Parameters.AddWithValue("@Id", ro.Id);
    //        cmd.Parameters.AddWithValue("@Description", ro.Description);
    //        cmd.Parameters.AddWithValue("@Changed_Action", ro.Changed_Action);
    //        cmd.Parameters.AddWithValue("@Changed_By", ro.Changed_By);
    //        cmd.Parameters.AddWithValue("@Active", ro.Active);
    //        cmd.Parameters.AddWithValue("@Material_Group", ro.Material_Group);
    //        cmd.Parameters.AddWithValue("@Material_Type", ro.Material_Type);
    //        cmd.Parameters.AddWithValue("@DescriptionText", ro.DescriptionText);
    //        cmd.Parameters.AddWithValue("@Can", ro.Can);
    //        cmd.Parameters.AddWithValue("@LidType", ro.LidType);
    //        cmd.Parameters.AddWithValue("@ContainerType", ro.ContainerType);
    //        cmd.Parameters.AddWithValue("@DescriptionType", ro.DescriptionType);
    //        cmd.Parameters.AddWithValue("@user_name", ro.user_name);
    //        cmd.Parameters.AddWithValue("@fn", ro.fn);
    //        cmd.Parameters.AddWithValue("@FirstName", ro.FirstName);
    //        cmd.Parameters.AddWithValue("@LastName", ro.LastName);
    //        cmd.Parameters.AddWithValue("@Email", ro.Email);
    //        cmd.Parameters.AddWithValue("@Authorize_ChangeMaster", ro.Authorize_ChangeMaster);
    //        cmd.Parameters.AddWithValue("@PrimaryCode", ro.PrimaryCode);
    //        cmd.Parameters.AddWithValue("@GroupStyle", ro.GroupStyle);
    //        cmd.Parameters.AddWithValue("@PackingStyle", ro.PackingStyle);
    //        cmd.Parameters.AddWithValue("@RefStyle", ro.RefStyle);
    //        cmd.Parameters.AddWithValue("@Packsize", ro.Packsize);
    //        cmd.Parameters.AddWithValue("@BaseUnit", ro.BaseUnit);
    //        cmd.Parameters.AddWithValue("@RegisteredNo", ro.RegisteredNo);
    //        cmd.Parameters.AddWithValue("@Address", ro.Address);
    //        cmd.Parameters.AddWithValue("@Plant", ro.Plant);

    //        cmd.Parameters.AddWithValue("@Product_Group", ro.Product_Group);
    //        cmd.Parameters.AddWithValue("@Product_GroupDesc", ro.Product_GroupDesc);
    //        cmd.Parameters.AddWithValue("@PRD_Plant", ro.PRD_Plant);

    //        cmd.Parameters.AddWithValue("@WHNumber", ro.WHNumber);
    //        cmd.Parameters.AddWithValue("@StorageType", ro.StorageType);
    //        cmd.Parameters.AddWithValue("@LE_Qty", ro.LE_Qty);
    //        cmd.Parameters.AddWithValue("@Storage_UnitType", ro.Storage_UnitType);

    //        cmd.Connection = con;
    //        con.Open();
    //        DataTable dt = new DataTable();
    //        SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
    //        oAdapter.Fill(dt);
    //        con.Close();
    //        Context.Response.Write(JsonConvert.SerializeObject(dt));
    //    }
    //}

    [WebMethod]
    public void UpdateTransMaster(string json)
    {
        List<objTransMaster> _x = JsonConvert.DeserializeObject<List<objTransMaster>>(json);
        objTransMaster ro = _x[0];
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spUpdateTransMaster";
            cmd.Parameters.AddWithValue("@Changed_Tabname", ro.Changed_Tabname);
            cmd.Parameters.AddWithValue("@Changed_Charname", ro.Changed_Charname);
            cmd.Parameters.AddWithValue("@Old_Id", ro.Old_Id);
            cmd.Parameters.AddWithValue("@Id", ro.Id);
            cmd.Parameters.AddWithValue("@Old_Description", ro.Old_Description);
            cmd.Parameters.AddWithValue("@Description", ro.Description);
            cmd.Parameters.AddWithValue("@Changed_Action", ro.Changed_Action);
            cmd.Parameters.AddWithValue("@Changed_By", ro.Changed_By);
            cmd.Parameters.AddWithValue("@Active", ro.Active);
            cmd.Parameters.AddWithValue("@Material_Group", ro.Material_Group);
            cmd.Parameters.AddWithValue("@Material_Type", ro.Material_Type);
            cmd.Parameters.AddWithValue("@DescriptionText", ro.DescriptionText);
            cmd.Parameters.AddWithValue("@Can", ro.Can);
            cmd.Parameters.AddWithValue("@LidType", ro.LidType);
            cmd.Parameters.AddWithValue("@ContainerType", ro.ContainerType);
            cmd.Parameters.AddWithValue("@DescriptionType", ro.DescriptionType);
            cmd.Parameters.AddWithValue("@user_name", ro.user_name);
            cmd.Parameters.AddWithValue("@fn", ro.fn);
            cmd.Parameters.AddWithValue("@FirstName", ro.FirstName);
            cmd.Parameters.AddWithValue("@LastName", ro.LastName);
            cmd.Parameters.AddWithValue("@Email", ro.Email);
            cmd.Parameters.AddWithValue("@Authorize_ChangeMaster", ro.Authorize_ChangeMaster);
            cmd.Parameters.AddWithValue("@PrimaryCode", ro.PrimaryCode);
            cmd.Parameters.AddWithValue("@GroupStyle", ro.GroupStyle);
            cmd.Parameters.AddWithValue("@PackingStyle", ro.PackingStyle);
            cmd.Parameters.AddWithValue("@RefStyle", ro.RefStyle);
            cmd.Parameters.AddWithValue("@Packsize", ro.Packsize);
            cmd.Parameters.AddWithValue("@BaseUnit", ro.BaseUnit);
            cmd.Parameters.AddWithValue("@TypeofPrimary", ro.TypeofPrimary);
            cmd.Parameters.AddWithValue("@RegisteredNo", ro.RegisteredNo);
            cmd.Parameters.AddWithValue("@Address", ro.Address);
            cmd.Parameters.AddWithValue("@Plant", ro.Plant);

            cmd.Parameters.AddWithValue("@Product_Group", ro.Product_Group);
            cmd.Parameters.AddWithValue("@Product_GroupDesc", ro.Product_GroupDesc);
            cmd.Parameters.AddWithValue("@PRD_Plant", ro.PRD_Plant);

            cmd.Parameters.AddWithValue("@WHNumber", ro.WHNumber);
            cmd.Parameters.AddWithValue("@StorageType", ro.StorageType);
            cmd.Parameters.AddWithValue("@LE_Qty", ro.LE_Qty);
            cmd.Parameters.AddWithValue("@Storage_UnitType", ro.Storage_UnitType);

            cmd.Parameters.AddWithValue("@Changed_Reason", ro.Changed_Reason);

            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    [WebMethod]
    public void GetComplateMaterial2(string name, string sCondition, string sUser, string FrDt, string ToDt)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spComplateMaterial2";
            cmd.Parameters.AddWithValue("@Material", name);
            cmd.Parameters.AddWithValue("@Condition", sCondition);
            cmd.Parameters.AddWithValue("@User", sUser);
            cmd.Parameters.AddWithValue("@FrDt", FrDt);
            cmd.Parameters.AddWithValue("@ToDt", ToDt);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    [WebMethod]
    public void GetTrackingReport(string name, string sCondition, string sUser, string FrDt, string ToDt)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetTrackingReport";
            cmd.Parameters.AddWithValue("@where", name);
            cmd.Parameters.AddWithValue("@Condition", sCondition);
            cmd.Parameters.AddWithValue("@User", sUser);
            cmd.Parameters.AddWithValue("@FrDt", FrDt);
            cmd.Parameters.AddWithValue("@ToDt", ToDt);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            if (name == "")
                Context.Response.Write(JsonConvert.SerializeObject(dt));
            else
            {
                DataTable dtTarget = new DataTable();
                dtTarget = dt.Clone();
                DataRow[] rowsToCopy;
                rowsToCopy = dt.Select("Material LIKE '%" + name.ToString().Trim() + "%' or DMSNo LIKE '%" + name.ToString().Trim() + "%' or DocumentNo LIKE '%" + name.ToString().Trim() + "%'");
                foreach (DataRow temp in rowsToCopy)
                {
                    dtTarget.ImportRow(temp);
                }
                Context.Response.Write(JsonConvert.SerializeObject(dtTarget));
            }
        }
    }
    //    [WebMethod]
    //    public void GetTrackingReport(string name, string sCondition, string sUser, string FrDt, string ToDt)
    //    {
    //        using (SqlConnection con = new SqlConnection(strConn))
    //        {
    //            SqlCommand cmd = new SqlCommand();
    //            cmd.CommandType = CommandType.StoredProcedure;
    //            cmd.CommandText = "spGetTrackingReport";
    //            cmd.Parameters.AddWithValue("@Material", name);
    //            cmd.Parameters.AddWithValue("@Condition", sCondition);
    //            cmd.Parameters.AddWithValue("@User", sUser);
    //            cmd.Parameters.AddWithValue("@FrDt", FrDt);
    //            cmd.Parameters.AddWithValue("@ToDt", ToDt);
    //            cmd.Connection = con;
    //            con.Open();
    //            DataTable dt = new DataTable();
    //            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
    //            oAdapter.Fill(dt);
    //            con.Close();
    //            Context.Response.Write(JsonConvert.SerializeObject(dt));
    //        }
    //    }
    [WebMethod]
    public void GetMasterDataLog(string name, string FrDt, string ToDt, string sUser, string Shortname)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spMasterLog";
            cmd.Parameters.AddWithValue("@Material", name);
            //cmd.Parameters.AddWithValue("@Condition", sCondition);
            cmd.Parameters.AddWithValue("@FrDt", FrDt);
            cmd.Parameters.AddWithValue("@ToDt", ToDt);
            cmd.Parameters.AddWithValue("@User", sUser);
            cmd.Parameters.AddWithValue("@Shortname", Shortname);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    [WebMethod]
    public void GetHistoryLog(string name, string FrDt, string ToDt)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spHistoryLog";
            cmd.Parameters.AddWithValue("@material", name);
            //cmd.Parameters.AddWithValue("@Condition", sCondition);
            cmd.Parameters.AddWithValue("@FrDt", FrDt);
            cmd.Parameters.AddWithValue("@ToDt", ToDt);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    [WebMethod]
    public void GetKPILog(string LayOut, string FrDt, string ToDt)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetKPILog2";
            cmd.Parameters.AddWithValue("@LayOut", LayOut);
            //cmd.Parameters.AddWithValue("@Condition", sCondition);
            cmd.Parameters.AddWithValue("@FrDt", FrDt);
            cmd.Parameters.AddWithValue("@ToDt", ToDt);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    [WebMethod]
    public void GetKPILog_Summarize(string LayOut, string FrDt, string ToDt)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetKPILog_Summarize";
            cmd.Parameters.AddWithValue("@LayOut", LayOut);
            //cmd.Parameters.AddWithValue("@Condition", sCondition);
            cmd.Parameters.AddWithValue("@FrDt", FrDt);
            cmd.Parameters.AddWithValue("@ToDt", ToDt);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    [WebMethod]
    public void GetSelectCrossColumn2(string sName, string FrDt, string ToDt)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "WebAPP_SelectCrossColumn2";
            cmd.Parameters.AddWithValue("@material", sName);
            //cmd.Parameters.AddWithValue("@Condition", sCondition);
            cmd.Parameters.AddWithValue("@FrDt", FrDt);
            cmd.Parameters.AddWithValue("@ToDt", ToDt);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    [WebMethod]
    public void Getulogin(string username, string password)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spulogin";
            cmd.Parameters.AddWithValue("@username", username);
            cmd.Parameters.AddWithValue("@password", password);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    [WebMethod]
    public void Getreason(string Id, string fn)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spreasonandrejection";
            cmd.Parameters.AddWithValue("@Id", Id);
            cmd.Parameters.AddWithValue("@fn", fn);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    //void Update ( string StatusApp, string Id , string User)
    //{
    //    using (SqlConnection oConn = new SqlConnection(strConn))
    //    {
    //        oConn.Open();
    //        string strQuery = "update TransApprove set StatusApp=@StatusApp Where Matdoc=@Id and fn = (select fn from ulogin where [user_name]=@user)";
    //        SqlCommand cmd = new SqlCommand(strQuery,oConn);
    //        cmd.Parameters.AddWithValue("@Id", Id);
    //        cmd.Parameters.AddWithValue("@User", User);
    //        cmd.Parameters.AddWithValue("@StatusApp", StatusApp);
    //        int rows = cmd.ExecuteNonQuery();
    //        oConn.Close();
    //    }
    //}
    [WebMethod]
    public void saveaprrove(AppObject ro)
    {
        //string datapath = "~/FileTest/" + data + ".json";
        //using (StreamReader sr = new StreamReader(Server.MapPath(datapath)))
        //{
        //string json = sr.ReadToEnd();
        //List<AppObject> ro = JsonConvert.DeserializeObject<List<AppObject>>(json);
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spupdateApprove";
            cmd.Parameters.AddWithValue("@ActiveBy", ro.ActiveBy);
            cmd.Parameters.AddWithValue("@Id", ro.Id);
            cmd.Parameters.AddWithValue("@fn", ro.fn);
            cmd.Parameters.AddWithValue("@StatusApp", ro.StatusApp);
            cmd.Parameters.AddWithValue("@Remark", ro.Remark);
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
        Context.Response.Write("success");
        //}
        //deletefile(datapath);
    }
    [WebMethod]
    public void saveActive(string data)
    {
        string datapath = "~/FileTest/" + data + ".json";
        using (StreamReader sr = new StreamReader(Server.MapPath(datapath)))
        {
            string json = sr.ReadToEnd();
            List<AppObject> ro = JsonConvert.DeserializeObject<List<AppObject>>(json);

            //json = json.Replace('@', '&');
            //List<AppObject> _x = JsonConvert.DeserializeObject<List<AppObject>>(json);
            //AppObject ro = _x[0];
            using (SqlConnection con = new SqlConnection(strConn))
            {
                SqlCommand cmd = new SqlCommand();
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "spupdateApprove";
                cmd.Parameters.AddWithValue("@ActiveBy", ro[0].ActiveBy);
                cmd.Parameters.AddWithValue("@Id", ro[0].Id);
                cmd.Parameters.AddWithValue("@fn", ro[0].fn);
                cmd.Parameters.AddWithValue("@StatusApp", ro[0].StatusApp);
                cmd.Parameters.AddWithValue("@Remark", ro[0].Remark);
                cmd.Connection = con;
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
            }
            Context.Response.Write("success");
            //}
            //deletefile(datapath);
        }
    }
    [WebMethod]
    public void saveapprove(string json)
    {
        json = json.Replace('@', '&');
        List<AppObject> _x = JsonConvert.DeserializeObject<List<AppObject>>(json);
        AppObject ro = _x[0];
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spupdateApprove";
            cmd.Parameters.AddWithValue("@ActiveBy", ro.ActiveBy);
            cmd.Parameters.AddWithValue("@Id", ro.Id);
            cmd.Parameters.AddWithValue("@fn", ro.fn);
            cmd.Parameters.AddWithValue("@StatusApp", ro.StatusApp);
            cmd.Parameters.AddWithValue("@Remark", ro.Remark);
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
        Context.Response.Write("success");
        //}
        //deletefile(datapath);
    }
    [WebMethod]
    public void GetDelDocument(string Id)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spDelDocument";
            cmd.Parameters.AddWithValue("@Id", Id);
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
            Context.Response.Write("success");
        }
    }
    [WebMethod]
    public void GetsaveInfoGroup(string Id, string InfoGroup, string user, string Check_PChanged)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spsaveInfoGroup";
            cmd.Parameters.AddWithValue("@Id", Id);
            cmd.Parameters.AddWithValue("@InfoGroup", InfoGroup);
            cmd.Parameters.AddWithValue("@user", user);
            cmd.Parameters.AddWithValue("@Check_PChanged", Check_PChanged);
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
            Context.Response.Write("success");
        }
    }
 
    [WebMethod]
    public void savecopy(string Id, string copy, string user)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spsaverefnumber";
            cmd.Parameters.AddWithValue("@Id", Id);
            cmd.Parameters.AddWithValue("@copy", copy);
            cmd.Parameters.AddWithValue("@CreateBy", user);
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
            Context.Response.Write("success");
        }
    }
    //[WebMethod]
    //public void ApproveStep(string user, string Id, string action, string condition, string satusapp)
    //{
    //    int i = 1;
    //    switch (i)
    //    {
    //        case 1:
    //            break;
    //    }
    //     string StatusApp="";
    //    DataTable dt= new DataTable(); 
    //    dt = builditems("select * from sapmaterial where Id='"+ Id +"'");
    //    foreach ( DataRow dr in dt.Rows)
    //    {
    //        if (dr["StatusApp"].ToString() == "0")
    //        {
    //            if (action == "Send Approve" && dr["Condition"].ToString() == "1")
    //            { Update("1", Id, user); StatusApp = "0"; }
    //            else { StatusApp = "1"; }
    //        } else if (dr["StatusApp"].ToString() == "1") {
    //            if (action == "Approve" && dr["Condition"].ToString() == "1")
    //            { Update("2", Id, user); StatusApp = "0"; }
    //            else { StatusApp = "2"; } 
    //        }
    //    }
    //    using (SqlConnection con = new SqlConnection(strConn))
    //    {
    //         SqlCommand cmd = new SqlCommand();
    //        cmd.CommandType = CommandType.StoredProcedure;
    //        cmd.CommandText = "spApproveStep";
    //        cmd.Parameters.AddWithValue("@user", user);
    //        cmd.Parameters.AddWithValue("@RequestNo", Id);
    //        cmd.Parameters.AddWithValue("@StatusApp", StatusApp);
    //        cmd.Connection = con;
    //        con.Open();
    //        dt = new DataTable();
    //        SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
    //        oAdapter.Fill(dt);
    //        con.Close();
    //        Context.Response.Write(JsonConvert.SerializeObject(dt));
    //    }
    //}
    [WebMethod]
    public void Getflagchangeresult(string sName)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spflagchangeresult";
            cmd.Parameters.AddWithValue("@Id", sName);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    [WebMethod]
    public void GetHistory(string sName)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spHistory";
            cmd.Parameters.AddWithValue("@Id", sName);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }

    [WebMethod]
    public void GetMasterLog(string sName, string sUser, string FrDt, string ToDt)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spMasterLog";
            cmd.Parameters.AddWithValue("@Changed_Charname", sName);
            cmd.Parameters.AddWithValue("@sUser", sUser);
            cmd.Parameters.AddWithValue("@FrDt", FrDt);
            cmd.Parameters.AddWithValue("@ToDt", ToDt);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }

    [WebMethod]
    public void GetMatStatus(string name)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spMatStatus";
            cmd.Parameters.AddWithValue("@name", name);
            //cmd.Parameters.AddWithValue("@Condition", sCondition);
            //cmd.Parameters.AddWithValue("@FrDt", FrDt);
            //cmd.Parameters.AddWithValue("@ToDt", ToDt);
            //cmd.Parameters.AddWithValue("@User", sUser);
            //cmd.Parameters.AddWithValue("@Shortname", Shortname);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    [WebMethod]
    public void GetMatStatusAll(string name)
    {
        dynamic dynJson = JsonConvert.DeserializeObject(name);
        foreach (var item in dynJson)
        {
            using (SqlConnection con = new SqlConnection(strConn))
            {
                SqlCommand cmd = new SqlCommand();
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "spMatStatusAll";
                cmd.Parameters.AddWithValue("@name", item.By.ToString());
                cmd.Parameters.AddWithValue("@Condition", item.Keyword.ToString());
                //cmd.Parameters.AddWithValue("@FrDt", FrDt);
                //cmd.Parameters.AddWithValue("@ToDt", ToDt);
                //cmd.Parameters.AddWithValue("@User", sUser);
                //cmd.Parameters.AddWithValue("@Shortname", Shortname);
                cmd.Connection = con;
                con.Open();
                DataTable dt = new DataTable();
                SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
                oAdapter.Fill(dt);
                con.Close();
                Context.Response.Write(JsonConvert.SerializeObject(dt));
            }
        }
    }
    [WebMethod]
    public void GetPlantRegisteredNo(string sName)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spPlantRegisteredNo";
            cmd.Parameters.AddWithValue("@RegisteredNo", sName);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    [WebMethod]
    public void GetSapMaterial(string Id)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetSapMaterial";
            cmd.Parameters.AddWithValue("@Id", Id);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    [WebMethod]
    public void Getsapmaterial2(string name, string sCondition)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spSapMaterial";
            cmd.Parameters.AddWithValue("@material", name);
            cmd.Parameters.AddWithValue("@Condition", sCondition);

            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    [WebMethod]
    public void Getselectall(string user, string where)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetselectall";
            cmd.Parameters.AddWithValue("@user", user);
            cmd.Parameters.AddWithValue("@where", where);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    [WebMethod]
    public void Getinfogroup(string where)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spinfogroup";
            //cmd.Parameters.AddWithValue("@user", user);
            cmd.Parameters.AddWithValue("@where", where);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    [WebMethod]
    public void Getpersonal(objpersonal ro)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetpersonal2";
            cmd.Parameters.AddWithValue("@user", ro.user);
            cmd.Parameters.AddWithValue("@where", ro.where);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    [WebMethod]
    public void GetPendingJob(string user, string where)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetPendingJob";
            cmd.Parameters.AddWithValue("@user", user);
            cmd.Parameters.AddWithValue("@where", where);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    [WebMethod]
    public void GetSearchresults(string table, string where)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spSearchresults";
            cmd.Parameters.AddWithValue("@table", table);
            cmd.Parameters.AddWithValue("@where", where);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    [WebMethod]
    public string GetDetails(string name, int age)
    {
        return string.Format("Name: {0}{2}Age: {1}{2}TimeStamp: {3}", name, age, Environment.NewLine, DateTime.Now.ToString());
    }
    //[WebMethod()]
    //public DataSet Getreadxml(string symbol)
    //{
    //    using (SqlConnection con = new SqlConnection(strConn))
    //    {
    //        SqlCommand cmd = new SqlCommand();
    //        cmd.CommandType = CommandType.StoredProcedure;
    //        cmd.CommandText = "spInsertMultipleRows";
    //        cmd.Parameters.AddWithValue("@xmlData", symbol);
    //        cmd.Connection = con;
    //        con.Open();
    //        DataSet oDataset = new DataSet();
    //        SqlDataAdapter da = new SqlDataAdapter(cmd);
    //        da.Fill(oDataset);
    //        return oDataset;
    //        con.Close();
    //    }
    //}
    //[WebMethod()]
    //public string[] QueryData(string strtable)
    //{
    //    int i = 0;
    //    string sql = null;
    //    SqlConnection oConn = new SqlConnection(strConn);
    //    oConn.Open();
    //    sql = "SELECT Name FROM " + strtable;
    //    System.Data.DataSet oDataset = new System.Data.DataSet();
    //    SqlDataAdapter oAdapter = new SqlDataAdapter(sql, oConn);
    //    oAdapter.Fill(oDataset);
    //    string[] s = new string[oDataset.Tables[0].Rows.Count];
    //    for (i = 0; i <= oDataset.Tables[0].Rows.Count - 1; i++)
    //    {
    //        s[i] = oDataset.Tables[0].Rows[i].ItemArray[0].ToString();
    //    }
    //    oConn.Close();
    //    return s;
    //}

    [WebMethod()]
    public DataSet GetDataSet(string sName)
    {
        using (SqlConnection oConn = new SqlConnection(strConn))
        {
            oConn.Open();
            string strQuery = "select * from " + sName;
            DataSet oDataset = new DataSet();
            SqlDataAdapter oAdapter = new SqlDataAdapter(strQuery, oConn);
            // Fill the dataset.
            oAdapter.Fill(oDataset);
            oConn.Close();
            return oDataset;
        }
    }
    [WebMethod()]
    public void Getjson2(string json)
    {
        string strQuery = "";
        dynamic dynJson = JsonConvert.DeserializeObject(json);
        foreach (var ro in dynJson)
        {
            //string json2 = convertthai(ro.Keyword);
            string value = ro.table.ToString();
            switch (value)
            {

                case "MasStyleofPrinting":
                case "MasScientificName":
                case "MasSpecie":
                    strQuery = @"select * from " + value + " Where ( Id like N'%" + ro.Keyword + "%' or Description like N'%" + ro.Keyword + "%')";
                    if (!string.IsNullOrEmpty(ro.Group.ToString()))
                        strQuery = strQuery + " and isnull(Inactive,'') <> 'X' and (MaterialGroup like '%" + ro.Group + "%')";
                    break;
                case "MasFAOZone":
                case "MasSymbol":
                case "MasCatchingperiodDate":
                case "MasCatchingMethod":
                case "MasCatchingArea":
                case "MasBrand":
                case "MasTypeofPrimary":
                    strQuery = @"select * from " + value + " where ( Id like N'%" +
                    ro.Keyword + "%' or Description like N'%" + ro.Keyword + "%')" + (!string.IsNullOrEmpty(ro.Group.ToString()) ? " and isnull(Inactive,'') <> 'X' " : "");
                    break;
                case "MasTypeofCarton2":
                case "MasTypeofCarton":
                    string _type = value.ToString() == "MasTypeofCarton" ? "'0'" : "'2'";
                    strQuery = @"select * from MasTypeofCarton Where ( Id like '%" + ro.Keyword + "%' or Description like N'%" + ro.Keyword + "%' or DescriptionText like N'%" + ro.Keyword + "%')";
                    if (string.IsNullOrEmpty(ro.Group.ToString()))
                        strQuery = strQuery + " or MaterialGroup like '%" + ro.Keyword
                        + "%' or MaterialType like '%" + ro.Keyword + "%'";
                    else
                        strQuery = strQuery + " and isnull(Inactive,'') <> 'X' and (MaterialGroup like '%" + ro.Group + "%' and MaterialType = " + _type + ")";
                    break;

                //                case "MasTypeofCarton2":
                //                    strQuery = @"select * from MasTypeofCarton Where isnull(Inactive,'') <> 'X' and (Id like "
                //                        + "'%" + ro.Keyword + "%' or Description like N'%" + ro.Keyword + "%' or DescriptionText like N'%" + ro.Keyword + "%') and (MaterialGroup like '%" + ro.Group + "%' and MaterialType = '2')";
                //                    break;
                case "MasPMSColour":
                case "MasProcessColour":
                case "MasTotalColour":
                case "MasGradeofCarton":
                case "MasFlute":
                    if (string.IsNullOrEmpty(ro.Group.ToString()))
                        strQuery = @"select * from " + value + " Where id like '%" + ro.Keyword + "%' or Description like N'%" + ro.Keyword + "%' or MaterialGroup like '%" + ro.Keyword + "%'";
                    else
                        strQuery = @"select * from " + value + " Where(Inactive <> 'X' or Inactive is null) and Id like "
                        + "'%" + ro.Keyword + "%' or Description like N'%" + ro.Keyword + "%' and (MaterialGroup like '%" + ro.Group + "%')";
                    break;
                case "MasPrimarySize":
                    strQuery = @"select * from " + value + " Where (Code like '%" + ro.Keyword + "%' or Description like '%" + ro.Keyword + "%' or Can like '%"
                + ro.Keyword + "%' or LidType like '%" + ro.Keyword + "%' or ContainerType like '%" + ro.Keyword + "%' or DescriptionType like '%" + ro.Keyword + "%')";
                    if (!string.IsNullOrEmpty(ro.Group.ToString()))
                        strQuery = strQuery + " and isnull(Inactive,'') <> 'X' ";
                    strQuery = strQuery + " order by Id";
                    break;
                case "MasPackingStyle":
                    strQuery = @"select * from " + value + " Where PrimaryCode like '%" + ro.Keyword + "%' or GroupStyle like '%" + ro.Keyword + "%' or PackingStyle like '%"
                + ro.Keyword + "%' or RefStyle like '%" + ro.Keyword + "%' or PackSize like '%" + ro.Keyword + "%' or BaseUnit like '%" + ro.Keyword + "%' or TypeofPrimary like '%" + ro.Keyword + "%' Order by RefStyle";
                    break;
                case "MasPackingStyle2":
                    strQuery = @"select * from MasPackingStyle Where (TypeofPrimary = '" + ro.Group + "') and (PrimaryCode like '%" + ro.Keyword + "%' or GroupStyle like '%" + ro.Keyword + "%' or PackingStyle like '%"
                + ro.Keyword + "%' or RefStyle like '%" + ro.Keyword + "%' or PackSize like '%" + ro.Keyword + "%' or BaseUnit like '%" + ro.Keyword + "%' or TypeofPrimary like '%" + ro.Keyword + "%') Order by RefStyle";
                    break;
                case "PlantRegistered":
                case "MasPlantRegisteredNo":
                    strQuery = @"select Id,RegisteredNo,Address,Plant,isnull(Inactive,'') as 'Inactive' from MasPlantRegisteredNo Where id like '%"
                    + ro.Keyword + "%' or RegisteredNo like '%" + ro.Keyword + "%' or Address like '%" + ro.Keyword + "%' or Plant like '%" + ro.Keyword + "%'";
                    break;
                case "MasVendor":
                    strQuery = @"select * from " + value + " Where Id like '%" + ro.Keyword + "%' or Code like '%" + ro.Keyword + "%' or Name like '%" + ro.Keyword + "%'";
                    break;
                case "ulogin":
                    strQuery = @"select * from ulogin Where user_name like '%" + ro.Keyword + "%' or fn like '%"
                        + ro.Keyword + "%' or FirstName like '%" + ro.Keyword + "%' or LastName like '%" + ro.Keyword + "%' or Email like '%" + ro.Keyword + "%'";
                    break;
                case "MasProductGroup":
                    strQuery = @"select * from MasProductGroup Where Product_Group like '%" + ro.Keyword + "%' or Product_GroupDesc like '%" + ro.Keyword + "%' or Product_GroupDesc like '%"
                        + ro.Keyword + "%' or PRD_Plant like '%" + ro.Keyword + "%'";
                    break;
                case "MasLogistics":
                    strQuery = @"select [Id],[ProductGroup],[Description],[Plant],[WHNumber],[StorageType],[LE_Qty],[Storage_UnitType],isnull(Inactive,'') as 'Inactive' from MasLogistics Where ProductGroup like '%"
                    + ro.Keyword + "%' or Description like N'%" + ro.Keyword + "%' or Plant like '%" + ro.Keyword + "%' or WHNumber like '%"
                    + ro.Keyword + "%' or StorageType like '%" + ro.Keyword + "%' or LE_Qty like '%" + ro.Keyword + "%' or Storage_UnitType like '%" + ro.Keyword + "%'";
                    break;
            }
            //Context.Response.Write(ro.table);
        }
        if (!string.IsNullOrEmpty(strQuery))
        {
            using (SqlConnection oConn = new SqlConnection(strConn))
            {
                oConn.Open();
                //sName = sName.Replace('@', '%');
                DataTable dt = new DataTable();
                SqlDataAdapter oAdapter = new SqlDataAdapter(strQuery, oConn);
                // Fill the dataset.
                oAdapter.Fill(dt);
                oConn.Close();
                Context.Response.Write(JsonConvert.SerializeObject(dt));
            }
        }
    }
    [WebMethod()]
    public void Getjson(string sName)
    {
        sName = sName.Replace('@', '%');
        sName = convertthai(sName);
        using (SqlConnection oConn = new SqlConnection(strConn))
        {
            oConn.Open();
            //sName = "PrimarySize Where Description like '%307%'";
            string strQuery = "select * from " + sName;
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(strQuery, oConn);
            // Fill the dataset.
            oAdapter.Fill(dt);
            oConn.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    //[WebMethod()]
    //public void savedata(string data)
    //{
    //    string datapath = "~/FileTest/" + data + ".json";
    //    using (StreamReader sr = new StreamReader(Server.MapPath(datapath)))
    //    {
    //        string json = sr.ReadToEnd();
    //        string sp = "spinsertTechnical";
    //        List<RootObject> ro = JsonConvert.DeserializeObject<List<RootObject>>(json);
    //        //Context.Response.Write(ro[0].Requester);
    //        using (SqlConnection con = new SqlConnection(strConn))
    //        {
    //            using (SqlCommand cmd = new SqlCommand(sp))
    //            {
    //                using (SqlDataAdapter sda = new SqlDataAdapter())
    //                {
    //                    DataTable dt = new DataTable();
    //                    cmd.CommandType = CommandType.StoredProcedure;
    //                    cmd.Parameters.AddWithValue("@RequestNo", ro[0].RequestNo);
    //                    cmd.Parameters.AddWithValue("@Requester", ro[0].Requester);
    //                    cmd.Parameters.AddWithValue("@Company", ro[0].Company);
    //                    cmd.Parameters.AddWithValue("@From", ro[0].from);
    //                    cmd.Parameters.AddWithValue("@To", ro[0].to);
    //                    cmd.Parameters.AddWithValue("@PetCategory", ro[0].PetCategory);
    //                    cmd.Parameters.AddWithValue("@PetFoodType", ro[0].PetFoodType);
    //                    cmd.Parameters.AddWithValue("@CompliedWith", ro[0].CompliedWith);
    //                    cmd.Parameters.AddWithValue("@NutrientProfile", ro[0].NutrientProfile);
    //                    cmd.Parameters.AddWithValue("@Requestfor", ro[0].Requestfor);
    //                    cmd.Parameters.AddWithValue("@ProductType", ro[0].ProductType);
    //                    cmd.Parameters.AddWithValue("@ProductStlye", ro[0].ProductStyle);
    //                    cmd.Parameters.AddWithValue("@Media", ro[0].Media);
    //                    cmd.Parameters.AddWithValue("@ChunkType", ro[0].ChunkType);
    //                    cmd.Parameters.AddWithValue("@NetWeight", ro[0].NetWeight);
    //                    cmd.Parameters.AddWithValue("@PackSize", ro[0].PackSize);
    //                    cmd.Parameters.AddWithValue("@Primary", ro[0].Packaging);
    //                    cmd.Parameters.AddWithValue("@Material", ro[0].Material);
    //                    cmd.Parameters.AddWithValue("@PackageType", ro[0].PackType);
    //                    cmd.Parameters.AddWithValue("@Design", ro[0].PackDesign);
    //                    cmd.Parameters.AddWithValue("@Color", ro[0].PackColor);
    //                    cmd.Parameters.AddWithValue("@Lid", ro[0].PackLid);
    //                    cmd.Parameters.AddWithValue("@PackagingShape", ro[0].PackShape);
    //                    cmd.Parameters.AddWithValue("@Lacquer", ro[0].PackLacquer);
    //                    cmd.Parameters.AddWithValue("@SellingUnit", ro[0].SellingUnit);
    //                    cmd.Parameters.AddWithValue("@Marketingnumber", ro[0].MarketingNumber);
    //                    cmd.Connection = con;
    //                    sda.SelectCommand = cmd;
    //                    sda.Fill(dt);
    //                    Context.Response.Write(JsonConvert.SerializeObject(dt));
    //                }
    //            }
    //        }
    //    }
    //    deletefile(datapath);
    //}

    //[WebMethod]
    //[ScriptMethod(UseHttpGet = true, ResponseFormat = ResponseFormat.Json)]
    //public void selectitemsdata(string id)
    //{
    //    using (SqlConnection con = new SqlConnection(strConn))
    //    {
    //        using (SqlCommand cmd = new SqlCommand("spselectapdata"))
    //        {
    //            using (SqlDataAdapter sda = new SqlDataAdapter())
    //            {
    //                DataTable dt = new DataTable();
    //                cmd.CommandType = CommandType.StoredProcedure;
    //                cmd.Parameters.AddWithValue("@id", id);
    //                cmd.Connection = con;
    //                sda.SelectCommand = cmd;
    //                sda.Fill(dt);
    //                Context.Response.Write(JsonConvert.SerializeObject(dt));
    //            }
    //        }
    //    }
    //}
    [WebMethod]
    [ScriptMethod(UseHttpGet = true, ResponseFormat = ResponseFormat.Json)]
    public void selectitems(string id)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            using (SqlCommand cmd = new SqlCommand("spselectap"))
            {
                using (SqlDataAdapter sda = new SqlDataAdapter())
                {
                    DataTable dt = new DataTable();
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.Connection = con;
                    sda.SelectCommand = cmd;
                    sda.Fill(dt);
                    Context.Response.Write(JsonConvert.SerializeObject(dt));
                }
            }
        }
    }
    [WebMethod]
    [ScriptMethod(UseHttpGet = true, ResponseFormat = ResponseFormat.Json)]
    public void Requestitems(string sName)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            using (SqlCommand cmd = new SqlCommand("sprequestRate"))
            {
                using (SqlDataAdapter sda = new SqlDataAdapter())
                {
                    DataTable dt = new DataTable();
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id", sName);
                    cmd.Connection = con;
                    sda.SelectCommand = cmd;
                    sda.Fill(dt);
                    Context.Response.Write(JsonConvert.SerializeObject(dt));
                }
            }
        }
        //List<Material> ListMaterial = new List<Material>();
        //using (JsonDatEntities context = new JsonDatEntities())
        //{
        //    var obj = (from r in context.Materials where r.Oldcode == id select r).ToList();
        //    ListMaterial = obj;
        //    JavaScriptSerializer js = new JavaScriptSerializer();
        //    Context.Response.Write(js.Serialize(ListMaterial));
        //}
    }
    [WebMethod()]
    public void Assign(string json)
    {
        //string datapath = "~/FileTest/" + data + ".json";
        //using (StreamReader sr = new StreamReader(Server.MapPath(datapath)))
        //{
        //string json = sr.ReadToEnd();
        json = json.Replace('@', '&');
        string sp = "spAssignDocument";
        List<Assign> _x = JsonConvert.DeserializeObject<List<Assign>>(json);
        Assign ro = _x[0];
        //Context.Response.Write(ro[0].Requester);
        using (SqlConnection con = new SqlConnection(strConn))
        {
            using (SqlCommand cmd = new SqlCommand(sp))
            {
                using (SqlDataAdapter sda = new SqlDataAdapter())
                {
                    DataTable dt = new DataTable();
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@Assignee", ro.Assignee);
                    cmd.Parameters.AddWithValue("@Id", ro.Id);
                    cmd.Connection = con;
                    sda.SelectCommand = cmd;
                    sda.Fill(dt);
                    Context.Response.Write(JsonConvert.SerializeObject(dt));
                }
            }
        }
        //}
        //deletefile(datapath);
    }
    public void artworkAssign(Assign ro)
    {
        using (SqlConnection cn = new SqlConnection(strConn))
        {
            using (SqlCommand cmd = new SqlCommand("spAssignDocument"))
            {
                cmd.Connection = cn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Assignee", ro.Assignee);
                cmd.Parameters.AddWithValue("@Id", ro.Id);
                cn.Open();
                cmd.ExecuteNonQuery();
                cn.Close();
            }
        }
    }
    [WebMethod()]

    public void saveCreateroot(string data)
    {
        List<CreateDocument> _x = JsonConvert.DeserializeObject<List<CreateDocument>>(data);
        CreateDocument ro = _x[0];
        using (SqlConnection con = new SqlConnection(strConn))
        {
            using (SqlCommand cmd = new SqlCommand("spCreateDocument"))
            {
                using (SqlDataAdapter sda = new SqlDataAdapter())
                {
                    DataTable dt = new DataTable();
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@CreateBy", ro.CreateBy);
                    cmd.Parameters.AddWithValue("@Code", ro.Code);
                    cmd.Parameters.AddWithValue("@Condition", ro.Condition);
                    cmd.Connection = con;
                    sda.SelectCommand = cmd;
                    sda.Fill(dt);
                    Context.Response.Write(JsonConvert.SerializeObject(dt));
                }
            }
        }
    }
    [WebMethod()]

    public void saveRequestDocument(RequestDocument ro)
    {
        //string datapath = "~/FileTest/" + data + ".json";
        //using (StreamReader sr = new StreamReader(Server.MapPath(datapath)))
        //{
        //string json = sr.ReadToEnd();
        string sp = "spRequestDocument";
        //List<RequestDocument> ro = JsonConvert.DeserializeObject<List<RequestDocument>>(json);
        //Context.Response.Write(ro[0].Requester);
        using (SqlConnection con = new SqlConnection(strConn))
        {
            using (SqlCommand cmd = new SqlCommand(sp))
            {
                using (SqlDataAdapter sda = new SqlDataAdapter())
                {
                    DataTable dt = new DataTable();
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@CreateBy", ro.CreateBy);
                    cmd.Parameters.AddWithValue("@Condition", ro.Condition);
                    cmd.Parameters.AddWithValue("@ProductCode", ro.ProductCode);
                    cmd.Parameters.AddWithValue("@Material", ro.Material);
                    cmd.Parameters.AddWithValue("@ProductGroup", ro.ProductGroup);
                    cmd.Connection = con;
                    sda.SelectCommand = cmd;
                    sda.Fill(dt);
                    Context.Response.Write(JsonConvert.SerializeObject(dt));
                }
            }
        }
        //}
        //deletefile(datapath);
    }
    public string Getvalue(List<InboundArtwork> _itemsArtwork, string Characteristic)
    {
        string _result = "";
        foreach (var p in _itemsArtwork)
        {
            if (p.Characteristic == Characteristic)
                _result = p.Value;
        }
        return _result;
    }
    public ArtworkObject builsartworkObject(string d)
    {
        ArtworkObject a = new ArtworkObject();
        string xmlfile = @"\\192.168.1.170\FileTest\dArtwork" + d + ".xml";
        XmlTextReader xmlreader = new XmlTextReader(xmlfile);
        DataSet ds = new DataSet();
        ds.ReadXml(xmlreader);
        xmlreader.Close();
        DataTable dt = ds.Tables[0];
        a = dt.AsEnumerable().Select(row =>
    new ArtworkObject
    {
        PAUserName = row.Field<string>("PAUserName"),
        Plant = row.Field<string>("Plant"),
        ArtworkNumber = row.Field<string>("ArtworkNumber"),
        PrintingStyleofPrimary = row.Field<string>("PrintingStyleofPrimary"),
        PrintingStyleofSecondary = row.Field<string>("PrintingStyleofSecondary"),
        CustomersDesign = row.Field<string>("CustomersDesign"),
        CustomersDesignDetail = row.Field<string>("CustomersDesignDetail"),
        CustomersSpec = row.Field<string>("CustomersSpec"),
        CustomersSpecDetail = row.Field<string>("CustomersSpecDetail"),
        CustomersSize = row.Field<string>("CustomersSize"),
        CustomersSizeDetail = row.Field<string>("CustomersSizeDetail"),
        CustomerNominatesVendor = row.Field<string>("CustomerNominatesVendor"),
        CustomerNominatesVendorDetail = row.Field<string>("CustomerNominatesVendorDetail"),
        CustomerNominatesColorPantone = row.Field<string>("CustomerNominatesColorPantone"),
        CustomerNominatesColorPantoneDetail = row.Field<string>("CustomerNominatesColorPantoneDetail"),
        CustomersBarcodeScanable = row.Field<string>("CustomersBarcodeScanable"),
        CustomersBarcodeScanableDetail = row.Field<string>("CustomersBarcodeScanableDetail"),
        CustomersBarcodeSpec = row.Field<string>("CustomersBarcodeSpec"),
        CustomersBarcodeSpecDetail = row.Field<string>("CustomersBarcodeSpecDetail"),
        FirstInfoGroup = row.Field<string>("FirstInfoGroup"),
        SONumber = row.Field<string>("SONumber"),
        PICMKT = row.Field<string>("PICMKT"),
        SOPlant = row.Field<string>("SOPlant"),
        Destination = row.Field<string>("Destination"),
        RemarkNoteofPA = row.Field<string>("RemarkNoteofPA"),
        FinalInfoGroup = row.Field<string>("FinalInfoGroup"),
        RemarkNoteofPG = row.Field<string>("RemarkNoteofPG")
    }).FirstOrDefault();
        return a;
    }
    [WebMethod]
    public void updateArtworkNumber(string d)
    {
        //open the tender xml file  
        //d = "1MMK202101050039";
        string xmlfile = @"\\192.168.1.170\FileTest\dArtwork" + d + ".xml";
        ArtworkObject _artworkObject = builsartworkObject(d);
        XmlTextReader xmlreader = new XmlTextReader(xmlfile);
        //reading the xml data  
        DataSet ds = new DataSet();
        ds.ReadXml(xmlreader);
        xmlreader.Close();
        //if ds is not empty 
        var dtArtwork = cs.builditems("select * from SapMaterial Where DocumentNo='" + d + "'");
        foreach (DataRow value in dtArtwork.Rows)
        {
            string Keys = string.Format("{0}", value["DocumentNo"]);

            string _group = "";
            if (ds.Tables.Count != 0)
            {
                DataTable dt = ds.Tables[0];
                List<string> list = new List<string>();
                string _LidType = "", _ContainerType = "", _ChangePoint = "",_ProductCode="", _PlantRegisteredNo="", _CompanyAddress="";
                DataTable destination = new DataTable(dt.TableName);
                destination = dt.Clone();
                foreach (DataRow r in dt.Rows)
                {
                    if (r["Characteristic"].ToString() == "ZPKG_SEC_CONTAINER_TYPE")
                        _ContainerType = string.Format("{0}", r["Value"]);
                    if (r["Characteristic"].ToString() == "ZPKG_SEC_LID_TYPE")
                        _LidType = string.Format("{0}", r["Value"]);
                    if (r["Characteristic"].ToString() == "ZPKG_SEC_CHANGE_POINT")
                        _ChangePoint = string.Format("{0}", r["Value"]);
                    if (r["Characteristic"].ToString() == "ZPKG_SEC_PRODUCT_CODE" && !list.Equals(r["Value"]))
                        list.Add(string.Format("{0}", r["Value"]));

                    DataRow dr = destination.Select("Characteristic='" + r["Characteristic"].ToString() + "'").FirstOrDefault();
                    if (dr != null)
                    {
                        if (string.IsNullOrEmpty(dr["Value"].ToString()))
                            dr["Value"] = r["Value"]; //changes the Product_name
                        else
                            dr["Value"] += string.Format(";{0}", r["Value"]);
                    }
                    else
                        destination.ImportRow(r);
                    //
                    //return String.Join(";", list.ToArray());
                }
                string[] userlevel = { "PA", "PG" };
                foreach (string data in userlevel)
                {
                    using (SqlConnection con = new SqlConnection(strConn))
                    {
                        SqlCommand cmd = new SqlCommand();
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandText = "spInsertMultipleRows";
                        cmd.Parameters.AddWithValue(@"@Description", string.Format("{0}", value["Description"].ToString()));
                        cmd.Parameters.AddWithValue("@Brand", "");
                        cmd.Parameters.AddWithValue("@Primarysize", "");
                        cmd.Parameters.AddWithValue("@Version", value["Version"].ToString());
                        cmd.Parameters.AddWithValue("@ChangePoint", _ChangePoint.ToString());
                        cmd.Parameters.AddWithValue("@MaterialGroup", "");
                        cmd.Parameters.AddWithValue("@CreateBy", _artworkObject.PAUserName);
                        cmd.Parameters.AddWithValue("@RequestNo", value["ID"].ToString());
                        cmd.Parameters.AddWithValue("@userlevel", data);
                        cmd.Parameters.AddWithValue("@PackingStyle", "");
                        cmd.Parameters.AddWithValue("@Packing", "");
                        cmd.Parameters.AddWithValue("@StyleofPrinting", "");
                        cmd.Parameters.AddWithValue("@ContainerType", "");
                        cmd.Parameters.AddWithValue("@LidType", "");
                        cmd.Parameters.AddWithValue("@TotalColour", "");
                        cmd.Parameters.AddWithValue("@StatusApp", string.Format("{0}", 0));
                        cmd.Parameters.AddWithValue("@ProductCode", "");
                        cmd.Parameters.AddWithValue("@FAOZone", "");
                        //cmd.Parameters.AddWithValue("@Plant", string.Format("{0}", _artworkObject.Plant.Replace(',', ';')));
                        cmd.Parameters.AddWithValue("@Plant", string.Format("{0}", _artworkObject.Plant.ToString()));
                        cmd.Parameters.AddWithValue("@Processcolour", "");
                        cmd.Parameters.AddWithValue("@PlantRegisteredNo", "");
                        cmd.Parameters.AddWithValue("@CompanyNameAddress", "");
                        cmd.Parameters.AddWithValue("@PMScolour", "");
                        cmd.Parameters.AddWithValue("@Symbol", "");
                        cmd.Parameters.AddWithValue("@CatchingArea", "");
                        cmd.Parameters.AddWithValue("@CatchingPeriodDate", "");
                        cmd.Parameters.AddWithValue("@Grandof", "");
                        cmd.Parameters.AddWithValue("@Flute", "");
                        cmd.Parameters.AddWithValue("@Vendor", "");
                        cmd.Parameters.AddWithValue("@Dimension", "");
                        cmd.Parameters.AddWithValue("@RSC", "");
                        cmd.Parameters.AddWithValue("@Accessories", "");
                        cmd.Parameters.AddWithValue("@PrintingStyleofPrimary", string.Format("{0}", _artworkObject.PrintingStyleofPrimary));
                        cmd.Parameters.AddWithValue("@PrintingStyleofSecondary", string.Format("{0}", _artworkObject.PrintingStyleofSecondary));
                        cmd.Parameters.AddWithValue("@CustomerDesign", _artworkObject.CustomersDesign + "|" + _artworkObject.CustomersDesignDetail);
                        cmd.Parameters.AddWithValue("@CustomerSpec", _artworkObject.CustomersSpec + "|" + _artworkObject.CustomersSpecDetail);
                        cmd.Parameters.AddWithValue("@CustomerSize", _artworkObject.CustomersSize + "|" + _artworkObject.CustomersSizeDetail);
                        cmd.Parameters.AddWithValue("@CustomerVendor", _artworkObject.CustomerNominatesVendor + "|" + _artworkObject.CustomerNominatesVendorDetail);
                        cmd.Parameters.AddWithValue("@CustomerColor", _artworkObject.CustomerNominatesColorPantone + "|" + _artworkObject.CustomerNominatesColorPantoneDetail);
                        cmd.Parameters.AddWithValue("@CustomerScanable", _artworkObject.CustomersBarcodeScanable + "|" + _artworkObject.CustomersBarcodeScanableDetail);
                        cmd.Parameters.AddWithValue("@CustomerBarcodeSpec", _artworkObject.CustomersBarcodeSpec + "|" + _artworkObject.CustomersBarcodeSpecDetail);
                        cmd.Parameters.AddWithValue("@FirstInfoGroup", string.Format("{0}", _artworkObject.FirstInfoGroup));
                        cmd.Parameters.AddWithValue("@SO", string.Format("{0}", _artworkObject.SONumber));
                        cmd.Parameters.AddWithValue("@PICMkt", string.Format("{0}", _artworkObject.PICMKT));
                        cmd.Parameters.AddWithValue("@SOPlant", string.Format("{0}", _artworkObject.SOPlant));
                        cmd.Parameters.AddWithValue("@Destination", string.Format("{0}", _artworkObject.Destination));
                        cmd.Parameters.AddWithValue("@Remark", string.Format("{0}", _artworkObject.RemarkNoteofPA));
                        cmd.Parameters.AddWithValue("@GrossWeight", "");
                        cmd.Parameters.AddWithValue("@FinalInfoGroup", string.Format("{0}", _artworkObject.FinalInfoGroup));
                        cmd.Parameters.AddWithValue("@Note", string.Format("{0}", _artworkObject.RemarkNoteofPG));
                        cmd.Parameters.AddWithValue("@SheetSize", "");
                        cmd.Parameters.AddWithValue("@Typeof", "");
                        cmd.Parameters.AddWithValue("@TypeofCarton2", "");
                        cmd.Parameters.AddWithValue("@DMSNo", _artworkObject.ArtworkNumber);

                        cmd.Parameters.AddWithValue("@TypeofPrimary", "");
                        cmd.Parameters.AddWithValue("@PrintingSystem", "");
                        cmd.Parameters.AddWithValue("@Direction", "");
                        cmd.Parameters.AddWithValue("@RollSheet", "");
                        cmd.Parameters.AddWithValue("@RequestType", "");
                        cmd.Parameters.AddWithValue("@PlantAddress", "");

                        cmd.Parameters.AddWithValue("@Fixed_Desc", "");
                        cmd.Parameters.AddWithValue("@Inactive", "");
                        cmd.Parameters.AddWithValue("@Catching_Method", "");
                        cmd.Parameters.AddWithValue("@Scientific_Name", "");
                        cmd.Parameters.AddWithValue("@Specie", "");
                        cmd.Connection = con;
                        con.Open();
                        cmd.ExecuteNonQuery();
                        con.Close();
                    }
                }
                foreach (DataRow item in destination.Rows)
                {
                    string charac = item["Characteristic"].ToString();
                    if (charac.ToString().Contains("ZPKG_SEC_PRIMARY_SIZE"))
                    {
                        item["Value"] = buildprimaty(item, _ContainerType.ToUpper(), _LidType.ToUpper(), String.Join(";", list.ToArray()));
                    }
                    if (charac.ToString().Contains("ZPKG_SEC_ACCESSORIES"))
                    {
                        item["Value"] = Insert(item["Value"].ToString());
                    }
                    //Initialize SQL Server Connection
                    SqlConnection cn = new SqlConnection(strConn);
                    SqlCommand cmd = new SqlCommand("spUpdateSapMaterial", cn);
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@Description", string.Format("{0}", item["Description"]));
                    //cmd.Parameters.AddWithValue("@ArtworkNumber", string.Format("{0}", _artworkObject.ArtworkNumber));
                    //cmd.Parameters.AddWithValue("@Date", string.Format("{0}", item.Date));
                    cmd.Parameters.AddWithValue("@Value", string.Format("{0}", item["Value"]));
                    cmd.Parameters.AddWithValue("@Group", string.Format("{0}", _group));
                    cmd.Parameters.AddWithValue("@Characteristic", string.Format("{0}", charac));
                    cmd.Parameters.AddWithValue("@Keys", string.Format("{0}", value["ID"]));
                    // Running the query.
                    cn.Open();
                    cmd.ExecuteNonQuery();
                    cn.Close();
                }
                using (SqlConnection CN = new SqlConnection(strConn))
                {
                    string qry = "spUpdateArtwork";
                    SqlCommand SqlCom = new SqlCommand(qry, CN);
                    SqlCom.CommandType = CommandType.StoredProcedure;
                    SqlCom.Parameters.Add(new SqlParameter("@Keys", value["ID"].ToString()));
                    CN.Open();
                    SqlCom.ExecuteNonQuery();
                    CN.Close();
                }
            }
        }
    }
    public string buildprimaty(DataRow item, string _ContainerType, string _LidType, string product)
    {
        using (SqlConnection CN = new SqlConnection(strConn))
        {
            string qry = "spGetPrimary";
            SqlCommand SqlCom = new SqlCommand(qry, CN);
            SqlCom.CommandType = CommandType.StoredProcedure;
            SqlCom.Parameters.Add(new SqlParameter("@Value", item["Value"].ToString()));
            SqlCom.Parameters.Add(new SqlParameter("@ContainerType", _ContainerType.ToString()));
            SqlCom.Parameters.Add(new SqlParameter("@LidType", _LidType.ToString()));
            SqlCom.Parameters.Add(new SqlParameter("@Product", product.ToString()));
            CN.Open();
            //SqlCom.ExecuteNonQuery();
            var getValue = SqlCom.ExecuteScalar();
            CN.Close();
            return getValue == null ? "" : getValue.ToString();
        }
    }
    [WebMethod()]
    public void savedocument2(string data)
    {
        //json = json.Replace('@', '&');
        //List<sapmaterial> _x = JsonConvert.DeserializeObject<List<sapmaterial>>(json);
        //sapmaterial ro = _x[0];
        //        using (FileStream fs = new FileStream(Server.MapPath("~/FileTest/ro.xml"), FileMode.Create))
        //        {
        //            new XmlSerializer(typeof(sapmaterial)).Serialize(fs, ro);
        //        }
        string datapath = "~/FileTest/" + data + ".json";
        using (StreamReader sr = new StreamReader(Server.MapPath(datapath)))
        {
            string json = sr.ReadToEnd();
            List<sapmaterial> _x = JsonConvert.DeserializeObject<List<sapmaterial>>(json);
            foreach (var ro in _x)
            {
                if (ro.RequestNo.ToString() != "" || ro.RequestNo.ToString() != "0") goto Jumpto;
                using (SqlConnection con = new SqlConnection(strConn))
                {
                    using (SqlCommand cmd = new SqlCommand("spCreateDocument"))
                    {
                        using (SqlDataAdapter sda = new SqlDataAdapter())
                        {
                            DataTable dt = new DataTable();
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@CreateBy", ro.CreateBy);
                            cmd.Parameters.AddWithValue("@Code", ro._root.Code);
                            cmd.Parameters.AddWithValue("@Condition", ro._root.Condition);
                            cmd.Connection = con;
                            sda.SelectCommand = cmd;
                            sda.Fill(dt);
                            foreach (DataRow dr in dt.Rows)
                            {
                                ro.RequestNo = dr["ID"].ToString();
                            }
                        }
                    }
                }
            Jumpto:
                //Assignee
                //        if (ro.Assignee.ToString() != "")
                //        {
                //            StringBuilder sb = new StringBuilder();
                //            DataTable result = cs.builditems("select * from SapMaterial where Id='" + ro.RequestNo + "'");
                //            foreach (DataRow r in result.Rows)
                //            {
                //                if (ro.Assignee.ToString() != r["Assignee"].ToString())
                //                {
                //                    Assign _assign = new Assign();
                //                    _assign.Id = r["ID"].ToString();
                //                    _assign.Assignee = ro.Assignee.ToString();
                //                    artworkAssign(_assign);
                //                    string Subject = "SEC PKG Template is created No. : " + r["DocumentNo"].ToString() + " /" + ro.Description.ToString() + "/" + ro.RequestType.ToString();
                //                    DataTable dt = cs.builditems("select ActiveBy from TransApprove Where fn in ('PA','PA_Approve') and MatDoc='" + ro.RequestNo + "' and ActiveBy <>''");
                //                    foreach (DataRow dr in dt.Rows)
                //                    {
                //                        sb.Append(cs.Getuser(dr["ActiveBy"].ToString(), "email") + ",");
                //                    }
                //                    cs.sendemail(@cs.Getuser(ro.Assignee, "email"), sb.ToString() 
                //                        , "SEC PKG Template is created No. : " + r["DocumentNo"] + " <br /><br /> Mail Assign PG <br />Assinee : " + cs.Getuser(ro.Assignee, "fn"), Subject, "");
                //                }
                //            }
                //		}
                //string datapath = "~/FileTest/" + data + ".json";
                //using (StreamReader sr = new StreamReader(Server.MapPath(datapath)))
                //{
                //string test = cs.ReadItems(string.Format(" select '{0}' value", ro.StyleofPrinting));
                //string json = sr.ReadToEnd();
                string sp = "spInsertMultipleRows";
                //List<sapmaterial> ro = JsonConvert.DeserializeObject<List<sapmaterial>>(json);
                //Context.Response.Write(ro[0].Requester);
                using (SqlConnection con = new SqlConnection(strConn))
                {
                    using (SqlCommand cmd = new SqlCommand(sp))
                    {
                        using (SqlDataAdapter sda = new SqlDataAdapter())
                        {
                            DataTable dt = new DataTable();
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@Description", ro.Description);
                            cmd.Parameters.AddWithValue("@Brand", ro.Brand);
                            cmd.Parameters.AddWithValue("@Primarysize", ro.PrimarySize);
                            cmd.Parameters.AddWithValue("@Version", ro.Version);
                            cmd.Parameters.AddWithValue("@ChangePoint", ro.ChangePoint);
                            cmd.Parameters.AddWithValue("@MaterialGroup", ro.MaterialGroup);
                            cmd.Parameters.AddWithValue("@CreateBy", ro.CreateBy);
                            cmd.Parameters.AddWithValue("@RequestNo", ro.RequestNo);
                            cmd.Parameters.AddWithValue("@userlevel", ro.Userlevel);
                            cmd.Parameters.AddWithValue("@PackingStyle", ro.PackingStyle);
                            cmd.Parameters.AddWithValue("@Packing", ro.Packing);
                            cmd.Parameters.AddWithValue("@StyleofPrinting", ro.StyleofPrinting);
                            cmd.Parameters.AddWithValue("@ContainerType", ro.ContainerType);
                            cmd.Parameters.AddWithValue("@LidType", ro.LidType);
                            cmd.Parameters.AddWithValue("@TotalColour", ro.TotalColour);
                            cmd.Parameters.AddWithValue("@StatusApp", ro.StatusApp);
                            cmd.Parameters.AddWithValue("@ProductCode", ro.ProductCode);
                            cmd.Parameters.AddWithValue("@FAOZone", ro.FAOZone);
                            cmd.Parameters.AddWithValue("@Plant", ro.Plant);
                            cmd.Parameters.AddWithValue("@Processcolour", ro.Processcolour);
                            cmd.Parameters.AddWithValue("@PlantRegisteredNo", ro.PlantRegisteredNo);
                            cmd.Parameters.AddWithValue("@CompanyNameAddress", ro.CompanyNameAddress);
                            cmd.Parameters.AddWithValue("@PMScolour", ro.PMScolour);
                            cmd.Parameters.AddWithValue("@Symbol", ro.Symbol);
                            cmd.Parameters.AddWithValue("@CatchingArea", ro.CatchingArea);
                            cmd.Parameters.AddWithValue("@CatchingPeriodDate", ro.CatchingPeriodDate);
                            cmd.Parameters.AddWithValue("@Grandof", ro.Grandof);
                            cmd.Parameters.AddWithValue("@Flute", ro.Flute);
                            cmd.Parameters.AddWithValue("@Vendor", ro.Vendor);
                            cmd.Parameters.AddWithValue("@Dimension", ro.Dimension);
                            cmd.Parameters.AddWithValue("@RSC", ro.RSC);
                            cmd.Parameters.AddWithValue("@Accessories", ro.Accessories);
                            cmd.Parameters.AddWithValue("@PrintingStyleofPrimary", ro.PrintingStyleofPrimary);
                            cmd.Parameters.AddWithValue("@PrintingStyleofSecondary", ro.PrintingStyleofSecondary);
                            cmd.Parameters.AddWithValue("@CustomerDesign", ro.CustomerDesign);
                            cmd.Parameters.AddWithValue("@CustomerSpec", ro.CustomerSpec);
                            cmd.Parameters.AddWithValue("@CustomerSize", ro.CustomerSize);
                            cmd.Parameters.AddWithValue("@CustomerVendor", ro.CustomerVendor);
                            cmd.Parameters.AddWithValue("@CustomerColor", ro.CustomerColor);
                            cmd.Parameters.AddWithValue("@CustomerScanable", ro.CustomerScanable);
                            cmd.Parameters.AddWithValue("@CustomerBarcodeSpec", ro.CustomerBarcodeSpec);
                            cmd.Parameters.AddWithValue("@FirstInfoGroup", ro.FirstInfoGroup);
                            cmd.Parameters.AddWithValue("@SO", ro.SO);
                            cmd.Parameters.AddWithValue("@PICMkt", ro.PICMkt);
                            cmd.Parameters.AddWithValue("@SOPlant", ro.SOPlant);
                            cmd.Parameters.AddWithValue("@Destination", ro.Destination);
                            cmd.Parameters.AddWithValue("@Remark", ro.Remark);
                            cmd.Parameters.AddWithValue("@GrossWeight", ro.GrossWeight);
                            cmd.Parameters.AddWithValue("@FinalInfoGroup", ro.FinalInfoGroup);
                            cmd.Parameters.AddWithValue("@Note", ro.Note);
                            cmd.Parameters.AddWithValue("@SheetSize", ro.SheetSize);
                            cmd.Parameters.AddWithValue("@Typeof", ro.Typeof);
                            cmd.Parameters.AddWithValue("@TypeofCarton2", ro.TypeofCarton2);
                            cmd.Parameters.AddWithValue("@DMSNo", ro.DMSNo);

                            cmd.Parameters.AddWithValue("@TypeofPrimary", ro.TypeofPrimary);
                            cmd.Parameters.AddWithValue("@PrintingSystem", ro.PrintingSystem);
                            cmd.Parameters.AddWithValue("@Direction", ro.Direction);
                            cmd.Parameters.AddWithValue("@RollSheet", ro.RollSheet);
                            cmd.Parameters.AddWithValue("@RequestType", ro.RequestType);
                            cmd.Parameters.AddWithValue("@PlantAddress", ro.PlantAddress);

                            cmd.Parameters.AddWithValue("@Fixed_Desc", ro.Fixed_Desc);
                            cmd.Parameters.AddWithValue("@Inactive", ro.Inactive);
                            cmd.Parameters.AddWithValue("@Catching_Method", ro.Catching_Method);
                            cmd.Parameters.AddWithValue("@Scientific_Name", ro.Scientific_Name);
                            cmd.Parameters.AddWithValue("@Specie", ro.Specie);
                            //cmd.Parameters.AddWithValue("@ReferenceMaterial", string.Format("{0}", ro.ReferenceMaterial));
                            cmd.Connection = con;
                            sda.SelectCommand = cmd;
                            sda.Fill(dt);
                            foreach (DataRow dr in dt.Rows)
                            {
                                if (dr["StatusApp"].ToString() == "5")
                                {
                                    OutboundArtwork(dr["DocumentNo"].ToString());
                                }
                            }
                            Context.Response.Write(JsonConvert.SerializeObject(dt));
                        }
                    }
                }
            }
            //deletefile(datapath);
        }
    }
    [WebMethod()]
    public void savedocument(string json)
    {
        json = json.Replace('@', '+');
        List<sapmaterial> _x = JsonConvert.DeserializeObject<List<sapmaterial>>(json);
        sapmaterial ro = _x[0];
        //        using (FileStream fs = new FileStream(Server.MapPath("~/FileTest/ro.xml"), FileMode.Create))
        //        {
        //            new XmlSerializer(typeof(sapmaterial)).Serialize(fs, ro);
        //        }
        if (ro.RequestNo.ToString() != "" || ro.RequestNo.ToString() != "0") goto Jumpto;
        using (SqlConnection con = new SqlConnection(strConn))
        {
            using (SqlCommand cmd = new SqlCommand("spCreateDocument"))
            {
                using (SqlDataAdapter sda = new SqlDataAdapter())
                {
                    DataTable dt = new DataTable();
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@CreateBy", ro.CreateBy);
                    cmd.Parameters.AddWithValue("@Code", ro._root.Code);
                    cmd.Parameters.AddWithValue("@Condition", ro._root.Condition);
                    cmd.Connection = con;
                    sda.SelectCommand = cmd;
                    sda.Fill(dt);
                    foreach (DataRow dr in dt.Rows)
                    {
                        ro.RequestNo = dr["ID"].ToString();
                    }
                }
            }
        }
    Jumpto:
        //Assignee
        //        if (ro.Assignee.ToString() != "")
        //        {
        //            StringBuilder sb = new StringBuilder();
        //            DataTable result = cs.builditems("select * from SapMaterial where Id='" + ro.RequestNo + "'");
        //            foreach (DataRow r in result.Rows)
        //            {
        //                if (ro.Assignee.ToString() != r["Assignee"].ToString())
        //                {
        //                    Assign _assign = new Assign();
        //                    _assign.Id = r["ID"].ToString();
        //                    _assign.Assignee = ro.Assignee.ToString();
        //                    artworkAssign(_assign);
        //                    string Subject = "SEC PKG Template is created No. : " + r["DocumentNo"].ToString() + " /" + ro.Description.ToString() + "/" + ro.RequestType.ToString();
        //                    DataTable dt = cs.builditems("select ActiveBy from TransApprove Where fn in ('PA','PA_Approve') and MatDoc='" + ro.RequestNo + "' and ActiveBy <>''");
        //                    foreach (DataRow dr in dt.Rows)
        //                    {
        //                        sb.Append(cs.Getuser(dr["ActiveBy"].ToString(), "email") + ",");
        //                    }
        //                    cs.sendemail(@cs.Getuser(ro.Assignee, "email"), sb.ToString() 
        //                        , "SEC PKG Template is created No. : " + r["DocumentNo"] + " <br /><br /> Mail Assign PG <br />Assinee : " + cs.Getuser(ro.Assignee, "fn"), Subject, "");
        //                }
        //            }
        //		}
        //string datapath = "~/FileTest/" + data + ".json";
        //using (StreamReader sr = new StreamReader(Server.MapPath(datapath)))
        //{
        //string json = sr.ReadToEnd();
        string sp = "spInsertMultipleRows";
        //List<sapmaterial> ro = JsonConvert.DeserializeObject<List<sapmaterial>>(json);
        //Context.Response.Write(ro[0].Requester);
        using (SqlConnection con = new SqlConnection(strConn))
        {
            using (SqlCommand cmd = new SqlCommand(sp))
            {
                using (SqlDataAdapter sda = new SqlDataAdapter())
                {
                    DataTable dt = new DataTable();
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@Description", ro.Description);
                    cmd.Parameters.AddWithValue("@Brand", ro.Brand);
                    cmd.Parameters.AddWithValue("@Primarysize", ro.PrimarySize);
                    cmd.Parameters.AddWithValue("@Version", ro.Version);
                    cmd.Parameters.AddWithValue("@ChangePoint", ro.ChangePoint);
                    cmd.Parameters.AddWithValue("@MaterialGroup", ro.MaterialGroup);
                    cmd.Parameters.AddWithValue("@CreateBy", ro.CreateBy);
                    cmd.Parameters.AddWithValue("@RequestNo", ro.RequestNo);
                    cmd.Parameters.AddWithValue("@userlevel", ro.Userlevel);
                    cmd.Parameters.AddWithValue("@PackingStyle", ro.PackingStyle);
                    cmd.Parameters.AddWithValue("@Packing", ro.Packing);
                    cmd.Parameters.AddWithValue("@StyleofPrinting", string.Format("{0}", ro.StyleofPrinting));
                    cmd.Parameters.AddWithValue("@ContainerType", ro.ContainerType);
                    cmd.Parameters.AddWithValue("@LidType", ro.LidType);
                    cmd.Parameters.AddWithValue("@TotalColour", ro.TotalColour);
                    cmd.Parameters.AddWithValue("@StatusApp", ro.StatusApp);
                    cmd.Parameters.AddWithValue("@ProductCode", ro.ProductCode);
                    cmd.Parameters.AddWithValue("@FAOZone", ro.FAOZone);
                    cmd.Parameters.AddWithValue("@Plant", ro.Plant);
                    cmd.Parameters.AddWithValue("@Processcolour", ro.Processcolour);
                    cmd.Parameters.AddWithValue("@PlantRegisteredNo", ro.PlantRegisteredNo);
                    cmd.Parameters.AddWithValue("@CompanyNameAddress", ro.CompanyNameAddress);
                    cmd.Parameters.AddWithValue("@PMScolour", ro.PMScolour);
                    cmd.Parameters.AddWithValue("@Symbol", ro.Symbol);
                    cmd.Parameters.AddWithValue("@CatchingArea", ro.CatchingArea);
                    cmd.Parameters.AddWithValue("@CatchingPeriodDate", ro.CatchingPeriodDate);
                    cmd.Parameters.AddWithValue("@Grandof", ro.Grandof);
                    cmd.Parameters.AddWithValue("@Flute", ro.Flute);
                    cmd.Parameters.AddWithValue("@Vendor", ro.Vendor);
                    cmd.Parameters.AddWithValue("@Dimension", ro.Dimension);
                    cmd.Parameters.AddWithValue("@RSC", ro.RSC);
                    cmd.Parameters.AddWithValue("@Accessories", ro.Accessories);
                    cmd.Parameters.AddWithValue("@PrintingStyleofPrimary", ro.PrintingStyleofPrimary);
                    cmd.Parameters.AddWithValue("@PrintingStyleofSecondary", ro.PrintingStyleofSecondary);
                    cmd.Parameters.AddWithValue("@CustomerDesign", ro.CustomerDesign);
                    cmd.Parameters.AddWithValue("@CustomerSpec", ro.CustomerSpec);
                    cmd.Parameters.AddWithValue("@CustomerSize", ro.CustomerSize);
                    cmd.Parameters.AddWithValue("@CustomerVendor", ro.CustomerVendor);
                    cmd.Parameters.AddWithValue("@CustomerColor", ro.CustomerColor);
                    cmd.Parameters.AddWithValue("@CustomerScanable", ro.CustomerScanable);
                    cmd.Parameters.AddWithValue("@CustomerBarcodeSpec", ro.CustomerBarcodeSpec);
                    cmd.Parameters.AddWithValue("@FirstInfoGroup", ro.FirstInfoGroup);
                    cmd.Parameters.AddWithValue("@SO", ro.SO);
                    cmd.Parameters.AddWithValue("@PICMkt", ro.PICMkt);
                    cmd.Parameters.AddWithValue("@SOPlant", ro.SOPlant);
                    cmd.Parameters.AddWithValue("@Destination", ro.Destination);
                    cmd.Parameters.AddWithValue("@Remark", ro.Remark);
                    cmd.Parameters.AddWithValue("@GrossWeight", ro.GrossWeight);
                    cmd.Parameters.AddWithValue("@FinalInfoGroup", ro.FinalInfoGroup);
                    cmd.Parameters.AddWithValue("@Note", ro.Note);
                    cmd.Parameters.AddWithValue("@SheetSize", ro.SheetSize);
                    cmd.Parameters.AddWithValue("@Typeof", ro.Typeof);
                    cmd.Parameters.AddWithValue("@TypeofCarton2", ro.TypeofCarton2);
                    cmd.Parameters.AddWithValue("@DMSNo", ro.DMSNo);

                    cmd.Parameters.AddWithValue("@TypeofPrimary", ro.TypeofPrimary);
                    cmd.Parameters.AddWithValue("@PrintingSystem", ro.PrintingSystem);
                    cmd.Parameters.AddWithValue("@Direction", ro.Direction);
                    cmd.Parameters.AddWithValue("@RollSheet", ro.RollSheet);
                    cmd.Parameters.AddWithValue("@RequestType", ro.RequestType);
                    cmd.Parameters.AddWithValue("@PlantAddress", ro.PlantAddress);

                    cmd.Parameters.AddWithValue("@Fixed_Desc", ro.Fixed_Desc);
                    cmd.Parameters.AddWithValue("@Inactive", ro.Inactive);
                    cmd.Parameters.AddWithValue("@Catching_Method", ro.Catching_Method);
                    cmd.Parameters.AddWithValue("@Scientific_Name", ro.Scientific_Name);
                    cmd.Parameters.AddWithValue("@Specie", ro.Specie);
                    cmd.Connection = con;
                    sda.SelectCommand = cmd;
                    sda.Fill(dt);
                    Context.Response.Write(JsonConvert.SerializeObject(dt));
                }
            }
        }
        //}
        //deletefile(datapath);
    }
    [WebMethod]
    public void savebyte(attachment ro)
    {
        //string datapath = "~/FileTest/" + data + ".json";
        //using (StreamReader sr = new StreamReader(Server.MapPath(datapath)))
        //{
        //string json = sr.ReadToEnd();
        //string sp = "spInsertMultipleRows";
        //List<attachment> ro = JsonConvert.DeserializeObject<List<attachment>>(json);
        //byte[] bytes = br.ReadBytes((Int32)fs.Length);
        using (SqlConnection con = new SqlConnection(strConn))
        {
            string query = "insert into tblFiles values (@Name,@ContentType,@Data,@MatDoc,@ActiveBy)";
            using (SqlCommand cmd = new SqlCommand(query))
            {
                cmd.Connection = con;
                cmd.Parameters.AddWithValue("@Name", ro.Name);
                cmd.Parameters.AddWithValue("@ContentType", ro.ContentType);
                cmd.Parameters.AddWithValue("@Data", ro.Data);
                cmd.Parameters.AddWithValue("@MatDoc", ro.MatDoc);
                cmd.Parameters.AddWithValue("@ActiveBy", ro.ActiveBy);
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
            }
        }
        //}
        Context.Response.Write("success");
        //deletefile(datapath);
    }
    //[WebMethod]
    //public void savebytencp(string data)
    //{
    //    string datapath = "~/FileTest/" + data + ".json";
    //    using (StreamReader sr = new StreamReader(Server.MapPath(datapath)))
    //    {
    //        string json = sr.ReadToEnd();
    //        //string sp = "spInsertMultipleRows";
    //        List<attachment> ro = JsonConvert.DeserializeObject<List<attachment>>(json);
    //        //byte[] bytes = br.ReadBytes((Int32)fs.Length);
    //        using (SqlConnection con = new SqlConnection(strConn))
    //        {
    //            string query = "insert into tbl_Files values (@Name,@ContentType,@Data,@MatDoc,@ActiveBy)";
    //            using (SqlCommand cmd = new SqlCommand(query))
    //            {
    //                cmd.Connection = con;
    //                cmd.Parameters.AddWithValue("@Name", ro[0].Name);
    //                cmd.Parameters.AddWithValue("@ContentType", ro[0].ContentType);
    //                cmd.Parameters.AddWithValue("@Data", ro[0].Data);
    //                cmd.Parameters.AddWithValue("@MatDoc", ro[0].MatDoc);
    //                cmd.Parameters.AddWithValue("@ActiveBy", ro[0].ActiveBy);
    //                con.Open();
    //                cmd.ExecuteNonQuery();
    //                con.Close();
    //            }
    //        }

    //    }
    //    Context.Response.Write("success");
    //    deletefile(datapath);
    //}
    [WebMethod]
    public void savechangeresult(string Name, string Result, string Matdoc, string Activeby)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spSaveChangeResult";
            cmd.Parameters.AddWithValue("@Name", Name);
            cmd.Parameters.AddWithValue("@Result", Result);
            cmd.Parameters.AddWithValue("@MatDoc", Matdoc);
            cmd.Parameters.AddWithValue("@ActiveBy", Activeby);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }

    //[WebMethod()]
    //public void savedocumentncp(string data)
    //{
    //    string datapath = "~/FileTest/" + data + ".json";
    //    using (StreamReader sr = new StreamReader(Server.MapPath(datapath)))
    //    {
    //        string json = sr.ReadToEnd();
    //        string sp = "spsavedocumentncp";
    //        List<ncpObject> ro = JsonConvert.DeserializeObject<List<ncpObject>>(json);
    //        //Context.Response.Write(ro[0].Requester);
    //        using (SqlConnection con = new SqlConnection(strConn))
    //        {
    //            using (SqlCommand cmd = new SqlCommand(sp))
    //            {
    //                using (SqlDataAdapter sda = new SqlDataAdapter())
    //                {
    //                    DataTable dt = new DataTable();
    //                    cmd.CommandType = CommandType.StoredProcedure;
    //                    cmd.Parameters.AddWithValue("@Id", ro[0].ID);
    //                    cmd.Parameters.AddWithValue("@ncptype", ro[0].ncptype);
    //                    cmd.Parameters.AddWithValue("@ncpid", ro[0].ncpid);
    //                    cmd.Parameters.AddWithValue("@Problem", ro[0].Problem);
    //                    cmd.Parameters.AddWithValue("@FirstDecision", ro[0].FirstDecision);
    //                    cmd.Parameters.AddWithValue("@Decision", ro[0].Decision);
    //                    cmd.Parameters.AddWithValue("@KeyDate", ro[0].KeyDate);
    //                    cmd.Parameters.AddWithValue("@Location", ro[0].Location);
    //                    cmd.Parameters.AddWithValue("@Plant", ro[0].Plant);
    //                    cmd.Parameters.AddWithValue("@MaterialType", ro[0].MaterialType);
    //                    cmd.Parameters.AddWithValue("@BatchCode", ro[0].BatchCode);
    //                    cmd.Parameters.AddWithValue("@Product", ro[0].Product);
    //                    cmd.Parameters.AddWithValue("@Batchsap", ro[0].Batchsap);
    //                    cmd.Parameters.AddWithValue("@Active", ro[0].Active);
    //                    cmd.Parameters.AddWithValue("@Material", ro[0].Material);
    //                    cmd.Parameters.AddWithValue("@ProductionDate", ro[0].ProductionDate);
    //                    cmd.Parameters.AddWithValue("@Quatity", ro[0].Quatity);
    //                    cmd.Parameters.AddWithValue("@Shift", ro[0].Shift);
    //                    cmd.Parameters.AddWithValue("@HoldQuatity", ro[0].HoldQuatity);
    //                    cmd.Parameters.AddWithValue("@Action", ro[0].Action);
    //                    cmd.Parameters.AddWithValue("@Remark", ro[0].Remark);
    //                    cmd.Parameters.AddWithValue("@Approve", ro[0].Approve);
    //                    cmd.Parameters.AddWithValue("@Approvefinal", ro[0].Approvefinal);
    //                    cmd.Connection = con;
    //                    sda.SelectCommand = cmd;
    //                    sda.Fill(dt);
    //                    //System.IO.File.Delete(Server.MapPath(datapath));
    //                    //string Path = Server.MapPath(datapath);
    //                    //	if (File.Exists(Path))
    //                    //	{
    //                    //		File.Delete(Path);
    //                    //	}
    //                    Context.Response.Write(JsonConvert.SerializeObject(dt));
    //                }
    //            }
    //        }
    //    }
    //    deletefile(datapath);
    //}

    [WebMethod()]
    public void savemaster(string data)
    {
        //data = data.Replace('@', '+');
        //List<masterObject> _x = JsonConvert.DeserializeObject<List<masterObject>>(data);
        //masterObject ro = _x[0];
        string datapath = "~/FileTest/" + data + ".json";
        using (StreamReader sr = new StreamReader(Server.MapPath(datapath)))
        {
            string json = sr.ReadToEnd();
            string sp = "spUpdateTransMaster";
            List<masterObject> ro = JsonConvert.DeserializeObject<List<masterObject>>(json);
            //Context.Response.Write(ro[0].Requester);
            using (SqlConnection con = new SqlConnection(strConn))
            {
                using (SqlCommand cmd = new SqlCommand(sp))
                {
                    using (SqlDataAdapter sda = new SqlDataAdapter())
                    {
                        DataTable dt = new DataTable();
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@Changed_Tabname", ro[0].Changed_Tabname);
                        cmd.Parameters.AddWithValue("@Changed_Charname", ro[0].Changed_Charname);
                        cmd.Parameters.AddWithValue("@Old_Id", ro[0].Old_Id);
                        cmd.Parameters.AddWithValue("@Id", ro[0].Id);
                        cmd.Parameters.AddWithValue("@Old_Description", string.Format("{0}", ro[0].Old_Description));
                        cmd.Parameters.AddWithValue("@Description", string.Format("{0}", ro[0].Description));
                        cmd.Parameters.AddWithValue("@Changed_Action", ro[0].Changed_Action);
                        cmd.Parameters.AddWithValue("@Changed_By", ro[0].Changed_By);
                        cmd.Parameters.AddWithValue("@Active", ro[0].Active);
                        cmd.Parameters.AddWithValue("@Material_Group", ro[0].Material_Group);
                        cmd.Parameters.AddWithValue("@Material_Type", ro[0].Material_Type);
                        cmd.Parameters.AddWithValue("@DescriptionText", string.Format("{0}", ro[0].DescriptionText));
                        cmd.Parameters.AddWithValue("@Can", ro[0].Can);
                        cmd.Parameters.AddWithValue("@LidType", ro[0].LidType);
                        cmd.Parameters.AddWithValue("@ContainerType", ro[0].ContainerType);
                        cmd.Parameters.AddWithValue("@DescriptionType", ro[0].DescriptionType);
                        cmd.Parameters.AddWithValue("@user_name", ro[0].user_name);
                        cmd.Parameters.AddWithValue("@fn", ro[0].fn);
                        cmd.Parameters.AddWithValue("@FirstName", ro[0].FirstName);
                        cmd.Parameters.AddWithValue("@LastName", ro[0].LastName);
                        cmd.Parameters.AddWithValue("@Email", ro[0].Email);
                        cmd.Parameters.AddWithValue("@Authorize_ChangeMaster", ro[0].Authorize_ChangeMaster);
                        cmd.Parameters.AddWithValue("@PrimaryCode", ro[0].PrimaryCode);
                        cmd.Parameters.AddWithValue("@GroupStyle", ro[0].GroupStyle);
                        cmd.Parameters.AddWithValue("@PackingStyle", ro[0].PackingStyle);
                        cmd.Parameters.AddWithValue("@RefStyle", ro[0].RefStyle);
                        cmd.Parameters.AddWithValue("@Packsize", ro[0].Packsize);
                        cmd.Parameters.AddWithValue("@BaseUnit", ro[0].BaseUnit);
                        cmd.Parameters.AddWithValue("@TypeofPrimary", ro[0].TypeofPrimary);
                        cmd.Parameters.AddWithValue("@RegisteredNo", ro[0].RegisteredNo);
                        cmd.Parameters.AddWithValue("@Address", ro[0].Address);
                        cmd.Parameters.AddWithValue("@Plant", ro[0].Plant);

                        cmd.Parameters.AddWithValue("@Product_Group", ro[0].Product_Group);
                        cmd.Parameters.AddWithValue("@Product_GroupDesc", ro[0].Product_GroupDesc);
                        cmd.Parameters.AddWithValue("@PRD_Plant", ro[0].PRD_Plant);

                        cmd.Parameters.AddWithValue("@WHNumber", ro[0].WHNumber);
                        cmd.Parameters.AddWithValue("@StorageType", ro[0].StorageType);
                        cmd.Parameters.AddWithValue("@LE_Qty", ro[0].LE_Qty);
                        cmd.Parameters.AddWithValue("@Storage_UnitType", ro[0].Storage_UnitType);

                        cmd.Parameters.AddWithValue("@Changed_Reason", ro[0].Changed_Reason);

                        cmd.Parameters.AddWithValue("@SAP_EDPUsername", ro[0].SAP_EDPUsername);
                        //cmd.Parameters.AddWithValue("@SAP_EDPPassword", ro[0].SAP_EDPPassword);	

                        cmd.Connection = con;
                        sda.SelectCommand = cmd;
                        sda.Fill(dt);
                        //System.IO.File.Delete(Server.MapPath(datapath));
                        //string Path = Server.MapPath(datapath);
                        //	if (File.Exists(Path))
                        //	{
                        //		File.Delete(Path);
                        //	}
                        Context.Response.Write(JsonConvert.SerializeObject(dt));
                    }
                }
            }
        }
        //deletefile(datapath);
    }
    //	public void DeleteFileFromFolder(string StrFilename)
    //	{
    //		string datapath = "~/FileTest/" + StrFilename + ".json";
    //		string strPhysicalFolder = Server.MapPath("..\\");
    //
    //		string strFileFullPath = strPhysicalFolder + StrFilename;
    //
    //		if (IO.File.Exists(strFileFullPath)) {
    //			IO.File.Delete(strFileFullPath);
    //		}
    //
    //	}
    [WebMethod()]
    public void testsavemaster2(string data)
    {
         List<masterObject> _x = JsonConvert.DeserializeObject<List<masterObject>>(data);
        masterObject ro = _x[0];
        var ascii = Encoding.UTF8.GetBytes(ro.Changed_Tabname);
        var text = Encoding.UTF8.GetString(ascii);
        Context.Response.Write(JsonConvert.SerializeObject(ro));
    }
        [WebMethod()]
    public void savemaster2(string data)
    {
        //data = data.Replace('@', '+');
        List<masterObject> _x = JsonConvert.DeserializeObject<List<masterObject>>(data);
        masterObject ro = _x[0];
        //string datapath = "~/FileTest/" + data + ".json";
        //using (StreamReader sr = new StreamReader(Server.MapPath(datapath)))
        //{
        //string json = sr.ReadToEnd();
        string sp = "spUpdateTransMaster";
        //List<masterObject> ro = JsonConvert.DeserializeObject<List<masterObject>>(json);
        //Context.Response.Write(ro[0].Requester);
        using (SqlConnection con = new SqlConnection(strConn))
        {
            using (SqlCommand cmd = new SqlCommand(sp))
            {
                using (SqlDataAdapter sda = new SqlDataAdapter())
                {
                    DataTable dt = new DataTable();
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@Changed_Tabname", ro.Changed_Tabname);
                    cmd.Parameters.AddWithValue("@Changed_Charname", ro.Changed_Charname);
                    cmd.Parameters.AddWithValue("@Old_Id", ro.Old_Id);
                    cmd.Parameters.AddWithValue("@Id", ro.Id);
                    cmd.Parameters.AddWithValue("@Old_Description", string.Format("{0}", ro.Old_Description));
                    cmd.Parameters.AddWithValue("@Description", string.Format("{0}", ro.Description));
                    cmd.Parameters.AddWithValue("@Changed_Action", ro.Changed_Action);
                    cmd.Parameters.AddWithValue("@Changed_By", ro.Changed_By);
                    cmd.Parameters.AddWithValue("@Active", ro.Active);
                    cmd.Parameters.AddWithValue("@Material_Group", ro.Material_Group);
                    cmd.Parameters.AddWithValue("@Material_Type", ro.Material_Type);
                    cmd.Parameters.AddWithValue("@DescriptionText", string.Format("{0}", ro.DescriptionText));
                    cmd.Parameters.AddWithValue("@Can", ro.Can);
                    cmd.Parameters.AddWithValue("@LidType", ro.LidType);
                    cmd.Parameters.AddWithValue("@ContainerType", ro.ContainerType);
                    cmd.Parameters.AddWithValue("@DescriptionType", ro.DescriptionType);
                    cmd.Parameters.AddWithValue("@user_name", ro.user_name);
                    cmd.Parameters.AddWithValue("@fn", ro.fn);
                    cmd.Parameters.AddWithValue("@FirstName", ro.FirstName);
                    cmd.Parameters.AddWithValue("@LastName", ro.LastName);
                    cmd.Parameters.AddWithValue("@Email", ro.Email);
                    cmd.Parameters.AddWithValue("@Authorize_ChangeMaster", ro.Authorize_ChangeMaster);
                    cmd.Parameters.AddWithValue("@PrimaryCode", ro.PrimaryCode);
                    cmd.Parameters.AddWithValue("@GroupStyle", ro.GroupStyle);
                    cmd.Parameters.AddWithValue("@PackingStyle", ro.PackingStyle);
                    cmd.Parameters.AddWithValue("@RefStyle", ro.RefStyle);
                    cmd.Parameters.AddWithValue("@Packsize", ro.Packsize);
                    cmd.Parameters.AddWithValue("@BaseUnit", ro.BaseUnit);
                    cmd.Parameters.AddWithValue("@TypeofPrimary", ro.TypeofPrimary);
                    cmd.Parameters.AddWithValue("@RegisteredNo", ro.RegisteredNo);
                    cmd.Parameters.AddWithValue("@Address", ro.Address);
                    cmd.Parameters.AddWithValue("@Plant", ro.Plant);

                    cmd.Parameters.AddWithValue("@Product_Group", ro.Product_Group);
                    cmd.Parameters.AddWithValue("@Product_GroupDesc", ro.Product_GroupDesc);
                    cmd.Parameters.AddWithValue("@PRD_Plant", ro.PRD_Plant);

                    cmd.Parameters.AddWithValue("@WHNumber", ro.WHNumber);
                    cmd.Parameters.AddWithValue("@StorageType", ro.StorageType);
                    cmd.Parameters.AddWithValue("@LE_Qty", ro.LE_Qty);
                    cmd.Parameters.AddWithValue("@Storage_UnitType", ro.Storage_UnitType);

                    cmd.Parameters.AddWithValue("@Changed_Reason", ro.Changed_Reason);

                    cmd.Parameters.AddWithValue("@SAP_EDPUsername", ro.SAP_EDPUsername);
                    //cmd.Parameters.AddWithValue("@SAP_EDPPassword", ro.SAP_EDPPassword);	

                    cmd.Connection = con;
                    sda.SelectCommand = cmd;
                    sda.Fill(dt);
                    //System.IO.File.Delete(Server.MapPath(datapath));
                    //string Path = Server.MapPath(datapath);
                    //	if (File.Exists(Path))
                    //	{
                    //		File.Delete(Path);
                    //	}
                    Context.Response.Write(JsonConvert.SerializeObject(dt));
                }
            }
        }
        //}
        //deletefile(datapath);
    }
    //	public void DeleteFileFromFolder(string StrFilename)
    //	{
    //		string datapath = "~/FileTest/" + StrFilename + ".json";
    //		string strPhysicalFolder = Server.MapPath("..\\");
    //
    //		string strFileFullPath = strPhysicalFolder + StrFilename;
    //
    //		if (IO.File.Exists(strFileFullPath)) {
    //			IO.File.Delete(strFileFullPath);
    //		}
    //
    //	}
    [WebMethod]
    public void saveinfogroup(string Id, string Result)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            string query = "Update SapMaterial set FinalInfoGroup = @Result Where Id=@Id";
            using (SqlCommand cmd = new SqlCommand(query))
            {
                cmd.Connection = con;
                cmd.Parameters.AddWithValue("@Id", Id);
                cmd.Parameters.AddWithValue("@Result", Result);
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
            }
        }
        Context.Response.Write("success");
    }
    [WebMethod]
    public void GetImpactedMatDesc(string data)
    {
        List<objImpactedMatDesc> _x = JsonConvert.DeserializeObject<List<objImpactedMatDesc>>(data);
        objImpactedMatDesc ro = _x[0];
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spGetImpactedMatDesc";
            cmd.Parameters.AddWithValue("@Material", ro.name);
            cmd.Parameters.AddWithValue("@User", ro.User);
            cmd.Parameters.AddWithValue("@FrDt", ro.FrDt);
            cmd.Parameters.AddWithValue("@ToDt", ro.ToDt);
            cmd.Parameters.AddWithValue("@Status", ro.Status);
            cmd.Parameters.AddWithValue("@MasterName", ro.MasterName);
            cmd.Parameters.AddWithValue("@Action", ro.Action);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    [WebMethod]
    public void saveImpactedMatDesc(string Id, string Reason, string Status)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spSaveImpactedMatDesc";
            cmd.Parameters.AddWithValue("@Id", Id);
            cmd.Parameters.AddWithValue("@Reason", Reason);
            cmd.Parameters.AddWithValue("@Status", Status);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    [WebMethod]
    public void UpdateImpactedmat2(string Changed_Id, string Changed_Action, string Material, string Description, string DMSNo, string New_Material, string New_Description, string Status, string Reason, string NewMat_JobId, string Char_Name, string Char_OldValue, string Char_NewValue)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spUpdateImpactedmat";

            cmd.Parameters.AddWithValue("@Changed_Id", Changed_Id);
            cmd.Parameters.AddWithValue("@Changed_Action", Changed_Action);
            cmd.Parameters.AddWithValue("@Material", Material);
            cmd.Parameters.AddWithValue("@Description", Description);
            cmd.Parameters.AddWithValue("@DMSNo", DMSNo);
            cmd.Parameters.AddWithValue("@New_Material", New_Material);
            cmd.Parameters.AddWithValue("@New_Description", New_Description);
            cmd.Parameters.AddWithValue("@Status", Status);
            cmd.Parameters.AddWithValue("@Reason", Reason);
            cmd.Parameters.AddWithValue("@NewMat_JobId", NewMat_JobId);
            cmd.Parameters.AddWithValue("@Char_Name", Char_Name);
            cmd.Parameters.AddWithValue("@Char_OldValue", Char_OldValue);
            cmd.Parameters.AddWithValue("@Char_NewValue", Char_NewValue);


            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }


    }
    [WebMethod]
    public void UpdateImpactedmat(string data)
    {
        List<objImpactedmat> _x = JsonConvert.DeserializeObject<List<objImpactedmat>>(data);
        objImpactedmat ro = _x[0];
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spUpdateImpactedmat";

            cmd.Parameters.AddWithValue("@Changed_Id", ro.Changed_Id);
            cmd.Parameters.AddWithValue("@Changed_Action", ro.Changed_Action);
            cmd.Parameters.AddWithValue("@Material", ro.Material);
            cmd.Parameters.AddWithValue("@Description", ro.Description);
            cmd.Parameters.AddWithValue("@DMSNo", ro.DMSNo);
            cmd.Parameters.AddWithValue("@New_Material", ro.New_Material);
            cmd.Parameters.AddWithValue("@New_Description", ro.New_Description);
            cmd.Parameters.AddWithValue("@Status", ro.Status);
            cmd.Parameters.AddWithValue("@Reason", ro.Reason);
            cmd.Parameters.AddWithValue("@NewMat_JobId", ro.NewMat_JobId);
            cmd.Parameters.AddWithValue("@Char_Name", ro.Char_Name);
            cmd.Parameters.AddWithValue("@Char_OldValue", ro.Char_OldValue);
            cmd.Parameters.AddWithValue("@Char_NewValue", ro.Char_NewValue);


            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }

    [WebMethod]
    public void Delete_UnusedJob(string Id)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spDeleteUnusedJob";
            cmd.Parameters.AddWithValue("@Id", Id);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }

    [WebMethod]
    public void ReUpload(string Id)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spReUpload";
            cmd.Parameters.AddWithValue("@Id", Id);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    [WebMethod]
    public void Extend(string Id)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "update SapMaterial set StatusApp = '8' 	where id = @Id";
            cmd.Parameters.AddWithValue("@Id", Id);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }
    [WebMethod]
    public void InactiveMat(string Material)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spInactiveMat";
            cmd.Parameters.AddWithValue("@Material", Material);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }

    [WebMethod]
    public void ReactiveMat(string Material)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spReactiveMat";
            cmd.Parameters.AddWithValue("@Material", Material);
            cmd.Connection = con;
            con.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
            oAdapter.Fill(dt);
            con.Close();
            Context.Response.Write(JsonConvert.SerializeObject(dt));
        }
    }

    //	[WebMethod]
    //    public void Update_ulogin(string user_name,string SAP_EDPPassword)
    //    {
    //        using (SqlConnection con = new SqlConnection(strConn))
    //        {
    //            SqlCommand cmd = new SqlCommand();
    //            cmd.CommandType = CommandType.StoredProcedure;
    //            cmd.CommandText = "spUpdate_ulogin";
    //			cmd.Parameters.AddWithValue("@user_name", user_name);
    //			cmd.Parameters.AddWithValue("@SAP_EDPPassword", SAP_EDPPassword);
    //            cmd.Connection = con;
    //            con.Open();
    //            DataTable dt = new DataTable();
    //            SqlDataAdapter oAdapter = new SqlDataAdapter(cmd);
    //            oAdapter.Fill(dt);
    //            con.Close();
    //            Context.Response.Write(JsonConvert.SerializeObject(dt));
    //        }
    //    }
}
public class Assign
{
    public string Id { get; set; }
    public string Assignee { get; set; }
}
public class RequestDocument
{
    public string CreateBy { get; set; }
    public string Condition { get; set; }
    public string ProductCode { get; set; }
    public string Material { get; set; }
    public string ProductGroup { get; set; }
}
public class CreateDocument
{
    public string CreateBy { get; set; }
    public string Condition { get; set; }
    public string Code { get; set; }
}
public class AppObject
{
    public string ActiveBy { get; set; }
    public string Id { get; set; }
    public string fn { get; set; }
    public string StatusApp { get; set; }
    public string Remark { get; set; }
}
public class attachment
{
    public string Name { get; set; }
    public string ContentType { get; set; }
    public byte[] Data { get; set; }
    public int MatDoc { get; set; }
    public string ActiveBy { get; set; }
}
public class sapmaterial
{
    //public string Assignee { get; set; }
    public string PlantAddress { get; set; }
    public string RequestType { get; set; }
    public string LidType { get; set; }
    public string ContainerType { get; set; }
    public string StyleofPrinting { get; set; }
    public string Packing { get; set; }
    public string PackingStyle { get; set; }
    public string Description { get; set; }
    public string Brand { get; set; }
    public string PrimarySize { get; set; }
    public string Version { get; set; }
    public string ChangePoint { get; set; }
    public string MaterialGroup { get; set; }
    public string CreateBy { get; set; }
    public string RequestNo { get; set; }
    public string StatusApp { get; set; }
    public string Userlevel { get; set; }
    public string SheetSize { get; set; }
    public string TotalColour { get; set; }
    public string ProductCode { get; set; }
    public string FAOZone { get; set; }
    public string Plant { get; set; }
    public string Processcolour { get; set; }
    public string PlantRegisteredNo { get; set; }
    public string CompanyNameAddress { get; set; }
    public string PMScolour { get; set; }
    public string Symbol { get; set; }
    public string CatchingArea { get; set; }
    public string CatchingPeriodDate { get; set; }
    public string Grandof { get; set; }
    public string Flute { get; set; }
    public string Vendor { get; set; }
    public string Dimension { get; set; }
    public string RSC { get; set; }
    public string Accessories { get; set; }
    public string PrintingStyleofPrimary { get; set; }
    public string PrintingStyleofSecondary { get; set; }
    public string CustomerDesign { get; set; }
    public string CustomerSpec { get; set; }
    public string CustomerSize { get; set; }
    public string CustomerVendor { get; set; }
    public string CustomerColor { get; set; }
    public string CustomerScanable { get; set; }
    public string CustomerBarcodeSpec { get; set; }
    public string FirstInfoGroup { get; set; }
    public string SO { get; set; }
    public string PICMkt { get; set; }
    public string SOPlant { get; set; }
    public string Destination { get; set; }
    public string Remark { get; set; }
    public string GrossWeight { get; set; }
    public string FinalInfoGroup { get; set; }
    public string Note { get; set; }
    public string TypeofCarton2 { get; set; }
    public string Typeof { get; set; }
    public string DMSNo { get; set; }
    public string TypeofPrimary { get; set; }
    public string PrintingSystem { get; set; }
    public string Direction { get; set; }
    public string RollSheet { get; set; }
    public string Fixed_Desc { get; set; }
    public string Inactive { get; set; }
    public string Catching_Method { get; set; }
    public string Scientific_Name { get; set; }
    public string Specie { get; set; }
    //public string ReferenceMaterial { get; set; }
    public CreateDocument _root { get; set; }
}
//public class RootObject
//{
//    public string RequestNo { get; set; }
//    public string Requester { get; set; }
//    public string Company { get; set; }
//    public DateTime from { get; set; }
//    public DateTime to { get; set; }
//    public string PetCategory { get; set; }
//    public string PetFoodType { get; set; }
//    public string CompliedWith { get; set; }
//    public string NutrientProfile { get; set; }
//    public string Requestfor { get; set; }
//    public string ProductType { get; set; }
//    public string ProductStyle { get; set; }
//    public string Media { get; set; }
//    public string ChunkType { get; set; }
//    public string NetWeight { get; set; }
//    public string PackSize { get; set; }
//    public string Packaging { get; set; }
//    public string Material { get; set; }
//    public string PackType { get; set; }
//    public string PackDesign { get; set; }
//    public string PackColor { get; set; }
//    public string PackLid { get; set; }
//    public string PackShape { get; set; }
//    public string PackLacquer { get; set; }
//    public string SellingUnit { get; set; }
//    public string MarketingNumber { get; set; }
//}
//public class ncpObject
//{
//    public string ID { get; set; }
//    public string ncptype { get; set; }
//    public string ncpid { get; set; }
//    public string Problem { get; set; }
//    public string FirstDecision { get; set; }
//    public string Decision { get; set; }
//    public string KeyDate { get; set; }
//    public string Location { get; set; }
//    public string Plant { get; set; }
//    public string BatchCode { get; set; }
//    public string Product { get; set; }
//    public string Batchsap { get; set; }
//    public string Active { get; set; }
//    public string Material { get; set; }
//    public string MaterialType { get; set; }
//    public string ProductionDate { get; set; }
//    public string Quatity { get; set; }
//    public string Shift { get; set; }
//    public string HoldQuatity { get; set; }
//    public string Action { get; set; }
//    public string Remark { get; set; }
//    public string Approve { get; set; }
//    public string Approvefinal { get; set; }
//}
public class masterObject
{
    public string Changed_Tabname { get; set; }
    public string Changed_Charname { get; set; }
    public string Old_Id { get; set; }
    public string Id { get; set; }
    public string Old_Description { get; set; }
    public string Description { get; set; }
    public string Changed_Action { get; set; }
    public string Changed_By { get; set; }
    public string Active { get; set; }
    public string Material_Group { get; set; }
    public string Material_Type { get; set; }
    public string DescriptionText { get; set; }
    public string Can { get; set; }
    public string LidType { get; set; }
    public string ContainerType { get; set; }
    public string DescriptionType { get; set; }
    public string user_name { get; set; }
    public string fn { get; set; }
    public string FirstName { get; set; }
    public string LastName { get; set; }
    public string Email { get; set; }
    public string Authorize_ChangeMaster { get; set; }
    public string PrimaryCode { get; set; }
    public string GroupStyle { get; set; }
    public string PackingStyle { get; set; }
    public string RefStyle { get; set; }
    public string Packsize { get; set; }
    public string BaseUnit { get; set; }
    public string TypeofPrimary { get; set; }
    public string RegisteredNo { get; set; }
    public string Address { get; set; }
    public string Plant { get; set; }
    public string Product_Group { get; set; }
    public string Product_GroupDesc { get; set; }
    public string PRD_Plant { get; set; }
    public string WHNumber { get; set; }
    public string StorageType { get; set; }
    public string LE_Qty { get; set; }
    public string Storage_UnitType { get; set; }
    public string Changed_Reason { get; set; }

    public string SAP_EDPUsername { get; set; }
    //public string SAP_EDPPassword { get; set; }

}
public class InboundArtwork
{
    public string ArtworkNumber { get; set; }
    public string Date { get; set; }
    public string Time { get; set; }
    public string Characteristic { get; set; }
    public string Value { get; set; }
    public string Description { get; set; }
}
public class ArtworkObject
{
    public string ArtworkNumber { get; set; }
    public string Date { get; set; }
    public string Time { get; set; }
    public string RecordType { get; set; }
    public string MaterialNumber { get; set; }
    public string MaterialDescription { get; set; }
    public string MaterialCreatedDate { get; set; }
    public string ArtworkURL { get; set; }
    public string Status { get; set; }
    public string PAUserName { get; set; }
    public string PGUserName { get; set; }
    public string ReferenceMaterial { get; set; }
    public string Plant { get; set; }
    public string PrintingStyleofPrimary { get; set; }
    public string PrintingStyleofSecondary { get; set; }
    public string CustomersDesign { get; set; }
    public string CustomersDesignDetail { get; set; }
    public string CustomersSpec { get; set; }
    public string CustomersSpecDetail { get; set; }
    public string CustomersSize { get; set; }
    public string CustomersSizeDetail { get; set; }
    public string CustomerNominatesVendor { get; set; }
    public string CustomerNominatesVendorDetail { get; set; }
    public string CustomerNominatesColorPantone { get; set; }
    public string CustomerNominatesColorPantoneDetail { get; set; }
    public string CustomersBarcodeScanable { get; set; }
    public string CustomersBarcodeScanableDetail { get; set; }
    public string CustomersBarcodeSpec { get; set; }
    public string CustomersBarcodeSpecDetail { get; set; }
    public string FirstInfoGroup { get; set; }
    public string SONumber { get; set; }
    public string SOitem { get; set; }
    public string SOPlant { get; set; }
    public string PICMKT { get; set; }
    public string Destination { get; set; }
    public string RemarkNoteofPA { get; set; }
    public string FinalInfoGroup { get; set; }
    public string RemarkNoteofPG { get; set; }
    public string CompleteInfoGroup { get; set; }
    public string ProductionExpirydatesystem { get; set; }
    public string Seriousnessofcolorprinting { get; set; }
    public string CustIngreNutritionAnalysis { get; set; }
    public string ShadeLimit { get; set; }
    public string PackageQuantity { get; set; }
    public string WastePercent { get; set; }

    public string SustainMaterial { get; set; }
    public string SustainPlastic { get; set; }
    public string SustainReuseable { get; set; }
    public string SustainRecyclable { get; set; }
    public string SustainComposatable { get; set; }
    public string SustainCertification { get; set; }
    public string SustainCertSourcing { get; set; }

    public string SustainOther { get; set; }
    public string SusSecondaryPKGWeight { get; set; }
    public string SusRecycledContent { get; set; }
}
public class objresp
{
    public string msg { get; set; }
    public string status { get; set; }
}
public class objpersonal
{
    public string Group { get; set; }
    public string By { get; set; }
    public string From { get; set; }
    public string To { get; set; }
    public string Product { get; set; }
    public string Material { get; set; }
    public string Workflow { get; set; }
    public string StatusApp { get; set; }
    public string where { get; set; }
    public string Brand { get; set; }
    public string Version { get; set; }
    public string Plant { get; set; }
    public string PrimarySize { get; set; }
    public string PlantRegistered { get; set; }
    public string TypeOf { get; set; }
    public string CreateBy { get; set; }
    public string PackingStyle { get; set; }
    public string DocumentNo { get; set; }
    public string Record { get; set; }
    public string user { get; set; }
}
public class objImpactedmat
{
    public string Changed_Id { get; set; }
    public string Changed_Action { get; set; }
    public string Material { get; set; }
    public string Description { get; set; }
    public string DMSNo { get; set; }
    public string New_Material { get; set; }
    public string New_Description { get; set; }
    public string Status { get; set; }
    public string Reason { get; set; }
    public string NewMat_JobId { get; set; }
    public string Char_Name { get; set; }
    public string Char_OldValue { get; set; }
    public string Char_NewValue { get; set; }
}
public class objImpactedMatDesc
{
    public string name { get; set; }
    public string User { get; set; }
    public string FrDt { get; set; }
    public string ToDt { get; set; }
    public string Status { get; set; }
    public string MasterName { get; set; }
    public string Action { get; set; }
}
public class objTransMaster
{
    public string Changed_Tabname { get; set; }
    public string Changed_Charname { get; set; }
    public string Old_Id { get; set; }
    public string Id { get; set; }
    public string Old_Description { get; set; }
    public string Description { get; set; }
    public string Changed_Action { get; set; }
    public string Changed_By { get; set; }
    public string Active { get; set; }
    public string Material_Group { get; set; }
    public string Material_Type { get; set; }
    public string DescriptionText { get; set; }
    public string Can { get; set; }
    public string LidType { get; set; }
    public string ContainerType { get; set; }
    public string DescriptionType { get; set; }
    public string user_name { get; set; }
    public string fn { get; set; }
    public string FirstName { get; set; }
    public string LastName { get; set; }
    public string Email { get; set; }
    public string Authorize_ChangeMaster { get; set; }
    public string PrimaryCode { get; set; }
    public string GroupStyle { get; set; }
    public string PackingStyle { get; set; }
    public string RefStyle { get; set; }
    public string Packsize { get; set; }
    public string BaseUnit { get; set; }
    public string TypeofPrimary { get; set; }
    public string RegisteredNo { get; set; }
    public string Address { get; set; }
    public string Plant { get; set; }
    public string Product_Group { get; set; }
    public string Product_GroupDesc { get; set; }
    public string PRD_Plant { get; set; }
    public string WHNumber { get; set; }
    public string StorageType { get; set; }
    public string LE_Qty { get; set; }
    public string Storage_UnitType { get; set; }
    public string Changed_Reason { get; set; }
}
public class objPrimarySize
{
    public string Item { get; set; }
    public string Id { get; set; }
    public string Description { get; set; }
    public string Can { get; set; }
    public string LidType { get; set; }
    public string ContainerType { get; set; }
    public string DescriptionType { get; set; }
    public string Changed_Action { get; set; }
    public string Changed_By { get; set; }
    public string Active { get; set; }
}