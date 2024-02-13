using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Web;

/// <summary>
/// Summary description for MyDataModule
/// </summary>
public class MyDataModule
{
    public string strConn = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
    public static string user = HttpContext.Current.User.Identity.Name.Replace(@"THAIUNION\", @"");
    public MyDataModule()
    {
        //
        // TODO: Add constructor logic here
        //
    }
    public DataTable builditems(string data)
    {
        using (SqlConnection oConn = new SqlConnection(strConn))
        {
            oConn.Open();
            string strQuery = data;
            DataTable dt = new DataTable();
            SqlDataAdapter oAdapter = new SqlDataAdapter(strQuery, oConn);
            // Fill the dataset.
            oAdapter.Fill(dt);
            oConn.Close();
            oConn.Dispose();
            return dt;
        }
    }
    public void GetExecuteNonQuery(string StoredProcedure, object[] Parameters)
    {
        var Results = new DataTable();
        try
        {
            using (SqlConnection conn = new SqlConnection(strConn))
            {
                using (SqlCommand cmd = new SqlCommand(StoredProcedure, conn))
                {
                    conn.Open();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddRange(Parameters);
                    cmd.ExecuteNonQuery();
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }
    public string splittext(string data,int n)
    {
		int i = n; i++; 
        string CustomerDesign = string.Format("{0}", data);
        string[] words = CustomerDesign.Split('|');
		if (i > words.Length) return "";
        return words[n].ToString();
    }
    public void ClearCache()
    {
        HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache);
        HttpContext.Current.Response.Cache.SetExpires(DateTime.Now);
        HttpContext.Current.Response.Cache.SetNoServerCaching();
        HttpContext.Current.Response.Cache.SetNoStore();
        HttpContext.Current.Response.Cookies.Clear();
        HttpContext.Current.Request.Cookies.Clear();
    }

    public void clearchachelocalall()
    {
        string GooglePath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\AppData\Local\Google\Chrome\User Data\Default\";
        string MozilaPath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\AppData\Roaming\Mozilla\Firefox\";
        string Opera1 = Environment.GetEnvironmentVariable("USERPROFILE") + @"\AppData\Local\Opera\Opera";
        string Opera2 = Environment.GetEnvironmentVariable("USERPROFILE") + @"\AppData\Roaming\Opera\Opera";
        string Safari1 = Environment.GetEnvironmentVariable("USERPROFILE") + @"\AppData\Local\Apple Computer\Safari";
        string Safari2 = Environment.GetEnvironmentVariable("USERPROFILE") + @"\AppData\Roaming\Apple Computer\Safari";
        string IE1 = Environment.GetEnvironmentVariable("USERPROFILE") + @"\AppData\Local\Microsoft\Intern~1";
        string IE2 = Environment.GetEnvironmentVariable("USERPROFILE") + @"\AppData\Local\Microsoft\Windows\History";
        string IE3 = Environment.GetEnvironmentVariable("USERPROFILE") + @"\AppData\Local\Microsoft\Windows\Tempor~1";
        string IE4 = Environment.GetEnvironmentVariable("USERPROFILE") + @"\AppData\Roaming\Microsoft\Windows\Cookies";
        string Flash = Environment.GetEnvironmentVariable("USERPROFILE") + @"\AppData\Roaming\Macromedia\Flashp~1";

        //Call This Method ClearAllSettings and Pass String Array Param
        ClearAllSettings(new string[] { GooglePath, MozilaPath, Opera1, Opera2, Safari1, Safari2, IE1, IE2, IE3, IE4, Flash });

    }

    public void ClearAllSettings(string[] ClearPath)
    {
        foreach (string HistoryPath in ClearPath)
        {
            if (Directory.Exists(HistoryPath))
            {
                DoDelete(new DirectoryInfo(HistoryPath));
            }

        }
    }
    public SqlDataReader executeProcedure(string commandName,  Dictionary<string, object> _params)
    {
        SqlConnection conn = new SqlConnection(strConn);
        conn.Open();
        SqlCommand comm = conn.CreateCommand();
        comm.CommandType = CommandType.StoredProcedure;
        comm.CommandText = commandName;
        if (_params != null)
    {
            foreach (KeyValuePair<string, object> kvp in _params)
            comm.Parameters.Add(new SqlParameter(kvp.Key, kvp.Value));
        }
        return comm.ExecuteReader(System.Data.CommandBehavior.CloseConnection);
    }
    public DataTable GetRelatedResources(string StoredProcedure, object[] Parameters)
    {
        var Results = new DataTable();
        try
        {
            using (SqlConnection conn = new SqlConnection(strConn))
            {
                using (SqlCommand cmd = new SqlCommand(StoredProcedure, conn))
                {
                    conn.Open();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddRange(Parameters);

                    SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                    adapter.Fill(Results);
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
        catch (Exception ex)
        {
            throw ex;
        }
        return Results;
    }
    void DoDelete(DirectoryInfo folder)
    {
        try
        {

            foreach (FileInfo file in folder.GetFiles())
            {
                try
                {
                    file.Delete();
                }
                catch
                { }

            }
            foreach (DirectoryInfo subfolder in folder.GetDirectories())
            {
                DoDelete(subfolder);
            }
        }
        catch
        {
        }
    }
   public void sendemail(string MailTo, 
        string MailCc, 
        string _Body, 
        string _Subject,string _Attachments)
    {
        //MailTo = "voravut.somboornpong@thaiunion.com"; MailCc = "";
		/*insertsendmail(MailTo,MailCc,_Body,_Subject);
        MailMessage msg = new MailMessage();
        if (string.IsNullOrEmpty(MailTo)) return;
        string[] words = MailTo.Split(';');
        foreach (string word in words)
        {
            if (!string.IsNullOrEmpty(word))
                msg.To.Add(new MailAddress(word));
        }
        string[] c = MailCc.Split(';');
        foreach (string s in c)
            if (!string.IsNullOrEmpty(s))
                msg.CC.Add(new MailAddress(s));
        msg.From = new MailAddress("wshuttleadm@thaiunion.com");
        msg.Subject = _Subject;
        msg.Body = _Body;// "Material  " + _Material.ToString() + " Created sap Complate";
        if(!string.IsNullOrEmpty(_Attachments)) { 
            //@"C:\\inetpub\wwwroot\WebService_jQuery_Ajax\FileTest\textfile.log"
            msg.Attachments.Add(new Attachment(_Attachments));
            } 
        msg.IsBodyHtml = true;
        SmtpClient client = new SmtpClient();
        client.UseDefaultCredentials = false;
        client.Credentials = new System.Net.NetworkCredential("wshuttleadm@thaiunion.com", "WSP@ss2018");
        client.Port = 587; // You can use Port 25 if 587 is blocked (mine is!)
        client.Host = "smtp.office365.com";
        client.DeliveryMethod = SmtpDeliveryMethod.Network;
        client.EnableSsl = true;
        client.Send(msg);*/
    }
	public string insertsendmail(string MailTo, string MailCc, string _Body, string _Subject)
    {
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.Text;
            //SELECT FirstName + '.' + LastName + '@thaiunion.com' AS Email"
            cmd.CommandText = "Insert into MailData values(@Sender,@To,@Cc,'',@Subject,@Body,getdate(),1,getdate(),'TEXT',1,0)";
            cmd.Parameters.AddWithValue("@Sender", String.Format("{0}",10));
            cmd.Parameters.AddWithValue("@To", MailTo.ToString());
            cmd.Parameters.AddWithValue("@Cc", MailCc.ToString());
            cmd.Parameters.AddWithValue("@Subject", _Subject.ToString());
            cmd.Parameters.AddWithValue("@Body", _Body.ToString());
            cmd.Connection = con;
            con.Open();
            var getValue = cmd.ExecuteScalar();
            con.Close();
            return ((string)getValue == null) ? string.Empty : getValue.ToString();
        }
    }
    public string ReadItems(string strQuery)
    {
        string result = "";
        // (ByVal FieldName As String, ByVal TableName As String, ByVal Cur As String, ByVal Value As String) As String
        DataTable dt = new DataTable();
        SqlConnection con = new SqlConnection(strConn);
        SqlDataAdapter sda = new SqlDataAdapter();
        SqlCommand cmd = new SqlCommand(strQuery);
        cmd.CommandType = CommandType.Text;
        cmd.Connection = con;
        con.Open();
        sda.SelectCommand = cmd;
        sda.Fill(dt);
        con.Close();
        con.Dispose();
        StringBuilder sb = new StringBuilder();
        if (dt.Rows.Count > 0)
        {
            foreach (DataRow row in dt.Rows)
            {
                sb.Append(row[0] + ",");
            }
            if (result.Length < 2)
            {
                result = sb.ToString();
                result = result.Substring(0, (result.Length - 1));
            }
        }
        return result;
    }

    public string Getuser(string user_name, string type)
    {
        string strData = "";
        string strSQL = @"select * from ulogin 
            where [user_name]='" + string.Format("{0}", user_name) + "' and isnull(Inactive,'')<>'X'";
        if (strSQL == "") return strData;
        using (SqlConnection con = new SqlConnection(strConn))
        {
            SqlDataAdapter da = new SqlDataAdapter(strSQL, strConn);
            DataSet ds = new DataSet();
            da.Fill(ds);
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                switch (type.ToLower())
                {
                    case "fullname": case "fn":
                        strData = string.Format("{0} {1}", dr["FirstName"], dr["LastName"]);
                        break;
                    case "email":
                        strData = dr["email"].ToString();
                        break;
                }
            }
        }
        return strData;
    }
}