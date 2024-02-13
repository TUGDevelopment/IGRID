using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Web;
using System.Web.Script.Serialization;
using System.Web.UI;
using System.Web.UI.WebControls;


public partial class _Default : System.Web.UI.Page
{
    WebService myclass = new WebService();
	MyDataModule cs = new MyDataModule();
    protected void Page_Load(object sender, EventArgs e)
    {
        //var xmlString = File.ReadAllText("C:\\temp\\myjson.txt");
        //myclass.savedocument(xmlString);
        //myclass.massinfogroup();
        ////loop details
        //myService.CHARACTERISTICS list = new myService.CHARACTERISTICS();

        //list.CHARACTERISTICSMember = new myService.CHARACTERISTIC[1];

        //myService.CHARACTERISTIC item = new myService.CHARACTERISTIC();
        //item.NAME = "NAME1";
        //item.DESCRIPTION = "DESCRIPTION1";
        //item.VALUE = "VALUE1";
        //list.CHARACTERISTICSMember[0] = item;


        //myService.MM72Client client = new myService.MM72Client();

        //var res = client.MM72_OUTBOUND_MATERIAL_CHARACTERISTIC(list);
        //var msg = res.msg;
        //var status = res.status;
        ////++++++++++++++++++++++++++++++++
        ////header
        //ServiceReference.IGRID_OUTBOUND_MODEL myh = new ServiceReference.IGRID_OUTBOUND_MODEL();
        ////myh.OUTBOUND_HEADERS = new ServiceReference.IGRID_OUTBOUND_HEADER_MODEL;

        //ServiceReference.IGRID_OUTBOUND_HEADER_MODEL result = new ServiceReference.IGRID_OUTBOUND_HEADER_MODEL();
        //result.ArtworkNumber = "";
        //result.Date = "";
        //myh.OUTBOUND_HEADERS[0] = result;

        //ServiceReference.IGRID_OUTBOUND_ITEM_MODEL detail = new ServiceReference.IGRID_OUTBOUND_ITEM_MODEL();
        //detail.ArtworkNumber = "";
        //myh.OUTBOUND_ITEMS[0] = detail;


        //ServiceReference.MM73Client h = new ServiceReference.MM73Client();
        //var value = h.MM73_OUTBOUND_MATERIAL_NUMBER(myh);
        //WebReference.ServiceCS test = new WebReference.ServiceCS();
        //WebReference.ArtworkObject _art = new WebReference.ArtworkObject
        //{
        //    ArtworkNumber = "1",
        //    WastePercent = ""
        //};
        //var _inbound = new WebReference.InboundArtwork[1];
        //string _output = test.inputArtworkNumber(_art, _inbound);
        //Response.Write(_output);
        //+++++++++++++++++++++++++++++++
        //myclass.updateArtworkNumber("2MMK202308172646");
        //InboundArtwork();
        //outb();
    }
    public void outb()
    {
		var _table = cs.builditems(@"select * from SapMaterial  where StatusApp=4 and id in  (select MatDoc from TransApprove b where cast(b.SubmitDate as date)='20210330' group by MatDoc)");
        foreach (DataRow dr in _table.Rows){
        myclass.OutboundArtwork(dr["DocumentNo"].ToString());
		}
    }
    public void InboundArtwork()
    {
        ArtworkObject _artwork = new ArtworkObject();
        using (StreamReader sr = new StreamReader(Server.MapPath(@"~/FileTest/h.json")))
        {
            string json = sr.ReadToEnd();
            //string sp = "spInsertMultipleRows";
            List<ArtworkObject> _test = JsonConvert.DeserializeObject<List<ArtworkObject>>(json);
            _artwork = _test[0];
        }
        List<InboundArtwork> _itemsArtwork = new List<InboundArtwork>();
        InboundArtwork item = new InboundArtwork();
        var s="";
        string datapath = "~/FileTest/test.json";
        using (StreamReader sr = new StreamReader(Server.MapPath(datapath)))
        {
            string json = sr.ReadToEnd();
            //string sp = "spInsertMultipleRows";
            List<InboundArtwork> _test = JsonConvert.DeserializeObject<List<InboundArtwork>>(json);
            s = myclass.inputArtworkNumber(_artwork, _test);
        }
        //_itemsArtwork.Add(new InboundArtwork { Characteristic = "ZPKG_SEC_GROUP", Value = "F" });
        //_itemsArtwork.Add(new InboundArtwork { Characteristic = "ZPKG_SEC_BRAND", Value = "198" });
        //_itemsArtwork.Add(new InboundArtwork { Characteristic = "ZPKG_SEC_PRODUCT_CODE", Value = "3AAOSBABJAMN5INNS5" });

        Response.Write(s.ToString());
    }
}
public class obj{
    public string NAME { get; set; }
    public string DESCRIPTION { get; set; }
    public string VALUE { get; set; }
    }