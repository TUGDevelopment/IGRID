using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using SAP.Middleware.Connector;
using System.Data;
/// <summary>
/// Summary description for SAPCONNECT
/// </summary>
public class SAPConnection
{
    public SAPConnection()
    {
        //
        // TODO: Add constructor logic here
        //
    }
    public RfcDestination SVIPRD()
    {
        return GetRfcDesctination("EDP", "WINSHUTADM", "WST_edp910", "910", "EN", "10.20.1.45", "00", "10", "60");
    }
    public RfcDestination GetRfcDesctination(string name,
    string username, string password, string client,
    string language, string appServerHost, string systemNumber,
    string maxPollSize, string idleTime)
    {
        RfcConfigParameters parameters = new RfcConfigParameters();
        parameters.Add(RfcConfigParameters.Name, name);
        parameters.Add(RfcConfigParameters.User, username);
        parameters.Add(RfcConfigParameters.Password, password);
        parameters.Add(RfcConfigParameters.Client, client);
        parameters.Add(RfcConfigParameters.Language, language);
        parameters.Add(RfcConfigParameters.AppServerHost, appServerHost);
        parameters.Add(RfcConfigParameters.SystemNumber, systemNumber);
        parameters.Add(RfcConfigParameters.PeakConnectionsLimit, maxPollSize);
        parameters.Add(RfcConfigParameters.IdleTimeout, idleTime);
        return RfcDestinationManager.GetDestination(parameters);
    }
    // class สำหรับ convert sap table    เป็น datatable
    public DataTable GetDataTableFromRFCTable(IRfcTable lrfcTable)
    {
        //sapnco_util
        DataTable loTable = new DataTable();
        //... Create ADO.Net table.
        for (int liElement = 0; liElement < lrfcTable.ElementCount; liElement++)
        {
            RfcElementMetadata metadata = lrfcTable.GetElementMetadata(liElement);
            loTable.Columns.Add(metadata.Name);
        }
        //... Transfer rows from lrfcTable to ADO.Net table.
        foreach (IRfcStructure row in lrfcTable)
        {
            DataRow ldr = loTable.NewRow();
            for (int liElement = 0; liElement < lrfcTable.ElementCount; liElement++)
            {
                RfcElementMetadata metadata = lrfcTable.GetElementMetadata(liElement);
                ldr[metadata.Name] = row.GetString(metadata.Name);
            }
            loTable.Rows.Add(ldr);
        }
        return loTable;
    }
    //Class สำหรับ Convert structure เป็น Datatable
    public DataTable ConvertStructure(IRfcStructure myrefcTable)
    {
        DataTable rowTable = new DataTable();
        for (int i = 0; i <= myrefcTable.ElementCount - 1; i++)
        {
            rowTable.Columns.Add(myrefcTable.GetElementMetadata(i).Name);
        }
        DataRow row = rowTable.NewRow();
        for (int j = 0; j <= myrefcTable.ElementCount - 1; j++)
        {
            row[j] = myrefcTable.GetValue(j);
        }
        rowTable.Rows.Add(row);
        return rowTable;
    }
}