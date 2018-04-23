using System;
using System.Data;
using System.Data.OleDb;

public partial class _Default : System.Web.UI.Page {
    protected void Page_Init(object sender, EventArgs e) {
        string fileName = Server.MapPath("~/App_Data/Customers.xls");
        gv.DataSource = OpenExcelFile(fileName);
        gv.DataBind();
    }

    protected object OpenExcelFile(string fileName) {
        DataTable dataTable = new DataTable();
        string connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", fileName);
        OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", connectionString);
        adapter.Fill(dataTable);
        return dataTable;
    }
}