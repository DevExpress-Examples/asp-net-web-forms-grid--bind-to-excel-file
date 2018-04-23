Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.OleDb

Partial Public Class _Default
	Inherits System.Web.UI.Page
	Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
		Dim fileName As String = Server.MapPath("~/App_Data/Customers.xls")
		gv.DataSource = OpenExcelFile(fileName)
		gv.DataBind()
	End Sub

	Protected Function OpenExcelFile(ByVal fileName As String) As Object
		Dim dataTable As New DataTable()
		Dim connectionString As String = String.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", fileName)
		Dim adapter As New OleDbDataAdapter("SELECT * FROM [Sheet1$]", connectionString)
		adapter.Fill(dataTable)
		Return dataTable
	End Function
End Class