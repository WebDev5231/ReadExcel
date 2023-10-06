using System;
using System.Text;
using System.Web.UI;
using System.Web.UI.HtmlControls;

namespace ReadExcel
{
    public partial class ReadRow : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                string rowData = Request.QueryString["data"];
                string[] columnValues = rowData.Split(',');

                StringBuilder detailsTable = new StringBuilder();

                detailsTable.Append("<table class=\"table table-striped table-hover\">"); // Adicionando as classes aqui

                foreach (string columnValue in columnValues)
                {
                    string[] parts = columnValue.Split(':');
                    if (parts.Length == 2)
                    {
                        detailsTable.Append("<tr>");
                        detailsTable.Append("<td><b>" + parts[0] + "</b></td>");
                        detailsTable.Append("<td>" + parts[1] + "</td>");
                        detailsTable.Append("</tr>");
                    }
                }
                detailsTable.Append("</table>");

                detailsCell.InnerHtml = detailsTable.ToString();
            }
        }
    }
}
