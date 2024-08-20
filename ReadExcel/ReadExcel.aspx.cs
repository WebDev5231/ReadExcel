using System;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Web;
using System.Web.UI;
using OfficeOpenXml;
using Dapper;
using ReadExcel.Models;
using System.Text;
using System.Globalization;

namespace ReadExcel
{
    public partial class ReadExcel : Page
    {
        private string uploadedFilePath;

        protected void Page_Load(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            if (fileUpload.HasFile && (Path.GetExtension(fileUpload.FileName) == ".xls" || Path.GetExtension(fileUpload.FileName) == ".xlsx"))
            {
                string filePath = Server.MapPath("~/Uploads/") + fileUpload.FileName;
                fileUpload.SaveAs(filePath);
                uploadedFilePath = filePath;
                ViewState["UploadedFilePath"] = uploadedFilePath;

                try
                {
                    using (var package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        if (package.Workbook.Worksheets.Count > 0)
                        {
                            var worksheet = package.Workbook.Worksheets[0];

                            var colCount = Math.Min(worksheet.Dimension.Columns, 3001);
                            var rowCount = Math.Min(worksheet.Dimension.Rows, 3001);

                            StringBuilder htmlTable = new StringBuilder();
                            htmlTable.Append("<table class='table table-bordered table-striped table-hover'>");

                            htmlTable.Append("<thead class='table-dark'>");
                            htmlTable.Append("<tr>");

                            for (int col = 1; col <= colCount; col++)
                            {
                                string columnHeader = worksheet.Cells[1, col].Text;
                                htmlTable.Append("<th scope='col'>" + columnHeader + "</th>");
                            }

                            htmlTable.Append("</tr>");
                            htmlTable.Append("</thead>");

                            htmlTable.Append("<tbody>");

                            for (int row = 2; row <= rowCount; row++)
                            {
                                bool isEmptyRow = true;

                                for (int col = 1; col <= colCount; col++)
                                {
                                    if (!string.IsNullOrWhiteSpace(worksheet.Cells[row, col].Text))
                                    {
                                        isEmptyRow = false;
                                        break;
                                    }
                                }

                                if (isEmptyRow)
                                {
                                    continue;
                                }

                                htmlTable.Append("<tr>");
                                for (int col = 1; col <= colCount; col++)
                                {
                                    string cellValue = worksheet.Cells[row, col].Text;
                                    htmlTable.Append("<td>" + cellValue + "</td>");
                                }

                                htmlTable.Append("</tr>");
                            }

                            htmlTable.Append("</tbody>");
                            htmlTable.Append("</table>");

                            Session["FileUploaded"] = true;
                            resultPlaceholder.Controls.Add(new LiteralControl(htmlTable.ToString()));
                        }
                        else
                        {
                            Response.Write("O arquivo Excel não contém nenhuma planilha.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    string errorMessage = ex.Message.Replace("'", "\\'");
                    ScriptManager.RegisterStartupScript(this, GetType(), "MostrarAlerta", $"excelError('{errorMessage}');", true);
                }
            }
            else
            {
                ScriptManager.RegisterStartupScript(this, GetType(), "MostrarAlerta", "excelInvalido();", true);
            }
        }

        protected void btnInsertData_Click(object sender, EventArgs e)
        {
            string filePath = ViewState["UploadedFilePath"] as string;

            if (!string.IsNullOrEmpty(filePath))
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    if (package.Workbook.Worksheets.Count > 0)
                    {
                        var worksheet = package.Workbook.Worksheets[0];

                        try
                        {
                            var insertOperacoes = new dbOperacoes();

                            int insertCount = insertOperacoes.insertSolicitacoes(worksheet);

                            string script = $"sucessoImportacao({insertCount});"; 
                            ScriptManager.RegisterStartupScript(this, GetType(), "SuccessAlert", script, true);
                        }
                        catch (Exception ex)
                        {
                            string errorMessage = ex.Message.Replace("'", "\\'");
                            ScriptManager.RegisterStartupScript(this, GetType(), "MostrarAlerta", $"erroInsert('{errorMessage}');", true);
                        }
                    }
                    else
                    {
                        ScriptManager.RegisterStartupScript(this, GetType(), "MostrarAlerta", "excelInvalido();", true);
                    }
                }
            }
            else
            {
                ScriptManager.RegisterStartupScript(this, GetType(), "MostrarAlerta", "excelInvalido();", true);
            }
        }
    }
}
