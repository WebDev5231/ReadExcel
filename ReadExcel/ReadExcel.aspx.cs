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
using System.Web.UI.WebControls;
using System.Globalization;

namespace ReadExcel
{
    public partial class ReadExcel : System.Web.UI.Page
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


                            ProcessWorksheet(worksheet);


                            int colCount = worksheet.Dimension.Columns;
                            int rowCount = worksheet.Dimension.Rows;

                            StringBuilder htmlTable = new StringBuilder();
                            htmlTable.Append("<table class='table table-bordered table-striped table-hover'>");

                            htmlTable.Append("<thead class='table-dark'>");
                            htmlTable.Append("<tr>");

                            int[] columnWidths = new int[colCount];
                            for (int col = 1; col <= colCount; col++)
                            {
                                columnWidths[col - 1] = 150;
                                string columnHeader = worksheet.Cells[1, col].Text;
                                htmlTable.Append("<th scope='col'>" + columnHeader + "</th>");
                            }

                            htmlTable.Append("<th scope='col'>Ação</th>");
                            htmlTable.Append("</tr>");
                            htmlTable.Append("</thead>");

                            // Tbody (Table Body)
                            htmlTable.Append("<tbody>");
                            for (int row = 2; row <= rowCount; row++)
                            {
                                htmlTable.Append("<tr>");
                                for (int col = 1; col <= colCount; col++)
                                {
                                    string cellValue = worksheet.Cells[row, col].Text;
                                    htmlTable.Append("<td style='min-width: " + columnWidths[col - 1] + "px;'>" + cellValue + "</td>");
                                }
                                string rowData = GetRowDataAsQueryString(worksheet, row, colCount);
                                htmlTable.Append("<td><button type='button' class='btn btn-primary detailsButton' data-rowdata='" + HttpUtility.UrlEncode(rowData) + "' onclick=\"window.open('ReadRow.aspx?data=" + HttpUtility.UrlEncode(rowData) + "','_blank');\">Detalhes</button></td>");
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
                catch
                {
                    ScriptManager.RegisterStartupScript(this, GetType(), "MostrarAlerta", "excelError();", true);
                    return;
                }
            }
            else
            {
                ScriptManager.RegisterStartupScript(this, GetType(), "MostrarAlerta", "excelInvalido();", true);
            }
        }
        protected void btnInsertData_Click(object sender, EventArgs e)
        {
            try
            {
                string filePath = ViewState["UploadedFilePath"] as string;
                if (!string.IsNullOrEmpty(filePath))
                {
                    string connectionString = ConfigurationManager.ConnectionStrings["MinhaConexao"].ConnectionString;

                    using (var package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        if (package.Workbook.Worksheets.Count > 0)
                        {
                            var worksheet = package.Workbook.Worksheets[0];

                            using (SqlConnection connection = new SqlConnection(connectionString))
                            {
                                connection.Open();
                                sqlQuery query = new sqlQuery(connectionString);

                                int rowCount = worksheet.Dimension.Rows;

                                string insertQuery = "INSERT INTO solicitacoes (Data, ID_Empresa, Codigo_Atual, Chassi, Faturado, UF_Fat, " +
                                     "Marca_Modelo, Qt_Eixos, Num_Cambio, Num_Motor, Tipo_Veiculo, Esp_Veiculo, Tipo_Carroceria, " +
                                     "Cor, Modelo, Combustivel, Potencia, Cilindradas, Ano_Fab, Ano_Mod, Cap_Passageiros, " +
                                     "Cap_Carga, Cmt, Pbt, Observacoes, Responsavel, Telefone, Ramal, Email, Impresso, " +
                                     "Mes_Fab, usuario_imprime, cadastrado, tanque_compartimento, tipo_solicitacao, " +
                                     "cod_receita, data_desembaraco, num_di) " +
                                     "VALUES (@Data, @ID_Empresa, @Codigo_Atual, @Chassi, @Faturado, @UF_Fat, @Marca_Modelo, " +
                                     "@Qt_Eixos, @Num_Cambio, @Num_Motor, @Tipo_Veiculo, @Esp_Veiculo, @Tipo_Carroceria, " +
                                     "@Cor, @Modelo, @Combustivel, @Potencia, @Cilindradas, @Ano_Fab, @Ano_Mod, " +
                                     "@Cap_Passageiros, @Cap_Carga, @Cmt, @Pbt, @Observacoes, @Responsavel, @Telefone, " +
                                     "@Ramal, @Email, @Impresso, @Mes_Fab, @usuario_imprime, @cadastrado, @tanque_compartimento, " +
                                     "@tipo_solicitacao, @cod_receita, @data_desembaraco, @num_di)";


                                int insertCount = 0;
                                for (int row = 2; row <= rowCount; row++)
                                {
                                    var chassiValue = worksheet.Cells[row, 2].Text;

                                    if (string.IsNullOrWhiteSpace(chassiValue))
                                    {
                                        continue;
                                    }

                                    System.Globalization.CultureInfo culture = new System.Globalization.CultureInfo("pt-BR");
                                    System.Threading.Thread.CurrentThread.CurrentCulture = culture;

                                    var nomeCor = worksheet.Cells[row, 12].Text;
                                    var idCor = query.GetIdByDescricao("cor", "corDESC", "corCOD", nomeCor);

                                    var tipoCarroceria = worksheet.Cells[row, 11].Text;
                                    var idCarroceria = query.GetIdByDescricao("tipocarroceria", "tcDESC", "tcCOD", tipoCarroceria);

                                    var tipoVeiculo = worksheet.Cells[row, 9].Text;
                                    var idTipoVeiculo = query.GetIdByDescricao("tipoveiculo", "tvDESC", "tvCOD", tipoVeiculo);

                                    var especieVeiculo = worksheet.Cells[row, 10].Text;
                                    var idEspecieVeiculo = query.GetIdByDescricao("especie_veiculo", "espvDESC", "espvCOD", especieVeiculo);

                                    var combustivel = worksheet.Cells[row, 14].Text;
                                    var idCombustivel = query.GetIdByDescricao("combustivel", "cmbCOMBUST", "cmbCOD", combustivel);

                                    solicitacoes novaSolicitacao = new solicitacoes();
                                    novaSolicitacao.Data = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                    novaSolicitacao.ID_Empresa = 4460;
                                    novaSolicitacao.Codigo_Atual = "I";
                                    novaSolicitacao.Chassi = worksheet.Cells[row, 2].Text;
                                    novaSolicitacao.Faturado = worksheet.Cells[row, 3].Text;
                                    novaSolicitacao.UF_Fat = worksheet.Cells[row, 4].Text;
                                    novaSolicitacao.Marca_Modelo = worksheet.Cells[row, 5].Text;
                                    novaSolicitacao.Qt_Eixos = Convert.ToInt32(worksheet.Cells[row, 6].Text);
                                    novaSolicitacao.Num_Cambio = worksheet.Cells[row, 7].Text;
                                    novaSolicitacao.Num_Motor = worksheet.Cells[row, 8].Text;
                                    novaSolicitacao.Tipo_Veiculo = idTipoVeiculo;
                                    novaSolicitacao.Esp_Veiculo = idEspecieVeiculo;
                                    novaSolicitacao.Tipo_Carroceria = idCarroceria;
                                    novaSolicitacao.Cor = idCor;
                                    novaSolicitacao.Modelo = worksheet.Cells[row, 13].Text;
                                    novaSolicitacao.Combustivel = idCombustivel;
                                    novaSolicitacao.Potencia = Convert.ToInt32(worksheet.Cells[row, 15].Text);
                                    novaSolicitacao.Cilindradas = Convert.ToInt32(worksheet.Cells[row, 16].Text);
                                    novaSolicitacao.Ano_Fab = worksheet.Cells[row, 17].Text;
                                    novaSolicitacao.Ano_Mod = worksheet.Cells[row, 18].Text;
                                    novaSolicitacao.Cap_Passageiros = Convert.ToInt32(worksheet.Cells[row, 19].Text);
                                                                      

                                    novaSolicitacao.Cap_Carga = decimal.TryParse(worksheet.Cells[row, 20].Text.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal parsedValueCapCarga) ? parsedValueCapCarga : 0.0m;
                                    novaSolicitacao.Cmt = decimal.TryParse(worksheet.Cells[row, 21].Text.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out var parsedValueCmt) ? parsedValueCmt : 0.0m;
                                    novaSolicitacao.Pbt = decimal.TryParse(worksheet.Cells[row, 22].Text.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal parsedValuePbt) ? parsedValuePbt : 0.0m;


                                    novaSolicitacao.Observacoes = worksheet.Cells[row, 23].Text;
                                    novaSolicitacao.Responsavel = worksheet.Cells[row, 24].Text;
                                    novaSolicitacao.Telefone = worksheet.Cells[row, 25].Text;
                                    novaSolicitacao.Ramal = worksheet.Cells[row, 26].Text;
                                    novaSolicitacao.Email = worksheet.Cells[row, 27].Text;
                                    novaSolicitacao.Impresso = false;
                                    novaSolicitacao.Mes_Fab = Convert.ToInt32(worksheet.Cells[row, 28].Text);
                                    novaSolicitacao.usuario_imprime = null;
                                    novaSolicitacao.cadastrado = false;
                                    novaSolicitacao.tanque_compartimento = null;
                                    novaSolicitacao.tipo_solicitacao = 2;
                                    novaSolicitacao.cod_receita = worksheet.Cells[row, 29].Text;
                                    novaSolicitacao.data_desembaraco = DateTime.ParseExact(worksheet.Cells[row, 30].Text, "dd/MM/yyyy", culture);
                                    novaSolicitacao.num_di = worksheet.Cells[row, 31].Text;

                                    int rowsAffected = connection.Execute(insertQuery, novaSolicitacao);
                                    insertCount++;
                                }

                                string script = "sucessoImportacao(" + insertCount + ");";
                                ClientScript.RegisterStartupScript(this.GetType(), "SuccessAlert", script, true);

                                File.Delete(filePath);
                                connection.Close();
                            }
                        }
                        else
                        {
                            ScriptManager.RegisterStartupScript(this, GetType(), "MostrarAlerta", "excelError();", true);
                        }
                    }
                }
                else
                {
                    ScriptManager.RegisterStartupScript(this, GetType(), "MostrarAlerta", "excelInvalido();", true);

                }
            }
            catch (Exception ex)
            {
                string errorMessage = ex.Message.Replace("'", "\\'");
                ScriptManager.RegisterStartupScript(this, GetType(), "MostrarAlerta", $"erroInsert('{errorMessage}');", true);
            }
        }

        private void ProcessWorksheet(ExcelWorksheet worksheet)
        {
            try
            {
                if (worksheet != null && worksheet.Dimension != null)
                {
                    int colCount = worksheet.Dimension.Columns;
                    int rowCount = worksheet.Dimension.Rows;
                }
            }
            catch (Exception ex)
            {
                Response.Write(ex);
            }
        }

        private string GetRowDataAsQueryString(ExcelWorksheet worksheet, int row, int colCount)
        {
            StringBuilder rowData = new StringBuilder();

            for (int col = 1; col <= colCount; col++)
            {
                if (col > 1)
                {
                    rowData.Append(",");
                }
                string columnName = worksheet.Cells[1, col].Text;
                string cellValue = worksheet.Cells[row, col].Text;
                rowData.Append(columnName + ":" + cellValue);
            }

            return rowData.ToString();
        }
        public class sqlQuery
        {
            private readonly string _connectionString;

            public sqlQuery(string connectionString)
            {
                _connectionString = connectionString;
            }

            public int GetIdByDescricao(string tableName, string descricaoColumnName, string idColumnName, string descricaoValue)
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    string query = $"SELECT {idColumnName} FROM {tableName} WHERE {descricaoColumnName} LIKE '%' + @Descricao + '%'";
                    int id = connection.QuerySingleOrDefault<int>(query, new { Descricao = descricaoValue });

                    connection.Close();

                    return id;
                }
            }

        }
    }
}
