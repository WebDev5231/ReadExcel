using Dapper;
using System;
using System.Configuration;
using System.Data.SqlClient;
using OfficeOpenXml;
using System.Globalization;
using System.Collections.Generic;
using ReadExcel.Models;
using ReadExcel.Data;

namespace ReadExcel
{
    public class dbOperacoes
    {
        internal int insertSolicitacoes(ExcelWorksheet worksheet)
        {
            using (SqlConnection connection = new SqlConnection(Database.ConnectionString))
            {
                connection.Open();
                using (SqlTransaction transaction = connection.BeginTransaction())
                {
                    int rowCount = worksheet.Dimension.Rows;
                    int totalFound = 0;
                    int totalInserted = 0;

                    List<(string chassi, string reason, int row)> failedInserts = new List<(string, string, int)>();

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


                    var nomeCorMapeamento = new colorMapper();
                    var queryGetId = new getValuesById();

                    try
                    {
                        for (int row = 2; row <= rowCount; row++)
                        {
                            var chassiValue = worksheet.Cells[row, 2].Text;

                            if (string.IsNullOrWhiteSpace(chassiValue))
                            {
                                continue;
                            }

                            totalFound++;

                            System.Globalization.CultureInfo culture = new System.Globalization.CultureInfo("pt-BR");
                            System.Threading.Thread.CurrentThread.CurrentCulture = culture;

                            var nomeCor = worksheet.Cells[row, 12].Text;
                            nomeCor = nomeCorMapeamento.MapperColor(nomeCor);
                            var idCor = queryGetId.GetIdByDescricao("cor", "corDESC", "corCOD", nomeCor);

                            var tipoCarroceria = worksheet.Cells[row, 11].Text;
                            var idCarroceria = queryGetId.GetIdByDescricao("tipocarroceria", "tcDESC", "tcCOD", tipoCarroceria);

                            var tipoVeiculo = worksheet.Cells[row, 9].Text;
                            var idTipoVeiculo = queryGetId.GetIdByDescricao("tipoveiculo", "tvDESC", "tvCOD", tipoVeiculo);

                            var especieVeiculo = worksheet.Cells[row, 10].Text;
                            var idEspecieVeiculo = queryGetId.GetIdByDescricao("especie_veiculo", "espvDESC", "espvCOD", especieVeiculo);

                            var combustivel = worksheet.Cells[row, 14].Text;
                            var idCombustivel = queryGetId.GetIdByDescricao("combustivel", "cmbCOMBUST", "cmbCOD", combustivel);

                            string GetValueOrThrow(int column, string fieldName)
                            {
                                string value = worksheet.Cells[row, column].Text;
                                if (string.IsNullOrWhiteSpace(value))
                                {
                                    throw new Exception($"Erro na linha {row}: '{fieldName}' está vazio.");
                                }
                                return value;
                            }

                            solicitacoes novaSolicitacao = new solicitacoes
                            {
                                Data = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                                ID_Empresa = 4460,
                                Codigo_Atual = "I",
                                Chassi = chassiValue,
                                Faturado = GetValueOrThrow(3, "Faturado"),
                                UF_Fat = GetValueOrThrow(4, "UF_Fat"),
                                Marca_Modelo = GetValueOrThrow(5, "Marca_Modelo"),
                                Qt_Eixos = Int32.TryParse(GetValueOrThrow(6, "Qt_Eixos"), out int qt_Eixos) ? qt_Eixos : 0,
                                Num_Cambio = "0",
                                Num_Motor = GetValueOrThrow(8, "Num_Motor"),
                                Tipo_Veiculo = idTipoVeiculo,
                                Esp_Veiculo = idEspecieVeiculo,
                                Tipo_Carroceria = idCarroceria,
                                Cor = idCor,
                                Modelo = GetValueOrThrow(13, "Modelo"),
                                Combustivel = idCombustivel,
                                Potencia = Int32.TryParse(GetValueOrThrow(15, "Potencia"), out int potencia) ? potencia : 0,
                                Cilindradas = Int32.TryParse(GetValueOrThrow(16, "Cilindradas"), out int cilindradas) ? cilindradas : 0,
                                Ano_Fab = GetValueOrThrow(17, "Ano_Fab"),
                                Ano_Mod = GetValueOrThrow(18, "Ano_Mod"),
                                Cap_Passageiros = Int32.TryParse(GetValueOrThrow(19, "Cap_Passageiros"), out int capPassageiros) ? capPassageiros : 0,
                                Cap_Carga = decimal.TryParse(GetValueOrThrow(20, "Cap_Carga").Replace(",", "."), NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out decimal capCarga) ? capCarga : 0.0m,
                                Cmt = decimal.TryParse(GetValueOrThrow(21, "Cmt").Replace(",", "."), NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out decimal cmt) ? cmt : 0.0m,
                                Pbt = decimal.TryParse(GetValueOrThrow(22, "Pbt").Replace(",", "."), NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out decimal pbt) ? pbt : 0.0m,
                                Observacoes = null,
                                Responsavel = GetValueOrThrow(24, "Responsavel"),
                                Telefone = GetValueOrThrow(25, "Telefone"),
                                Ramal = null,
                                Email = GetValueOrThrow(27, "Email"),
                                Impresso = false,
                                Mes_Fab = Int32.TryParse(GetValueOrThrow(28, "Mes_Fab"), out int mes_Fab) ? mes_Fab : 0,
                                usuario_imprime = null,
                                cadastrado = false,
                                tanque_compartimento = null,
                                tipo_solicitacao = 2,
                                cod_receita = GetValueOrThrow(29, "cod_receita"),
                                data_desembaraco = DateTime.ParseExact(GetValueOrThrow(30, "data_desembaraco"), "dd/MM/yyyy", culture),
                                num_di = GetValueOrThrow(31, "num_di")
                            };

                            try
                            {
                                int rowsAffected = connection.Execute(insertQuery, novaSolicitacao, transaction: transaction);

                                if (rowsAffected > 0)
                                {
                                    totalInserted++;
                                }
                                else
                                {
                                    failedInserts.Add((chassiValue, "Falha desconhecida na inserção.", row));
                                }
                            }
                            catch (Exception ex)
                            {
                                failedInserts.Add((chassiValue, ex.Message, row));
                                throw new Exception($"Erro ao inserir o registro na linha {row}: {ex.Message}");
                            }
                        }

                        transaction.Commit();
                        return totalInserted;

                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        string errorMessage = $"Erro ao importar dados: {ex.Message}";

                        Console.WriteLine(errorMessage);
                        throw new Exception(errorMessage);
                    }

                }
            }
        }
    }
}
