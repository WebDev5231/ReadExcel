using OfficeOpenXml;
using ReadExcel.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using Dapper;
using ReadExcel.Data;

namespace ReadExcel
{
    public class dbOperacoes
    {
        internal int insertSolicitacoes(ExcelWorksheet worksheet, string selectedVehicleType)
        {
            var nomeCorMapeamento = new colorMapper();
            var queryGetId = new getValuesById();
            var solicitacaoFactory = new solicitacaoFactory(nomeCorMapeamento, queryGetId);

            using (SqlConnection connection = new SqlConnection(Database.ConnectionString))
            {
                connection.Open();
                using (SqlTransaction transaction = connection.BeginTransaction())
                {
                    int rowCount = worksheet.Dimension.Rows;
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
                    try
                    {
                        for (int row = 2; row <= rowCount; row++)
                        {
                            var chassiValue = worksheet.Cells[row, 2].Text;

                            if (string.IsNullOrWhiteSpace(chassiValue))
                            {
                                continue;
                            }

                            try
                            {
                                var novaSolicitacao = solicitacaoFactory.CriarSolicitacao(worksheet, row, selectedVehicleType);

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
