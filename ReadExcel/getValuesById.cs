using Dapper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace ReadExcel
{
    public class getValuesById
    {

        string connectionString = ConfigurationManager.ConnectionStrings["MinhaConexao"].ConnectionString;

        public int GetIdByDescricao(string tableName, string descricaoColumnName, string idColumnName, string descricaoValue)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                descricaoValue = NormalizeString(descricaoValue);

                string query = $"SELECT {idColumnName} FROM {tableName} WHERE UPPER({descricaoColumnName}) LIKE '%' + @Descricao + '%'";

                try
                {
                    int id = connection.QuerySingleOrDefault<int>(query, new { Descricao = descricaoValue });

                    if (id == 0)
                    {
                        Console.WriteLine($"Nenhum ID encontrado para a descrição '{descricaoValue}' na tabela '{tableName}'");
                    }
                    return id;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erro ao executar consulta: {ex.Message}");
                    return 0;
                }
            }
        }

        private string NormalizeString(string value)
        {
            return value.Trim().ToUpper();
        }

    }
}