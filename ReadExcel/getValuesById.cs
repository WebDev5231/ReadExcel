using Dapper;
using ReadExcel.Data;
using System;
using System.Data.SqlClient;

namespace ReadExcel
{
    public class getValuesById
    {
        public int GetIdByDescricao(string tableName, string descricaoColumnName, string idColumnName, string descricaoValue)
        {
            using (SqlConnection connection = new SqlConnection(Database.ConnectionString))
            {
                descricaoValue = descricaoValue.Trim().ToUpper();

                string query = $"SELECT {idColumnName} FROM {tableName} WHERE UPPER({descricaoColumnName}) LIKE '%' + @Descricao + '%'";

                try
                {
                    int id = connection.QuerySingleOrDefault<int>(query, new { Descricao = descricaoValue });

                    if (id == 0)
                    {
                        throw new Exception($"Nenhum ID encontrado para a descrição '{descricaoValue}' na tabela '{tableName}'");
                    }

                    return id;
                }
                catch (Exception ex)
                {
                    throw new Exception($"Erro: {ex.Message}");
                }
            }
        }
    }
}