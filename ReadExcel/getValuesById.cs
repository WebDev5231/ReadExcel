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
                descricaoValue.Trim().ToUpper();

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
                    //Console.WriteLine($"Erro ao executar consulta: {ex.Message}");
                    //return 0;

                    throw new Exception($"Erro: {ex.Message}");
                }
            }
        }
    }
}