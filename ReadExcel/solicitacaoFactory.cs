using OfficeOpenXml;
using ReadExcel.Models;
using System;
using System.Globalization;
using System.IO;

namespace ReadExcel
{
    public class solicitacaoFactory
    {
        private readonly colorMapper _nomeCorMapeamento;
        private readonly getValuesById _queryGetId;
        private readonly CultureInfo _culture = new CultureInfo("pt-BR");
        private readonly string _logFilePath = @"C:\inetpub\wwwroot\Sistema\readExcel\log.txt";

        public solicitacaoFactory(colorMapper nomeCorMapeamento, getValuesById queryGetId)
        {
            _nomeCorMapeamento = nomeCorMapeamento;
            _queryGetId = queryGetId;
        }

        private void WriteLog(string message)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(_logFilePath, true))
                {
                    sw.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}");
                }
            }
            catch (Exception ex)
            {
                using (StreamWriter sw = new StreamWriter(_logFilePath, true))
                {
                    sw.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {ex}");
                }
            }
        }

        public solicitacoes CriarSolicitacao(ExcelWorksheet worksheet, int row, string selectedVehicleType)
        {
            try
            {
                string GetValueOrThrow(int column, string fieldName)
                {
                    string value = worksheet.Cells[row, column].Text;

                    bool allowNull = false;
                    if (string.IsNullOrWhiteSpace(value) && !allowNull)
                    {
                        throw new Exception($"Erro na linha {row}: '{fieldName}' está vazio.");
                    }
                    return value;
                }

                // Início do processamento
                WriteLog($"Iniciando criação da solicitação para a linha {row}");

                var chassi = GetValueOrThrow(2, "Chassi");
                WriteLog($"Importando chassi: {chassi} - {DateTime.Now: yyyy-MM-dd HH:mm:ss}");

                var nomeCor = worksheet.Cells[row, 12].Text;
                nomeCor = _nomeCorMapeamento.MapperColor(nomeCor);
                var idCor = _queryGetId.GetIdByDescricao("cor", "corDESC", "corCOD", nomeCor);

                var tipoCarroceria = worksheet.Cells[row, 11].Text;
                int idCarroceria;

                if (tipoCarroceria == "ABERTA/CABINE DUPLA")
                {
                    idCarroceria = 134;
                }
                else
                {
                    idCarroceria = _queryGetId.GetIdByDescricao("tipocarroceria", "tcDESC", "tcCOD", tipoCarroceria);
                }

                var tipoVeiculo = worksheet.Cells[row, 9].Text;
                var idTipoVeiculo = _queryGetId.GetIdByDescricao("tipoveiculo", "tvDESC", "tvCOD", tipoVeiculo);

                var especieVeiculo = worksheet.Cells[row, 10].Text;
                var idEspecieVeiculo = _queryGetId.GetIdByDescricao("especie_veiculo", "espvDESC", "espvCOD", especieVeiculo);

                var combustivel = worksheet.Cells[row, 14].Text;
                var idCombustivel = _queryGetId.GetIdByDescricao("combustivel", "cmbCOMBUST", "cmbCOD", combustivel);

                bool isImportado = selectedVehicleType == "IMPORTADO";

                var solicitacoes = new solicitacoes
                {
                    Data = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    ID_Empresa = 4460,
                    Codigo_Atual = "I",
                    Chassi = chassi,
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
                    tipo_solicitacao = 2
                };                

                solicitacoes.cod_receita = GetValueOrThrow(29, "cod_receita");

                string dataDesembaracoString = GetValueOrThrow(30, "data_desembaraco");
                if (DateTime.TryParseExact(dataDesembaracoString, "dd/MM/yyyy", _culture, DateTimeStyles.None, out DateTime dataDesembaraco) ||
                    DateTime.TryParseExact(dataDesembaracoString, "M/d/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dataDesembaraco) ||
                    DateTime.TryParseExact(dataDesembaracoString, "M/d/yyyy", _culture, DateTimeStyles.None, out dataDesembaraco))
                {
                    solicitacoes.data_desembaraco = dataDesembaraco;
                }
                else
                {
                    throw new Exception($"Erro na linha {row}: formato inválido para a data de desembarque: {dataDesembaracoString}");
                }

                solicitacoes.num_di = GetValueOrThrow(31, "num_di");

                if (!isImportado)
                {
                    solicitacoes.cod_receita = "0";
                    solicitacoes.data_desembaraco = DateTime.Now;
                    solicitacoes.num_di = "0";
                }

                WriteLog($"Solicitação criada com sucesso para a linha {row}");

                return solicitacoes;
            }
            catch (Exception ex)
            {
                WriteLog($"Erro ao criar solicitação para a linha {row}: {ex.Message}");
                throw;
            }
        }
    }
}