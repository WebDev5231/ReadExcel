using OfficeOpenXml;
using ReadExcel.Models;
using System;
using System.Globalization;

namespace ReadExcel
{
    public class solicitacaoFactory
    {
        private readonly colorMapper _nomeCorMapeamento = new colorMapper();
        private readonly getValuesById _queryGetId = new getValuesById();
        private readonly CultureInfo _culture = new CultureInfo("pt-BR");

        public solicitacaoFactory(colorMapper nomeCorMapeamento, getValuesById queryGetId)
        {
            _nomeCorMapeamento = nomeCorMapeamento;
            _queryGetId = queryGetId;
        }

        public solicitacoes CriarSolicitacao(ExcelWorksheet worksheet, int row, bool allowNull = false)
        {
            string GetValueOrThrow(int column, string fieldName)
            {
                string value = worksheet.Cells[row, column].Text;
                if (string.IsNullOrWhiteSpace(value) && !allowNull)
                {
                    throw new Exception($"Erro na linha {row}: '{fieldName}' está vazio.");
                }
                return value;
            }

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

            var solicitacoes = new solicitacoes();

            solicitacoes.Data = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            solicitacoes.ID_Empresa = 4460;
            solicitacoes.Codigo_Atual = "I";
            solicitacoes.Chassi = GetValueOrThrow(2, "Chassi");
            solicitacoes.Faturado = GetValueOrThrow(3, "Faturado");
            solicitacoes.UF_Fat = GetValueOrThrow(4, "UF_Fat");
            solicitacoes.Marca_Modelo = GetValueOrThrow(5, "Marca_Modelo");
            solicitacoes.Qt_Eixos = Int32.TryParse(GetValueOrThrow(6, "Qt_Eixos"), out int qt_Eixos) ? qt_Eixos : 0;
            solicitacoes.Num_Cambio = "0";
            solicitacoes.Num_Motor = GetValueOrThrow(8, "Num_Motor");
            solicitacoes.Tipo_Veiculo = idTipoVeiculo;
            solicitacoes.Esp_Veiculo = idEspecieVeiculo;
            solicitacoes.Tipo_Carroceria = idCarroceria;
            solicitacoes.Cor = idCor;
            solicitacoes.Modelo = GetValueOrThrow(13, "Modelo");
            solicitacoes.Combustivel = idCombustivel;
            solicitacoes.Potencia = Int32.TryParse(GetValueOrThrow(15, "Potencia"), out int potencia) ? potencia : 0;
            solicitacoes.Cilindradas = Int32.TryParse(GetValueOrThrow(16, "Cilindradas"), out int cilindradas) ? cilindradas : 0;
            solicitacoes.Ano_Fab = GetValueOrThrow(17, "Ano_Fab");
            solicitacoes.Ano_Mod = GetValueOrThrow(18, "Ano_Mod");
            solicitacoes.Cap_Passageiros = Int32.TryParse(GetValueOrThrow(19, "Cap_Passageiros"), out int capPassageiros) ? capPassageiros : 0;
            solicitacoes.Cap_Carga = decimal.TryParse(GetValueOrThrow(20, "Cap_Carga").Replace(",", "."), NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out decimal capCarga) ? capCarga : 0.0m;
            solicitacoes.Cmt = decimal.TryParse(GetValueOrThrow(21, "Cmt").Replace(",", "."), NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out decimal cmt) ? cmt : 0.0m;
            solicitacoes.Pbt = decimal.TryParse(GetValueOrThrow(22, "Pbt").Replace(",", "."), NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out decimal pbt) ? pbt : 0.0m;
            solicitacoes.Observacoes = null;
            solicitacoes.Responsavel = GetValueOrThrow(24, "Responsavel");
            solicitacoes.Telefone = GetValueOrThrow(25, "Telefone");
            solicitacoes.Ramal = null;
            solicitacoes.Email = GetValueOrThrow(27, "Email");
            solicitacoes.Impresso = false;
            solicitacoes.Mes_Fab = Int32.TryParse(GetValueOrThrow(28, "Mes_Fab"), out int mes_Fab) ? mes_Fab : 0;
            solicitacoes.usuario_imprime = null;
            solicitacoes.cadastrado = false;
            solicitacoes.tanque_compartimento = null;
            solicitacoes.tipo_solicitacao = 2;
            solicitacoes.cod_receita = GetValueOrThrow(29, "cod_receita");
            solicitacoes.data_desembaraco = DateTime.ParseExact(GetValueOrThrow(30, "data_desembaraco"), "dd/MM/yyyy", _culture);
            solicitacoes.num_di = GetValueOrThrow(31, "num_di");

            return solicitacoes;
        }
    }
}
