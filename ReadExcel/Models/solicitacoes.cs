using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ReadExcel.Models
{
    [Table("solicitacoes")]
    public class solicitacoes
    {
        [Key]
        public int ID_Solicitacao { get; set; }
        public string Data { get; set; }
        public int ID_Empresa { get; set; }
        public string Codigo_Atual { get; set; }
        public string Chassi { get; set; }
        public string Faturado { get; set; }
        public string UF_Fat { get; set; }
        public string Marca_Modelo { get; set; }
        public int Qt_Eixos { get; set; }
        public string Num_Cambio { get; set; }
        public string Num_Motor { get; set; }
        public string Num_EixoTraseiro { get; set; }
        public string Num_3Eixo { get; set; }
        public string Num_Carroceria { get; set; }
        public int Tipo_Veiculo { get; set; }
        public int Esp_Veiculo { get; set; }
        public int Tipo_Carroceria { get; set; }
        public int Cor { get; set; }
        public string Modelo { get; set; }
        public int Combustivel { get; set; }
        public int Potencia { get; set; }
        public int Cilindradas { get; set; }
        public string Ano_Fab { get; set; }
        public string Ano_Mod { get; set; }
        public int Cap_Passageiros { get; set; }
        public decimal Cap_Carga { get; set; }
        public decimal Cmt { get; set; }
        public decimal Pbt { get; set; }
        public string Observacoes { get; set; }
        public string Responsavel { get; set; }
        public string Telefone { get; set; }
        public string Ramal { get; set; }
        public string Email { get; set; }
        public bool Impresso { get; set; }
        public string ip { get; set; }
        public string simrav { get; set; }
        public int Mes_Fab { get; set; }
        public string usuario_imprime { get; set; }
        public bool cadastrado { get; set; }
        public string email_alternativo { get; set; }
        public int? tanque_compartimento { get; set; }
        public int tipo_solicitacao { get; set; }
        public string cod_receita { get; set; }
        public DateTime data_desembaraco { get; set; }
        public string num_di { get; set; }
        public string chassi_rev { get; set; }
        public int cap_volumetrica { get; set; }
        public decimal comprimento { get; set; }
        public int quant_paletes { get; set; }
        public decimal cap_carga_litros { get; set; }
        public string modelo_bau_frigo { get; set; }
    }
}
