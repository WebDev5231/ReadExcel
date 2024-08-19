using ReadExcel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcel
{
    public class colorMapper
    {
        internal string MapperColor(string nomeCor)
        {
            switch (nomeCor.Trim().ToUpper())
            {
                case "AMARELA":
                    return "AMARELA";
                case "AMARELO":
                    return "AMARELA";
                case "AZUL":
                    return "AZUL";
                case "BEGE":
                    return "BEGE";
                case "BRANCO":
                    return "BRANCA";
                case "BRANCA":
                    return "BRANCA";
                case "CINZA":
                    return "CINZA";
                case "DOURADA":
                    return "DOURADA";
                case "GRENA":
                    return "GRENA";
                case "LARANJA":
                    return "LARANJA";
                case "MARRON":
                    return "MARROM";
                case "PRATA":
                    return "PRATA";
                case "PRETO":
                    return "PRETA";
                case "PRETA":
                    return "PRETA";
                case "ROSA":
                    return "ROSA";
                case "ROXA":
                    return "ROXA";
                case "VERDE":
                    return "VERDE";
                case "VERMELHA":
                    return "VERMELHA";
                case "VERMELHO":
                    return "VERMELHA";
                case "FANTASIA":
                    return "FANTASIA";
                default:
                    throw new Exception($"Cor desconhecida: {nomeCor}");
            }
        }
    }
}