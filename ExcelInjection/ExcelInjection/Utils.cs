using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelInjection
{
    public enum TypeData
    {
        Decimal,
        Varchar,
        DateTime,
        Int
    }
    class Utils
    {
        public static List<string> columnsInt = new List<string>();
        public static List<string> columnsDecimal = new List<string>();
        public static List<string> columnsDate = new List<string>();
        public static List<string> columnsString = new List<string>();



        public static TypeData GetType(string data)
        {
            if (data.ToUpper().Replace(" ", "") == "Periodo".ToUpper().Replace(" ", "")
                                   ||data.ToUpper().Replace(" ", "") == "Regimen emisor".ToUpper().Replace(" ", "")
                                   || data.ToUpper().Replace(" ", "") == "Tipo de cambio".ToUpper().Replace(" ", "")
                                   || data.ToUpper().Replace(" ", "") == "Forma pago".ToUpper().Replace(" ", "") 
                                   || columnsInt.Contains(data.ToUpper().Replace(" ", "")))
            {
                return TypeData.Int;
            }
            else if (data.ToUpper().Replace(" ", "") == "Version".ToUpper().Replace(" ", "")
                || data.ToUpper().Replace(" ", "") == "SubTotal".ToUpper().Replace(" ", "")
                || data.ToUpper().Replace(" ", "") == "Descuento".ToUpper().Replace(" ", "")
                || data.ToUpper().Replace(" ", "") == "IVA Trasladado 0".ToUpper().Replace(" ", "")
                || data.ToUpper().Replace(" ", "") == "IVA Trasladado 16".ToUpper().Replace(" ", "")
                || data.ToUpper().Replace(" ", "") == "IVA Retenido".ToUpper().Replace(" ", "")
                || data.ToUpper().Replace(" ", "") == "ISR Retenido".ToUpper().Replace(" ", "")
                || data.ToUpper().Replace(" ", "") == "IEPS Trasladado".ToUpper().Replace(" ", "")
                || data.ToUpper().Replace(" ", "") == "Local retenido".ToUpper().Replace(" ", "")
                || data.ToUpper().Replace(" ", "") == "Local trasladado".ToUpper().Replace(" ", "")
                || data.ToUpper().Replace(" ", "") == "Total".ToUpper().Replace(" ", "")
                || data.ToUpper().Replace(" ", "") == "IVA Trasladado".ToUpper().Replace(" ", "")
                || columnsDecimal.Contains(data.ToUpper().Replace(" ", "")))
            {
                return TypeData.Decimal;

            }
            else if (data.ToUpper().Replace(" ", "") == "Fecha emision".ToUpper().Replace(" ", "")
                || data.ToUpper().Replace(" ", "") == "Fecha certificacion".ToUpper().Replace(" ", "")
                || data.ToUpper().Replace(" ", "") == "Fecha proceso cancelacion".ToUpper().Replace(" ", "")
                || columnsDate.Contains(data.ToUpper().Replace(" ", "")))
            {

                return TypeData.DateTime;
            }
            else
            {
                return TypeData.Varchar;
            }
        }
    }

    public class NewColumn
    {
        public string name;
        public string type;
    }
}
