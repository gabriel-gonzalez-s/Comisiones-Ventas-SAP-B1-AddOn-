using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComisionesVentas.CapaNegocios
{
    public static class NPeriodo
    {
        public static string NombrePeriodo { get; set; }
        public static DateTime FechaInicio { get; set; }
        public static DateTime FechaFin { get; set; }
        public static SAPbouiCOM.DataTable VendedoresConComisiones { get; set; }
        public static SAPbouiCOM.DataTable VendedoresConPremios { get; set; }
    }
}
