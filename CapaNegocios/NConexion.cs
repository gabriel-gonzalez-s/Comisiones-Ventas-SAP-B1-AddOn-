using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using FuncionalidadesSDKB1;

namespace ComisionesVentas.CapaNegocios
{
    public class NConexion
    {
        public static SAPbobsCOM.Company oCompany;
        public static string sCodUsuActual;
        public static string sAliasUsuActual;
        public static string sNomUsuActual;
        public static string sCurrentCompanyDB;
        public static SAPbobsCOM.Company Conectar_Aplicacion()
        {
            CapaDatos.DConexion oConexion = new CapaDatos.DConexion();

            oCompany = oConexion.oCompany;

            sCodUsuActual = oCompany.UserSignature.ToString();
            sNomUsuActual = ApplicationExtensions.CurrentUserName();
            sAliasUsuActual = oCompany.UserName;
            sCurrentCompanyDB = oCompany.CompanyDB;

            return oCompany;

        }

        public static SAPbobsCOM.Company Verifica_Conexion(SAPbobsCOM.Company oCurrentCompany)
        {
            if (oCurrentCompany.Connected)
                return oCurrentCompany;
            else
                return Conectar_Aplicacion();
        }
    }
}
