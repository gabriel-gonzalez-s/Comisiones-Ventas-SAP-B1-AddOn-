using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;
using FuncionalidadesSDKB1;
using SAPbobsCOM;

namespace ComisionesVentas.CapaDatos
{
    public class DConexion
    {
        private static SAPbobsCOM.Company _oCompany;

        public DConexion()
        {
            Conectar_Aplicacion();
        }

        public Company oCompany
        {
            get
            {
                return _oCompany;
            }

            set
            {
                _oCompany = value;
            }
        }

        public void Conectar_Aplicacion()
        {

            SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();
            this.oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();

            //SAPbouiCOM.SboGuiApi SboGuiApi = new SAPbouiCOM.SboGuiApi();
            //oCompany = new SAPbobsCOM.Company();
            //SboGuiApi.Connect("0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056");
            //oApplication = SboGuiApi.GetApplication();
            //oCompany = (SAPbobsCOM.Company)oApplication.Company.GetDICompany();

            
        }
    }
}
