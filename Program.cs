using System;
using System.Collections.Generic;
using SAPbouiCOM.Framework;
using System.Globalization;
using FuncionalidadesSDKB1;

namespace ComisionesVentas
{
    public class Program
    {

        public static NumberFormatInfo oNumberFormatInfo = new NumberFormatInfo();
        public static SAPbouiCOM.Application SBO_Application = null;
        //'Variables de Conexion a las BD DI API
        //'------------------------------------------------------------
        public static SAPbouiCOM.SboGuiApi SboGuiApi = null;
        public static SAPbouiCOM.Application SBO_App = null;
        public static SAPbobsCOM.Company oCompany = null;
        public static SAPbobsCOM.Recordset oRsSUers = null;
        public static SAPbobsCOM.SBObob oSBObob = null;

        public static SAPbouiCOM.Form oForm = null;
        public static SAPbouiCOM.StaticText oStatic = null;
        

        public static string sCodUsuActual   = "";
        public static string sAliasUsuActual   = "";
        public static string sNomUsuActual   = "";
        public static string sCurrentCompanyDB = "";
        public static bool   esPantallaModal = false ;           // Variable para Controlar el Estado Modal de un Form Activo
        public static string IDPantallaModal = "";             // Variable que Contiene el ID de la Pantalla Modal Activa

        public static string sBDComercial = "SBO_COMERCIAL";


        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;
                if (args.Length < 1)
                {
                    oApp = new Application();
                }
                else
                {
                    oApp = new Application(args[0]);
                }
                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();
                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                SBO_Application = Application.SBO_Application;

                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                Application.SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
                Application.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                Application.SBO_Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClickEvent);

                oCompany = CapaNegocios.NConexion.Conectar_Aplicacion();
                    
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    SBO_Application.Menus.RemoveEx("ComisionesVentas");
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    System.Windows.Forms.Application.Exit();
                    break;
                default:
                    break;
            }
        }

        static void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                oForm = Application.SBO_Application.Forms.ActiveForm;

                switch (oForm.TypeEx)
                {
                    case "ComisionesVentas.Parametros":
                        Parametros.Parametros_MenuEvent(ref pVal, out BubbleEvent);
                        break;
                    case "ComisionesVentas.Comisiones":
                        Comisiones.Comisiones_MenuEvent(ref pVal, out BubbleEvent);
                        break;
                }

            }
            catch (Exception )
            {
            }

        }

        static void SBO_Application_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {

            BubbleEvent = true;

            try
            {
                oForm = Application.SBO_Application.Forms.ActiveForm;

                switch (oForm.TypeEx)
                {
                    case "ComisionesVentas.Parametros":
                        Parametros.Parametros_RightClickEvent(ref eventInfo, out BubbleEvent);
                        break;
                    case "ComisionesVentas.Comisiones":
                        Comisiones.Comisiones_RightClickEvent(ref eventInfo, out BubbleEvent);        
                        break;
                }
            }
            catch (Exception)
            {
            }

        }

        static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if ((pVal.FormTypeEx != null) && (pVal.FormTypeEx == "ComisionesVentas.Parametros") 
                   && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK 
                       || pVal.EventType == SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK
                       || pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST))
                {
                    Parametros.Parametros_ItemEvent(FormUID, ref pVal, oCompany, Application.SBO_Application, out BubbleEvent);
                }


                if ((pVal.FormTypeEx != null) && (pVal.FormTypeEx == "ComisionesVentas.Comisiones") 
                    && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED 
                        || pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK 
                        || pVal.EventType == SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK
                        || pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST 
                        || pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE))
                {
                    Comisiones.Comisiones_ItemEvent(FormUID, ref pVal, oCompany, Application.SBO_Application, out BubbleEvent);
                }


            }
            catch (Exception) { }

            //'------------------------------------------------------------------------------------------------------------------------------------------------
            //'  ESTOS EVENTOS MANEJAN LA CONDICION MODAL DE LAS PANTALLAS DONDE ReportType = "Modal"
            //'------------------------------------------------------------------------------------------------------------------------------------------------

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE && pVal.BeforeAction == false)
            {
                try
                {
                    oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

                    if (oForm.ReportType == "Modal")
                    {
                        esPantallaModal = true;
                        IDPantallaModal = pVal.FormUID;
                    }
                }
                catch (Exception)
                {                }
            }

            if (esPantallaModal && FormUID != IDPantallaModal && IDPantallaModal.Trim().Length > 0 &&
               (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE || pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE || pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
            {
                    oForm = Application.SBO_Application.Forms.Item(IDPantallaModal);
                    oForm.Select(); // Selecciona la pantalla modal
                    BubbleEvent = false;
            }

            //' If the modal from is closed...
            if (FormUID == IDPantallaModal && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE && esPantallaModal)
            {
                esPantallaModal = false;
                IDPantallaModal = "";
            }
            //'------------------------------------------------------------------------------------------------------------------------------------------------


            //if (pVal.FormTypeEx == "139" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && pVal.BeforeAction == false)
            //{
            //    SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
            //    string sCFL_ID = oCFLEvento.ChooseFromListUID;
            //    oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            //    SAPbouiCOM.ChooseFromList oCFL = (SAPbouiCOM.ChooseFromList)oForm.ChooseFromLists.Item(sCFL_ID);

            //    SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;
            //    string val = "";
            //    try
            //    {
            //        val = oDataTable.GetValue(0, 0).ToString();
            //        oForm.DataSources.UserDataSources.Item("UDS_CFLPRO").ValueEx = val;
            //    }
            //    catch (Exception)
            //    {
            //    }
            //    //string XX= ((SAPbouiCOM.IChooseFromListEvent)pVal).ChooseFromListUID;
            //}

        }


    }
}
