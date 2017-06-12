using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;
using System.Drawing;
using System.Globalization;

using FuncionalidadesSDKB1;
using ComisionesVentas.CapaNegocios;

namespace ComisionesVentas
{
    [FormAttribute("ComisionesVentas.Comisiones", "CapaPresentacion/Comisiones.b1f")]
    class Comisiones : UserFormBase
    {

        private static SAPbouiCOM.Form oForm = null;
        private static SAPbobsCOM.Company oCompany = null;
        private static SAPbouiCOM.UserDataSource oUSerDataSource;
        private static SAPbouiCOM.ProgressBar oProgBar;
        private static SAPbouiCOM.DataTable oDTTable;
        private static SAPbouiCOM.Grid oGrid;
        private static SAPbouiCOM.ComboBox oComboBox;
        private static SAPbouiCOM.EditText oEdit;
        private static decimal Col_DatAnt;
        private static decimal Col_DatAct;
        private static decimal Col_DatAntPag;
        private static decimal Col_DatActPag;
        private static string sCol_DatAnt;
        private static string sCol_DatAct;
        private static string sCol_DatAntPag;
        private static string sCol_DatActPag;

        public Comisiones()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("Item_1").Specific));
            this.Folder1 = ((SAPbouiCOM.Folder)(this.GetItem("Item_2").Specific));
            this.Folder2 = ((SAPbouiCOM.Folder)(this.GetItem("Item_3").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_4").Specific));
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.DT_SQL = ((SAPbouiCOM.DataTable)(this.UIAPIRawForm.DataSources.DataTables.Item("DT_SQL")));
            this.DT_SQL2 = ((SAPbouiCOM.DataTable)(this.UIAPIRawForm.DataSources.DataTables.Item("DT_SQL2")));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_6").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("Item_8").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_9").Specific));
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_10").Specific));
            this.Grid0.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid0_ClickAfter);
            this.Grid1 = ((SAPbouiCOM.Grid)(this.GetItem("Item_11").Specific));
            this.Grid1.LostFocusAfter += new SAPbouiCOM._IGridEvents_LostFocusAfterEventHandler(this.Grid1_LostFocusAfter);
            this.Grid1.GotFocusAfter += new SAPbouiCOM._IGridEvents_GotFocusAfterEventHandler(this.Grid1_GotFocusAfter);
            this.Grid1.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid1_ClickAfter);
            this.ComboBox1 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_12").Specific));
            this.ComboBox1.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox1_ComboSelectAfter);
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_13").Specific));
            this.Grid2 = ((SAPbouiCOM.Grid)(this.GetItem("Item_14").Specific));
            this.Grid2.LostFocusAfter += new SAPbouiCOM._IGridEvents_LostFocusAfterEventHandler(this.Grid2_LostFocusAfter);
            this.Grid2.GotFocusAfter += new SAPbouiCOM._IGridEvents_GotFocusAfterEventHandler(this.Grid2_GotFocusAfter);
            this.Grid2.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid2_ClickAfter);
            this.ComboBox2 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_15").Specific));
            this.ComboBox2.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox2_ComboSelectAfter);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_16").Specific));
            this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.ComboBox3 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_17").Specific));
            this.ComboBox3.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox3_ComboSelectAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_18").Specific));
            this.Button1.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button1_ClickAfter);
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_19").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_20").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_21").Specific));
            this.Folder3 = ((SAPbouiCOM.Folder)(this.GetItem("Item_24").Specific));
            this.Folder4 = ((SAPbouiCOM.Folder)(this.GetItem("Item_25").Specific));
            this.ComboBox4 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_26").Specific));
            this.ComboBox4.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox4_ComboSelectAfter);
            this.ComboBox5 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_27").Specific));
            this.ComboBox5.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox5_ComboSelectAfter);
            this.Grid3 = ((SAPbouiCOM.Grid)(this.GetItem("Item_28").Specific));
            this.Grid3.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid3_ClickAfter);
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_29").Specific));
            this.Folder5 = ((SAPbouiCOM.Folder)(this.GetItem("Item_30").Specific));
            this.Folder6 = ((SAPbouiCOM.Folder)(this.GetItem("Item_32").Specific));
            this.Folder7 = ((SAPbouiCOM.Folder)(this.GetItem("Item_33").Specific));
            this.Folder8 = ((SAPbouiCOM.Folder)(this.GetItem("Item_34").Specific));
            this.Grid4 = ((SAPbouiCOM.Grid)(this.GetItem("Item_35").Specific));
            this.Grid4.GridSortBefore += new SAPbouiCOM._IGridEvents_GridSortBeforeEventHandler(this.Grid4_GridSortBefore);
            this.Grid4.GridSortAfter += new SAPbouiCOM._IGridEvents_GridSortAfterEventHandler(this.Grid4_GridSortAfter);
            this.Grid4.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid4_ClickAfter);
            this.Grid5 = ((SAPbouiCOM.Grid)(this.GetItem("Item_22").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("Item_23").Specific));
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_36").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_37").Specific));
            this.ComboBox6 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_38").Specific));
            this.Grid6 = ((SAPbouiCOM.Grid)(this.GetItem("Item_39").Specific));
            this.Grid7 = ((SAPbouiCOM.Grid)(this.GetItem("Item_40").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("Item_41").Specific));
            this.StaticText10 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_42").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("Item_43").Specific));
            this.StaticText11 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_44").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("Item_46").Specific));
            this.EditText5.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.EditText5_ChooseFromListBefore);
            this.StaticText12 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_48").Specific));
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("Item_49").Specific));
            this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("Item_50").Specific));
            this.Button4.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button4_ClickAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_45").Specific));
            this.Button2.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button2_ClickAfter);
            this.Grid8 = ((SAPbouiCOM.Grid)(this.GetItem("Item_47").Specific));
            this.StaticText14 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_52").Specific));
            this.StaticText15 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_53").Specific));
            this.CheckBox0 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_54").Specific));
            this.Button5 = ((SAPbouiCOM.Button)(this.GetItem("Item_55").Specific));
            this.Button5.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button5_ClickAfter);
            this.Button6 = ((SAPbouiCOM.Button)(this.GetItem("Item_56").Specific));
            this.Button6.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button6_ClickAfter);
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("Item_51").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
            //  Application.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(this.form_ItemEvent);
            this.ResizeAfter += new ResizeAfterHandler(this.Form_ResizeAfter);

        }
        private void OnCustomInitialize()
        {
            oForm = Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
            oCompany = NConexion.Verifica_Conexion(Program.oCompany);
            
            //Cargar Combo Periodos
            string Sql = "select Code, U_Nombre from [" + Program.sBDComercial.Trim() + "].[dbo].[@Z_COMI_PER] where LTRIM(RTRIM(U_Visible)) = 'Y'";
            ComboBoxExtensions.LoadComboQueryDataTable( ComboBox0, DT_SQL, Sql, 0, 1, false);
            
            ComboBox0.Select(ComboBox0.ValidValues.Count - 1 , SAPbouiCOM.BoSearchKey.psk_Index);

            Folder0.Item.Click();

        }
        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
        }
        public static void Comisiones_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, SAPbobsCOM.Company oCompany,  SAPbouiCOM.Application oApplication, out bool bBubbleEvent)
        {
            bBubbleEvent = true;
            oCompany = Program.oCompany;

            try
            {
                switch (pVal.BeforeAction)
                {
                    case true:  //BeforeAction == true
                        if (pVal.ItemUID == "Item_10" || pVal.ItemUID == "Item_11" || pVal.ItemUID == "Item_14" || pVal.ItemUID == "Item_28" || pVal.ItemUID == "Item_23")
                        {

                            switch (pVal.EventType)
                            {
                                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
                                    try
                                    {
                                        oGrid = (SAPbouiCOM.Grid)oForm.Items.Item(pVal.ItemUID).Specific;
                                        string sCodigo = Convert.ToString(oGrid.DataTable.GetValue(pVal.ColUID, pVal.Row));
                                        
                                        switch (pVal.ColUID)
                                        {
                                            case "Numero PCV":
                                                ApplicationExtensions.LinkedObjectForm("139", "2050", "8", sCodigo);
                                                bBubbleEvent = false;
                                                break;
                                            case "Numero Pago":
                                                ApplicationExtensions.LinkedObjectForm("170", "2817", "3", sCodigo);
                                                bBubbleEvent = false;
                                                break;
                                            case "Numero Documento":
                                                string sTipo = Convert.ToString(oGrid.DataTable.GetValue("Tipo", pVal.Row));
                                                switch (sTipo)
                                                {
                                                    case "FACT":
                                                        ApplicationExtensions.LinkedObjectForm("133", "2053", "8", sCodigo);
                                                        break;
                                                    case "N/CR":
                                                        ApplicationExtensions.LinkedObjectForm("179", "2055", "8", sCodigo);
                                                        break;
                                                }
                                                bBubbleEvent = false;
                                                break;
                                        }
                                    }
                                    catch
                                    {                                    }
                                    break;
                                case SAPbouiCOM.BoEventTypes.et_CLICK:
                                    if (pVal.ItemUID == "Item_14" || pVal.ItemUID == "Item_11")
                                    {
                                        oForm.EnableMenu("1292", true);
                                        oForm.EnableMenu("1293", true);
                                    }
                                    else
                                    {
                                        oForm.EnableMenu("1292", false);
                                        oForm.EnableMenu("1293", false);
                                    }
                                    break;
                                case SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK:
                                    if (pVal.ItemUID == "Item_14" || pVal.ItemUID == "Item_11")
                                    {
                                        oForm.EnableMenu("1292", true);
                                        oForm.EnableMenu("1293", true);
                                    }
                                    else
                                    {
                                        oForm.EnableMenu("1292", false);
                                        oForm.EnableMenu("1293", false);
                                    }
                                    break;
                             
                                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        
                                    
                                    break;


                            }
                        }
                        break;

                    case false: //BeforeAction == false

                        if (pVal.ItemUID == "Item_46")
                            switch (pVal.EventType)
                            {
                                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                                    oDTTable = DataTableExtensions.GetDataTableFromCLF(pVal, oForm);
                                    oEdit = (SAPbouiCOM.EditText)oForm.Items.Item("Item_46").Specific;
                                    oEdit.Value = oDTTable.GetValue(1, 0).ToString();
                                    break;
                            }
                            break;
                }
            }
            catch (Exception)
            {            }

        }
        public static void Parametros_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.ComboBox oComboBox = ( SAPbouiCOM.ComboBox)oForm.Items.Item("Item_15").Specific;
            // Menu Agregar Linea Grid Comisiones
            if (pVal.MenuUID == "1292" && pVal.BeforeAction == true)
            {
                BubbleEvent = false;
                SAPbouiCOM.UserDataSource oUDS = null;

                switch (oForm.PaneLevel)
                {
                    case 2: //GRID PAGOS
                            CapaPresentacion.BuscaPago oBuscaPago = new CapaPresentacion.BuscaPago();

                            oUDS = oBuscaPago.UIAPIRawForm.DataSources.UserDataSources.Item("UD_0");
                            oUDS.ValueEx = oForm.UniqueID;

                            oBuscaPago.Show();
                        break;
                    case 3: // GRID COMISIONES
                        if (oComboBox.Value.Trim().Length > 0)
                        {
                            CapaPresentacion.BuscarPedido activeForm = new CapaPresentacion.BuscarPedido();

                            oUDS = activeForm.UIAPIRawForm.DataSources.UserDataSources.Item("UD_0");
                            oUDS.ValueEx = oForm.UniqueID;

                            oUDS = activeForm.UIAPIRawForm.DataSources.UserDataSources.Item("UD_1");
                            oUDS.ValueEx = oComboBox.Selected.Value;

                            oUDS = activeForm.UIAPIRawForm.DataSources.UserDataSources.Item("UD_2");
                            oUDS.ValueEx = oComboBox.Selected.Description;

                            activeForm.Show();
                        }
                        else
                        {
                            Application.SBO_Application.MessageBox("Debe Estar Seleccionado unicamente un Vendedor");
                        }

                        break;
                }      


            }
            //Menu Eliminar Linea Grid Comisiones
            if (pVal.MenuUID == "1293" && pVal.BeforeAction == true)
            {
                BubbleEvent = false;

                try
                {
                    oForm.Freeze(true);
                    int nRow = 0;
                    SAPbouiCOM.DataTable ooDTTableGRID = null;
                    SAPbouiCOM.StaticText oStatic = null;

                    switch (oForm.PaneLevel)
                    {
                        case 2: //GRID PAGOS
                             oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_11").Specific;

                            nRow = oGrid.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                            oGrid.DataTable.Rows.Remove(nRow);

                            ooDTTableGRID = oForm.DataSources.DataTables.Item("DT_PAG");
                            oDTTable = oForm.DataSources.DataTables.Item("DT_SQL");

                            //Se Borra y vuelve a cargar el DT_COM para que actualice los totales
                            oDTTable.CopyFrom(ooDTTableGRID);
                            ooDTTableGRID.Clear();
                            ooDTTableGRID.CopyFrom(oDTTable);

                            Formatear_Grid_PagosVentas();
                            oStatic = (SAPbouiCOM.StaticText)oForm.Items.Item("Item_20").Specific;
                            oStatic.Caption = "Total Registros: " + oDTTable.Rows.Count.ToString().Trim();
                            GridExtensions.RowNumberGrid(oGrid);

                            break;
                        case 3: // GRID COMISIONES
                            oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_14").Specific;

                            nRow = oGrid.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                            oGrid.DataTable.Rows.Remove(nRow);

                            ooDTTableGRID = oForm.DataSources.DataTables.Item("DT_COM");
                            oDTTable = oForm.DataSources.DataTables.Item("DT_SQL");

                            //Se Borra y vuelve a cargar el DT_COM para que actualice los totales
                            oDTTable.CopyFrom(ooDTTableGRID);
                            ooDTTableGRID.Clear();
                            ooDTTableGRID.CopyFrom(oDTTable);

                            Formatear_Grid_Comisiones();
                             oStatic = (SAPbouiCOM.StaticText)oForm.Items.Item("Item_21").Specific;
                            oStatic.Caption = "Total Registros: " + oDTTable.Rows.Count.ToString().Trim();
                            GridExtensions.RowNumberGrid(oGrid);


                            break;
                    }

                }
                catch (Exception)
                {                }
                finally
                {
                    oForm.Freeze(false);
                }

            }

           
            if (pVal.BeforeAction == true)
            {
                switch (pVal.MenuUID)
                {
                    case "Asignar Porcentaje Comision sin DMP":  // Menu Asignar Porcentaje Comision sin DMP Grid Comisiones
                        try
                        {
                            BubbleEvent = false;
                            oForm.Freeze(true);
                            if (oComboBox.Value.Trim().Length > 0)
                            {
                                NumberFormatInfo nfi = new NumberFormatInfo();
                                nfi.NumberDecimalSeparator = ".";

                                oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_14").Specific;

                                int dtRow = oGrid.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);

                                Calcular_Valor_Comision("Porc Comi", dtRow, 1);

                            }
                            else
                            {
                                Application.SBO_Application.MessageBox("Debe Estar Seleccionado unicamente un Vendedor");
                            }
                        }
                        catch (Exception)
                        {                        }
                        finally
                        {   oForm.Freeze(false); }
                        break;

                    case "Asignar Proyecto a Documento Pagado": // Menu Asignar Proyecto a Documento en GRID PAGOS
                        try
                        {
                            BubbleEvent = false;
                            oForm.Freeze(true);

                            CapaPresentacion.BuscaProyecto activeForm = new CapaPresentacion.BuscaProyecto();

                            SAPbouiCOM.UserDataSource oUDS = activeForm.UIAPIRawForm.DataSources.UserDataSources.Item("UD_0");
                            oUDS.ValueEx = oForm.UniqueID;

                            oUDS = activeForm.UIAPIRawForm.DataSources.UserDataSources.Item("UD_1");
                            oUDS.ValueEx = "Pagos";

                            activeForm.Show();
                        }
                        catch (Exception)
                        {                        }
                        finally
                        { oForm.Freeze(false);   }
                        break;

                    case "MARCAR: Pago Sin PCV Registrada": // Menu AMARCAR: Pago Sin PCV Registrada GRID PAGOS
                        try
                        {
                            BubbleEvent = false;
                            oForm.Freeze(true);

                            oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_11").Specific;

                            int iRow = oGrid.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                            int dtRow = oGrid.GetDataTableRowIndex(iRow);

                            string TipoPag = oGrid.DataTable.GetValue("Tipo Pago", dtRow).ToString();
                            string ComentAct = oGrid.DataTable.GetValue("Comentario",dtRow).ToString();

                            if (TipoPag != "NOPCV")
                            {
                                oGrid.DataTable.SetValue("Comentario", dtRow, "*** Pago Sin PCV Registrada ***  " + ComentAct);
                                oGrid.DataTable.SetValue("Tipo Pago", dtRow, "NOPCV");

                                string dValor = Convert.ToString(Commons.Color_RGB_SAP(117, 255, 75));
                                oGrid.DataTable.SetValue(23, dtRow, dValor); //Valor del Color

                                oGrid.CommonSetting.SetRowBackColor(iRow + 1, Convert.ToInt32(dValor));

                            }
                            else
                            {
                                oGrid.DataTable.SetValue("Comentario", dtRow, ComentAct.Replace("*** Pago Sin PCV Registrada ***",""));
                                oGrid.DataTable.SetValue("Tipo Pago", dtRow, "");
                                ColorRowPagoMarca(oGrid,dtRow,iRow);
                            }

                            Calcular_Totales_Pagos();
                            ((SAPbouiCOM.Button)oForm.Items.Item("Item_45").Specific).Item.Enabled = true;

                        }
                        catch (Exception)
                        {                      }
                        finally
                        { oForm.Freeze(false); }
                        break;

                    case "MARCAR: Pago NO APLICABLE a Comision": // Menu MARCAR: Pago NO APLICABLE a Comision GRID PAGOS
                        try
                        {
                            BubbleEvent = false;
                            oForm.Freeze(true);

                            oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_11").Specific;

                            int iRow = oGrid.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                            int dtRow = oGrid.GetDataTableRowIndex(iRow);

                            string TipoPag = oGrid.DataTable.GetValue("Tipo Pago", dtRow).ToString();
                            string ComentAct = oGrid.DataTable.GetValue("Comentario", dtRow).ToString();

                            if (TipoPag != "NOCOMI")
                            {
                                oGrid.DataTable.SetValue("Comentario", dtRow, "*** Pago NO APLICABLE a Comision ***  " + ComentAct);
                                oGrid.DataTable.SetValue("Tipo Pago", dtRow, "NOCOMI");

                                string dValor = Convert.ToString(Commons.Color_RGB_SAP(240, 128, 128));
                                oGrid.DataTable.SetValue(23, dtRow, dValor); //Valor del Color

                                oGrid.CommonSetting.SetRowBackColor(iRow + 1, Convert.ToInt32(dValor));
                            }
                            else
                            {
                                oGrid.DataTable.SetValue("Comentario", dtRow, ComentAct.Replace("*** Pago NO APLICABLE a Comision ***", ""));
                                oGrid.DataTable.SetValue("Tipo Pago", dtRow, "");
                                ColorRowPagoMarca(oGrid, dtRow, iRow);
                            }

                            Calcular_Totales_Pagos();
                            ((SAPbouiCOM.Button)oForm.Items.Item("Item_45").Specific).Item.Enabled = true;

                        }
                        catch (Exception)
                        {                        }
                        finally
                        { oForm.Freeze(false); }
                        break;

                }
            }
            
        }
        public static void ColorRowPagoMarca(SAPbouiCOM.Grid oGrid,int dtRow, int iRow)
        {
            string NuevaL = oGrid.DataTable.GetValue(22, dtRow).ToString();
            decimal PagAct = Convert.ToDecimal(oGrid.DataTable.GetValue(14, dtRow));
            decimal PagIni = Convert.ToDecimal(oGrid.DataTable.GetValue(24, dtRow));

            //Cambiar Color de Linea
            if ((PagAct == PagIni) && NuevaL.Trim() != "Y")
            {
                int Color = Commons.Color_RGB_SAP(255, 255, 255);
                oGrid.DataTable.SetValue(23, dtRow, "0");
                GridExtensions.ColorBackColorRowGrid(oGrid, iRow + 1, Color);
                for (int iCols = 9; iCols <= 17; iCols++)
                {
                    if (iCols != 14)
                        oGrid.CommonSetting.SetCellBackColor(iRow + 1, iCols + 1, Commons.Color_RGB_SAP(250, 250, 210));
                }
                oGrid.CommonSetting.SetCellBackColor(iRow + 1, 25 + 1, Commons.Color_RGB_SAP(250, 250, 210));
            }
            else if (NuevaL.Trim() != "Y")
            {
                int Color = Commons.Color_RGB_SAP(250, 235, 215);
                oGrid.DataTable.SetValue(23, dtRow, Color);
                GridExtensions.ColorBackColorRowGrid(oGrid, iRow + 1, Color);
            }
            else if (NuevaL.Trim() == "Y")
            {
                int Color = Commons.Color_RGB_SAP(250, 250, 210);
                oGrid.DataTable.SetValue(23, dtRow, Color);
                GridExtensions.ColorBackColorRowGrid(oGrid, iRow + 1, Color);
            }
        }
        public static void Parametros_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (eventInfo.ItemUID == "Item_14")
            {
                oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_14").Specific;
                oGrid.Rows.SelectedRows.Add(eventInfo.Row);
            } 

            if (eventInfo.ItemUID == "Item_14" && eventInfo.ColUID == "Porc Comi") //GRID COMISIONES
            {
                try
                {
                    if (eventInfo.BeforeAction == true)
                            MenuExtensions.CreateContextualMenu("Asignar Porcentaje Comision sin DMP", "Asignar Porcentaje Comision sin DMP",-1);
                    else
                            Application.SBO_Application.Menus.RemoveEx("Asignar Porcentaje Comision sin DMP");
                }
                catch (Exception)
                {                }

            }

            if (eventInfo.ItemUID == "Item_11") //GRID PAGOS
            {
                try
                {
                    if (eventInfo.BeforeAction == true)
                    {
                        MenuExtensions.CreateContextualMenu("Asignar Proyecto a Documento Pagado", "Asignar Proyecto a Documento Pagado",-1);
                        MenuExtensions.CreateContextualMenu("MARCAR: Pago Sin PCV Registrada", "MARCAR: Pago Sin PCV Registrada",-1);
                        MenuExtensions.CreateContextualMenu("MARCAR: Pago NO APLICABLE a Comision", "MARCAR: Pago NO APLICABLE a Comision",-1);
                    }
                    else
                    {
                            Application.SBO_Application.Menus.RemoveEx("Asignar Proyecto a Documento Pagado");
                            Application.SBO_Application.Menus.RemoveEx("MARCAR: Pago Sin PCV Registrada");
                            Application.SBO_Application.Menus.RemoveEx("MARCAR: Pago NO APLICABLE a Comision");
                    }
                }
                catch (Exception)
                {                }
            }
        }
        private void Form_ResizeAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.Item oItem = oForm.Items.Item("Item_0");
            oItem.Top = 25;
            oItem.Left = 16;
            oItem.Height = oForm.Height - 95;
            oItem.Width = oForm.Width - 39;

            oItem = oForm.Items.Item("Item_31");
            oItem.Top = 54;
            oItem.Left = 33;
            oItem.Height = oForm.Height - 145;
            oItem.Width = oForm.Width - 69;

            SAPbouiCOM.Grid oGridPag = (SAPbouiCOM.Grid)oForm.Items.Item("Item_11").Specific;
            oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_47").Specific;
            oGrid.Item.Top = oGridPag.Item.Top + oGridPag.Item.Height + 1;//oForm.Height - 118;
            oGrid.Item.Left = oForm.Width - 863;
            oGrid.Item.Height = 37;
            oGrid.Item.Width = 822;

            for (int i = 0; i < oGrid.Columns.Count; i++)
            {
                oGrid.Columns.Item(i).Width = 150;
            }

            SAPbouiCOM.RowHeaders oHeader = oGrid.RowHeaders;
            oHeader.Width = 40;

        }
        private void ComboBox0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                bool bProgBar = true;

                //********************* PROGRESS BAR
                ProgressBarExtensions.Create_ProgressBar(ref oProgBar, "Cargando / Calculando Datos de Ventas PCV", 6);
                
                oForm.Freeze(true);

                Cargar_Datos_Periodo(ComboBox0.Selected.Description);

                DateTime dDesde = Convert.ToDateTime(oForm.DataSources.UserDataSources.Item("UD_DESDE").Value);
                DateTime dHasta = Convert.ToDateTime(oForm.DataSources.UserDataSources.Item("UD_HASTA").Value);

                int PaneAct = oForm.PaneLevel;
                //Cargar Combo Vendedores Ventas
                oForm.PaneLevel = 1;
                ComboBoxExtensions.CleanComboBox(ComboBox1);
                Cargar_Combo_Vendedores_Periodo(ComboBox1, dDesde, dHasta);
                ComboBox1.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                bProgBar = ProgressBarExtensions.Increment_ProgressBar(ref oProgBar, 1);
                bProgBar = ProgressBarExtensions.ChangeText_ProgressBar(ref oProgBar, "Cargando / Calculando Pagos del Periodo");
                //Cargar Combo Vendedores Pagos
                oForm.PaneLevel = 2;
                ComboBoxExtensions.CleanComboBox(ComboBox3);
                Cargar_Combo_Vendedores_Pagos_Periodo(ComboBox3, dDesde, dHasta);
                ComboBox3.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.PaneLevel = PaneAct;

                bProgBar = ProgressBarExtensions.Increment_ProgressBar(ref oProgBar, 1);
                bProgBar = ProgressBarExtensions.ChangeText_ProgressBar(ref oProgBar, "Cargando / Calculando Comisiones del Periodo");
                //Cargar Combo Vendedores Comisiones
                oForm.PaneLevel = 3;
                ComboBoxExtensions.CleanComboBox(ComboBox2);
                Cargar_Combo_Vendedores_Periodo(ComboBox2, dDesde, dHasta);
                ComboBox2.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.PaneLevel = PaneAct;

                bProgBar = ProgressBarExtensions.Increment_ProgressBar(ref oProgBar, 1);
                bProgBar = ProgressBarExtensions.ChangeText_ProgressBar(ref oProgBar, "Cargando / Calculando Modificaciones en PCV del Periodo");
                //Cargar Combo Vendedores Modificaciones
                oForm.PaneLevel = 4;
                ComboBoxExtensions.CleanComboBox(ComboBox4);
                Cargar_Combo_Vendedores_Periodo(ComboBox4, dDesde, dHasta);
                if (!CheckBox0.Checked)
                    ComboBox4.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                else
                    Grid3.DataTable.Clear();
                oForm.PaneLevel = PaneAct;
                
                bProgBar = ProgressBarExtensions.Increment_ProgressBar(ref oProgBar, 1);
                bProgBar = ProgressBarExtensions.ChangeText_ProgressBar(ref oProgBar, "Cargando / Calculando Pagos de Premios del Periodo");
                //Cargar Combo Vendedores Pagos
                oForm.PaneLevel = 5;
                ComboBoxExtensions.CleanComboBox(ComboBox5);
                Cargar_Combo_Vendedores_Pagos_Periodo(ComboBox5, dDesde, dHasta);
                ComboBox5.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.PaneLevel = PaneAct;

                bProgBar = ProgressBarExtensions.Increment_ProgressBar(ref oProgBar, 1);
                bProgBar = ProgressBarExtensions.Close_ProgressBar(ref oProgBar);


            }
            catch (Exception)
            {            }
            finally
            {
                oForm.Freeze(false);
            }

        }
        private void ComboBox1_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            DateTime dDesde = Convert.ToDateTime(oForm.DataSources.UserDataSources.Item("UD_DESDE").Value);
            DateTime dHasta = Convert.ToDateTime(oForm.DataSources.UserDataSources.Item("UD_HASTA").Value);
            string cVend = ComboBox1.Selected.Value;

            Cargar_Grid_VentasPCV(dDesde, dHasta, cVend);
        }
        private void ComboBox2_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            string cVend = ComboBox2.Selected.Value;

            Cargar_Grid_Comisiones(cVend, ComboBox0.Selected.Description);

        }
        private void ComboBox3_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            DateTime dDesde = Convert.ToDateTime(oForm.DataSources.UserDataSources.Item("UD_DESDE").Value);
            DateTime dHasta = Convert.ToDateTime(oForm.DataSources.UserDataSources.Item("UD_HASTA").Value);
            string cVend = ComboBox3.Selected.Value;

            Cargar_Grid_PagosVentas(dDesde, dHasta, ComboBox0.Selected.Description, cVend);

        }
        private void ComboBox4_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            DateTime dDesde = Convert.ToDateTime(oForm.DataSources.UserDataSources.Item("UD_DESDE").Value);
            DateTime dHasta = Convert.ToDateTime(oForm.DataSources.UserDataSources.Item("UD_HASTA").Value);
            string cVend = ComboBox4.Selected.Value;

            Cargar_Grid_Modificaciones(dDesde, dHasta, cVend);

        }
        private void ComboBox5_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            string cVend = ComboBox5.Selected.Value;

            Cargar_Grid_Pagos_Premios(cVend, ComboBox0.Selected.Description);

        }
        private void Grid0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                Grid0.Rows.SelectedRows.Add(pVal.Row);
            }
            catch (Exception)
            {            }

        }
        private void Grid1_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                Grid1.Rows.SelectedRows.Add(pVal.Row);
            }
            catch (Exception)
            {            }
        }
        private void Grid2_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                Grid2.Rows.SelectedRows.Add(pVal.Row);
            }
            catch (Exception)
            {            }
        }
        private void Grid3_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                Grid3.Rows.SelectedRows.Add(pVal.Row);
            }
            catch (Exception)
            {            }
        }
        private void Grid4_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                Grid4.Rows.SelectedRows.Add(pVal.Row);
            }
            catch (Exception)
            { }

        }
        private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            switch (oForm.PaneLevel)
            {
                case 1:
                    DataTableExtensions.DatatableSAP_a_Excel(oForm.DataSources.DataTables.Item("DT_VTAS"));
                    break;
                case 2:
                    DataTableExtensions.DatatableSAP_a_Excel(oForm.DataSources.DataTables.Item("DT_PAG"));
                    break;
                case 3:
                    DataTableExtensions.DatatableSAP_a_Excel(oForm.DataSources.DataTables.Item("DT_COM"));
                    break;
            }

        }
        private void Button4_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            CapaPresentacion.BuscaProyecto activeForm = new CapaPresentacion.BuscaProyecto();

            SAPbouiCOM.UserDataSource oUDS = activeForm.UIAPIRawForm.DataSources.UserDataSources.Item("UD_0");
            oUDS.ValueEx = oForm.UniqueID;

            oUDS = activeForm.UIAPIRawForm.DataSources.UserDataSources.Item("UD_1");
            oUDS.ValueEx = "Consultas";

            activeForm.Show();

        }
        private void Button1_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if(Button1.Item.Enabled)
            {
                string sql = "";
                oDTTable = oForm.DataSources.DataTables.Item("DT_COM");

                if (Button1.Caption.Trim() == "Registrar Comisiones")
                {
                    if (ComboBox2.Selected.Value != "0") //Si esta seleccionado un Vendedor, se registran todos y luego se actualiza el vendedor
                    {
                        sql = "EXEC  [dbo].[Min_Comisiones_Calculo_Periodo] 0,'" + ComboBox0.Selected.Description + "'";
                        DT_SQL2.ExecuteQuery(sql);

                        Insertar_Comisiones(DT_SQL2);

                        Eliminar_Comisiones(ComboBox2.Selected.Value, ComboBox0.Selected.Description);

                        Insertar_Comisiones(oDTTable);

                        Button1.Caption = "Actualizar Comisiones";
                        Button1.Item.Enabled = false;

                        Application.SBO_Application.StatusBar.SetText("Registro de Comisiones Exitoso", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                    else
                    {
                        Insertar_Comisiones(oDTTable);
                    }
                }

                if (Button1.Caption.Trim() == "Actualizar Comisiones")
                {
                    Insertar_Comisiones_Historico(oForm.DataSources.DataTables.Item("DT_COMINI"));

                    Eliminar_Comisiones(ComboBox2.Selected.Value, ComboBox0.Selected.Description);

                    Insertar_Comisiones(oDTTable);

                    Button1.Item.Enabled = false;

                    Application.SBO_Application.StatusBar.SetText("Actualizacion de Comisiones Exitoso", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
            }
        }
        private void Button2_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (Button2.Item.Enabled)
                {
                    string sql = "";
                    oDTTable = oForm.DataSources.DataTables.Item("DT_PAG");

                    switch (Button2.Caption.Trim())
                    {
                        case "Registrar Pagos":
                            if (ComboBox3.Selected.Value.Trim() != "") //Si esta seleccionado un Vendedor, se registran todos y luego se actualiza el vendedor
                            {
                                sql = "EXEC  [dbo].[Min_Comisiones_Consultar_Pagos_Venta_Periodo] 0,'" + ComboBox0.Selected.Description + "'";
                                DT_SQL2.ExecuteQuery(sql);

                                Insertar_Pagos(DT_SQL2);

                                Eliminar_Pagos(ComboBox3.Selected.Value, ComboBox0.Selected.Description);

                                Insertar_Pagos(oDTTable);
                            }
                            else
                            {
                                Insertar_Pagos(oDTTable);
                            }

                            Button2.Caption = "Actualizar Pagos";
                            Button2.Item.Enabled = false;

                            Application.SBO_Application.StatusBar.SetText("Registro de Pagos Exitoso", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                            break;
                        case "Actualizar Pagos":
                            //Insertar_Comisiones_Historico(oForm.DataSources.DataTables.Item("DT_PAGOINI"));

                            Eliminar_Pagos(ComboBox3.Selected.Value, ComboBox0.Selected.Description);

                            Insertar_Pagos(oDTTable);

                            Button2.Item.Enabled = false;

                            Application.SBO_Application.StatusBar.SetText("Actualizacion de Pagos Exitoso", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                            break;
                    }
                }
            }
            catch (Exception)
            {            }
        }
        private void Button5_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            DateTime dDesde = Convert.ToDateTime(oForm.DataSources.UserDataSources.Item("UD_DESDE").Value);
            DateTime dHasta = Convert.ToDateTime(oForm.DataSources.UserDataSources.Item("UD_HASTA").Value);
            string cVend = ComboBox4.Selected.Value;

            ProgressBarExtensions.Create_ProgressBar(ref oProgBar, "Cargando / Calculando Modificaciones en PCV del Periodo", 2);
            Cargar_Grid_Modificaciones(dDesde, dHasta, cVend);
            ProgressBarExtensions.Increment_ProgressBar(ref oProgBar, 1);
            ProgressBarExtensions.Close_ProgressBar(ref oProgBar);

        }
        private void Button6_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (Button6.Item.Enabled)
                {
                    string sql = "";
                    oDTTable = oForm.DataSources.DataTables.Item("DT_PREMI");

                    switch (Button6.Caption.Trim())
                    {
                        case "Registrar Premios":
                            if (ComboBox5.Selected.Value.Trim() != "") //Si esta seleccionado un Vendedor, se registran todos y luego se actualiza el vendedor seleccionado
                            {
                                sql = "EXEC  [dbo].[Min_Comisiones_Pagos_Premios_Periodo] 0 ";
                                DT_SQL2.ExecuteQuery(sql);

                                Insertar_Pagos_Premios(DT_SQL2);

                                Eliminar_Pagos_Premios(ComboBox5.Selected.Value, ComboBox0.Selected.Description);

                                Insertar_Pagos_Premios(oDTTable);
                            }
                            else
                            {
                                Insertar_Pagos_Premios(oDTTable);
                            }

                            Button6.Caption = "Actualizar Premios";
                            Button6.Item.Enabled = false;

                            Application.SBO_Application.StatusBar.SetText("Registro de Pagos Premios Exitoso", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                            break;
                        case "Actualizar Premios":
                            //Insertar_Comisiones_Historico(oForm.DataSources.DataTables.Item("DT_PAGOINI"));

                            Eliminar_Pagos_Premios(ComboBox5.Selected.Value, ComboBox0.Selected.Description);

                            Insertar_Pagos_Premios(oDTTable);

                            Button6.Item.Enabled = false;

                            Application.SBO_Application.StatusBar.SetText("Actualizacion de Pagos Premios Exitoso", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                            break;
                    }
                }
            }
            catch (Exception)
            { }

        }
        private void Grid1_GotFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                int dtRow = Grid2.GetDataTableRowIndex(pVal.Row);

                switch (pVal.ColUID)
                {
                    case "Pago Total":
                        Col_DatAntPag = Math.Round(Convert.ToDecimal(Grid1.DataTable.GetValue(pVal.ColUID, dtRow)), 2);
                        break;
                    case "Comentario":
                        sCol_DatAntPag = Grid1.DataTable.GetValue(pVal.ColUID, dtRow).ToString();
                        break;
                }
            }
            catch (Exception)
            {            }

        }
        private void Grid1_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                oForm.Freeze(true);

                int dtRow = Grid1.GetDataTableRowIndex(pVal.Row);
                switch (pVal.ColUID)
                {
                    case "Pago Total":
                        Col_DatActPag = Math.Round(Convert.ToDecimal(Grid1.DataTable.GetValue(pVal.ColUID, dtRow)), 2);
                        break;
                    case "Comentario":
                        sCol_DatActPag = Grid1.DataTable.GetValue(pVal.ColUID, dtRow).ToString();
                        break;
                }

                if (Col_DatAntPag != Col_DatActPag)
                {
                    Button2.Item.Enabled = true;

                    string NuevaL = Grid1.DataTable.GetValue(22, dtRow).ToString();

                    decimal PagAct = Convert.ToDecimal(Grid1.DataTable.GetValue(14, dtRow));
                    decimal PagIni = Convert.ToDecimal(Grid1.DataTable.GetValue(24, dtRow));

                    //Cambiar Color de Linea
                    if ((PagAct == PagIni) && NuevaL.Trim() != "Y")
                    {
                        int Color =  Commons.Color_RGB_SAP(255, 255, 255);
                        Grid1.DataTable.SetValue(23, dtRow, "0");
                        GridExtensions.ColorBackColorRowGrid(Grid1, dtRow + 1, Color);
                        for (int iCols = 9; iCols <= 17; iCols++)
                        {
                            if (iCols != 14)
                                Grid1.CommonSetting.SetCellBackColor(dtRow + 1, iCols + 1, Commons.Color_RGB_SAP(250, 250, 210));
                        }
                    }
                    else if (NuevaL.Trim() != "Y")
                    {
                        int Color = Commons.Color_RGB_SAP(250, 235, 215);//(255, 160, 122);
                        Grid1.DataTable.SetValue(23, dtRow, Color);
                        GridExtensions.ColorBackColorRowGrid(Grid1, dtRow + 1, Color);
                    }

                    //Cambiar Color de Fuentes en Datos Modificados (10, 11, 13, 14)
                    if (PagAct != PagIni)
                    {
                        int ColorF = Commons.Color_RGB_SAP(255, 0, 0);
                        int ColorB = Commons.Color_RGB_SAP(250, 250, 210);
                        GridExtensions.ColorFontCellGrid(Grid1, 14 + 1, dtRow + 1, ColorF, ColorB);
                    }
                    else
                    {
                        int ColorF = Commons.Color_RGB_SAP(0, 0, 0);
                        int ColorB = NuevaL.Trim() != "Y" ? Commons.Color_RGB_SAP(240, 255, 255) : Convert.ToInt32(Grid1.DataTable.GetValue(20, dtRow));
                        GridExtensions.ColorFontCellGrid(Grid1, 14 + 1, dtRow + 1, ColorF, ColorB);
                    }

                    Calcular_Totales_Pagos();
                }

                if (sCol_DatAntPag != sCol_DatActPag)
                    Button2.Item.Enabled = true;

            }
            catch (Exception)
            {            }
            finally
            {  oForm.Freeze(false);            }

        }
        private void Grid2_GotFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                int dtRow = Grid2.GetDataTableRowIndex(pVal.Row);

                if (pVal.ColUID == "Comentario")
                { sCol_DatAnt = Convert.ToString(Grid2.DataTable.GetValue(pVal.ColUID, dtRow)); }
                else
                { Col_DatAnt = Math.Round(Convert.ToDecimal(Grid2.DataTable.GetValue(pVal.ColUID, dtRow)), 2); }
            }
            catch (Exception)
            {            }
        }
        private void Grid2_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                oForm.Freeze(true);

                NumberFormatInfo nfi = new NumberFormatInfo();
                nfi.NumberDecimalSeparator = ".";

                int dtRow = Grid2.GetDataTableRowIndex(pVal.Row);
                if (pVal.ColUID == "Comentario")
                    sCol_DatAct = Convert.ToString(Grid2.DataTable.GetValue(pVal.ColUID, dtRow));
                else
                    Col_DatAct = Math.Round(Convert.ToDecimal(Grid2.DataTable.GetValue(pVal.ColUID, dtRow)), 2);

                if (Col_DatAnt != Col_DatAct || sCol_DatAnt != sCol_DatAct)
                {
                    Calcular_Valor_Comision(pVal.ColUID, dtRow , 0);

                    Button1.Item.Enabled = true;
                }


            }
            catch (Exception)
            {            }
            finally
            {  oForm.Freeze(false);            }

        }
        private void Grid4_GridSortAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                // oForm.Freeze(true);
                GridExtensions.RowNumberGrid(Grid4);
                int iPaneAct = oForm.PaneLevel;
                oForm.PaneLevel = iPaneAct - 1;
                oForm.PaneLevel = iPaneAct;
            }
            catch (Exception) { }
            finally
            { oForm.Freeze(false); }
        }
        private void Grid4_GridSortBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            oForm.Freeze(true);

        }
        private void EditText5_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

        }


        ///-------------------------------------------------------------------------------------------------------------------------------------------------------------
        ///-------------------------------------------------------------------------------------------------------------------------------------------------------------
        ///  PROCEDIMIENTOS Y FUNCIONES
        ///-------------------------------------------------------------------------------------------------------------------------------------------------------------
        ///-------------------------------------------------------------------------------------------------------------------------------------------------------------

        private void Cargar_Datos_Periodo(string Periodo)
        {

            string Sql = @" select DATEADD(MONTH,-1,DATEFROMPARTS(LEFT([U_Nombre],4),RIGHT([U_Nombre],2),U_Inicio)) as Inicio
                                  ,DATEFROMPARTS(LEFT([U_Nombre],4),RIGHT([U_Nombre],2),U_Fin) as Fin
                                  ,U_Valor_UF
								  ,U_Valor_DMP
                            from [" + Program.sBDComercial.Trim() + "].[dbo].[@Z_COMI_PER] where U_Nombre = '" + Periodo  + "'";

            try
            {
                DT_SQL.ExecuteQuery(Sql);

                oUSerDataSource = oForm.DataSources.UserDataSources.Item("UD_DESDE");
                oUSerDataSource.Value = Convert.ToDateTime(DT_SQL.GetValue("Inicio", 0)).ToString("yyyyMMdd");
                oUSerDataSource = oForm.DataSources.UserDataSources.Item("UD_HASTA");
                oUSerDataSource.Value = Convert.ToDateTime(DT_SQL.GetValue("Fin", 0)).ToString("yyyyMMdd");
                oUSerDataSource = oForm.DataSources.UserDataSources.Item("UD_UF");
                oUSerDataSource.Value = DT_SQL.GetValue("U_Valor_UF", 0).ToString();
                oUSerDataSource = oForm.DataSources.UserDataSources.Item("UD_DMP");
                oUSerDataSource.Value = DT_SQL.GetValue("U_Valor_DMP", 0).ToString();
            }
            catch 
            {            }
        }
        private void Cargar_Combo_Vendedores_Periodo(SAPbouiCOM.ComboBox oCombo, DateTime fDesde, DateTime fHasta)
        {
            string sql = @"select DISTINCT V.SlpCode , '(' + LTRIM(RTRIM(V.Memo)) + ') - ' + LTRIM(RTRIM(V.SlpName))   
                            from OSLP V join ORDR P ON V.SlpCode = P.SlpCode
                            where V.Active  = 'Y' AND V.SlpCode != -1 AND series = 27
                            AND P.CreateDate BETWEEN  '" + fDesde.ToString("MM-dd-yyyy") + "' AND '" + fHasta.ToString("MM-dd-yyyy") + "' ";

            ComboBoxExtensions.LoadComboQueryDataTable(oCombo, DT_SQL, sql, 0, 1, true);
        }
        private void Cargar_Combo_Vendedores_Pagos_Periodo(SAPbouiCOM.ComboBox oCombo, DateTime fDesde, DateTime fHasta)
        {
            string sql = @"select DISTINCT V.SlpCode , '(' + LTRIM(RTRIM(V.Memo)) + ') - ' + LTRIM(RTRIM(V.SlpName))   
                            from OSLP V join ORDR P ON V.SlpCode = P.SlpCode
                            where V.Active  = 'Y' --AND V.SlpCode != -1 AND series = 27
                            --AND P.CreateDate BETWEEN  '" + fDesde.ToString("MM-dd-yyyy") + "' AND '" + fHasta.ToString("MM-dd-yyyy") + "' ";

            ComboBoxExtensions.LoadComboQueryDataTable(oCombo, DT_SQL, sql, 0, 1, true);
        }
        private void Cargar_Grid_VentasPCV(DateTime fDesde, DateTime fHasta, string sVend)
        {
            try
            {
                oForm.Freeze(true);
                string sql = "EXEC  [dbo].[Min_Comisiones_PCV_Ventas_Periodo] '" + fDesde.ToString("MM-dd-yyyy") + "','" + fHasta.ToString("MM-dd-yyyy") + "','" + sVend + "'";
                oDTTable = oForm.DataSources.DataTables.Item("DT_VTAS");
                oDTTable.ExecuteQuery(sql);

                StaticText4.Caption = "Total Registros: " + oDTTable.Rows.Count.ToString().Trim();

                SAPbouiCOM.EditTextColumn oEditCol = (SAPbouiCOM.EditTextColumn)Grid0.Columns.Item(0);
                oEditCol.LinkedObjectType = "17";

                Grid0.Columns.Item(5).Width = 200;
                Grid0.Columns.Item(14).Width = 200;

                List<int> ColumnasJustificadas = new List<int>(new int[] { 0, 1, 2, 6, 8, 9, 10, 11, 12, 15 });
                List<int> ColumnasEnfasis = new List<int>(new int[] { 8, 9, 10, 11, 12, 15 });

                for (int iCols = 0; iCols <= Grid0.Columns.Count - 1; iCols++)
                {
                    Grid0.Columns.Item(iCols).Editable = false;
                    Grid0.Columns.Item(iCols).BackColor = Commons.Color_RGB_SAP(255, 255, 255);
                    int White = Color.FromArgb(255, 255, 255).ToArgb();
                    if (ColumnasJustificadas.Contains(iCols))
                    {
                        Grid0.Columns.Item(iCols).RightJustified = true;
                    }
                    if (ColumnasEnfasis.Contains(iCols))
                    {
                        Grid0.Columns.Item(iCols).BackColor = Commons.Color_RGB_SAP(250, 250, 210);
                    }
                }

                GridExtensions.RowNumberGrid(Grid0);
            }
            catch (Exception ex)
            { Application.SBO_Application.StatusBar.SetText(ex.Message);  }
            finally
            {  oForm.Freeze(false);            }

        }
        private void Cargar_Grid_PagosVentas(DateTime fDesde, DateTime fHasta , string Periodo, string cVend)
        {
            try
            {
                oForm.Freeze(true);
                cVend = cVend.Trim().Length == 0 ? "0": cVend;
                string sql = "EXEC  [dbo].[Min_Comisiones_Consultar_Pagos_Venta_Periodo] 0,'" + Periodo + "'";
                oDTTable = oForm.DataSources.DataTables.Item("DT_PAG");
                oDTTable.ExecuteQuery(sql);

                if (oDTTable.Rows.Count > 0 && !oDTTable.IsEmpty)
                {
                    sql = "EXEC  [dbo].[Min_Comisiones_Consultar_Pagos_Venta_Periodo] '" + cVend + "','" + Periodo + "'";
                    oDTTable.ExecuteQuery(sql);
                    Button2.Caption = "Actualizar Pagos";
                    Button2.Item.Enabled = false;
                }
                else
                {
                    sql = "EXEC  [dbo].[Min_Comisiones_Pagos_Ventas_Periodo] '" + fDesde.ToString("MM-dd-yyyy") + "','" + fHasta.ToString("MM-dd-yyyy") + "','" + cVend + "'";
                    oDTTable.ExecuteQuery(sql);
                    Button2.Caption = "Registrar Pagos";
                    Button2.Item.Enabled = true;
                }

                StaticText5.Caption = "Total Registros: " + oDTTable.Rows.Count.ToString().Trim();

                Formatear_Grid_PagosVentas();

                GridExtensions.RowNumberGrid(Grid1);

                Calcular_Totales_Pagos();
            }
            catch
            {            }
            finally
            {  oForm.Freeze(false);            }

        }
        private void Cargar_Grid_Comisiones(string sVend, string Periodo)
        {
            try
            {
                oForm.Freeze(true);



                string sql = "EXEC  [dbo].[Min_Comisiones_Consultar_Comisiones_Periodo] 0,'" + Periodo + "'";
                oDTTable = oForm.DataSources.DataTables.Item("DT_COM");
                oDTTable.ExecuteQuery(sql);

                if (oDTTable.Rows.Count > 0 && !oDTTable.IsEmpty)
                {
                    sql = "EXEC  [dbo].[Min_Comisiones_Consultar_Comisiones_Periodo] '" + sVend + "','" + Periodo + "'";
                    oDTTable.ExecuteQuery(sql);
                    Button1.Caption = "Actualizar Comisiones";
                    Button1.Item.Enabled = false;
                }
                else
                {
                    sql = "EXEC  [dbo].[Min_Comisiones_Calculo_Periodo] '" + sVend + "','" + Periodo + "'";
                    oDTTable.ExecuteQuery(sql);
                    Button1.Caption = "Registrar Comisiones";
                    Button1.Item.Enabled = true;
                }

                oForm.DataSources.DataTables.Item("DT_COMINI").CopyFrom(oDTTable); //Se  copia la Consulta Inicial para Registrarla como Historico en caso de Modificacion.

                StaticText6.Caption = "Total Registros: " + oDTTable.Rows.Count.ToString().Trim();

                Formatear_Grid_Comisiones();

                GridExtensions.RowNumberGrid(Grid2);

            }
            catch (Exception e)
            {            }
            finally
            {  oForm.Freeze(false);            }

        }
        private void Cargar_Grid_Modificaciones(DateTime fDesde, DateTime fHasta, string sVend)
        {
            try
            {
                oForm.Freeze(true);

                oCompany = NConexion.Verifica_Conexion(oCompany);
                SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string sql = "SET NOCOUNT ON;SET FMTONLY OFF;EXEC [SBO_COMERCIAL].[dbo].[Min_Comisiones_Calculo_PCV] 0,1";
                try
                { oRecordset.DoQuery(sql); }
                catch (Exception)
                {                }

                sql = "EXEC  [dbo].[Min_Comisiones_Modificaciones_Periodo] '" + fDesde.ToString("MM-dd-yyyy") + "','" + fHasta.ToString("MM-dd-yyyy") + "','" + sVend + "'";
                oDTTable = oForm.DataSources.DataTables.Item("DT_MOD");
                oDTTable.ExecuteQuery(sql);

                StaticText9.Caption = "Total Registros: " + oDTTable.Rows.Count.ToString().Trim();

                SAPbouiCOM.EditTextColumn oEditCol = (SAPbouiCOM.EditTextColumn)Grid3.Columns.Item(0);
                oEditCol.LinkedObjectType = "17";

                Grid3.Columns.Item(6).Width = 200;
                Grid3.Columns.Item(19).Width = 200;

                List<int> ColumnasJustificadas = new List<int>(new int[] { 0, 1, 2, 7, 9, 10, 11, 12, 13, 14, 16, 17});

                for (int iCols = 0; iCols <= Grid3.Columns.Count - 1; iCols++)
                {
                    Grid3.Columns.Item(iCols).Editable = false;
                    Grid3.Columns.Item(iCols).BackColor = Commons.Color_RGB_SAP(255, 255, 255);
                    int White = Color.FromArgb(255, 255, 255).ToArgb();
                    if (ColumnasJustificadas.Contains(iCols))
                    {
                        Grid3.Columns.Item(iCols).RightJustified = true;
                    }
                }

                int ColorF = Commons.Color_RGB_SAP(255, 0, 0);
                int ColorB = Commons.Color_RGB_SAP(255, 223, 0);
                int Negro = Commons.Color_RGB_SAP(0, 0, 0);
                int Blanco = Commons.Color_RGB_SAP(255, 255, 255);
                for (int iRow = 0; iRow <= Grid3.Rows.Count - 1; iRow++)
                {
                    try
                    {
                        //string Modif = "";
                        if (Convert.ToString(Grid3.DataTable.GetValue(22, iRow)) == "S")
                        {
                            Grid3.CommonSetting.SetRowBackColor(iRow + 1, ColorB);
                            GridExtensions.ColorFontCellGrid(Grid3, 13 + 1, iRow + 1, ColorF, ColorB);
                            GridExtensions.ColorFontCellGrid(Grid3, 14 + 1, iRow + 1, ColorF, ColorB);
                        }
                        else
                        {
                            Grid3.CommonSetting.SetRowBackColor(iRow + 1, Blanco);
                            GridExtensions.ColorFontCellGrid(Grid3, 13 + 1, iRow + 1, Negro, Blanco);
                            GridExtensions.ColorFontCellGrid(Grid3, 14 + 1, iRow + 1, Negro, Blanco);
                        }
                    }
                    catch (Exception)
                    {                    }
                }

                GridExtensions.RowNumberGrid(Grid3);
            }
            catch (Exception e)
            {            }
            finally
            {  oForm.Freeze(false); }

        }
        private void Cargar_Grid_Pagos_Premios(string sVend, string Periodo)
        {
            try
            {
                oForm.Freeze(true);

                sVend = sVend.Trim().Length == 0 ? "0" : sVend;
                string sql = "EXEC  [dbo].[Min_Comisiones_Consultar_Pagos_Premios_Periodo] 0,'" + Periodo + "'";
                oDTTable = oForm.DataSources.DataTables.Item("DT_PREMI");
                oDTTable.ExecuteQuery(sql);

                if (oDTTable.Rows.Count > 0 && !oDTTable.IsEmpty)
                {
                    sql = "EXEC  [dbo].[Min_Comisiones_Consultar_Pagos_Premios_Periodo] '" + sVend + "','" + Periodo + "'";
                    oDTTable.ExecuteQuery(sql);
                    Button2.Caption = "Actualizar Premios";
                    Button2.Item.Enabled = false;
                }
                else
                {
                    sql = "SET NOCOUNT ON;SET FMTONLY OFF;EXEC [SBO_COMERCIAL].[dbo].[Min_Comisiones_Calculo_PCV] 0,1";
                    oDTTable.ExecuteQuery(sql);

                    sql = "EXEC  [dbo].[Min_Comisiones_Pagos_Premios_Periodo] '" + sVend + "'";
                    oDTTable = oForm.DataSources.DataTables.Item("DT_PREMI");
                    oDTTable.ExecuteQuery(sql);

                    Button6.Caption = "Registrar Premios";
                    Button6.Item.Enabled = true;
                }

               StaticText14.Caption = "Total Registros: " + oDTTable.Rows.Count.ToString().Trim();

                Formatear_Grid_PagosPremios();
               GridExtensions.RowNumberGrid(Grid4);
            }
            catch (Exception){}
            finally
            { oForm.Freeze(false); }
        }
        private static void Formatear_Grid_PagosPremios()
        {
            try
            {
                oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_35").Specific;

                //SAPbouiCOM.EditTextColumn oEditCol = (SAPbouiCOM.EditTextColumn)Grid0.Columns.Item(0);
                //oEditCol.LinkedObjectType = "17";

                //Grid0.Columns.Item(5).Width = 200;
                //Grid0.Columns.Item(14).Width = 200;

                List<int> ColumnasJustificadas = new List<int>(new int[] { 2, 4, 5, 6, 7, 8, 9 });
                List<int> ColumnasEditables = new List<int>(new int[] { 5, 6, 8 });
                List<int> ColumnasEnfasis = new List<int>(new int[] { 2, 4, 7, 11 });

                oGrid.Columns.Item(0).TitleObject.Sortable = true;
                oGrid.Columns.Item(2).TitleObject.Sortable = true;


                for (int iCols = 0; iCols <= oGrid.Columns.Count - 1; iCols++)
                {
                    oGrid.Columns.Item(iCols).Editable = false;
                    oGrid.Columns.Item(iCols).BackColor = Commons.Color_RGB_SAP(255, 255, 255);

                    if (ColumnasEditables.Contains(iCols))
                    {
                        oGrid.Columns.Item(iCols).Editable = true;
                        oGrid.Columns.Item(iCols).BackColor = Commons.Color_RGB_SAP(240, 255, 255);
                    }
                    else
                        oGrid.Columns.Item(iCols).Editable = false;

                    if (ColumnasJustificadas.Contains(iCols))
                    {
                        oGrid.Columns.Item(iCols).RightJustified = true;
                    }
                    if (ColumnasEnfasis.Contains(iCols))
                    {
                        oGrid.Columns.Item(iCols).BackColor = Commons.Color_RGB_SAP(250, 250, 210);
                    }
                }
            }
            catch { }
            
        }
        private static void Formatear_Grid_PagosVentas()
        {
            try
            {

                oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_11").Specific;

                SAPbouiCOM.EditTextColumn oEditCol = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item(0);
                oEditCol.LinkedObjectType = "24";

                oEditCol = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("Numero Documento");
                oEditCol.LinkedObjectType = "13";

                oGrid.Columns.Item(0).TitleObject.Sortable = true;
                oGrid.Columns.Item(5).Width = 200;
                oGrid.Columns.Item(14).Width = 100;

                List<int> ColumnasJustificadas = new List<int>(new int[] { 0, 1, 4, 6, 7, 10, 11, 12, 13, 14, 15, 16, 25 });
                List<int> ColumnasEnfasis = new List<int>(new int[] { 9, 10, 11, 12, 13, 14, 15, 16, 17, 25 });
                List<int> ColumnasNoVisibles = new List<int>(new int[] { 18, 19, 20, 21, 22, 23, 24, 27 });

                for (int iCols = 0; iCols <= oGrid.Columns.Count - 1; iCols++)
                {
                    oGrid.Columns.Item(iCols).Editable = false;
                    oGrid.Columns.Item(iCols).BackColor = Commons.Color_RGB_SAP(255, 255, 255);
                    int White = Color.FromArgb(255, 255, 255).ToArgb();
                    oGrid.Columns.Item(iCols).ForeColor = Commons.Color_RGB_SAP(0, 0, 0);
                    if (ColumnasJustificadas.Contains(iCols))
                    {
                        oGrid.Columns.Item(iCols).RightJustified = true;
                    }
                    if (ColumnasEnfasis.Contains(iCols))
                    {
                        oGrid.Columns.Item(iCols).BackColor = Commons.Color_RGB_SAP(250, 250, 210);
                    }
                    if (ColumnasNoVisibles.Contains(iCols))
                    {
                        oGrid.Columns.Item(iCols).Visible = false;
                    }

                }

                // Columna Total Documento Pago
                oGrid.Columns.Item(4).BackColor = Commons.Color_RGB_SAP(240, 255, 255);   
                // Columna Pago Total a FACT / NCR / ANT
                oGrid.Columns.Item(14).Editable = true;
                oGrid.Columns.Item(14).BackColor = Commons.Color_RGB_SAP(240, 255, 255);
                // Columna Comentario 
                oGrid.Columns.Item(26).Editable = true;
                oGrid.Columns.Item(26).Width = 300;
                oGrid.Columns.Item(26).BackColor = Commons.Color_RGB_SAP(240, 255, 255);
                
                for (int iRow = 0; iRow <= oGrid.Rows.Count - 1; iRow++)
                {
                    try
                    {
                        int iColor = Convert.ToInt32(oGrid.DataTable.GetValue(23, iRow));
                        if (iColor > 0)
                        {
                            oGrid.CommonSetting.SetRowBackColor(iRow + 1, iColor);
                        }

                        //Cambiar Colores GRID
                        decimal PagAct = Convert.ToDecimal(oGrid.DataTable.GetValue(14, iRow));
                        decimal PagIni = Convert.ToDecimal(oGrid.DataTable.GetValue(24, iRow));
                        string ProyAct = Convert.ToString(oGrid.DataTable.GetValue(15, iRow));
                        string ProyIni = Convert.ToString(oGrid.DataTable.GetValue(25, iRow));

                        string NuevaL = oGrid.DataTable.GetValue(22, iRow).ToString();

                        //Cambiar Color de Linea
                        if ((PagAct != PagIni) && NuevaL.Trim() != "Y")
                        {
                            int Color = Commons.Color_RGB_SAP(250, 235, 215);//(255, 160, 122);
                            GridExtensions.ColorBackColorRowGrid(oGrid, iRow + 1, Color);
                        }

                        //Cambiar Color de Fuentes en Datos Modificados (14)
                        if (PagAct != PagIni)
                        {
                            int ColorF = Commons.Color_RGB_SAP(255, 0, 0);
                            int ColorB = Commons.Color_RGB_SAP(250, 250, 210);
                            GridExtensions.ColorFontCellGrid(oGrid, 14 + 1, iRow + 1, ColorF, ColorB);
                        }
                        else
                        {
                            int ColorF = Commons.Color_RGB_SAP(0, 0, 0);
                            int ColorB = NuevaL.Trim() != "Y" ? Commons.Color_RGB_SAP(240, 255, 255) : Convert.ToInt32(oGrid.DataTable.GetValue(23, iRow));
                            GridExtensions.ColorFontCellGrid(oGrid, 14 + 1, iRow + 1, ColorF, ColorB);
                        }

                        if (ProyAct != ProyIni)
                        {
                            int ColorF = Commons.Color_RGB_SAP(255, 0, 0);
                            int ColorB = Commons.Color_RGB_SAP(250, 250, 210);
                            GridExtensions.ColorFontCellGrid(oGrid, 15 + 1, iRow + 1, ColorF, ColorB);
                        }
                        else
                        {
                            int ColorF = Commons.Color_RGB_SAP(0, 0, 0);
                            int ColorB = NuevaL.Trim() != "Y" ? Commons.Color_RGB_SAP(250, 250, 210) : Convert.ToInt32(oGrid.DataTable.GetValue(23, iRow));
                            GridExtensions.ColorFontCellGrid(oGrid, 15 + 1, iRow + 1, ColorF, ColorB);
                        }

                        //Colocar en Rojo Pagos Negativos (N/CR)
                        if (PagAct < 0)
                            oGrid.CommonSetting.SetCellFontColor(iRow + 1, 14 + 1, Commons.Color_RGB_SAP(255, 0, 0));

                        //Actualiza el Contador
                        int ContDocuPago = Convert.ToInt32(oGrid.DataTable.GetValue(28, iRow));
                        oGrid.DataTable.SetValue(29, iRow, ContDocuPago);

                        if (ContDocuPago > 1)
                        {
                            for (int i = 1; i <= ContDocuPago - 1; i++)
                            {
                                oGrid.DataTable.SetValue(29, iRow - i, ContDocuPago);
                            }

                            for (int iCols = 0; iCols <= 8; iCols++)
                            {
                                oGrid.CommonSetting.SetCellFontColor(iRow + 1, iCols + 1, iColor > 0 ? iColor : Commons.Color_RGB_SAP(255, 255, 255));
                            }
                        }

                    }
                    catch (Exception)
                    {                    }
                }

                //oGrid.AutoResizeColumns();
                //oGrid.Columns.Item(10).BackColor = Funciones.Color_RGB_SAP(250, 250, 210);
            }
            catch (Exception)
            {            }

        }
        private static void Formatear_Grid_Comisiones()
        {
            try
            {
                oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_14").Specific;

                SAPbouiCOM.EditTextColumn oEditCol = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item(4);
                oEditCol.LinkedObjectType = "17";

                oGrid.Columns.Item(1).Width = 220;
                oGrid.Columns.Item(4).Width = 70;
                oGrid.Columns.Item(5).Width = 320;
                oGrid.Columns.Item(9).Width = 30;

                oGrid.Columns.Item(15).Visible = false;
                oGrid.Columns.Item(16).Visible = false;
                oGrid.Columns.Item(17).Visible = false;
                oGrid.Columns.Item(18).Visible = false;
                oGrid.Columns.Item(19).Visible = false;
                oGrid.Columns.Item(20).Visible = false;
                oGrid.Columns.Item(22).Visible = false;
                oGrid.Columns.Item(23).Visible = false;
                //Grid0.Columns.Item(14).Width = 200;

                List<int> ColumnasJustificadas = new List<int>(new int[] { 2, 3, 4, 6, 7, 8, 9, 10, 11, 12, 13, 14, 24 });
                List<int> ColumnasEditables = new List<int>(new int[] { 10, 11, 13, 14, 21 });
                List<int> ColumnasEnfasis = new List<int>(new int[] { 6, 7, 8, 9, 12, 24 });


                for (int iCols = 0; iCols <= oGrid.Columns.Count - 1; iCols++)
                {
                    if (ColumnasEditables.Contains(iCols)) oGrid.Columns.Item(iCols).Editable = true;
                    else oGrid.Columns.Item(iCols).Editable = false;

                    oGrid.Columns.Item(iCols).BackColor = Commons.Color_RGB_SAP(255, 255, 255);
                    int White = Color.FromArgb(255, 255, 255).ToArgb();
                    if (ColumnasJustificadas.Contains(iCols))
                    {
                        oGrid.Columns.Item(iCols).RightJustified = true;
                    }

                    if (ColumnasEnfasis.Contains(iCols))
                    {
                        oGrid.Columns.Item(iCols).BackColor = Commons.Color_RGB_SAP(250, 250, 210);
                    }

                }

                //oGrid.Columns.Item(10).BackColor = Funciones.Color_RGB_SAP(240, 255, 255);
                //oGrid.Columns.Item(11).BackColor = Funciones.Color_RGB_SAP(240, 255, 255);
                //oGrid.Columns.Item(13).BackColor = Funciones.Color_RGB_SAP(240, 255, 255);
                oGrid.Columns.Item(21).BackColor = Commons.Color_RGB_SAP(240, 255, 255);

                for (int iRow = 0; iRow <= oGrid.Rows.Count - 1; iRow++)
                {
                    try
                    {
                        int iColor = Convert.ToInt32(oGrid.DataTable.GetValue(20, iRow));
                        if (iColor > 0)
                        {
                            oGrid.CommonSetting.SetRowBackColor(iRow + 1, iColor);
                        }

                        //Cambiar Colores GRID
                        decimal PreAct = Convert.ToDecimal(oGrid.DataTable.GetValue("Comi Premio", iRow));
                        decimal PreIni = Convert.ToDecimal(oGrid.DataTable.GetValue("Comi Premio Ini", iRow));
                        decimal ComAct = Convert.ToDecimal(oGrid.DataTable.GetValue("Comision Vta", iRow));
                        decimal ComIni = Convert.ToDecimal(oGrid.DataTable.GetValue("Comision Vta Ini", iRow));
                        string NuevaL = oGrid.DataTable.GetValue(19, iRow).ToString();

                        //Cambiar Color de Fuentes en Datos Modificados (10, 11, 13, 14)
                        if (PreAct != PreIni)
                        {
                            int ColorF = Commons.Color_RGB_SAP(255, 0, 0);
                            int ColorB = Commons.Color_RGB_SAP(250, 250, 210);
                            GridExtensions.ColorFontCellGrid(oGrid, 10 + 1, iRow + 1, ColorF, ColorB);
                            GridExtensions.ColorFontCellGrid(oGrid, 11 + 1, iRow + 1, ColorF, ColorB);
                        }
                        else
                        {
                            int ColorF = Commons.Color_RGB_SAP(0, 0, 0);
                            int ColorB = Convert.ToInt32(oGrid.DataTable.GetValue(20, iRow));
                            ColorB = ColorB == 0 ? Commons.Color_RGB_SAP(240, 255, 255) : ColorB;
                            GridExtensions.ColorFontCellGrid(oGrid, 10 + 1, iRow + 1, ColorF, ColorB);
                            GridExtensions.ColorFontCellGrid(oGrid, 11 + 1, iRow + 1, ColorF, ColorB);
                        }

                        if (ComAct != ComIni)
                        {
                            int ColorF = Commons.Color_RGB_SAP(255, 0, 0);
                            int ColorB = Commons.Color_RGB_SAP(250, 250, 210);
                            GridExtensions.ColorFontCellGrid(oGrid, 13 + 1, iRow + 1, ColorF, ColorB);
                            GridExtensions.ColorFontCellGrid(oGrid, 14 + 1, iRow + 1, ColorF, ColorB);
                        }
                        else
                        {
                            int ColorF = Commons.Color_RGB_SAP(0, 0, 0);
                            int ColorB = Convert.ToInt32(oGrid.DataTable.GetValue(20, iRow));
                            ColorB = ColorB == 0 ? Commons.Color_RGB_SAP(240, 255, 255) : ColorB;
                            GridExtensions.ColorFontCellGrid(oGrid, 13 + 1, iRow + 1, ColorF, ColorB);
                            GridExtensions.ColorFontCellGrid(oGrid, 14 + 1, iRow + 1, ColorF, ColorB);
                        }

                    }
                    catch (Exception)
                    {                    }
                }

                SAPbouiCOM.EditTextColumn col = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item(6);
                col.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                col = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item(8);
                col.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                col = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item(11);
                col.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                col = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item(14);
                col.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

            }
            catch (Exception)
            {            }
        }
        private void Insertar_Pagos(SAPbouiCOM.DataTable oDT)
        {
            bool bProgBar = false;
            //********************* PROGRESS BAR
            bProgBar = ProgressBarExtensions.Create_ProgressBar(ref oProgBar,"Insertando Datos Pagos", 2);
           
            NumberFormatInfo nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";

            string sql = "";
            DateTime FechaHoraAct = DateTime.Now;

            for (int i = 0; i <= oDT.Rows.Count - 1; i++)
            {
                try
                {
                    sql = @"EXEC  [dbo].[Min_Comisiones_Insertar_Pagos_Ventas_Periodo] '" + oDT.GetValue(0, i) + "','"
                                                                                          + Convert.ToDateTime(oDT.GetValue(1, i)).ToString("yyyy-MM-dd HH:mm:ss") + "','"
                                                                                          + oDT.GetValue(2, i).ToString().Replace("'", "''") + "','"
                                                                                          + oDT.GetValue(3, i).ToString().Replace("'", "''") + "',"
                                                                                          + Convert.ToDecimal(oDT.GetValue(4, i)).ToString(nfi) + ",'"
                                                                                          + oDT.GetValue(5, i).ToString().Replace("'", "''") + "','"
                                                                                          + oDT.GetValue(6, i) + "','"
                                                                                          + Convert.ToDateTime(oDT.GetValue(7, i)).ToString("yyyy-MM-dd HH:mm:ss") + "','"
                                                                                          + oDT.GetValue(8, i) + "','"
                                                                                          + oDT.GetValue(9, i) + "','"
                                                                                          + oDT.GetValue(10, i) + "','"
                                                                                          + oDT.GetValue(11, i) + "','"
                                                                                          + Convert.ToDateTime(oDT.GetValue(12, i)).ToString("yyyy-MM-dd HH:mm:ss") + "','"
                                                                                          + Convert.ToDateTime(oDT.GetValue(13, i)).ToString("yyyy-MM-dd HH:mm:ss") + "',"
                                                                                          + Convert.ToDecimal(oDT.GetValue(14, i)).ToString(nfi) + ",'"
                                                                                          + oDT.GetValue(15, i) + "','"
                                                                                          + oDT.GetValue(16, i) + "','"
                                                                                          + oDT.GetValue(17, i) + "','"
                                                                                          + oDT.GetValue(18, i) + "','"
                                                                                          + NConexion.sCodUsuActual + "','"
                                                                                          + FechaHoraAct.ToString("yyyy-MM-dd HH:mm:ss") + "','"
                                                                                          + ComboBox0.Selected.Description + "','"
                                                                                          + oDT.GetValue(22, i) + "','"
                                                                                          + oDT.GetValue(23, i) + "',"
                                                                                          + Convert.ToDecimal(oDT.GetValue(24, i)).ToString(nfi) + ",'"
                                                                                          + oDT.GetValue(25, i) + "','"
                                                                                          + oDT.GetValue(26, i).ToString().Replace("'", "''") + "','"
                                                                                          + oDT.GetValue(27, i) + "'";

                    DT_SQL.ExecuteQuery(sql);
                }
                catch (Exception){}
            }
            // Aplicar Consulta para que Actualice tabla temporal de Pagos por Proyectos
            try
            {
                sql = "EXEC  [dbo].[Min_Comisiones_Consultar_Pagos_Venta_Periodo] 0,'" + ComboBox0.Selected.Description + "'";
                DT_SQL.ExecuteQuery(sql);
            }
            catch (Exception){}

            //********************* PROGRESS BAR
            bProgBar = ProgressBarExtensions.Increment_ProgressBar(ref oProgBar, 1);
            bProgBar = ProgressBarExtensions.Close_ProgressBar(ref oProgBar);
        }
        private void Insertar_Comisiones(SAPbouiCOM.DataTable oDT)
        {
            bool bProgBar = false;
            bProgBar = ProgressBarExtensions.Create_ProgressBar(ref oProgBar, "Insertando Datos Comisiones", 2);
                     
            NumberFormatInfo nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";

            string sql = "";
            DateTime FechaHoraAct = DateTime.Now;

            for (int i = 0; i <= oDT.Rows.Count - 1; i++)
            {
                try
                {
                    sql = @"EXEC  [dbo].[Min_Comisiones_Insertar_Comisiones_Periodo] '" + oDT.GetValue(0, i).ToString().Replace("'", "''") + "','"
                                                                                        + oDT.GetValue(1, i).ToString().Replace("'", "''") + "','"
                                                                                        + Convert.ToDateTime(oDT.GetValue(2, i)).ToString("yyyy-MM-dd HH:mm:ss") + "','"
                                                                                        + oDT.GetValue(3, i) + "','"
                                                                                        + oDT.GetValue(4, i) + "','"
                                                                                        + oDT.GetValue(5, i).ToString().Replace("'","''") + "',"
                                                                                        + Convert.ToDecimal(oDT.GetValue(6, i)).ToString(nfi) + ","
                                                                                        + Convert.ToDecimal(oDT.GetValue(7, i)).ToString(nfi) + ","
                                                                                        + Convert.ToDecimal(oDT.GetValue(8, i)).ToString(nfi) + ","
                                                                                        + Convert.ToDecimal(oDT.GetValue(10, i)).ToString(nfi) + ","
                                                                                        + Convert.ToDecimal(oDT.GetValue(11, i)).ToString(nfi) + ","
                                                                                        + Convert.ToDecimal(oDT.GetValue(12, i)).ToString(nfi) + ","
                                                                                        + Convert.ToDecimal(oDT.GetValue(13, i)).ToString(nfi) + ","
                                                                                        + Convert.ToDecimal(oDT.GetValue(14, i)).ToString(nfi) + ",'"
                                                                                        + NConexion.sCodUsuActual + "','"
                                                                                        + FechaHoraAct.ToString("yyyy-MM-dd HH:mm:ss") + "','"
                                                                                        + oDT.GetValue(17, i) + "','"
                                                                                        + oDT.GetValue(18, i) + "',"
                                                                                        + Convert.ToDecimal(oDT.GetValue(9, i)).ToString(nfi) + ",'"
                                                                                        + oDT.GetValue(19, i) + "','"
                                                                                        + oDT.GetValue(20, i) + "','"
                                                                                        + oDT.GetValue(21, i).ToString().Replace("'", "''") + "','"
                                                                                        + oDT.GetValue(22, i) + "','"
                                                                                        + oDT.GetValue(23, i) + "',"
                                                                                        + Convert.ToDecimal(oDT.GetValue(24, i)).ToString(nfi) + "";

                    DT_SQL.ExecuteQuery(sql);
                }
                catch (Exception)
                {                }
            }

            bProgBar = ProgressBarExtensions.Increment_ProgressBar(ref oProgBar, 1);
            bProgBar = ProgressBarExtensions.Close_ProgressBar(ref oProgBar);

        }
        private void Insertar_Pagos_Premios(SAPbouiCOM.DataTable oDT)
        {
            bool bProgBar = false;
            //********************* PROGRESS BAR
            bProgBar = ProgressBarExtensions.Create_ProgressBar(ref oProgBar, "Insertando Datos Pagos Premios", 2);

            NumberFormatInfo nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";

            string sql = "";
            DateTime FechaHoraAct = DateTime.Now;

            for (int i = 0; i <= oDT.Rows.Count - 1; i++)
            {
                try
                {
                    sql = @"EXEC  [dbo].[Min_Comisiones_Insertar_Pagos_Premios_Periodo] '" + oDT.GetValue(0, i) + "','"
                                                                                          + oDT.GetValue(1, i).ToString().Replace("'", "''") + "','"
                                                                                          + oDT.GetValue(2, i) + "','"
                                                                                          + oDT.GetValue(3, i).ToString().Replace("'", "''") + "','"
                                                                                          + Convert.ToDecimal(oDT.GetValue(4, i)).ToString(nfi) + ",'"
                                                                                          + Convert.ToDecimal(oDT.GetValue(5, i)).ToString(nfi) + ",'"
                                                                                          + Convert.ToDecimal(oDT.GetValue(6, i)).ToString(nfi) + ",'"
                                                                                          + Convert.ToDecimal(oDT.GetValue(7, i)).ToString(nfi) + ",'"
                                                                                          + Convert.ToDecimal(oDT.GetValue(8, i)).ToString(nfi) + ",'"
                                                                                          + oDT.GetValue(9, i) + "','"
                                                                                          + oDT.GetValue(10, i) + "','"
                                                                                          + oDT.GetValue(11, i) + "','"
                                                                                          + NConexion.sCodUsuActual + "','"
                                                                                          + FechaHoraAct.ToString("yyyy-MM-dd HH:mm:ss") + "'"
                                                                                          + ComboBox0.Selected.Description + "','" ;

                    DT_SQL.ExecuteQuery(sql);
                }
                catch (Exception) { }
            }
            // Aplicar Consulta para que Actualice tabla temporal de Premios por Proyectos
            try
            {
                sql = "EXEC  [dbo].[Min_Comisiones_Pagos_Premios_Periodo] 0";
                DT_SQL.ExecuteQuery(sql);
            }
            catch (Exception) { }

            //********************* PROGRESS BAR
            bProgBar = ProgressBarExtensions.Increment_ProgressBar(ref oProgBar, 1);
            bProgBar = ProgressBarExtensions.Close_ProgressBar(ref oProgBar);
        }
        private void Eliminar_Pagos(string sVend, string Periodo)
        {
            try
            {
                sVend =  sVend.Trim() == "" ? "0": sVend ;
                string sql = @"EXEC  [dbo].[Min_Comisiones_Eliminar_Pagos_Ventas_Periodo] '" + sVend + "','" + Periodo + "'";
                DT_SQL.ExecuteQuery(sql);
            }
            catch (Exception)
            {            }
        }
        private void Eliminar_Pagos_Premios(string sVend, string Periodo)
        {
            try
            {
                sVend = sVend.Trim() == "" ? "0" : sVend;
                string sql = @"EXEC  [dbo].[Min_Comisiones_Eliminar_Pagos_Premios_Periodo] '" + sVend + "','" + Periodo + "'";
                DT_SQL.ExecuteQuery(sql);
            }
            catch (Exception)
            { }
        }
        private void Eliminar_Comisiones(string sVend, string Periodo)
        {
            try
            {
                string sql = @"EXEC  [dbo].[Min_Comisiones_Eliminar_Comisiones_Periodo] '" + sVend + "','" + Periodo + "'";
                DT_SQL.ExecuteQuery(sql);
            }
            catch (Exception)
            {            }
        }
        private void Insertar_Comisiones_Historico(SAPbouiCOM.DataTable oDT)
        {

            NumberFormatInfo nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";

            string sql = "";
            DateTime FechaHoraAct = DateTime.Now;

            for (int i = 0; i <= oDT.Rows.Count - 1; i++)
            {
                try
                {
                    sql = @"EXEC  [dbo].[Min_Comisiones_Insertar_Comisiones_Historico] '" + oDT.GetValue(0, i).ToString().Replace("'", "''") + "','"
                                                                                        + oDT.GetValue(1, i).ToString().Replace("'", "''") + "','"
                                                                                        + Convert.ToDateTime(oDT.GetValue(2, i)).ToString("yyyy-MM-dd HH:mm:ss") + "','"
                                                                                        + oDT.GetValue(3, i) + "','"
                                                                                        + oDT.GetValue(4, i) + "','"
                                                                                        + oDT.GetValue(5, i).ToString().Replace("'", "''") + "',"
                                                                                        + Convert.ToDecimal(oDT.GetValue(6, i)).ToString(nfi) + ","
                                                                                        + Convert.ToDecimal(oDT.GetValue(7, i)).ToString(nfi) + ","
                                                                                        + Convert.ToDecimal(oDT.GetValue(8, i)).ToString(nfi) + ","
                                                                                        + Convert.ToDecimal(oDT.GetValue(10, i)).ToString(nfi) + ","
                                                                                        + Convert.ToDecimal(oDT.GetValue(11, i)).ToString(nfi) + ","
                                                                                        + Convert.ToDecimal(oDT.GetValue(12, i)).ToString(nfi) + ","
                                                                                        + Convert.ToDecimal(oDT.GetValue(13, i)).ToString(nfi) + ","
                                                                                        + Convert.ToDecimal(oDT.GetValue(14, i)).ToString(nfi) + ",'"
                                                                                        + oDT.GetValue(15, i) + "','"
                                                                                        + Convert.ToDateTime(oDT.GetValue(16, i)).ToString("yyyy-MM-dd HH:mm:ss") + "','"
                                                                                        + oDT.GetValue(17, i) + "','"
                                                                                        + oDT.GetValue(18, i) + "','"
                                                                                        + NConexion.sCodUsuActual + "','"
                                                                                        + FechaHoraAct.ToString("yyyy-MM-dd HH:mm:ss") + "',"
                                                                                        + Convert.ToDecimal(oDT.GetValue(9, i)).ToString(nfi) + ",'"
                                                                                        + oDT.GetValue(19, i) + "','"
                                                                                        + oDT.GetValue(20, i) + "','"
                                                                                        + oDT.GetValue(21, i).ToString().Replace("'", "''") + "','"
                                                                                        + oDT.GetValue(22, i) + "','"
                                                                                        + oDT.GetValue(23, i) + "',"
                                                                                        + Convert.ToDecimal(oDT.GetValue(24, i)).ToString(nfi) + "";

                    DT_SQL.ExecuteQuery(sql);
                }
                catch (Exception)
                {                }
            }

        }
        public static void Agregar_Documento_Grid_Comisiones(string nPCV, ref int TipoFila, int Vend)
        {
            try
            {
                oForm.Freeze(true);

                oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("Item_4").Specific;
                oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_14").Specific;
                SAPbouiCOM.StaticText oStatic = (SAPbouiCOM.StaticText)oForm.Items.Item("Item_21").Specific;
                SAPbouiCOM.RowHeaders oHeader = null;
                oHeader = oGrid.RowHeaders;

                SAPbouiCOM.DataTable ooDTTableGRID = oForm.DataSources.DataTables.Item("DT_COM");
                oDTTable = oForm.DataSources.DataTables.Item("DT_SQL");

                string sql = "EXEC [Min_Comisiones_Calculo_PCV] " + nPCV + ", " + TipoFila + ", " + Vend;
                oDTTable.ExecuteQuery(sql);
                decimal valor = Convert.ToDecimal(oDTTable.GetValue(23, oDTTable.Rows.Count - 1));
                if (!oDTTable.IsEmpty)
                {
                    DataTableExtensions.JoinDataTables(ooDTTableGRID, oDTTable);
                    valor = Convert.ToDecimal(ooDTTableGRID.GetValue(23, ooDTTableGRID.Rows.Count - 1));
                    //Se Borra y vuelve a cargar el DT_COM para que actualice los totales
                    oDTTable.CopyFrom(ooDTTableGRID);
                    ooDTTableGRID.Clear();
                    ooDTTableGRID.CopyFrom(oDTTable);

                    Formatear_Grid_Comisiones();

                    string dValor = Convert.ToString(Commons.Color_RGB_SAP(250, 250, 210));
                    ooDTTableGRID.SetValue(20, ooDTTableGRID.Rows.Count - 1, dValor); //Valor del Color
                    ooDTTableGRID.SetValue(19, ooDTTableGRID.Rows.Count - 1, "Y"); //Linea Nueva
                    ooDTTableGRID.SetValue(17, ooDTTableGRID.Rows.Count - 1, oComboBox.Selected.Description); //Valor del periodo Actual

                    oGrid.CommonSetting.SetRowBackColor(ooDTTableGRID.Rows.Count, Commons.Color_RGB_SAP(250, 250, 210));

                    oHeader.SetText(ooDTTableGRID.Rows.Count - 1, Convert.ToString(ooDTTableGRID.Rows.Count));

                    oStatic.Caption = "Total Registros: " + ooDTTableGRID.Rows.Count.ToString().Trim();

                    oForm.Items.Item("Item_18").Enabled = true;
                    valor = Convert.ToDecimal(ooDTTableGRID.GetValue(23, ooDTTableGRID.Rows.Count - 1));
                }
                else 
                { 
                    Application.SBO_Application.MessageBox("El Documento pertenece a un Periodo que No esta Registrado");
                    TipoFila = 0;
                }

            }
            catch (Exception)
            {            }
            finally
            {  oForm.Freeze(false);            }

        }
        public static void Agregar_Documento_Grid_Pagos(string nPago)
        {
            try
            {
                oForm.Freeze(true);

                oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("Item_4").Specific;
                oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_11").Specific;
                SAPbouiCOM.StaticText oStatic = (SAPbouiCOM.StaticText)oForm.Items.Item("Item_20").Specific;
                SAPbouiCOM.RowHeaders oHeader = null;
                oHeader = oGrid.RowHeaders;

                SAPbouiCOM.DataTable ooDTTableGRID = oForm.DataSources.DataTables.Item("DT_PAG");
                oDTTable = oForm.DataSources.DataTables.Item("DT_SQL");

                string sql = "EXEC [dbo].[Min_Comisiones_Pagos_Ventas_Individual] " + nPago ;
                oDTTable.ExecuteQuery(sql);
                //decimal valor = Convert.ToDecimal(oDTTable.GetValue(23, oDTTable.Rows.Count - 1));
                if (!oDTTable.IsEmpty)
                {
                    DataTableExtensions.JoinDataTables(ooDTTableGRID, oDTTable);
                    //valor = Convert.ToDecimal(ooDTTableGRID.GetValue(23, ooDTTableGRID.Rows.Count - 1));
                    //Se Borra y vuelve a cargar el DT_COM para que actualice los totales
                    oDTTable.CopyFrom(ooDTTableGRID);
                    ooDTTableGRID.Clear();
                    ooDTTableGRID.CopyFrom(oDTTable);

                    Formatear_Grid_PagosVentas();

                    //Cambiar color e informacion de lineas nuevas
                    int ContDocuPago = Convert.ToInt32(ooDTTableGRID.GetValue(28, ooDTTableGRID.Rows.Count - 1));
                    string dValor = Convert.ToString(Commons.Color_RGB_SAP(250, 250, 210));
                                       
                    for (int i = 1; i <= ContDocuPago ; i++)
                    {
                        int iRow = ooDTTableGRID.Rows.Count - i;

                        ooDTTableGRID.SetValue(23, iRow, dValor); //Valor del Color
                        ooDTTableGRID.SetValue(22, iRow, "Y"); //Linea Nueva
                        ooDTTableGRID.SetValue(21, iRow, oComboBox.Selected.Description); //Valor del periodo Actual

                        oGrid.CommonSetting.SetRowBackColor(iRow + 1, Commons.Color_RGB_SAP(250, 250, 210));

                        oHeader.SetText(iRow, Convert.ToString(iRow + 1));
                    }

                    Calcular_Totales_Pagos();
                    oStatic.Caption = "Total Registros: " + ooDTTableGRID.Rows.Count.ToString().Trim();

                    oForm.Items.Item("Item_45").Enabled = true;
                    //valor = Convert.ToDecimal(ooDTTableGRID.GetValue(23, ooDTTableGRID.Rows.Count - 1));
                }
                else
                {
                    Application.SBO_Application.MessageBox("El Documento pertenece a un Periodo que No esta Registrado");
                }

            }
            catch (Exception)
            {
            }
            finally
            {
                oForm.Freeze(false);
            }

        }
        public static void Agregar_Proyecto_Grid_Pagos(string sProyecto)
        {
            try
            {
                oForm.Freeze(true);

                string sql = "SELECT PR.U_SlpCode, SlpName, VN.Memo as 'Vendedor' FROM [@ZINFOP] PR JOIN [OSLP] VN ON PR.U_SlpCode = VN.SlpCode WHERE U_PrjCode = '" + sProyecto + "'";

                oDTTable = oForm.DataSources.DataTables.Item("DT_SQL");
                oDTTable.ExecuteQuery(sql);

                string IdVend = oDTTable.GetValue(0, 0).ToString();
                string NomVend = oDTTable.GetValue(2, 0).ToString();

                oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_11").Specific;
                int gridRow = oGrid.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                int dtRow = oGrid.GetDataTableRowIndex(gridRow);

                string sProyIni = oGrid.DataTable.GetValue("Proyecto Ini", dtRow).ToString();
                string NuevaL = oGrid.DataTable.GetValue(22, dtRow).ToString();

                //Asignar proyecto y Cambiar Color de Fuentes en Datos Modificados (Proyecto)
                if (sProyecto != sProyIni)
                {
                    oGrid.DataTable.SetValue("Proyecto", dtRow, sProyecto.Trim().Length == 0 ? "N/A" : sProyecto);
                    int ColorF = Commons.Color_RGB_SAP(255, 0, 0);
                    int ColorB = Commons.Color_RGB_SAP(250, 250, 210);
                    GridExtensions.ColorFontCellGrid(oGrid, 15 + 1, gridRow + 1, ColorF, ColorB);
                    oGrid.DataTable.SetValue("Vendedor", dtRow, sProyecto.Trim().Length == 0 ? "" : NomVend);
                    oGrid.DataTable.SetValue("Id Vend", dtRow, sProyecto.Trim().Length == 0 ? "" : IdVend);
                }
                else
                {
                    oGrid.DataTable.SetValue("Proyecto", dtRow, sProyecto);
                    int ColorF = Commons.Color_RGB_SAP(0, 0, 0);
                    int ColorB = NuevaL.Trim() != "Y" ? Commons.Color_RGB_SAP(250, 250, 210) : Convert.ToInt32(oGrid.DataTable.GetValue(20, dtRow));
                    GridExtensions.ColorFontCellGrid(oGrid, 15 + 1, gridRow + 1, ColorF, ColorB);
                    oGrid.DataTable.SetValue("Vendedor", dtRow, NomVend);
                    oGrid.DataTable.SetValue("Id Vend", dtRow, IdVend);
                }

                ((SAPbouiCOM.Button)oForm.Items.Item("Item_45").Specific).Item.Enabled = true;

            }
            catch (Exception)
            { }
            finally
            {  oForm.Freeze(false);            }
        }
        private static void Calcular_Valor_Comision(string ColUID, int dtRow, int AsignaBase )
        {
            oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_14").Specific;

            NumberFormatInfo nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";

            Double PorcDesc = 0.0;
            Double PorComiBase;
            Double dValor;
            Double Porc;
            Double Venta;
            Double Comision;
            String sValor;

            PorcDesc = Convert.ToDouble(oGrid.DataTable.GetValue("Porc Dscto", dtRow));
            PorComiBase = (1.6 - (1.6 * ((PorcDesc * 2) / 100)));
            switch (ColUID)
            {
                case "Porc Comi":
                    Porc = Math.Abs(Convert.ToDouble(oGrid.DataTable.GetValue(ColUID, dtRow)));
                    Porc = Porc == Math.Round(PorComiBase, 2) || AsignaBase == 1 ? PorComiBase : Porc;
                    Venta = Convert.ToDouble(oGrid.DataTable.GetValue("Monto Vta Neta", dtRow));
                    dValor = Math.Round((Venta * (Porc / 100)), 0);
                    sValor = dValor.ToString(nfi);
                    oGrid.DataTable.SetValue("Comision Vta", dtRow, sValor);
                    if (AsignaBase == 1)
                    {
                        sValor = PorComiBase.ToString(nfi);
                        oGrid.DataTable.SetValue("Porc Comi", dtRow, sValor);
                    }
                    break;
                case "Comision Vta":
                    Comision = Convert.ToDouble(oGrid.DataTable.GetValue(ColUID, dtRow));
                    Venta = Convert.ToDouble(oGrid.DataTable.GetValue("Monto Vta Neta", dtRow));
                    dValor = ((Comision * 100) / Venta);
                    sValor = dValor.ToString(nfi);
                    oGrid.DataTable.SetValue("Porc Comi", dtRow, sValor);
                    break;
                case "Porc Prem Com":
                    Porc = Math.Abs(Convert.ToDouble(oGrid.DataTable.GetValue(ColUID, dtRow)));
                    Venta = Convert.ToDouble(oGrid.DataTable.GetValue("Monto Vta Neta", dtRow));
                    dValor = Math.Round((Venta * (Porc / 100)), 0);
                    sValor = dValor.ToString(nfi);
                    oGrid.DataTable.SetValue("Comi Premio", dtRow, sValor);
                    break;
                case "Comi Premio":
                    Comision = Convert.ToDouble(oGrid.DataTable.GetValue(ColUID, dtRow));
                    Venta = Convert.ToDouble(oGrid.DataTable.GetValue("Monto Vta Neta", dtRow));
                    dValor = ((Comision * 100) / Venta);
                    sValor = dValor.ToString(nfi);
                    oGrid.DataTable.SetValue("Porc Prem Com", dtRow, sValor);
                    break;
            }

            //Cambiar Colores GRID
            decimal PreAct = Convert.ToDecimal(oGrid.DataTable.GetValue("Comi Premio", dtRow));
            decimal PreIni = Convert.ToDecimal(oGrid.DataTable.GetValue("Comi Premio Ini", dtRow));
            decimal ComAct = Convert.ToDecimal(oGrid.DataTable.GetValue("Comision Vta", dtRow));
            decimal ComIni = Convert.ToDecimal(oGrid.DataTable.GetValue("Comision Vta Ini", dtRow));
            string NuevaL = oGrid.DataTable.GetValue(19, dtRow).ToString();

            //Cambiar Color de Linea
            if ((PreAct == PreIni && ComAct == ComIni) && NuevaL.Trim() != "Y")
            {
                int Color = Commons.Color_RGB_SAP(255, 255, 255);
                oGrid.DataTable.SetValue(20, dtRow, Color);
                GridExtensions.ColorBackColorRowGrid(oGrid, dtRow + 1, Color);
            }
            else if (NuevaL.Trim() != "Y")
            {
                int Color = Commons.Color_RGB_SAP(250, 235, 215);//(255, 160, 122);
                oGrid.DataTable.SetValue(20, dtRow, Color);
                GridExtensions.ColorBackColorRowGrid(oGrid, dtRow + 1, Color);
            }

            //Cambiar Color de Fuentes en Datos Modificados (10, 11, 13, 14)
            if (PreAct != PreIni)
            {
                int ColorF = Commons.Color_RGB_SAP(255, 0, 0);
                int ColorB = Commons.Color_RGB_SAP(250, 250, 210);
                GridExtensions.ColorFontCellGrid(oGrid, 10 + 1, dtRow + 1, ColorF, ColorB);
                GridExtensions.ColorFontCellGrid(oGrid, 11 + 1, dtRow + 1, ColorF, ColorB);
                if (NuevaL.Trim() != "Y") oGrid.DataTable.SetValue(19, dtRow, "MP");
            }
            else
            {
                int ColorF = Commons.Color_RGB_SAP(0, 0, 0);
                int ColorB = NuevaL.Trim() != "Y" ? Commons.Color_RGB_SAP(240, 255, 255) : Convert.ToInt32(oGrid.DataTable.GetValue(20, dtRow));
                GridExtensions.ColorFontCellGrid(oGrid, 10 + 1, dtRow + 1, ColorF, ColorB);
                GridExtensions.ColorFontCellGrid(oGrid, 11 + 1, dtRow + 1, ColorF, ColorB);
                if (NuevaL.Trim() != "Y") oGrid.DataTable.SetValue(19, dtRow, "");
            }

            if (ComAct != ComIni)
            {
                int ColorF = Commons.Color_RGB_SAP(255, 0, 0);
                int ColorB = Commons.Color_RGB_SAP(250, 250, 210);
                GridExtensions.ColorFontCellGrid(oGrid, 13 + 1, dtRow + 1, ColorF, ColorB);
                GridExtensions.ColorFontCellGrid(oGrid, 14 + 1, dtRow + 1, ColorF, ColorB);
                if (NuevaL.Trim() != "Y") oGrid.DataTable.SetValue(19, dtRow, "MC");
            }
            else
            {
                int ColorF = Commons.Color_RGB_SAP(0, 0, 0);
                int ColorB = NuevaL.Trim() != "Y" ? Commons.Color_RGB_SAP(240, 255, 255) : Convert.ToInt32(oGrid.DataTable.GetValue(20, dtRow));
                GridExtensions.ColorFontCellGrid(oGrid, 13 + 1, dtRow + 1, ColorF, ColorB);
                GridExtensions.ColorFontCellGrid(oGrid, 14 + 1, dtRow + 1, ColorF, ColorB);
                if (NuevaL.Trim() != "Y") oGrid.DataTable.SetValue(19, dtRow, "");
            }

            if (PreAct != PreIni && ComAct != ComIni && NuevaL.Trim() != "Y")
            {
                oGrid.DataTable.SetValue(19, dtRow, "MA");
            }
        }
        private static void Calcular_Totales_Pagos()
        {
            try
            {
                oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_47").Specific;
                SAPbouiCOM.DataTable DT_PAG = oForm.DataSources.DataTables.Item("DT_PAG");
                SAPbouiCOM.DataTable DT_TOTPAG = oForm.DataSources.DataTables.Item("DT_TOTPAG");
                SAPbouiCOM.DataTable DT_SQL = oForm.DataSources.DataTables.Item("DT_SQL");
                //SAPbouiCOM.DataTable DT_SQL = oForm.DataSources.DataTables.Item("DT_SQL");

                Double TotalPagos = 0.0;
                Double TotalNOAplican = 0.0;
                Double TotalSinPCV = 0.0;
                Double TotalPagComi = 0.0;
                Int32 nProyectos = 0;

                String sql = "select COUNT( DISTINCT Proyecto) as nProy from [@Z_COMI_PAGOS_TEMP] where ISNULL(Proyecto,'') != '' ";
                DT_SQL.ExecuteQuery(sql);
                if (!DT_SQL.IsEmpty) nProyectos = Convert.ToInt32(DT_SQL.GetValue("nProy", 0));
                

                for (int i = 0; i < DT_PAG.Rows.Count; i++)
                {
                    TotalPagos += Convert.ToDouble(DT_PAG.GetValue("Pago Total", i));
                    TotalNOAplican += DT_PAG.GetValue("Tipo Pago", i).ToString().Trim() == "NOCOMI" ? Convert.ToDouble(DT_PAG.GetValue("Pago Total", i)) : 0;
                    TotalSinPCV += DT_PAG.GetValue("Tipo Pago", i).ToString().Trim() == "NOPCV" ? Convert.ToDouble(DT_PAG.GetValue("Pago Total", i)) : 0;
                    TotalPagComi += DT_PAG.GetValue("Tipo Pago", i).ToString().Trim() == "" ? Convert.ToDouble(DT_PAG.GetValue("Pago Total", i)) : 0;
                    //sProyecto = DT_PAG.GetValue("Proyecto", i).ToString().Trim();
                    //if (i > 0 )
                    //    nProyectos += sProyecto != DT_PAG.GetValue("Proyecto", i - 1).ToString().Trim() && sProyecto.Length > 0 ? 1 : 0;
                    //else
                    //    nProyectos += sProyecto.Length > 0 ? 1 : 0; 
                }

                try
                {
                    DT_TOTPAG.Columns.Add("Cant. Proyectos", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                    DT_TOTPAG.Columns.Add("Total Pagos", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                    DT_TOTPAG.Columns.Add("Pagos No aplicables a Comision", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                    DT_TOTPAG.Columns.Add("Pagos sin PCV Generada", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                    DT_TOTPAG.Columns.Add("Total Pagos para Comision", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                    DT_TOTPAG.Rows.Add();
                }
                catch (Exception) { }

                DT_TOTPAG.SetValue("Cant. Proyectos", 0, string.Format("{0:N0}",nProyectos));
                DT_TOTPAG.SetValue("Total Pagos", 0, string.Format("{0:C0}", TotalPagos));
                DT_TOTPAG.SetValue("Pagos No aplicables a Comision", 0, string.Format("{0:C0}", TotalNOAplican));
                DT_TOTPAG.SetValue("Pagos sin PCV Generada", 0, string.Format("{0:C0}", TotalSinPCV));
                DT_TOTPAG.SetValue("Total Pagos para Comision", 0, string.Format("{0:C0}", TotalPagComi));


                oGrid.CommonSetting.SetRowBackColor(1, Commons.Color_RGB_SAP(255, 255, 255)); ;
                for (int i = 0; i < oGrid.Columns.Count; i++)
                {
                    oGrid.Columns.Item(i).RightJustified = true;
                    oGrid.Columns.Item(i).Width = 150;
                }

                SAPbouiCOM.RowHeaders oHeader = oGrid.RowHeaders;
                oHeader.Width = 40;
                oHeader.SetText(0,"Totales");

                //oGrid.AutoResizeColumns();

            }
            catch (Exception){}
            

        }


        private SAPbouiCOM.Folder Folder0;
        private SAPbouiCOM.Folder Folder1;
        private SAPbouiCOM.Folder Folder2;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.DataTable DT_SQL;
        private SAPbouiCOM.DataTable DT_SQL2;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Grid Grid1;
        private SAPbouiCOM.ComboBox ComboBox1;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.Grid Grid2;
        private SAPbouiCOM.ComboBox ComboBox2;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.ComboBox ComboBox3;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.Folder Folder3;
        private SAPbouiCOM.Folder Folder4;
        private SAPbouiCOM.ComboBox ComboBox4;
        private SAPbouiCOM.ComboBox ComboBox5;
        private SAPbouiCOM.Grid Grid3;
        private SAPbouiCOM.StaticText StaticText9;
        private SAPbouiCOM.Folder Folder5;
        private SAPbouiCOM.Folder Folder6;
        private SAPbouiCOM.Folder Folder7;
        private SAPbouiCOM.Folder Folder8;
        private SAPbouiCOM.Grid Grid4;
        private SAPbouiCOM.Grid Grid5;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.StaticText StaticText7;
        private SAPbouiCOM.StaticText StaticText8;
        private SAPbouiCOM.ComboBox ComboBox6;
        private SAPbouiCOM.Grid Grid6;
        private SAPbouiCOM.Grid Grid7;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.StaticText StaticText10;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.StaticText StaticText11;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.StaticText StaticText12;
        private SAPbouiCOM.Button Button3;
        private SAPbouiCOM.Button Button4;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Grid Grid8;
        private SAPbouiCOM.StaticText StaticText14;
        private SAPbouiCOM.StaticText StaticText15;
        private SAPbouiCOM.CheckBox CheckBox0;
        private SAPbouiCOM.Button Button5;
        private SAPbouiCOM.Button Button6;
        private SAPbouiCOM.EditText EditText6;

        
    }
}