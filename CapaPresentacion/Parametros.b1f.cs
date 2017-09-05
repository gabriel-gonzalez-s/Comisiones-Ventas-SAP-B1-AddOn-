using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using SAPbobsCOM;
using System.Globalization;
using FuncionalidadesSDKB1;

namespace ComisionesVentas
{
    [FormAttribute("ComisionesVentas.Parametros", "CapaPresentacion/Parametros.b1f")]
    class Parametros : UserFormBase
    {
        private static SAPbouiCOM.Form oForm = null;
        private static SAPbouiCOM.DataTable oDTTable = null;
        private static SAPbouiCOM.EditText oEditText = null;
        private static SAPbouiCOM.Matrix oMatrix = null;
        private static string ItemActiveMenu = "";
        private static bool cCreacion = false;
        private static bool cModificacion = false;

        public Parametros()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Folder2 = ((SAPbouiCOM.Folder)(this.GetItem("Item_99").Specific));
            this.Folder3 = ((SAPbouiCOM.Folder)(this.GetItem("Item_100").Specific));
            this.StaticText30 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_76").Specific));
            this.StaticText31 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_78").Specific));
            this.EditText31 = ((SAPbouiCOM.EditText)(this.GetItem("Item_79").Specific));
            this.StaticText32 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_80").Specific));
            this.EditText32 = ((SAPbouiCOM.EditText)(this.GetItem("Item_81").Specific));
            this.StaticText33 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_82").Specific));
            this.EditText33 = ((SAPbouiCOM.EditText)(this.GetItem("Item_83").Specific));
            this.StaticText34 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_84").Specific));
            this.EditText34 = ((SAPbouiCOM.EditText)(this.GetItem("Item_85").Specific));
            this.EditText34.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.EditText34_LostFocusAfter);
            this.StaticText35 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_86").Specific));
            this.EditText35 = ((SAPbouiCOM.EditText)(this.GetItem("Item_87").Specific));
            this.EditText35.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.EditText35_LostFocusAfter);
            this.StaticText36 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_88").Specific));
            this.EditText36 = ((SAPbouiCOM.EditText)(this.GetItem("Item_89").Specific));
            this.StaticText37 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_90").Specific));
            this.EditText37 = ((SAPbouiCOM.EditText)(this.GetItem("Item_91").Specific));
            this.Button6 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button6.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button6_ClickBefore);
            this.Button6.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button6_ClickAfter);
            this.Button6.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button6_PressedAfter);
            this.Button7 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.Matrix4 = ((SAPbouiCOM.Matrix)(this.GetItem("0_U_G").Specific));
            this.Matrix4.MatrixLoadAfter += new SAPbouiCOM._IMatrixEvents_MatrixLoadAfterEventHandler(this.Matrix4_MatrixLoadAfter);
            this.Matrix4.GotFocusAfter += new SAPbouiCOM._IMatrixEvents_GotFocusAfterEventHandler(this.Matrix4_GotFocusAfter);
            this.Matrix4.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix4_ClickAfter);
            this.EditText38 = ((SAPbouiCOM.EditText)(this.GetItem("Item_95").Specific));
            this.StaticText38 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_96").Specific));
            this.EditText39 = ((SAPbouiCOM.EditText)(this.GetItem("Item_97").Specific));
            this.StaticText39 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_98").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_1").Specific));
            this.Button1.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button1_ClickAfter);
            this.PictureBox0 = ((SAPbouiCOM.PictureBox)(this.GetItem("Item_2").Specific));
            this.PictureBox0.ClickAfter += new SAPbouiCOM._IPictureBoxEvents_ClickAfterEventHandler(this.PictureBox0_ClickAfter);
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_0").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("Item_3").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.CheckBox0 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_7").Specific));
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("Item_9").Specific));
            this.CheckBox1 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_10").Specific));
            this.Folder1 = ((SAPbouiCOM.Folder)(this.GetItem("Item_11").Specific));
            this.Folder1.ClickBefore += new SAPbouiCOM._IFolderEvents_ClickBeforeEventHandler(this.Folder1_ClickBefore);
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("1_U_G").Specific));
            this.Matrix0.MatrixLoadAfter += new SAPbouiCOM._IMatrixEvents_MatrixLoadAfterEventHandler(this.Matrix0_MatrixLoadAfter);
            this.Matrix0.ComboSelectAfter += new SAPbouiCOM._IMatrixEvents_ComboSelectAfterEventHandler(this.Matrix0_ComboSelectAfter);
            this.Matrix0.GotFocusAfter += new SAPbouiCOM._IMatrixEvents_GotFocusAfterEventHandler(this.Matrix0_GotFocusAfter);
            this.Matrix0.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix0_ClickAfter);
            this.Matrix1 = ((SAPbouiCOM.Matrix)(this.GetItem("2_U_G").Specific));
            this.Matrix1.MatrixLoadAfter += new SAPbouiCOM._IMatrixEvents_MatrixLoadAfterEventHandler(this.Matrix1_MatrixLoadAfter);
            this.Matrix1.ClickBefore += new SAPbouiCOM._IMatrixEvents_ClickBeforeEventHandler(this.Matrix1_ClickBefore);
            this.Matrix1.GotFocusAfter += new SAPbouiCOM._IMatrixEvents_GotFocusAfterEventHandler(this.Matrix1_GotFocusAfter);
            this.Matrix1.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix1_ClickAfter);
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_14").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_12").Specific));
            this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("Item_13").Specific));
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_15").Specific));
            this.Grid0.KeyDownAfter += new SAPbouiCOM._IGridEvents_KeyDownAfterEventHandler(this.Grid0_KeyDownAfter);
            this.Grid0.ClickBefore += new SAPbouiCOM._IGridEvents_ClickBeforeEventHandler(this.Grid0_ClickBefore);
            this.Grid0.MatrixLoadAfter += new SAPbouiCOM._IGridEvents_MatrixLoadAfterEventHandler(this.Grid0_MatrixLoadAfter);
            this.Grid0.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid0_ClickAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_16").Specific));
            this.Button2.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button2_ClickAfter);
            this.Folder4 = ((SAPbouiCOM.Folder)(this.GetItem("Item_17").Specific));
            this.Grid1 = ((SAPbouiCOM.Grid)(this.GetItem("Item_18").Specific));
            this.Grid1.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid1_ClickAfter);
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_19").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("Item_20").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_21").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            //    Application.SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(form_MenuEvent);
            this.VisibleAfter += new SAPbouiCOM.Framework.FormBase.VisibleAfterHandler(this.Form_VisibleAfter);
            this.DataLoadAfter += new SAPbouiCOM.Framework.FormBase.DataLoadAfterHandler(this.Form_DataLoadAfter);
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);

        }
        private void Form_VisibleAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                //this.UIAPIRawForm.EnableMenu("1291", true);
                Application.SBO_Application.ActivateMenuItem("1291");
            }
            catch //(Exception)
            {
            }

        }
        private void OnCustomInitialize()
        {

            oForm = Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
            oForm.DataSources.DataTables.Add("DT_SQLPR");
            oForm.DataSources.DataTables.Add("DT_GRID");
            oForm.DataSources.DataTables.Add("DT_GRID2");

            Formatear_Items();
            Folder2.Select();
        }
        private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            UserFormBase activeForm = new GrupoComi();
            activeForm.Show();

        }
        private void Button1_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            CapaPresentacion.Feriados FormFeriados = new CapaPresentacion.Feriados();
            FormFeriados.Show();
            Application.SBO_Application.Forms.Item(FormFeriados.UIAPIRawForm.UniqueID).Select();

        }
        private void Button2_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (Button2.Item.Enabled && oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    Guardar_DT_GRID();
                    ((SAPbouiCOM.Button)oForm.Items.Item("1").Specific).Item.Click();
                    //oForm.Refresh();
                }
            }
            catch (Exception) { }

        }
        private void PictureBox0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {

            SAPbouiCOM.DBDataSource source = oForm.DataSources.DBDataSources.Item("@Z_COMI_PER");
            Asignar_Domingos_Feriados(Convert.ToString(source.GetValue("U_Codigo", 0)), Convert.ToInt32(source.GetValue("U_Inicio", 0)), Convert.ToInt32(source.GetValue("U_Fin", 0)));

        }
        private void Button6_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (cCreacion)
            {
                Application.SBO_Application.ActivateMenuItem("1291");
                cCreacion = false;
            }
            if (cModificacion)
            {
                Cargar_Combo_Grupo(Matrix0);
                cModificacion = false;
            }
        }
        private void Button6_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (Button6.Caption == "Crear")
                cCreacion = true;
            if (Button6.Caption == "Actualizar")
                cModificacion = true;

        }
        private void Matrix0_GotFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                oForm.EnableMenu("1292", true);
                oForm.EnableMenu("1293", true);
            }
            catch (Exception) { }
        }
        private void Matrix1_GotFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                oForm.EnableMenu("1292", true);
                oForm.EnableMenu("1293", true);
            }
            catch (Exception) { }
        }
        private void Matrix4_GotFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                oForm.EnableMenu("1292", true);
                oForm.EnableMenu("1293", true);
            }
            catch (Exception) { }
        }
        private void Matrix0_MatrixLoadAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Matrix0.AutoResizeColumns();
            if (Matrix0.RowCount >= 1)
            {
                MatrixExtensions.SelectMatrixRow(ref Matrix0, 1);
            }
        }
        private void Matrix1_MatrixLoadAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Matrix1.AutoResizeColumns();
            if (Matrix1.RowCount >= 1)
            {
                MatrixExtensions.SelectMatrixRow(ref Matrix1, 1);
                Cargar_Grid_Comisiones();
            }
        }
        private void Matrix4_MatrixLoadAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Matrix4.AutoResizeColumns();
        }
        private void Matrix0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                oForm.EnableMenu("1292", true);
                oForm.EnableMenu("1293", true);
                MatrixExtensions.SelectMatrixRow(ref Matrix0, pVal.Row);
                Cargar_Grid_Comisiones_Vend();
            }
            catch (Exception) { }
        }
        private void Matrix1_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            MatrixExtensions.SelectMatrixRow(ref Matrix1, pVal.Row);
            Cargar_Grid_Comisiones();
        }
        private void Matrix4_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            MatrixExtensions.SelectMatrixRow(ref Matrix4, pVal.Row);
        }
        private void Matrix1_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                oForm.EnableMenu("1292", true);
                oForm.EnableMenu("1293", true);
            }
            catch (Exception) { }
        }
        private void Matrix0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.ComboBox oComboBox = ComboBoxExtensions.GetComboBoxMatrix(Matrix0, "C_0_3", pVal.Row);
                string sVal = oComboBox.Selected.Description;
                MatrixExtensions.SetCellValue(Matrix0, "C_0_6", pVal.Row, sVal);

            }
            catch (Exception) { }
        }
        private void Grid0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                oForm.EnableMenu("1292", true);
                oForm.EnableMenu("1293", true);
            }
            catch (Exception) { }
        }
        private void Grid0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                Grid0.Rows.SelectedRows.Add(pVal.Row);
                oForm.EnableMenu("1292", true);
                oForm.EnableMenu("1293", true);
            }
            catch (Exception) { }
        }
        private void Grid0_MatrixLoadAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                Grid0.Rows.SelectedRows.Add(0);
            }
            catch (Exception) { }
        }
        private void Grid0_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                Button2.Item.Enabled = true;
        }
        private void Grid1_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                Grid1.Rows.SelectedRows.Add(pVal.Row);
            }
            catch (Exception) { }
        }
        private void Folder1_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.DataTable DT_GRID = oForm.DataSources.DataTables.Item("DT_GRID2");
            if (DT_GRID.IsEmpty)
                Cargar_Grid_Comisiones_Vend();
            //Cargar_Combo_Grupo(Matrix0);
        }
        private void EditText34_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            DateTime Desde = Convert.ToDateTime(EditText34.Value.Trim() + '/' + EditText32.Value.Substring(5, 2) + '/' + EditText32.Value.Substring(0, 4)).AddMonths(-1);
            StaticText0.Caption = Desde.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
            StaticText1.Caption = EditText35.Value.Trim() + '/' + EditText32.Value.Substring(5, 2) + '/' + EditText32.Value.Substring(0, 4);
        }
        private void EditText35_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            DateTime Desde = Convert.ToDateTime(EditText34.Value.Trim() + '/' + EditText32.Value.Substring(5, 2) + '/' + EditText32.Value.Substring(0, 4)).AddMonths(-1);
            StaticText0.Caption = Desde.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
            StaticText1.Caption = EditText35.Value.Trim() + '/' + EditText32.Value.Substring(5, 2) + '/' + EditText32.Value.Substring(0, 4);
        }
        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            try
            {
                DateTime Desde = Convert.ToDateTime(EditText34.Value.Trim() + '/' + EditText32.Value.Substring(5, 2) + '/' + EditText32.Value.Substring(0, 4)).AddMonths(-1);
                StaticText0.Caption = Desde.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                StaticText1.Caption = EditText35.Value.Trim() + '/' + EditText32.Value.Substring(5, 2) + '/' + EditText32.Value.Substring(0, 4);
                Cargar_Combo_Grupo(Matrix0);
            }
            catch (Exception)
            {
            }
        }
        public static void Parametros_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, SAPbobsCOM.Company oCompany, SAPbouiCOM.Application oApplication, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

            switch (pVal.BeforeAction)
            {
                case true:
                    if (pVal.ItemUID != "0_U_G" && pVal.ItemUID != "1_U_G" && pVal.ItemUID != "2_U_G")
                    {
                        oForm.EnableMenu("1292", false);
                        oForm.EnableMenu("1293", false);
                    }
                    break;
                case false:
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                    {
                        SAPbouiCOM.DataTable oDataTable = DataTableExtensions.GetDataTableFromCLF(pVal, oForm);

                        switch (pVal.ItemUID)
                        {
                            case "1_U_G":
                                try
                                {
                                    string sPreVal = "";
                                    string sCod = oDataTable.GetValue(0, 0).ToString();
                                    string sNom = oDataTable.GetValue(1, 0).ToString();
                                    string sMem = oDataTable.GetValue(2, 0).ToString();

                                    sNom = "(" + sMem.Trim() + ") - " + sNom.Trim();

                                    SAPbouiCOM.DBDataSource source = oForm.DataSources.DBDataSources.Item("@Z_COMI_VENDPER");
                                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("1_U_G").Specific;
                                    sPreVal = MatrixExtensions.GetCellValue(oMatrix, "C_0_1", pVal.Row);
                                    MatrixExtensions.SetCellValue(oMatrix, "C_0_1", pVal.Row, sCod);
                                    MatrixExtensions.SetCellValue(oMatrix, "C_0_4", pVal.Row, sNom);
                                    if (sPreVal != sCod)
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;

                                }
                                catch (Exception ) { }
                                break;
                            case "Item_13":
                                try
                                {
                                    string sCod1 = oDataTable.GetValue(13, 0).ToString();
                                    string sNom1 = oDataTable.GetValue(14, 0).ToString();
                                    string sCodeT = oDataTable.GetValue(0, 0).ToString();

                                    SAPbouiCOM.DBDataSource source1 = oForm.DataSources.DBDataSources.Item("@Z_COMI_GRUPO");
                                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("2_U_G").Specific;
                                    if (source1.GetValue("U_Codigo", source1.Size - 1).ToString().Trim().Length != 0)
                                        MatrixExtensions.AddLineMatrixDBDataSource(oMatrix, source1);
                                    oMatrix.FlushToDataSource();
                                    source1.SetValue("U_Codigo", source1.Size - 1, sCod1);
                                    source1.SetValue("U_Nombre", source1.Size - 1, sNom1);
                                    source1.SetValue("U_Code", source1.Size - 1, sCodeT);
                                    oMatrix.LoadFromDataSourceEx();
                                    MatrixExtensions.SelectMatrixRow(ref oMatrix, source1.Size);
                                    oForm.Select();
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("C_0_4").Cells.Item(source1.Size).Specific).Active = true;

                                    ((SAPbouiCOM.EditText)oForm.Items.Item("Item_13").Specific).Item.Visible=false;

                                }
                                catch (Exception ) { }
                                break;
                        }
                    }
                    break;
            }
        }
        public static void Parametros_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            oForm = Application.SBO_Application.Forms.ActiveForm;

            try
            {
                switch (pVal.MenuUID)
                {
                    case "1282": // Boton CREAR -------------------------------------------------------------------------------------------------------------------------------------
                        if (pVal.BeforeAction == false)
                        {
                            CargarNuevoRegistroParametros(oForm);
                        }
                        break;
                    case "1283"://Eliminar Registro
                        if (pVal.BeforeAction == true)
                        {
                            if (Application.SBO_Application.MessageBox("¿Eliminar Este Periodo?", 2, "Eliminar", "Cancelar") == 2)
                                BubbleEvent = false;
                        }
                        break;
                    case "1292"://Agregar Linea Matrix
                        if (pVal.BeforeAction == true)
                        {
                            switch (ItemActiveMenu)
                            {
                                case "0_U_G":
                                    MatrixExtensions.AddLineMatrixDBDataSource(((SAPbouiCOM.Matrix)oForm.Items.Item("0_U_G").Specific), oForm.DataSources.DBDataSources.Item("@Z_COMI_PAR"), "C_0_4");
                                    break;
                                case "1_U_G":
                                    MatrixExtensions.AddLineMatrixDBDataSource(((SAPbouiCOM.Matrix)oForm.Items.Item("1_U_G").Specific), oForm.DataSources.DBDataSources.Item("@Z_COMI_VENDPER"));
                                    break;
                                case "2_U_G":
                                    ((SAPbouiCOM.EditText)oForm.Items.Item("Item_13").Specific).Item.Visible = true;
                                    ((SAPbouiCOM.EditText)oForm.Items.Item("Item_13").Specific).Item.Click();
                                    Application.SBO_Application.Menus.Item("5377").Activate();
                                    break;
                                case "Item_15": //Grid
                                    SAPbouiCOM.DataTable DT_GRID = oForm.DataSources.DataTables.Item("DT_GRID");
                                    SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_15").Specific;
                                    Nuevo_Registro_Grid_Comisiones();
                                    oGrid.Rows.SelectedRows.Add(DT_GRID.Rows.Count - 1);
                                    ((SAPbouiCOM.Button)oForm.Items.Item("Item_16").Specific).Item.Enabled = true;
                                    break;
                            }
                        }
                        break;
                    case "1293"://Eliminar Linea Matrix
                        if (pVal.BeforeAction == true)
                        {
                            switch (ItemActiveMenu)
                            {
                                case "2_U_G":
                                    oMatrix = ((SAPbouiCOM.Matrix)oForm.Items.Item("2_U_G").Specific);
                                    SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_COMI_GRUPO");
                                    int nRow = (int)oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                    oMatrix.FlushToDataSource();
                                    oDBDataSource.RemoveRecord(nRow - 1);
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                    oMatrix.LoadFromDataSource();
                                    BubbleEvent = false;
                                    break;
                                case "Item_15":
                                    SAPbouiCOM.DataTable DT_GRID = oForm.DataSources.DataTables.Item("DT_GRID");
                                    if (!DT_GRID.IsEmpty)
                                    {
                                        try
                                        {
                                            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_15").Specific;

                                            nRow = oGrid.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);                                            //oGrid.DataTable.Rows.Remove(oGrid.GetDataTableRowIndex(nRow));
                                            DT_GRID.SetValue("LineId", nRow , -99);
                                            DT_GRID.SetValue("U_Codigo", nRow , -99);

                                            ((SAPbouiCOM.Button)oForm.Items.Item("Item_16").Specific).Item.Enabled = true;
                                        }
                                        catch (Exception) { }
                                    }
                                    break;
                            }
                        }
                        break;
                }
            }
            catch (Exception) { }
        }
        public static void Parametros_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            ItemActiveMenu = eventInfo.ItemUID;

            switch (eventInfo.BeforeAction)
            {
                case true:
                    if (eventInfo.ItemUID == "Item_15")
                    {
                        SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_15").Specific;
                        oGrid.Rows.SelectedRows.Add(eventInfo.Row);
                    }
                    break;
            }
           
        }
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------
        private static void CargarNuevoRegistroParametros(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.DBDataSource source = oForm.DataSources.DBDataSources.Item("@Z_COMI_PER");

                //Consultar y cargar proximos datos encabezado
                string Sql = @"SELECT [Code] + 1 as Code
                                            ,[Name]
                                            ,[DocEntry] + 1 as [DocEntry]
                                            ,[Canceled]
                                            ,[Object]
                                            ,[LogInst]
                                            ,[UserSign]
                                            ,[Transfered]
                                            ,[CreateDate]
                                            ,[CreateTime]
                                            ,[UpdateDate]
                                            ,[UpdateTime]
                                            ,[DataSource]
                                            ,IIF(RTRIM(RIGHT([U_Codigo],2)) = '12',[U_Codigo] + 89,[U_Codigo] + 1)  as [U_Codigo]
                                            ,IIF(RTRIM(RIGHT([U_Nombre],2)) = '12' ,STR(LEFT([U_Nombre],4)+1)+'-'+RIGHT('00' + RTRIM(RIGHT([U_Nombre],2)-11),2)
	                                                                                ,LEFT([U_Nombre],4)+'-'+RIGHT('00' + RTRIM(RIGHT([U_Nombre],2)+1),2)) as [U_Nombre]
                                            ,[U_Abierto]
                                            ,[U_Inicio]
                                            ,[U_Fin]
                                            ,[U_Valor_UF]
                                            ,[U_Valor_DMP]
                                            ,[U_Visible]
                                            ,[U_Status]
                                        FROM [dbo].[@Z_COMI_PER]
                                        WHERE code = (select max(CAST(code as INT)) from [SBO_COMERCIAL].[dbo].[@Z_COMI_PER])";

                oDTTable = oForm.DataSources.DataTables.Item("DT_SQLPR");
                oDTTable.ExecuteQuery(Sql);

                source.Clear();
                source.InsertRecord(source.Size);
                source.Offset = source.Size - 1;
                int Code = (int)oDTTable.GetValue("Code", 0);
                source.SetValue("Code", 0, Convert.ToString(oDTTable.GetValue("Code", 0)));
                source.SetValue("DocEntry", 0, Convert.ToString(oDTTable.GetValue("DocEntry", 0)));
                source.SetValue("U_Codigo", 0, Convert.ToString(oDTTable.GetValue("U_Codigo", 0)));
                source.SetValue("U_Nombre", 0, Convert.ToString(oDTTable.GetValue("U_Nombre", 0)));
                source.SetValue("U_Abierto", 0, "A");
                source.SetValue("U_Inicio", 0, Convert.ToString(oDTTable.GetValue("U_Inicio", 0)));
                source.SetValue("U_Fin", 0, Convert.ToString(oDTTable.GetValue("U_Fin", 0)));
                source.SetValue("U_Valor_UF", 0, Commons.GetStringFromDouble((double)oDTTable.GetValue("U_Valor_UF", 0)));
                source.SetValue("U_Valor_DMP", 0, Commons.GetStringFromDouble((double)oDTTable.GetValue("U_Valor_DMP", 0)));
                source.SetValue("U_Visible", 0, "Y");
                source.SetValue("U_Status", 0, "N");

                DateTime Desde = Convert.ToDateTime(oDTTable.GetValue("U_Inicio", 0).ToString().Trim() + '/' + oDTTable.GetValue("U_Nombre", 0).ToString().Substring(5, 2) + '/' + oDTTable.GetValue("U_Nombre", 0).ToString().Substring(0, 4)).AddMonths(-1);
                ((SAPbouiCOM.StaticText)oForm.Items.Item("Item_4").Specific).Caption = Desde.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                ((SAPbouiCOM.StaticText)oForm.Items.Item("Item_5").Specific).Caption = oDTTable.GetValue("U_Fin", 0).ToString().Trim() + '/' + oDTTable.GetValue("U_Nombre", 0).ToString().Substring(5, 2) + '/' + oDTTable.GetValue("U_Nombre", 0).ToString().Substring(0, 4);

                //Domingos y Feriados
                Asignar_Domingos_Feriados(Convert.ToString(oDTTable.GetValue("U_Codigo", 0)), Convert.ToInt32(oDTTable.GetValue("U_Inicio", 0)), Convert.ToInt32(oDTTable.GetValue("U_Fin", 0)));
                                
                //Consultar y cargar proximos datos Detalles Rango Comisiones
                source = oForm.DataSources.DBDataSources.Item("@Z_COMI_PAR");
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("0_U_G").Specific;
                oMatrix.FlushToDataSource();
                source.Clear();
                oMatrix.Clear();

                Sql = @"SELECT [Code] + 1 as Code
                                    ,[LineId] 
                                    ,[Object]
                                    ,[LogInst]
                                    ,[U_Codigo]
                                    ,[U_Periodo]
                                    , [U_Nombre]
                                    ,[U_Descripcion]
                                    ,[U_Correlativo]
                                    ,[U_Desde]
                                    ,[U_Hasta]
                                    ,[U_Comision]
                                    ,[U_Usuario_crea]
                                    ,[U_Fecha_crea]
                                    ,[U_Grupo]
                                FROM [SBO_COMERCIAL].[dbo].[@Z_COMI_PAR] 
                                WHERE code = (select max(CAST(code as INT)) from [SBO_COMERCIAL].[dbo].[@Z_COMI_PAR])
                                ORDER BY [LineId] ";

                oDTTable.ExecuteQuery(Sql);

                for (int i = 0; i < oDTTable.Rows.Count; i++)
                {
                    source.InsertRecord(source.Size);
                    source.Offset = source.Size - 1;
                    source.SetValue("U_Codigo", source.Size - 1, Convert.ToString(oDTTable.GetValue("U_Codigo", i)));
                    source.SetValue("U_Periodo", source.Size - 1, Convert.ToString(oDTTable.GetValue("U_Periodo", i)));
                    source.SetValue("U_Nombre", source.Size - 1, Convert.ToString(oDTTable.GetValue("U_Nombre", i)));
                    source.SetValue("U_Descripcion", source.Size - 1, Convert.ToString(oDTTable.GetValue("U_Descripcion", i)));
                    source.SetValue("U_Correlativo", source.Size - 1, Convert.ToString(oDTTable.GetValue("U_Correlativo", i)));
                    source.SetValue("U_Desde", source.Size - 1, Commons.GetStringFromDouble((double)oDTTable.GetValue("U_Desde", i)));
                    source.SetValue("U_Hasta", source.Size - 1, Commons.GetStringFromDouble((double)oDTTable.GetValue("U_Hasta", i)));
                    source.SetValue("U_Comision", source.Size - 1, Commons.GetStringFromDouble((double)oDTTable.GetValue("U_Comision", i)));
                    source.SetValue("U_Usuario_crea", source.Size - 1, Convert.ToString(Program.sCodUsuActual));
                    source.SetValue("U_Fecha_crea", source.Size - 1, DateTime.Now.ToString("yyyyMMdd"));
                    source.SetValue("U_Grupo", source.Size - 1, Convert.ToString(oDTTable.GetValue("U_Grupo", i)));
                }
                oMatrix.LoadFromDataSource();

                //Consultar y cargar proximos datos Grupos
                source = oForm.DataSources.DBDataSources.Item("@Z_COMI_GRUPO");
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("2_U_G").Specific;
                oMatrix.FlushToDataSource();
                source.Clear();
                oMatrix.Clear();

                Sql = @"SELECT [Code] + 1 as Code
                                    ,[LineId]
                                    ,[Object]
                                    ,[LogInst]
                                    ,[U_Codigo]
                                    ,[U_Periodo]
                                    ,[U_Nombre]
                                    ,[U_Descripcion]
                                    ,[U_Code]
                                FROM [dbo].[@Z_COMI_GRUPO]
                                WHERE code = (select max(CAST(code as INT)) from [SBO_COMERCIAL].[dbo].[@Z_COMI_PAR])
                                ORDER BY [LineId] ";

                oDTTable.ExecuteQuery(Sql);

                for (int i = 0; i < oDTTable.Rows.Count; i++)
                {
                    source.InsertRecord(source.Size);
                    source.Offset = source.Size - 1;
                    source.SetValue("U_Codigo", source.Size - 1, Convert.ToString(oDTTable.GetValue("U_Codigo", i)));
                    source.SetValue("U_Periodo", source.Size - 1, Convert.ToString(oDTTable.GetValue("U_Periodo", i)));
                    source.SetValue("U_Nombre", source.Size - 1, Convert.ToString(oDTTable.GetValue("U_Nombre", i)));
                    source.SetValue("U_Descripcion", source.Size - 1, Convert.ToString(oDTTable.GetValue("U_Descripcion", i)));
                    source.SetValue("U_Code", source.Size - 1, Convert.ToString(oDTTable.GetValue("U_Code", i)));

                }
                oMatrix.LoadFromDataSource();
                Cargar_Grid_Comisiones();

                //Consultar y cargar proximos datos Vendedores
                source = oForm.DataSources.DBDataSources.Item("@Z_COMI_VENDPER");
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("1_U_G").Specific;
                oMatrix.FlushToDataSource();
                source.Clear();
                oMatrix.Clear();

                Sql = @"SELECT [Code] + 1 as Code
                                    ,[LineId]
                                    ,[Object]
                                    ,[LogInst]
                                    ,[U_Codigo]
                                    ,[U_Periodo]
                                    ,[U_Grupo]
                                    ,[U_Nombre]
                                    ,[U_Comentario]
                                    ,[U_Nombre_Grupo]
                                FROM [SBO_COMERCIAL].[dbo].[@Z_COMI_VENDPER]
                                WHERE code = (select max(CAST(code as INT)) from [SBO_COMERCIAL].[dbo].[@Z_COMI_PAR])
                                ORDER BY [LineId] ";

                oDTTable.ExecuteQuery(Sql);

                for (int i = 0; i < oDTTable.Rows.Count; i++)
                {
                    source.InsertRecord(source.Size);
                    source.Offset = source.Size - 1;
                    source.SetValue("U_Codigo", source.Size - 1, Convert.ToString(oDTTable.GetValue("U_Codigo", i)));
                    source.SetValue("U_Periodo", source.Size - 1, Convert.ToString(oDTTable.GetValue("U_Periodo", i)));
                    source.SetValue("U_Grupo", source.Size - 1, Convert.ToString(oDTTable.GetValue("U_Grupo", i)));
                    source.SetValue("U_Nombre", source.Size - 1, Convert.ToString(oDTTable.GetValue("U_Nombre", i)));
                    source.SetValue("U_Comentario", source.Size - 1, Convert.ToString(oDTTable.GetValue("U_Comentario", i)));
                    source.SetValue("U_Nombre_Grupo", source.Size - 1, Convert.ToString(oDTTable.GetValue("U_Nombre_Grupo", i)));

                }
                oMatrix.LoadFromDataSource();

                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("2_U_G").Specific;
                MatrixExtensions.SelectMatrixRow(ref oMatrix, 1);

            }
            catch (Exception) { }
            finally
            {
                oForm.Freeze(false);
            }
        }
        private static void Asignar_Domingos_Feriados(string U_Periodo, int U_Desde, int U_Hasta)
        {
            string Sql = "exec [dbo].[Min_Comi_Domingos_Feriados] '" + U_Periodo.Trim() + "'," + U_Desde + "," + U_Hasta;

            try
            {
                oDTTable = oForm.DataSources.DataTables.Item("DT_SQLPR");
                oDTTable.ExecuteQuery(Sql);

                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("Item_95").Specific;
                oEditText.Value = Convert.ToString(oDTTable.GetValue("Per_Domingos", 0));
                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("Item_97").Specific;
                oEditText.Value = Convert.ToString(oDTTable.GetValue("Per_Feriados", 0));
            }
            catch //(Exception)
            {
            }

        }
        private void Formatear_Items()
        {
            Grid0.DataTable = oForm.DataSources.DataTables.Item("DT_GRID");
            Grid1.DataTable = oForm.DataSources.DataTables.Item("DT_GRID2");

            Matrix0.Columns.Item("C_0_2").Visible = false;
            Matrix1.Columns.Item("C_0_2").Visible = false;
            Matrix1.Columns.Item("C_0_5").Visible = false;
            Matrix1.Columns.Item("Col_0").Visible = false;
            Matrix1.Columns.Item("Col_1").Visible = false;
            //Matrix4.Columns.Item("C_0_1").Visible = false;
            //Matrix4.Columns.Item("C_0_2").Visible = false;
            //Matrix4.Columns.Item("C_0_3").Visible = false;

            StaticText30.Item.Visible = false;
            StaticText31.Item.Visible = false;
            StaticText33.Item.Visible = false;

            EditText0.Item.Visible = false;
            EditText1.Item.Visible = false;
            EditText31.Item.Visible = false;
            EditText33.Item.Visible = false;
            EditText3.Item.Visible = false;

            PictureBox0.Picture = @"\\fssapbo\SAPB1\Anexos\images\refresh yellow(2).png";
            CheckBox0.Item.Width = 56;
            CheckBox1.Item.Width = 60;

            //Cargar_Combo_Grupo(Matrix0);

            //SAPbouiCOM.Column sboCol = (SAPbouiCOM.Column)Matrix0.Columns.Item("C_0_3");
            //string sql = "SELECT Code, U_Codigo + ' - ' + U_Nombre FROM [SBO_COMERCIAL].[dbo].[@Z_COMI_COMGRP]";
            //ComboBoxExtensions.LoadComboQueryDataTable(sboCol, oForm.DataSources.DataTables.Item("DT_SQLPR"), sql, 0, 1, false);

        }
        private static void Cargar_Combo_Grupo(SAPbouiCOM.Matrix oMatrixp)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("2_U_G").Specific;
            string sGrupo = (oForm.DataSources.DBDataSources.Item("@Z_COMI_PER")).GetValue("Code", 0).ToString().Trim();

            SAPbouiCOM.Column sboCol = (SAPbouiCOM.Column)oMatrixp.Columns.Item("C_0_3");
            ComboBoxExtensions.CleanComboBox(sboCol);
            string sql = "SELECT LineId as Code, U_Codigo + ' - ' + U_Nombre from [@Z_COMI_GRUPO] WHERE Code = " + sGrupo;//"SELECT Code, U_Codigo + ' - ' + U_Nombre FROM [SBO_COMERCIAL].[dbo].[@Z_COMI_COMGRP]";
            ComboBoxExtensions.LoadComboQueryDataTable(sboCol, oForm.DataSources.DataTables.Item("DT_SQLPR"), sql, 0, 1, false);
        }
        private static void Cargar_Grid_Comisiones()
        {
            try
            {
                oForm = Application.SBO_Application.Forms.ActiveForm;
                oForm.Freeze(true);
                SAPbouiCOM.DataTable DT_GRID = oForm.DataSources.DataTables.Item("DT_GRID");
                SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_COMI_PAR");
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("2_U_G").Specific;
                int nRow = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                string sGrupo = nRow == -1 ? (oForm.DataSources.DBDataSources.Item("@Z_COMI_GRUPO")).GetValue("U_Code", 0)
                                           : ((SAPbouiCOM.EditText)oMatrix.Columns.Item("C_0_5").Cells.Item(nRow).Specific).Value;
                var Value = (string)null;
                string sfield;

                string sql = "SELECT * FROM [SBO_COMERCIAL].[dbo].[@Z_COMI_PAR] WHERE CODE = -99 ";
                DT_GRID.ExecuteQuery(sql);

                DT_GRID.Rows.Remove(0);

                for (int i = 0; i <= oDBDataSource.Size - 1; i++)
                {
                    string sVal = oDBDataSource.GetValue("U_Grupo", i);
                    if (sVal.Trim() == sGrupo.Trim())
                    {
                        DT_GRID.Rows.Add();
                        for (int x = 0; x <= oDBDataSource.Fields.Count - 1; x++)
                        {
                            try
                            {
                                sfield = oDBDataSource.Fields.Item(x).Name;
                                Value = oDBDataSource.GetValue(sfield, i);
                                Value = Value.Trim().Length == 0 ? null : Value;
                                DT_GRID.SetValue(sfield, DT_GRID.Rows.Count - 1, Value);
                            }
                            catch (Exception) { }
                        }
                    }
                }
                Formatear_Grid_Comisiones((SAPbouiCOM.Grid)oForm.Items.Item("Item_15").Specific);
                try { } //DT_GRID.Rows.Remove(DT_GRID.Rows.Count - 1); }
                catch (Exception) { }
                ((SAPbouiCOM.Grid)oForm.Items.Item("Item_15").Specific).Rows.SelectedRows.Add(0);
            }
            catch (Exception) { }
            finally { oForm.Freeze(false); }
        }
        private static void Cargar_Grid_Comisiones_Vend()
        {
            try
            {
                oForm = Application.SBO_Application.Forms.ActiveForm;
                oForm.Freeze(true);
                SAPbouiCOM.DataTable DT_GRID = oForm.DataSources.DataTables.Item("DT_GRID2");
                SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_COMI_PAR");
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("1_U_G").Specific;
                int nRow = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                string sGrupo = nRow == -1 ? (oForm.DataSources.DBDataSources.Item("@Z_COMI_GRUPO")).GetValue("U_Code", 0)
                                           : ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("C_0_3").Cells.Item(nRow).Specific).Value;
                var Value = (string)null;
                string sfield;

                string sql = "SELECT * FROM [SBO_COMERCIAL].[dbo].[@Z_COMI_PAR] WHERE CODE = -99 ";
                DT_GRID.ExecuteQuery(sql);

                DT_GRID.Rows.Remove(0);

                for (int i = 0; i <= oDBDataSource.Size - 1; i++)
                {
                    string sVal = oDBDataSource.GetValue("U_Grupo", i);
                    if (sVal.Trim() == sGrupo.Trim())
                    {
                        DT_GRID.Rows.Add();
                        for (int x = 0; x <= oDBDataSource.Fields.Count - 1; x++)
                        {
                            try
                            {
                                sfield = oDBDataSource.Fields.Item(x).Name;
                                Value = oDBDataSource.GetValue(sfield, i);
                                Value = Value.Trim().Length == 0 ? null : Value;
                                DT_GRID.SetValue(sfield, DT_GRID.Rows.Count - 1, Value);
                            }
                            catch (Exception) { }
                        }
                    }
                }
                Formatear_Grid_Comisiones((SAPbouiCOM.Grid)oForm.Items.Item("Item_18").Specific);
                try { } //DT_GRID.Rows.Remove(DT_GRID.Rows.Count - 1); }
                catch (Exception) { }
                ((SAPbouiCOM.Grid)oForm.Items.Item("Item_18").Specific).Rows.SelectedRows.Add(0);
            }
            catch (Exception) { }
            finally { oForm.Freeze(false); }
        }
        private static void Formatear_Grid_Comisiones(SAPbouiCOM.Grid oGrid)
        {
            List<int> ColumnasVisibles = new List<int>(new int[] { 9, 10, 11 });
            List<int> ColumnasJustificadas = new List<int>(new int[] { 9, 10, 11 });

            //SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_15").Specific;

            oGrid.Columns.Item(9).TitleObject.Caption = "Valor Desde UF";
            oGrid.Columns.Item(10).TitleObject.Caption = "Valor Hasta UF";
            oGrid.Columns.Item(11).TitleObject.Caption = "Valor Comision";

            for (int iCols = 0; iCols <= oGrid.Columns.Count - 1; iCols++)
            {
                if (!ColumnasVisibles.Contains(iCols))
                {
                    oGrid.Columns.Item(iCols).Visible = false;
                }
                if (ColumnasJustificadas.Contains(iCols))
                {
                    oGrid.Columns.Item(iCols).RightJustified = true;
                    if (oGrid.Item.UniqueID == "Item_18") 
                        oGrid.Columns.Item(iCols).Editable = false;
                }
            }
            oGrid.AutoResizeColumns();

        }
        private static void Nuevo_Registro_Grid_Comisiones()
        {
            try
            {
                oForm.Freeze(true);
                SAPbouiCOM.DataTable DT_GRID = oForm.DataSources.DataTables.Item("DT_GRID");
                int nDataRow = DT_GRID.Rows.Count - 1;

                if (DT_GRID.Rows.Count > 0)
                {
                    DT_GRID.Rows.Add();
                    for (int i = 0; i <= DT_GRID.Columns.Count - 1; i++)
                    {
                        try
                        {
                            var Valor = DT_GRID.GetValue(DT_GRID.Columns.Item(i).Name, nDataRow);
                            DT_GRID.SetValue(DT_GRID.Columns.Item(i).Name, DT_GRID.Rows.Count - 1, Valor);
                        }
                        catch (Exception) { }
                    }
                }
                else
                {
                    SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_COMI_PAR");
                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("2_U_G").Specific;
                    int nRow = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                    string sGrupo = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("C_0_5").Cells.Item(nRow).Specific).Value;

                    DT_GRID.Rows.Add();

                    for (int i = 0; i <= DT_GRID.Columns.Count - 1; i++)
                    {
                        try
                        {
                            var Valor = oDBDataSource.GetValue(DT_GRID.Columns.Item(i).Name, oDBDataSource.Size - 1);
                            DT_GRID.SetValue(DT_GRID.Columns.Item(i).Name, DT_GRID.Rows.Count - 1, Valor);
                        }
                        catch (Exception) { }
                    }
                    DT_GRID.SetValue("U_Grupo", DT_GRID.Rows.Count - 1, sGrupo);
                }

                int nProx = ((int)DT_GRID.GetValue("LineId", DT_GRID.Rows.Count - 1)) + 1;
                DT_GRID.SetValue("LineId", DT_GRID.Rows.Count - 1, nProx);
                DT_GRID.SetValue("U_Codigo", DT_GRID.Rows.Count - 1, nProx);
                DT_GRID.SetValue("U_Desde", DT_GRID.Rows.Count - 1, 0);
                DT_GRID.SetValue("U_Hasta", DT_GRID.Rows.Count - 1, 0);
                DT_GRID.SetValue("U_Comision", DT_GRID.Rows.Count - 1, 0);
            }
            catch (Exception) { }
            finally {oForm.Freeze(false);}

        }
        private void Guardar_DT_GRID()
        {
            try
            {
                oForm = Application.SBO_Application.Forms.ActiveForm;
                oForm.Freeze(true);
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_15").Specific;
                SAPbouiCOM.DataTable DT_GRID = oForm.DataSources.DataTables.Item("DT_GRID");
                SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_COMI_PAR");
                string sGrupo = DT_GRID.GetValue("U_Grupo", 0).ToString();
                List<int> aDel = new List<int>();

                Matrix4.FlushToDataSource();
                for (int i = 0; i <= oDBDataSource.Size - 1; i++) //Ubicar Registros con el U_Grupo correspondiente
                {
                    string sVal = oDBDataSource.GetValue("U_Grupo", i);
                    if (sVal.Trim() == sGrupo.Trim())
                    {
                        oDBDataSource.RemoveRecord(i);
                        i--;
                    }
                }

                var Value = (string)null;
                string sfield;
                List<string> ValoresDouble = new List<string>(new string[] { "U_Desde", "U_Hasta", "U_Comision" });

                for (int i = 0; i <= oGrid.Rows.Count - 1; i++)
                {
                    if(Convert.ToInt32(DT_GRID.GetValue("LineId", oGrid.GetDataTableRowIndex(i))) != -99)
                    {
                        oDBDataSource.InsertRecord(oDBDataSource.Size);
                        for (int x = 0; x <= oGrid.Columns.Count - 1; x++)
                        {
                            try
                            {
                                sfield = DT_GRID.Columns.Item(x).Name;
                                Value = DT_GRID.GetValue(sfield, oGrid.GetDataTableRowIndex(i)).ToString();
                                Value = Value.Trim().Length == 0 ? null : Value;
                                if (ValoresDouble.Contains(sfield))
                                    Value = Value.Replace(",", ".");
                                oDBDataSource.SetValue(sfield, oDBDataSource.Size - 1, Value);
                            }
                            catch (Exception) { }
                        }
                    }
                }
                Matrix4.LoadFromDataSource();
            }
            catch (Exception) { }
            finally { oForm.Freeze(false); }
        }
        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
           // throw new System.NotImplementedException();

        }

        private SAPbouiCOM.Folder Folder2;
        private SAPbouiCOM.Folder Folder3;
        private SAPbouiCOM.StaticText StaticText30;
        private SAPbouiCOM.StaticText StaticText31;
        private SAPbouiCOM.EditText EditText31;
        private SAPbouiCOM.StaticText StaticText32;
        private SAPbouiCOM.EditText EditText32;
        private SAPbouiCOM.StaticText StaticText33;
        private SAPbouiCOM.EditText EditText33;
        private SAPbouiCOM.StaticText StaticText34;
        private SAPbouiCOM.EditText EditText34;
        private SAPbouiCOM.StaticText StaticText35;
        private SAPbouiCOM.EditText EditText35;
        private SAPbouiCOM.StaticText StaticText36;
        private SAPbouiCOM.EditText EditText36;
        private SAPbouiCOM.StaticText StaticText37;
        private SAPbouiCOM.EditText EditText37;
        private SAPbouiCOM.Button Button6;
        private SAPbouiCOM.Button Button7;
        private SAPbouiCOM.Matrix Matrix4;
        private SAPbouiCOM.EditText EditText38;
        private SAPbouiCOM.StaticText StaticText38;
        private SAPbouiCOM.EditText EditText39;
        private SAPbouiCOM.StaticText StaticText39;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.PictureBox PictureBox0;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.CheckBox CheckBox0;
        private SAPbouiCOM.Folder Folder0;
        private SAPbouiCOM.CheckBox CheckBox1;
        private SAPbouiCOM.Folder Folder1;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.Matrix Matrix1;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Button Button2;



        private void Button6_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_COMI_GRUPO");
            //var Value = (string)null;
            //string sfield;

            //for (int i = 0; i <= oDBDataSource.Size - 1; i++)
            //{
            //    for (int x = 0; x <= oDBDataSource.Fields.Count - 1; x++)
            //    {
            //        try
            //        {
            //            sfield = oDBDataSource.Fields.Item(x).Name;
            //            Value = oDBDataSource.GetValue(sfield, i);
            //            Value = Value.Trim().Length == 0 ? null : Value;
            //        }
            //        catch (Exception) { }
            //    }
            //}

        }

        private SAPbouiCOM.Folder Folder4;
        private SAPbouiCOM.Grid Grid1;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.StaticText StaticText5;



        //private SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler form_DataEvent;

        //private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        //{
        //    if (Button0.Caption == "Crear")
        //    {
        //        SAPbouiCOM.DBDataSource oDBDataSource = this.UIAPIRawForm.DataSources.DBDataSources.Item("@Z_COMI_PER");

        //        string Code = oDBDataSource.GetValue("Code", oDBDataSource.Offset);

        //        oDBDataSource.SetValue("Code", oDBDataSource.Offset, "2");
        //    }

        //}


        //  private void Button1_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        //{

        // ////SAPbouiCOM.DBDataSource oDBDataSource = this.UIAPIRawForm.DataSources.DBDataSources.Item("@Z_COMI_PER");

        // ////string Code = oDBDataSource.GetValue("Code", oDBDataSource.Offset);

        // ////oDBDataSource.SetValue("Code", oDBDataSource.Offset, "2");


        // //SAPbobsCOM.GeneralService oGeneralService = null;
        // //SAPbobsCOM.GeneralData oGeneralData = null;
        // ////SAPbobsCOM.GeneralData oChild = null;
        // ////SAPbobsCOM.GeneralDataCollection oChildren = null;
        // //SAPbobsCOM.GeneralDataParams oGeneralParams = null;
        // //SAPbobsCOM.CompanyService oCompanyService = null;
        // //try
        // //{

        // //    oCompanyService = Program.oCompany.GetCompanyService();
        // //    // Get GeneralService (oCmpSrv is the CompanyService)
        // //    oGeneralService = oCompanyService.GetGeneralService("Z_COMI_PER");
        // //    // Create data for new row in main UDO
        // //    oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));
        // //    oGeneralData.SetProperty("Code", "1");
        // //    oGeneralData.SetProperty("Name", "1");
        // //    //oGeneralData.SetProperty("U_Codigo",EditText1.Value);
        // //    //oGeneralData.SetProperty("U_Nombre", EditText0.Value);
        // //    //oGeneralData.SetProperty("U_Abierto", EditText2.Value);
        // //    //  Handle child rows
        // //    //oChildren = oGeneralData.Child("SM_MOR1");
        // //    //int i = 0;
        // //    //for (i = 1; i <= lstMainDish.Items.Count; i++)
        // //    //{
        // //    //    // Create data for rows in the child table
        // //    //    oChild = oChildren.Add();
        // //    //    oChild.SetProperty("U_MainDish", lstMainDish.Items[i - 1]);
        // //    //    oChild.SetProperty("U_SideDish", lstSideDish.Items[i - 1]);
        // //    //    oChild.SetProperty("U_Drink", lstDrink.Items[i - 1]);
        // //    //}
        // //    // Add the new row, including children, to database
        // //    oGeneralParams = oGeneralService.Add(oGeneralData);
        // //}
        // //catch //(Exception)
        // //{
        // //    //Interaction.MsgBox(ex.Message, (Microsoft.VisualBasic.MsgBoxStyle)(0), null);
        // //}

        // }


    }
}
