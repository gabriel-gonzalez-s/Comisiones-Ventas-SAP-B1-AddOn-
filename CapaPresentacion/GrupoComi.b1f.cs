using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

using FuncionalidadesSDKB1;
using ComisionesVentas.CapaNegocios;

namespace ComisionesVentas
{
    [FormAttribute("ComisionesVentas.GrupoComi", "CapaPresentacion/GrupoComi.b1f")]
    class GrupoComi : UserFormBase
    {
        SAPbouiCOM.Form oForm = null;
        public GrupoComi()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_2").Specific));
            this.Matrix0.ClickBefore += new SAPbouiCOM._IMatrixEvents_ClickBeforeEventHandler(this.Matrix0_ClickBefore);
            this.Matrix0.MatrixLoadAfter += new SAPbouiCOM._IMatrixEvents_MatrixLoadAfterEventHandler(this.Matrix0_MatrixLoadAfter);
            this.Matrix0.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix0_ClickAfter);
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("Item_1").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.DataUpdateBefore += new SAPbouiCOM.Framework.FormBase.DataUpdateBeforeHandler(this.Form_DataUpdateBefore);
            this.DataUpdateAfter += new DataUpdateAfterHandler(this.Form_DataUpdateAfter);

        }
        private void OnCustomInitialize()
        {
            // Ready Matrix to populate data
            Matrix0.Clear();
            Matrix0.Columns.Item("Col_0").Visible = false;
            Matrix0.Columns.Item("Col_1").Visible = false;

            // Querying the DB Data source
            oForm = Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
            oForm.Select();

            LoadDataMatrix();

            oForm.EnableMenu("1281",false);
            oForm.EnableMenu("1282",false);

        }
        private void LoadDataMatrix()
        {
            SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_COMI_COMGRP");
            oDBDataSource.Query();
            oDBDataSource.InsertRecord(oDBDataSource.Size);

            //MatrixExtensions.AddLineMatrixDBDataSource(Matrix0, oDBDataSource);

            int nProxCode = UDOExtensions.GetNextCode("Z_COMI_COMGRP", NConexion.Verifica_Conexion(Program.oCompany));
            oDBDataSource.SetValue("Code", oDBDataSource.Size - 1, nProxCode.ToString());
            oDBDataSource.SetValue("DocEntry", oDBDataSource.Size - 1, nProxCode.ToString());

            Matrix0.LoadFromDataSource();
            Matrix0.CommonSetting.SetCellEditable(oDBDataSource.Size, 3, true);
            MatrixExtensions.SelectMatrixRowSetFocus(ref Matrix0, "Col_2", oDBDataSource.Size);
            
        }
        private void Matrix0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            MatrixExtensions.SelectMatrixRow(ref Matrix0, pVal.Row);
        }
        private void Matrix0_MatrixLoadAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Matrix0.AutoResizeColumns();
        }
        private void Matrix0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                oForm.EnableMenu("1283", false);
                oForm.EnableMenu("1284", false);
            }
            catch (Exception) { }
        }
        private void Form_DataUpdateBefore(ref SAPbouiCOM.BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_COMI_COMGRP");

            var ValidaNuevoReg = oDBDataSource.GetValue("U_Codigo", oDBDataSource.Size - 1);
            if (ValidaNuevoReg.Trim().Length == 0)
                oDBDataSource.RemoveRecord(oDBDataSource.Size - 1);

        }
        private void Form_DataUpdateAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            try
            {
                oForm.Freeze(true);
                Matrix0.Columns.Item("Col_2").Editable = false;
                LoadDataMatrix();
            }
            catch (Exception) { }
            finally { oForm.Freeze(false); }

        }

        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.Folder Folder0;


    }
}
