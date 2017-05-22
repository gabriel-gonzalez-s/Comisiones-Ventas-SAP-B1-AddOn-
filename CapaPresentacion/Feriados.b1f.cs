using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace ComisionesVentas.CapaPresentacion
{
    [FormAttribute("ComisionesVentas.CapaPresentacion.Feriados", "CapaPresentacion/Feriados.b1f")]
    class Feriados : UserFormBase
    {
        private static SAPbouiCOM.Form oForm = null;
        private static SAPbouiCOM.Matrix oMatrix = null;
        private static SAPbouiCOM.DataTable oDTTable = null;

        public Feriados()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("3").Specific));
            this.Matrix0.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix0_ClickAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.ResizeAfter += new SAPbouiCOM.Framework.FormBase.ResizeAfterHandler(this.Form_ResizeAfter);
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);

        }

        private SAPbouiCOM.Button Button0;

        private void OnCustomInitialize()
        {
            oForm = Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
            //oForm =  Application.SBO_Application.Forms.ActiveForm;
            oForm.EnableMenu("1281", false);
            oForm.EnableMenu("1282", false);
            oForm.EnableMenu("1288", false);
            oForm.EnableMenu("1289", false);
            oForm.EnableMenu("1290", false);
            oForm.EnableMenu("1291", false);
            Matrix0.Columns.Item("Code").Visible = false;
            Matrix0.Columns.Item("DocEntry").Visible = false;
            Cargar_Datos_Matrix();

        }

        private void Form_ResizeAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            
        }

        private static void Cargar_Datos_Matrix()
        {
            try
            {

                //oForm = Application.SBO_Application.Forms.ActiveForm;
                oDTTable = oForm.DataSources.DataTables.Item("DT_SQL1");
                SAPbouiCOM.DBDataSource source = oForm.DataSources.DBDataSources.Item("@ZDFER");

                oForm.Freeze(true);

                string sql = "SELECT * FROM [@ZDFER] ORDER BY U_Fecha";

                oDTTable.ExecuteQuery(sql);

                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;
                oMatrix.FlushToDataSource();
                source.Clear();
                oMatrix.Clear();

                for (int i = 0; i < oDTTable.Rows.Count; i++)
                {
                    source.InsertRecord(source.Size);
                    source.Offset = source.Size - 1;
                    source.SetValue("Code", source.Size - 1, Convert.ToString(oDTTable.GetValue("Code", i)));
                    //DateTime fecha = (DateTime)oDTTable.GetValue("U_Fecha", i);
                    source.SetValue("U_Fecha", source.Size - 1, Convert.ToDateTime(oDTTable.GetValue("U_Fecha", i)).ToString("yyyyMMdd"));
                    source.SetValue("U_Descripcion", source.Size - 1, Convert.ToString(oDTTable.GetValue("U_Descripcion", i)));
                    source.SetValue("DocEntry", source.Size - 1, Convert.ToString(oDTTable.GetValue("DocEntry", i)));
                }

                sql = "SELECT MAX(CAST(Code as int)) + 1 as Proximo FROM [@ZDFER]";
                oDTTable.ExecuteQuery(sql);

                source.InsertRecord(source.Size);
                source.Offset = source.Size - 1;
                source.SetValue("Code", source.Size - 1, Convert.ToString(oDTTable.GetValue("Proximo", 0)));
                source.SetValue("DocEntry", source.Size - 1, Convert.ToString(oDTTable.GetValue("Proximo", 0)));

                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
            }
            catch (Exception)
            {
            }
            finally
            {
                oForm.Freeze(false);
            }
        }



        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Matrix Matrix0;



        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //Cargar_Datos_Matrix();
            oDTTable = oForm.DataSources.DataTables.Item("DT_SQL1");

            string sql = "DELETE FROM [@ZDFER] WHERE U_Fecha is null";

            oDTTable.ExecuteQuery(sql);

            Cargar_Datos_Matrix();

        }

        private void Matrix0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            int rowNum = Matrix0.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
            try
            {
                Matrix0.SelectRow(pVal.Row, true, false);
            }
            catch (Exception)
            {
            }

        }

        private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //if (Button0.Caption == "Actualizar")
            //{
            //    Cargar_Datos_Matrix();
            //    //Matrix0.AddRow();
            //}
        }

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            throw new System.NotImplementedException();

        }
    }
}
