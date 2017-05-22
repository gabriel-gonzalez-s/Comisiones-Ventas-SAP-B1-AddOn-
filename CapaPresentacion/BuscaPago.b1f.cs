using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace ComisionesVentas.CapaPresentacion
{
    [FormAttribute("ComisionesVentas.CapaPresentacion.BuscaPago", "CapaPresentacion/BuscaPago.b1f")]
    class BuscaPago : UserFormBase
    {
        private static SAPbouiCOM.Form oForm = null;
        private static string Ultima_Busqueda = "";

        public BuscaPago()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("Item_1").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_3").Specific));
            this.EditText0.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.EditText0_KeyDownAfter);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_4").Specific));
            this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_5").Specific));
            this.Button1.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button1_ClickAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_6").Specific));
            this.Button2.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button2_ClickAfter);
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_7").Specific));
            this.Grid0.DoubleClickAfter += new SAPbouiCOM._IGridEvents_DoubleClickAfterEventHandler(this.Grid0_DoubleClickAfter);
            this.Grid0.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid0_ClickAfter);
            this.DT_SQL = ((SAPbouiCOM.DataTable)(this.UIAPIRawForm.DataSources.DataTables.Item("DT_SQL")));
            this.CheckBox0 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_8").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Folder Folder0;

        private void OnCustomInitialize()
        {
            oForm = Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
            BuscarDatos();
        }

        private void Grid0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                Grid0.Rows.SelectedRows.Add(pVal.Row);
            }
            catch (Exception)
            {
            }
        }

        private void Grid0_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Button1.Item.Click();
        }

        private void EditText0_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                switch (pVal.CharPressed)
                {
                    case 13:
                        if (EditText0.Value.Trim() == Ultima_Busqueda || (EditText0.Value.Trim().Length == 0 && Ultima_Busqueda == "°#TODOS LOS REGISTROS#°") && Grid0.Rows.SelectedRows.Count > 0)
                        { Button1.Item.Click(); }
                        else { Button0.Item.Click(); }
                        break;
                    case 40:
                        if (Grid0.Rows.Count > 0)
                        {
                            int nRow = Grid0.Rows.SelectedRows.Count == 0 ? -1 : Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                            Grid0.Rows.SelectedRows.Add(nRow < 0 ? 0 : nRow + 1);
                        }
                        break;
                    case 38:
                        if (Grid0.Rows.Count > 0)
                        {
                            int nRow = Grid0.Rows.SelectedRows.Count == 0 ? -1 : Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                            Grid0.Rows.SelectedRows.Add(nRow < 0 ? 0 : nRow - 1);
                        }
                        break;
                }

            }
            catch (Exception)
            {
            }

        }

        private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            BuscarDatos();
        }

        private void Button1_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (Grid0.Rows.SelectedRows.Count > 0)
                {
                    //oUDS = oForm.DataSources.UserDataSources.Item("UD_1");
                    string sPago = Convert.ToString(Grid0.DataTable.GetValue(0, Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)));
                    Comisiones.Agregar_Documento_Grid_Pagos(sPago);
                    oForm.Close(); 
                }

            }
            catch (Exception)
            {
            }

        }


        private void Button2_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            oForm.Close();
        }

        private void BuscarDatos()
        {
            try
            {
                string sql = "";
                string Condi = CheckBox0.Checked ? "Y" : "N";
                if (EditText0.Value.Trim().Length > 0)
                {
                    sql = @"SELECT
		                            ORCT.DocNum AS 'Numero Pago' 
		                            ,ORCT.TaxDate as 'Fecha PR'
		                            --,ORCT.CardName as 'Cliente'
		                            ,ORCT.Comments as 'Info Detallada'
		                            ,IIF(ORCT.NoDocSum > 0, ORCT.NoDocSum, ORCT.DocTotal) as 'Total del Documento'
		                            ,ORCT.JrnlMemo as 'Comentarios'
		                            ,LEFT(ORCT.Ref1,2)+'-'+SUBSTRING(ORCT.Ref1,3,2)+'-'+RIGHT(ORCT.Ref1,4) as 'Referencia 1'
		                            ,ORCT.CreateDate as 'Fecha Creacion'
                            FROM
	                            [" + Program.sBDComercial.Trim() + @"].DBO.ORCT
                            WHERE
	                            ORCT.Canceled = '" + Condi + @"'
                                and (CAST(DocNum as nvarchar(20)) like '%" + EditText0.Value.Trim() + @"%' 
                                        or CAST(Comments as nvarchar(MAX)) like '%" + EditText0.Value.Trim() + @"%'
                                        or CAST(JrnlMemo as nvarchar(MAX)) like '%" + EditText0.Value.Trim() + @"%')
                            ORDER BY
	                            1" ;

                    DT_SQL.ExecuteQuery(sql);

                    Ultima_Busqueda = EditText0.Value.Trim();

                }
                else
                {
                    sql = @"SELECT
		                            ORCT.DocNum AS 'Numero Pago' 
		                            ,ORCT.TaxDate as 'Fecha PR'
		                            --,ORCT.CardName as 'Cliente'
		                            ,ORCT.Comments as 'Info Detallada'
		                            ,IIF(ORCT.NoDocSum > 0, ORCT.NoDocSum, ORCT.DocTotal) as 'Total del Documento'
		                            ,ORCT.JrnlMemo as 'Comentarios'
		                            ,LEFT(ORCT.Ref1,2)+'-'+SUBSTRING(ORCT.Ref1,3,2)+'-'+RIGHT(ORCT.Ref1,4) as 'Referencia 1'
		                            ,ORCT.CreateDate as 'Fecha Creacion'
                            FROM
	                            [" + Program.sBDComercial.Trim() + @"].DBO.ORCT
                            WHERE
	                            ORCT.Canceled = '" + Condi + @"'
                            ORDER BY
	                            1";

                    DT_SQL.ExecuteQuery(sql);

                    Ultima_Busqueda = "°#TODOS LOS REGISTROS#°";
                }
            }
            catch (Exception)
            {
            }
        }

        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.DataTable DT_SQL;
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.CheckBox CheckBox0;

    }
}
