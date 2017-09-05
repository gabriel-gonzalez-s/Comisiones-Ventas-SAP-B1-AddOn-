using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using ComisionesVentas.CapaNegocios;

namespace ComisionesVentas.CapaPresentacion
{
    [FormAttribute("ComisionesVentas.CapaPresentacion.BuscaVendedor", "CapaPresentacion/BuscaVendedor.b1f")]
    class BuscaVendedor : UserFormBase
    {
        private static SAPbouiCOM.Form oForm = null;
        private static SAPbouiCOM.UserDataSource oUDS = null;
        public static string sOrigf { get; set; }
        public BuscaVendedor()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("Item_1").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_3").Specific));
            this.Button1.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button1_ClickAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_4").Specific));
            this.Button2.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button2_ClickAfter);
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_5").Specific));
            this.Grid0.DoubleClickAfter += new SAPbouiCOM._IGridEvents_DoubleClickAfterEventHandler(this.Grid0_DoubleClickAfter);
            this.Grid0.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid0_ClickAfter);
            this.DT_SQL = ((SAPbouiCOM.DataTable)(this.UIAPIRawForm.DataSources.DataTables.Item("DT_SQL")));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.ResizeAfter += new ResizeAfterHandler(this.Form_ResizeAfter);

        }

        private SAPbouiCOM.Folder Folder0;

        private void OnCustomInitialize()
        {
            try
            {
                oForm = Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);

                string sql = @"SELECT 
	                                SlpCode as 'Codigo'
	                                ,Memo  as 'Inicial' 
	                                ,SlpName as 'Nombre'
                                FROM 
	                                OSLP 
                                WHERE 
	                                Active = 'Y'
	                                AND SlpCode != -1
                                ORDER BY
	                                1";
                DT_SQL.ExecuteQuery(sql);
                oForm.ReportType = "Modal";
            }
            catch (Exception) { }
        }

        private void Form_ResizeAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            oUDS = oForm.DataSources.UserDataSources.Item("UD_1");
            sOrigf = oUDS.ValueEx;
            string sql;
            switch (sOrigf)
            {
                case "ComisionesA":
                    oForm.Title = "Agregar Vendedor a Comisiones";
                    sql = @"SELECT 
	                            SlpCode as 'Codigo'
	                            ,Memo  as 'Inicial' 
	                            ,SlpName as 'Nombre'
                            FROM 
	                            OSLP 
                            WHERE 
	                            Active = 'Y'
	                            AND SlpCode != -1
                                AND SlpCode NOT IN ( SELECT DISTINCT V.SlpCode 
                                                        FROM OSLP V join [@Z_COMI_VEND] P ON V.SlpCode = P.[Id Vend]
                                                        WHERE V.Active  = 'Y' 
                                                            AND V.SlpCode != -1  
                                                            AND Periodo = '" + NPeriodo.NombrePeriodo + @"') 
                            ORDER BY
	                            1";
                    DT_SQL.ExecuteQuery(sql);
                    break;
                case "ComisionesR":
                    oForm.Title = "Agregar Vendedor a Comisiones";
                    sql = @"SELECT 
	                            SlpCode as 'Codigo'
	                            ,Memo  as 'Inicial' 
	                            ,SlpName as 'Nombre'
                            FROM 
	                            OSLP 
                            WHERE 
	                            Active = 'Y'
	                            AND SlpCode != -1
                                AND SlpCode NOT IN ( SELECT DISTINCT V.SlpCode 
                                                        FROM OSLP V join ORDR P ON V.SlpCode = P.SlpCode
                                                        WHERE V.Active  = 'Y' AND V.SlpCode != -1 AND series = 27
                                                        AND P.CreateDate BETWEEN  '" + NPeriodo.FechaInicio.ToString("MM-dd-yyyy") + 
                                                                     @"' AND '" + NPeriodo.FechaFin.ToString("MM-dd-yyyy") + @"') 
                            ORDER BY
	                            1";
                    DT_SQL.ExecuteQuery(sql);
                    break;
            }
        }

        private void Grid0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                Grid0.Rows.SelectedRows.Add(pVal.Row);
            }
            catch (Exception) { }
        }

        private void Grid0_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Button1.Item.Click();
        }

        private void Button2_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            oForm.Close();
        }

        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.DataTable DT_SQL;

        private void Button1_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (Grid0.Rows.SelectedRows.Count > 0)
            {

                oUDS = oForm.DataSources.UserDataSources.Item("UD_0");
                int nRow = Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                string[] sVend = new string[2];
                sVend[0] = Convert.ToString(Grid0.DataTable.GetValue(0, nRow));
                sVend[1] = "(" + Convert.ToString(Grid0.DataTable.GetValue(1, nRow)) + ") - " 
                            + Convert.ToString(Grid0.DataTable.GetValue(2, nRow));
                SAPbouiCOM.Form oFormP = Application.SBO_Application.Forms.Item(oUDS.ValueEx);
                oForm.Close();

                switch (sOrigf)
                {
                    case "Pagos":
                        Comisiones.Asigna_Vendedor_Grid_Pagos(sVend[0]);
                        break;
                    case "Premios":
                        Comisiones.Asigna_Vendedor_Grid_Premios(sVend[0]);
                        break;
                    case "ComisionesA":
                    case "ComisionesR":
                        Comisiones.Agregar_Vendedor_Comisiones_Periodo(sVend);
                        break;

                }


            }

        }
    }
}
