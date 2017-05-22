using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace ComisionesVentas.CapaPresentacion
{
    [FormAttribute("ComisionesVentas.CapaPresentacion.BuscarPedido", "CapaPresentacion/BuscarPedido.b1f")]
    class BuscarPedido : UserFormBase
    {
        private static SAPbouiCOM.Form oForm = null;
        private static SAPbouiCOM.UserDataSource oUDS = null;
        private static string Ultima_Busqueda = "";

        public BuscarPedido()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_0").Specific));
            this.EditText0.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.EditText0_KeyDownAfter);
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_1").Specific));
            this.Grid0.DoubleClickAfter += new SAPbouiCOM._IGridEvents_DoubleClickAfterEventHandler(this.Grid0_DoubleClickAfter);
            this.Grid0.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid0_ClickAfter);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_2").Specific));
            this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_3").Specific));
            this.Button1.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button1_ClickAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_4").Specific));
            this.Button2.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button2_ClickAfter);
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.DT_SQL = ((SAPbouiCOM.DataTable)(this.UIAPIRawForm.DataSources.DataTables.Item("DT_SQL")));
            this.DT_PCV = ((SAPbouiCOM.DataTable)(this.UIAPIRawForm.DataSources.DataTables.Item("DT_PCV")));
            this.OptionBtn0 = ((SAPbouiCOM.OptionBtn)(this.GetItem("Item_7").Specific));
            this.OptionBtn1 = ((SAPbouiCOM.OptionBtn)(this.GetItem("Item_8").Specific));
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("Item_10").Specific));
            this.Folder1 = ((SAPbouiCOM.Folder)(this.GetItem("Item_12").Specific));
            this.CheckBox0 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_13").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.ResizeAfter += new ResizeAfterHandler(this.Form_ResizeAfter);

        }

        private SAPbouiCOM.EditText EditText0;

        private void OnCustomInitialize()
        {
            oForm = Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
            OptionBtn0.Item.Width = 65;
            OptionBtn1.GroupWith("Item_7");
            OptionBtn0.Selected = true;
        }

        private void Form_ResizeAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                oUDS = oForm.DataSources.UserDataSources.Item("UD_1");
                string sql = @"select Project as 'Proyecto', DocNum as 'Numero Pedido', CardName as 'Cliente Final', CreateDate as 'Fecha Creacion'
                                from ORDR 
                                where Series = 27 and SlpCode = " + oUDS.ValueEx + @"
                                order by Project, CreateDate";

                DT_SQL.ExecuteQuery(sql);
                DT_PCV.CopyFrom(DT_SQL);

                Grid0.Columns.Item(2).Width = 200;

                oUDS =  oForm.DataSources.UserDataSources.Item("UD_2");
                StaticText1.Caption = "VENDEDOR: " + oUDS.ValueEx ;

            }
            catch (Exception)
            {
            }

        }

        private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                string sql = "";
                if (EditText0.Value.Trim().Length > 0)
                {
                    oUDS = oForm.DataSources.UserDataSources.Item("UD_1");
                    if (CheckBox0.Checked)
                        sql = @"select Project as 'Proyecto', DocNum as 'Numero Pedido', CardName as 'Cliente Final', CreateDate as 'Fecha Creacion'
                                    from ORDR 
                                    where Series = 27 and ( DocNum = " + EditText0.Value.Trim() + " or Project = '" + EditText0.Value.Trim() + @"')  
                                    order by Project, CreateDate";
                    else
                        sql = @"select Project as 'Proyecto', DocNum as 'Numero Pedido', CardName as 'Cliente Final', CreateDate as 'Fecha Creacion'
                                    from ORDR 
                                    where Series = 27 and SlpCode = " + oUDS.ValueEx + " and ( DocNum = " + EditText0.Value.Trim() + " or Project = '" + EditText0.Value.Trim() + @"')  
                                    order by Project, CreateDate";

                    DT_SQL.ExecuteQuery(sql);
                    DT_PCV.CopyFrom(DT_SQL);

                    Grid0.Columns.Item(2).Width = 200;

                    Ultima_Busqueda = EditText0.Value.Trim();

                }
                else
                {
                    oUDS = oForm.DataSources.UserDataSources.Item("UD_1");
                    if (CheckBox0.Checked)
                        sql = @"select Project as 'Proyecto', DocNum as 'Numero Pedido', CardName as 'Cliente Final', CreateDate as 'Fecha Creacion'
                                    from ORDR 
                                    where Series = 27 
                                    order by Project, CreateDate";
                    else
                        sql = @"select Project as 'Proyecto', DocNum as 'Numero Pedido', CardName as 'Cliente Final', CreateDate as 'Fecha Creacion'
                                    from ORDR 
                                    where Series = 27 and SlpCode = " + oUDS.ValueEx + @"
                                    order by Project, CreateDate";

                    DT_SQL.ExecuteQuery(sql);
                    DT_PCV.CopyFrom(DT_SQL);

                    Grid0.Columns.Item(2).Width = 200;

                    Ultima_Busqueda = "°#TODOS LOS REGISTROS#°";
                }
            }
            catch (Exception)
            {
            }

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

        private void Button2_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            oForm.Close();
        }

        private void Button1_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (Grid0.Rows.SelectedRows.Count > 0)
                {
                    oUDS = oForm.DataSources.UserDataSources.Item("UD_1");
                    string sPedido = Convert.ToString(Grid0.DataTable.GetValue(1, Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)));
                    int TipoOperacion = OptionBtn0.Selected ? 1 : -1;
                    int Vend = CheckBox0.Checked ?  Convert.ToInt32(oUDS.ValueEx) : 0 ; //Si muestra todos los vendedores se pasa la info con el vendedor consultado.
                    Comisiones.Agregar_Documento_Grid_Comisiones(sPedido, ref TipoOperacion, Vend);
                    if (TipoOperacion != 0) { oForm.Close(); }
                }

            }
            catch (Exception)
            {
            }
        }


        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.DataTable DT_SQL;
        private SAPbouiCOM.DataTable DT_PCV;
        private SAPbouiCOM.OptionBtn OptionBtn0;
        private SAPbouiCOM.OptionBtn OptionBtn1;
        private SAPbouiCOM.Folder Folder0;
        private SAPbouiCOM.Folder Folder1;
        private SAPbouiCOM.CheckBox CheckBox0;








    }
}
