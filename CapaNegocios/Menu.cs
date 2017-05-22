using System;
using System.Collections.Generic;
using System.Text;
using SAPbouiCOM.Framework;

namespace ComisionesVentas
{
    class Menu
    {
        private static SAPbouiCOM.Application SBO_Application;

        public void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            oMenus = Application.SBO_Application.Menus;
            SBO_Application = Application.SBO_Application;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = Application.SBO_Application.Menus.Item("43520"); // moudles'


            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
            oCreationPackage.UniqueID = "ComisionesVentas";
            oCreationPackage.String = "Comisiones Ventas";
            oCreationPackage.Image = @"\\fssapbo\SAPB1\Anexos\images\coins_add.png";
            oCreationPackage.Enabled = true;
            oCreationPackage.Position = -1;

            oMenus = oMenuItem.SubMenus;

            if (SBO_Application.Menus.Exists("ComisionesVentas") == true)
            {
                SBO_Application.Menus.RemoveEx("ComisionesVentas");
            }

            try
            {
                //  If the manu already exists this code will fail
                oMenus.AddEx(oCreationPackage);
            }
            catch //(Exception e)
            {

            }

            try
            {
                // Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("ComisionesVentas");
                oMenus = oMenuItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "ComisionesVentas.Parametros";
                oCreationPackage.String = "Parametros";
                oCreationPackage.Position = 1;
                oMenus.AddEx(oCreationPackage);
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "ComisionesVentas.Comisiones";
                oCreationPackage.String = "Comisiones";
                oCreationPackage.Position = 2;
                oMenus.AddEx(oCreationPackage);

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                oCreationPackage.UniqueID = "Informes2";
                oCreationPackage.String = "Informes2";
                oCreationPackage.Position = -1;
                oMenus.AddEx(oCreationPackage);

                oMenuItem = Application.SBO_Application.Menus.Item("Informes2");
                oMenus = oMenuItem.SubMenus;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "ComisionesVentas.Comisiones2";
                oCreationPackage.String = "Menu 2";
                oCreationPackage.Position = 8;
                oMenus.AddEx(oCreationPackage);

            }
            catch //(Exception er)
            { //  Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        
            
            
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            UserFormBase activeForm = null;

            try
            {
                if (pVal.BeforeAction)
                switch (pVal.MenuUID)
                {
                    case "ComisionesVentas.Parametros":
                        activeForm = new Parametros();
                        activeForm.Show();

                        break;
                    case "ComisionesVentas.Comisiones":
                        activeForm = new Comisiones();
                        activeForm.Show();
                        break;
                } 

                //if (pVal.BeforeAction && pVal.MenuUID == "ComisionesVentas.Comisiones")
                //{
                //    Comisiones activeForm = new Comisiones();
                //    activeForm.Show();
                //}
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }

    }
}
