using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LiquidacionAutor.Formularios;
using SAPbobsCOM;
using SAPbouiCOM;

namespace LiquidacionAutor
{

    // 0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056

    public class ConexionAddOn
    {

        private SAPbouiCOM.Application SBO_Application;
        private SAPbobsCOM.Company oCompany;
        private AsisLiqRegaliasAut asisLiqRegaliasAut;
        private string Pais;
        
        public ConexionAddOn()
        {
            SetApplication();
            AgregarMenus();
        }

        private void SetApplication()
        {

            string strCookie;
            string strConnectionContext;
            int intError;
            string strError = "";

            SAPbouiCOM.SboGuiApi oSboGuiApi=new SboGuiApi();
            string strCon = Environment.GetCommandLineArgs().GetValue(1).ToString();

            oSboGuiApi.Connect(strCon);

            SBO_Application = oSboGuiApi.GetApplication();

            //Se inicializan los eventos que se requieren para el add-on
            SBO_Application.ItemEvent+=new _IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);

            SBO_Application.MenuEvent += new _IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);

            SBO_Application.AppEvent += new _IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);

            SBO_Application.FormDataEvent += new _IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);

            oCompany = new SAPbobsCOM.Company();

            strCookie = oCompany.GetContextCookie();

            strConnectionContext = SBO_Application.Company.GetConnectionContext(strCookie);

            if (oCompany.Connected)
            {
                oCompany.Disconnect();
            }

            oCompany.SetSboLoginContext(strConnectionContext);

            intError = oCompany.Connect();

            if (intError != 0)
            {
                oCompany.GetLastError(out intError, out strError);

                SBO_Application.StatusBar.SetText(strError,BoMessageTime.bmt_Medium,BoStatusBarMessageType.smt_Error);
            }
            else
            {
                new Estructuras(SBO_Application, oCompany).crearEstructuras();

                SBO_Application.StatusBar.SetText("AddOn Liquidacion de autor conectado", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);

                SetFilters();
                consultarPais();
            }

        }

        private void consultarPais()
        {
            Recordset oRecordset = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            SQL.SQL sql = new SQL.SQL("LiquidacionAutor.SQL.GetContrySociety.sql");
            oRecordset.DoQuery(string.Format(sql.getQuery()));

            if (oRecordset.RecordCount > 0)
            {
                this.Pais = oRecordset.Fields.Item("Country").Value.ToString();
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            oRecordset = null;
            GC.Collect();
        }

        private void SetFilters()
        {

            SAPbouiCOM.EventFilters oFilters=new EventFilters();
            SAPbouiCOM.EventFilter oFilter;

            oFilter = oFilters.Add(BoEventTypes.et_CLICK);
            oFilter.AddEx("HCO_ACRA");

            oFilter = oFilters.Add(BoEventTypes.et_ITEM_PRESSED);
            oFilter.AddEx("HCO_ACRA");        

            oFilter = oFilters.Add(BoEventTypes.et_CHOOSE_FROM_LIST);
            oFilter.AddEx("HCO_ACRA");

            oFilter = oFilters.Add(BoEventTypes.et_MENU_CLICK);
            oFilter.AddEx("HCO_MCRA");

            oFilter = oFilters.Add(BoEventTypes.et_FORM_LOAD);
            oFilter.AddEx("HCO_ACRA");

            oFilter = oFilters.Add(BoEventTypes.et_FORM_DATA_UPDATE);
            oFilter.AddEx("HCO_ACRA");

            oFilter = oFilters.Add(BoEventTypes.et_MATRIX_LINK_PRESSED);
            oFilter.AddEx("HCO_ACRA");

            SBO_Application.SetFilter(oFilters);

        }

        private void AgregarMenus()
        {

            SAPbouiCOM.Menus oMenus;
            SAPbouiCOM.MenuItem oMenuItem;
            SAPbouiCOM.MenuCreationParams oMenuCreationParams;

            oMenuCreationParams = (MenuCreationParams) SBO_Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);

            oMenuItem = SBO_Application.Menus.Item("1536");

            oMenuCreationParams.Type=BoMenuType.mt_STRING;
            oMenuCreationParams.UniqueID = "HCO_MCRA";
            oMenuCreationParams.String = "Asistente de liquidación de autor";
            oMenuCreationParams.Position = 15;

            oMenus = oMenuItem.SubMenus;

            if (!oMenus.Exists("HCO_MCRA"))
            {
                oMenus.AddEx(oMenuCreationParams);
            }
        }

        public void SBO_Application_ItemEvent(string strFormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.FormTypeEx == "HCO_ACRA")
                    asisLiqRegaliasAut.ManejarEventosItem(ref pVal, ref BubbleEvent);
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.Message);
            }                  
        }

        public void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo oBusinessInfo, out  bool BubbleEvent)
        {
            BubbleEvent = true;
            //try
            //{                
            //}
            //catch (Exception ex)
            //{
            //    throw ex;
            //}
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;//Se llama el método del evento para cada uno de las clases 

            if (pVal.MenuUID == "HCO_MCRA")
            {
                if (pVal.BeforeAction)
                {
                    this.asisLiqRegaliasAut = new AsisLiqRegaliasAut(SBO_Application, oCompany, this.Pais);
                    this.asisLiqRegaliasAut.CrearFormulario();
                }
            }
        }

        public void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes eventTypes)
        {

            try
            {

                SAPbouiCOM.Menus oMenus;

                if (eventTypes == BoAppEventTypes.aet_CompanyChanged || eventTypes == BoAppEventTypes.aet_FontChanged ||
                    eventTypes == BoAppEventTypes.aet_LanguageChanged || eventTypes == BoAppEventTypes.aet_ServerTerminition ||
                    eventTypes == BoAppEventTypes.aet_ShutDown)
                {

                    if (SBO_Application.Forms.Count > 0)
                    {
                        foreach (SAPbouiCOM.Form oForm in SBO_Application.Forms)
                        {
                            if (new string[] { "HCO_ACRA" }.Contains(oForm.TypeEx))
                            {
                                oForm.Close();
                            }
                        }
                    }                    

                    oMenus = SBO_Application.Menus;

                    if (oMenus.Exists("HCO_MCRA"))                    
                        oMenus.RemoveEx("HCO_MCRA");

                    if (oCompany.Connected)
                    {
                        oCompany.Disconnect();
                    }

                    System.Windows.Forms.Application.Exit();

                }

            }
            catch (Exception e)
            {
                SBO_Application.StatusBar.SetText(e.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
            }

        }

    }
}
