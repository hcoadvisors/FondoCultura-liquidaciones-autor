using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace LiquidacionAutor.Formularios
{
    class AsisLiqRegaliasAut
    {
        protected Application SBO_Application;
        protected Form oForm;
        protected SAPbobsCOM.Company oCompany;
        private string Pais;
        public AsisLiqRegaliasAut(Application sboaApplication, SAPbobsCOM.Company sboCompany, string pais)
        {
            this.SBO_Application = sboaApplication;
            this.oCompany = sboCompany;
            this.Pais = pais;
        }

        public void CrearFormulario()
        {
            CargarFormulario();
        }

        public void CargarFormulario()
        {
            try
            {
                bool blnFormOpen = false;

                if (!blnFormOpen)
                {

                    FormCreationParams oFormCreationParams;
                    XmlDocument oXmlDataDocument = new XmlDocument();
                    oXmlDataDocument.Load(System.Windows.Forms.Application.StartupPath + @"/FormulariosXml/AsisLiqRegaliasAut.xml");
                    oFormCreationParams = (SAPbouiCOM.FormCreationParams)SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                    oFormCreationParams.XmlData = oXmlDataDocument.InnerXml;
                    oForm = SBO_Application.Forms.AddEx(oFormCreationParams);
                    oForm.Visible = true;

                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        public void ManejarEventosItem(ref ItemEvent pVal, ref bool BubbelEvent)
        {
            if (pVal.BeforeAction)
            {
                this.oForm = SBO_Application.Forms.Item(pVal.FormUID);                
            }
            else
            {
                if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED)
                {
                    if (pVal.ItemUID.Equals("btnSig"))
                    {
                        switch (oForm.PaneLevel)
                        {
                            case 1:
                                if (validarCampos())
                                {
                                    cargarGridLiquidacion();
                                    ((StaticText)oForm.Items.Item("lblFecIni").Specific).Caption = "Fecha desde: " + oForm.DataSources.UserDataSources.Item("udFecIni").Value;
                                    ((StaticText)oForm.Items.Item("lblFecFin").Specific).Caption = "Fecha hasta: " + oForm.DataSources.UserDataSources.Item("udFecFin").Value;
                                    oForm.Items.Item("btnAnt").Visible = true;
                                    oForm.PaneLevel = 2;
                                }
                                break;
                            case 2:
                                cargarGridLiquidacionResumida();
                                ((StaticText)oForm.Items.Item("lblCab").Specific).Caption = "Resumen de liquidación autor";
                                ((Button)oForm.Items.Item("btnSig").Specific).Caption = "Contabilizar";
                                oForm.PaneLevel = 3;
                                break;
                            case 3:
                                if (string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("udFecCont").Value))
                                {
                                    SBO_Application.StatusBar.SetText("Debe seleccionar la fecha de contabilización", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    oForm.Items.Item("txtFecCont").Click(BoCellClickType.ct_Regular);
                                }
                                else
                                {                                    
                                    if (SBO_Application.MessageBox("Este proceso es irrebersible, ¿desa continuar?", 1, "Continuar", "Cancelar") == 1)
                                    {
                                        if(contabilizar())
                                            actualizarDocumentos();
                                        oForm.Items.Item("2").Visible = false;
                                        oForm.Items.Item("btnAnt").Visible = false;
                                        ((Button)oForm.Items.Item("btnSig").Specific).Caption = "Finalizar";
                                        oForm.PaneLevel = 4;
                                    }
                                }
                                break;
                            case 4:
                                oForm.Close();
                                break;
                        }
                    }
                    else if(pVal.ItemUID.Equals("btnAnt"))
                    {
                        switch (oForm.PaneLevel)
                        {
                            case 2:
                                ((StaticText)oForm.Items.Item("lblCab").Specific).Caption = "Resumen de liquidación autor - libro";
                                oForm.Items.Item("btnAnt").Visible = false;
                                oForm.PaneLevel = 1;
                                break;
                            case 3:
                                ((Button)oForm.Items.Item("btnSig").Specific).Caption = "Siguiente";
                                oForm.PaneLevel = 2;
                                break;
                        }
                    }
                }
            }
        }

        private bool validarCampos()
        {
            bool resultado = true;
            if (string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("udFecIni").Value))
            {
                SBO_Application.StatusBar.SetText("Debe seleccionar la fecha inicial", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Items.Item("txtFecIni").Click(BoCellClickType.ct_Regular);
                resultado = false;
            }
            else if (string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("udFecFin").Value))
            {
                SBO_Application.StatusBar.SetText("Debe seleccionar la fecha final", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Items.Item("txtFecFin").Click(BoCellClickType.ct_Regular);
                resultado = false;
            }

            return resultado;
        }

        private void cargarGridLiquidacion()
        {
            SQL.SQL sql = new SQL.SQL("LiquidacionAutor.SQL.GetLiquidacion.sql");
            DateTime fechaInicial = DateTime.ParseExact(oForm.DataSources.UserDataSources.Item("udFecIni").Value, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            DateTime fechaFinal = DateTime.ParseExact(oForm.DataSources.UserDataSources.Item("udFecFin").Value, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            var grCuotas = (Grid)oForm.Items.Item("grLiqui").Specific;
            grCuotas.DataTable.ExecuteQuery(string.Format(sql.getQuery(), fechaInicial.ToString("yyyyMMdd"), fechaFinal.ToString("yyyyMMdd")));


            grCuotas.Columns.Item("Autor").Editable = false;
            grCuotas.Columns.Item("ItemCode").Editable = false;
            grCuotas.Columns.Item("Dscription").Editable = false;
            grCuotas.Columns.Item("DocEntry").Editable = false;
            grCuotas.Columns.Item("DocEntry").Visible = false;
            grCuotas.Columns.Item("DocNum").Editable = false;
            grCuotas.Columns.Item("U_Porct_Obra").Editable = false;
            grCuotas.Columns.Item("DocDate").Editable = false;
            grCuotas.Columns.Item("DocCur").Editable = false;
            grCuotas.Columns.Item("Quantity").Editable = false;
            grCuotas.Columns.Item("LineTotal").Editable = false;
            grCuotas.Columns.Item("U_Porct_Cal").Editable = false;
            grCuotas.Columns.Item("MontoRegalia").Editable = false;
            grCuotas.Columns.Item("ObjType").Visible = false;
            grCuotas.Columns.Item("ObjType").Editable = false;
            grCuotas.Columns.Item("Tipo").Editable = false;
            grCuotas.Columns.Item("FolioPref").Editable = false;
            grCuotas.Columns.Item("FolioNum").Editable = false;
            grCuotas.Columns.Item("FolioPref").Visible = false;
            grCuotas.Columns.Item("FolioNum").Visible = false;
            grCuotas.Columns.Item("PTICode").Editable = false;
            grCuotas.Columns.Item("Letter").Editable = false;
            grCuotas.Columns.Item("FolNumFrom").Editable = false;
            grCuotas.Columns.Item("FolNumTo").Editable = false;
            grCuotas.Columns.Item("PTICode").Visible = false;
            grCuotas.Columns.Item("Letter").Visible = false;
            grCuotas.Columns.Item("FolNumFrom").Visible = false;
            grCuotas.Columns.Item("FolNumTo").Visible = false;
            grCuotas.Columns.Item("TotalFrgn").Editable = false;


            switch (Pais)
            {
                case Constantes.Chile:
                    grCuotas.Columns.Item("FolioPref").Visible = true;
                    grCuotas.Columns.Item("FolioNum").Visible = true;
                    break;
                case Constantes.Argentina:
                    grCuotas.Columns.Item("PTICode").Visible = true;
                    grCuotas.Columns.Item("Letter").Visible = true;
                    grCuotas.Columns.Item("FolNumFrom").Visible = true;
                    grCuotas.Columns.Item("FolNumTo").Visible = true;
                    break;
            }

            ((EditTextColumn)grCuotas.Columns.Item("ItemCode")).TitleObject.Caption = "Código de libro";
            ((EditTextColumn)grCuotas.Columns.Item("Dscription")).TitleObject.Caption = "Título";
            ((EditTextColumn)grCuotas.Columns.Item("U_Porct_Obra")).TitleObject.Caption = "% Obra";
            ((EditTextColumn)grCuotas.Columns.Item("DocEntry")).TitleObject.Caption = "N° Interno documento";
            ((EditTextColumn)grCuotas.Columns.Item("DocNum")).TitleObject.Caption = "N° documento";
            ((EditTextColumn)grCuotas.Columns.Item("DocDate")).TitleObject.Caption = "Fecha de venta";
            ((EditTextColumn)grCuotas.Columns.Item("DocCur")).TitleObject.Caption = "Moneda";
            ((EditTextColumn)grCuotas.Columns.Item("Quantity")).TitleObject.Caption = "Cantidad de venta";
            ((EditTextColumn)grCuotas.Columns.Item("LineTotal")).TitleObject.Caption = "Importe de documento (ML)";
            ((EditTextColumn)grCuotas.Columns.Item("U_Porct_Cal")).TitleObject.Caption = "Porcentaje de regalia";
            ((EditTextColumn)grCuotas.Columns.Item("MontoRegalia")).TitleObject.Caption = "Monto de regalia (ML)";
            ((EditTextColumn)grCuotas.Columns.Item("Tipo")).TitleObject.Caption = "Tipo de documento";
            ((EditTextColumn)grCuotas.Columns.Item("FolioPref")).TitleObject.Caption = "Prefijo de folio";
            ((EditTextColumn)grCuotas.Columns.Item("FolioNum")).TitleObject.Caption = "Numero de folio";
            ((EditTextColumn)grCuotas.Columns.Item("PTICode")).TitleObject.Caption = "Punto de emisión";
            ((EditTextColumn)grCuotas.Columns.Item("Letter")).TitleObject.Caption = "Letra";
            ((EditTextColumn)grCuotas.Columns.Item("FolNumFrom")).TitleObject.Caption = "Número de desde";
            ((EditTextColumn)grCuotas.Columns.Item("FolNumTo")).TitleObject.Caption = "Número de hasta";
            ((EditTextColumn)grCuotas.Columns.Item("TotalFrgn")).TitleObject.Caption = "Importe ME";

            grCuotas.CollapseLevel = 1;
            ((EditTextColumn)grCuotas.Columns.Item("Quantity")).ColumnSetting.SumType = BoColumnSumType.bst_Auto;
            ((EditTextColumn)grCuotas.Columns.Item("MontoRegalia")).ColumnSetting.SumType = BoColumnSumType.bst_Auto;
        }

        private void cargarGridLiquidacionResumida()
        {
            SQL.SQL sql = new SQL.SQL("LiquidacionAutor.SQL.GetLiquidacionResumida.sql");
            DateTime fechaInicial = DateTime.ParseExact(oForm.DataSources.UserDataSources.Item("udFecIni").Value, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            DateTime fechaFinal = DateTime.ParseExact(oForm.DataSources.UserDataSources.Item("udFecFin").Value, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            oForm.DataSources.DataTables.Item("dtLiquiR").ExecuteQuery(string.Format(sql.getQuery(), fechaInicial.ToString("yyyyMMdd"), fechaFinal.ToString("yyyyMMdd")));

            var grid = (Grid)oForm.Items.Item("grLiquiR").Specific;

            grid.Columns.Item("U_Code_Autor").Editable = false;
            grid.Columns.Item("U_Name_Autor").Editable = false;
            grid.Columns.Item("FechaVencimiento").Editable = false;
            grid.Columns.Item("Cantidad").Editable = false;
            grid.Columns.Item("ValorVendido").Editable = false;
            grid.Columns.Item("MontoRegalia").Editable = false;

            ((EditTextColumn)grid.Columns.Item("U_Code_Autor")).TitleObject.Caption = "Código de autor";
            ((EditTextColumn)grid.Columns.Item("U_Name_Autor")).TitleObject.Caption = "Nombre de autor";
            ((EditTextColumn)grid.Columns.Item("FechaVencimiento")).TitleObject.Caption = "Fecha de vencimiento";
            ((EditTextColumn)grid.Columns.Item("Cantidad")).TitleObject.Caption = "Cantidad vendida";
            ((EditTextColumn)grid.Columns.Item("ValorVendido")).TitleObject.Caption = "Valor vendido (ML)";
            ((EditTextColumn)grid.Columns.Item("MontoRegalia")).TitleObject.Caption = "Importe para pago (ML)";

            ((EditTextColumn)grid.Columns.Item("U_Code_Autor")).LinkedObjectType = "2";

            ((EditTextColumn)grid.Columns.Item("MontoRegalia")).ColumnSetting.SumType = BoColumnSumType.bst_Auto;

            oForm.DataSources.UserDataSources.Item("udFecCont").Value = DateTime.Today.ToString("dd/MM/yyyy");
        }

        private bool contabilizar()
        {

            bool retorno = true;
            var dtLiquiR = oForm.DataSources.DataTables.Item("dtLiquiR");

            var dtResult = oForm.DataSources.DataTables.Item("dtResult");
            try
            {
                dtResult.Clear();
                dtResult.Columns.Add("TransId", BoFieldsType.ft_AlphaNumeric);
                dtResult.Columns.Add("Descripcion", BoFieldsType.ft_AlphaNumeric);
                dtResult.Rows.Add();

                var cuentas = consultarCuentasRegalias();

                SAPbobsCOM.JournalEntries oJournal = (SAPbobsCOM.JournalEntries)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                DateTime fechaContabilizacion = DateTime.ParseExact(oForm.DataSources.UserDataSources.Item("udFecCont").Value, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                DateTime fechaInicial = DateTime.ParseExact(oForm.DataSources.UserDataSources.Item("udFecIni").Value, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                oJournal.ReferenceDate = fechaContabilizacion;
                oJournal.Memo = "Asiento derechos de autor " + fechaInicial.Month + "/" + fechaInicial.Year;

                double total = 0;
                for (int i = 0; i < dtLiquiR.Rows.Count; i++)
                {
                    oJournal.Lines.AccountCode = cuentas[1];
                    oJournal.Lines.Reference1 = dtLiquiR.GetValue("U_Code_Autor", i).ToString();
                    oJournal.Lines.Reference2 = dtLiquiR.GetValue("U_Name_Autor", i).ToString();
                    oJournal.Lines.Credit = (double)dtLiquiR.GetValue("MontoRegalia", i);
                    oJournal.Lines.Debit = 0;
                    oJournal.Lines.Add();
                    total += (double)dtLiquiR.GetValue("MontoRegalia", i);
                }

                oJournal.Lines.AccountCode = cuentas[0];
                oJournal.Lines.Credit = 0;
                oJournal.Lines.Debit = total;
                oJournal.Lines.Add();

                if (oJournal.Add() != 0)
                {
                    dtResult.SetValue("Descripcion", 0, "Error creando asiento - " + oCompany.GetLastErrorDescription());
                    retorno = false;
                }
                else
                {
                    dtResult.SetValue("TransId", 0, oCompany.GetNewObjectKey());
                    dtResult.SetValue("Descripcion", 0, "Asiento creado correctamente");
                }
            }
            catch (Exception ex)
            {

                dtResult.SetValue("Descripcion", 0, "Error creando asiento - " + ex.Message);
            }
            
            var grResult = (Grid)oForm.Items.Item("grResult").Specific;

            grResult.Columns.Item("TransId").Editable = false;
            grResult.Columns.Item("Descripcion").Editable = false;

            ((EditTextColumn)grResult.Columns.Item("TransId")).LinkedObjectType = "30";

            ((EditTextColumn)grResult.Columns.Item("TransId")).TitleObject.Caption = "Asiento";
            ((EditTextColumn)grResult.Columns.Item("Descripcion")).TitleObject.Caption = "Resultado";

            grResult.AutoResizeColumns();

            return retorno;
            
        }

        private string[] consultarCuentasRegalias()
        {
            string[] cuenta = new string[2];
            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SQL.SQL sql = new SQL.SQL("LiquidacionAutor.SQL.GetCuentasRegalias.sql");
            oRecordset.DoQuery(string.Format(sql.getQuery()));

            if (oRecordset.RecordCount > 0)
            {
                cuenta[0] = oRecordset.Fields.Item("U_HCO_LostAcct").Value.ToString();
                cuenta[1] = oRecordset.Fields.Item("U_HCO_PasiAcct").Value.ToString();
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            oRecordset = null;
            GC.Collect();

            return cuenta;
        }

        private void actualizarDocumentos()
        {
            var dtLiqui = oForm.DataSources.DataTables.Item("dtLiqui");

            for(int i = 0; i < dtLiqui.Rows.Count; i++)
            {
                if(dtLiqui.GetValue("DocEntry", i) != null)
                {
                    SAPbobsCOM.Documents oDocument = (SAPbobsCOM.Documents)oCompany.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Enum.Parse(typeof(SAPbobsCOM.BoObjectTypes), dtLiqui.GetValue("ObjType", i)));
                    oDocument.GetByKey((int)dtLiqui.GetValue("DocEntry", i));
                    oDocument.UserFields.Fields.Item("U_HCO_Liquidated").Value = "S";
                    if(oDocument.Update() != 0)
                    {
                        throw new Exception(oCompany.GetLastErrorDescription());
                    }
                }
            }
        }
    }
}
