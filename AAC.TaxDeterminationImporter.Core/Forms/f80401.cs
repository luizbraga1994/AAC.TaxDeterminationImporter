using SAPbouiCOM;
using SBO.Hub;
using SBO.Hub.Forms;
using SBO.Hub.Util;
using AAC.TaxDeterminationImporter.Core.BLL;
using AAC.TaxDeterminationImporter.Core.Model;
using System;
using System.Collections.Generic;

namespace AAC.TaxDeterminationImporter.Core.Forms
{
    /// <summary>
    /// Determinação de Código de Imposto
    /// </summary>
    public class f80401 : SystemForm
    {
        public static int SelectedKey;

        public f80401(ItemEvent itemEvent)
        {
            ItemEventInfo = itemEvent;
        }

        public override bool ItemEvent()
        {
            if (!ItemEventInfo.BeforeAction)
            {
                if (ItemEventInfo.EventType == BoEventTypes.et_FORM_LOAD)
                {
                    var itemBase = Form.Items.Item("2002");
                    var it_Import = Form.Items.Add("bt_Import", BoFormItemTypes.it_BUTTON);
                    it_Import.LinkTo = "2002";
                    it_Import.Top = itemBase.Top;
                    it_Import.Width = 80;
                    it_Import.Left = itemBase.Left - 220;
                    it_Import.Height = itemBase.Height;

                    Button bt_Import = (Button)it_Import.Specific;
                    bt_Import.Caption = "Importar";

                    var it_Remove = Form.Items.Add("bt_Remove", BoFormItemTypes.it_BUTTON);
                    it_Remove.LinkTo = "2002";
                    it_Remove.Top = itemBase.Top;
                    it_Remove.Width = 100;
                    it_Remove.Left = itemBase.Left - 120;
                    it_Remove.Height = itemBase.Height;

                    Button bt_Remove = (Button)it_Remove.Specific;
                    bt_Remove.Caption = "Remover Valores";

                    var it_Export = Form.Items.Add("bt_Export", BoFormItemTypes.it_BUTTON);
                    it_Export.LinkTo = "2002";
                    it_Export.Top = itemBase.Top;
                    it_Export.Width = 80;
                    it_Export.Left = itemBase.Left - 310;
                    it_Export.Height = itemBase.Height;

                    Button bt_Export = (Button)it_Export.Specific;
                    bt_Export.Caption = "Exportar";

                    var it_Help = Form.Items.Add("bt_Help", BoFormItemTypes.it_BUTTON);
                    it_Help.LinkTo = "2002";
                    it_Help.Top = itemBase.Top;
                    it_Help.Width = 60;
                    it_Help.Left = itemBase.Left - 380;
                    it_Help.Height = itemBase.Height;

                    Button bt_Help = (Button)it_Help.Specific;
                    bt_Help.Caption = "Ajuda";
                }

                if (ItemEventInfo.EventType == BoEventTypes.et_CLICK)
                {
                    if (ItemEventInfo.ItemUID == "bt_Help")
                    {
                        SBOApp.Application.MessageBox(
                            "AAC Tax Determination Importer - Instruções de Uso\r\n" +
                            "──────────────────────────────────────────────\r\n\r\n" +
                            "IMPORTAR\r\n" +
                            "1. Selecione uma linha na grade de determinação.\r\n" +
                            "2. Clique em 'Importar' e escolha um arquivo Excel (.xlsx).\r\n" +
                            "3. O arquivo deve conter os campos de chave nas colunas A-E,\r\n" +
                            "   datas nas colunas F-G e os códigos de imposto a partir da coluna H.\r\n" +
                            "4. Erros de validação serão exibidos em um log antes da importação.\r\n\r\n" +
                            "EXPORTAR\r\n" +
                            "1. Selecione uma linha na grade de determinação.\r\n" +
                            "2. Clique em 'Exportar' e escolha a pasta de destino.\r\n" +
                            "3. Um arquivo Excel será gerado com todos os valores da linha selecionada.\r\n\r\n" +
                            "REMOVER VALORES\r\n" +
                            "1. Selecione uma linha na grade de determinação.\r\n" +
                            "2. Clique em 'Remover Valores'.\r\n" +
                            "3. Na janela aberta, selecione as linhas a remover e confirme.\r\n\r\n" +
                            "Obs.: As ações de Importar, Exportar e Remover Valores\r\n" +
                            "estão disponíveis apenas com o formulário em modo de edição (OK).",
                            1, "Fechar");
                        return true;
                    }

                    if (ItemEventInfo.ItemUID == "bt_Import" || ItemEventInfo.ItemUID == "bt_Remove" || ItemEventInfo.ItemUID == "bt_Export")
                    {
                        if (Form.Mode == BoFormMode.fm_OK_MODE)
                        {
                            try
                            {
                                ComboBox cb_Type = (ComboBox)Form.Items.Item("2025").Specific;
                                if (cb_Type.Value.Trim() == "1")
                                {
                                    Matrix mt_Tax = (Matrix)Form.Items.Item("2003").Specific;
                                    SelectedKey = mt_Tax.GetNextSelectedRow();
                                    if (SelectedKey != -1)
                                    {
                                        if (ItemEventInfo.ItemUID == "bt_Import")
                                        {
                                            this.Import(SelectedKey);
                                        }
                                        else if(ItemEventInfo.ItemUID == "bt_Export")
                                        {
                                            this.Export(SelectedKey, ((EditText)mt_Tax.GetCellSpecific("2005", SelectedKey)).Value);
                                        }
                                        else if (ItemEventInfo.ItemUID == "bt_Remove")
                                        {
                                            this.Remove(SelectedKey);
                                        }
                                    }
                                    else
                                    {
                                        SBOApp.Application.SetStatusBarMessage("Nenhuma linha selecionada");
                                    }
                                }
                                else
                                {
                                    SBOApp.Application.SetStatusBarMessage("Ação indisponível para o tipo de determinação informada");
                                }
                            }
                            catch (Exception ex)
                            {
                                SBOApp.Application.SetStatusBarMessage(ex.Message);
                            }
                        }
                        else
                        {
                            SBOApp.Application.SetStatusBarMessage("Por favor, salve os dados antes de continuar");
                        }
                    }
                    //if (ItemEventInfo.ItemUID == "2003")
                    //{
                    //    Matrix mt_Tax = (Matrix)Form.Items.Item("2003").Specific;
                    //    SelectedKey = mt_Tax.GetNextSelectedRow();
                    //}
                }
            }

            return true;
        }

        private void Import(int selectedRow)
        {
            DialogUtil dialogUtil = new DialogUtil();
            string filePath = dialogUtil.OpenFileDialog("Arquivo Excel|*.xlsx");
            if (!String.IsNullOrEmpty(filePath))
            {
                TaxDeterminationBLL taxDeterminationBLL = new TaxDeterminationBLL();
                List<TaxDeterminationModel> list = taxDeterminationBLL.ImportData(filePath, selectedRow);
                if (list != null)
                {
                    FrmImportLog frmImportLog = new FrmImportLog();
                    frmImportLog.Show(list);
                }
                else
                {
                    SBOApp.Application.MessageBox("Importação finalizada com sucesso");
                    Form.Close();
                    SBOApp.Application.ActivateMenuItem("8497");
                }
            }
        }

        private void Export(int selectedRow, string description)
        {
            DialogUtil dialogUtil = new DialogUtil();
            string path = dialogUtil.FolderBrowserDialog();
            if (!String.IsNullOrEmpty(path))
            {
                TaxDeterminationBLL taxDeterminationBLL = new TaxDeterminationBLL();
                string error = taxDeterminationBLL.ExportData(path, selectedRow, description);
                if (!String.IsNullOrEmpty(error))
                {
                    SBOApp.Application.SetStatusBarMessage(error);
                }
                else
                {
                    int answer = SBOApp.Application.MessageBox("Exportação finalizada com sucesso. Deseja abrir o arquivo gerado?", 1, "Sim", "Não");
                    if (answer == 1)
                    {
                        System.Diagnostics.Process.Start(System.IO.Path.Combine(path, selectedRow + " - " + description + ".xlsx"));
                    }
                }
            }
        }

        private void Remove(int selectedRow)
        {
            Matrix mt_Tax = (Matrix)Form.Items.Item("2003").Specific;
            string field1 = ((ComboBox)mt_Tax.GetCellSpecific("2001", selectedRow)).Selected.Description;
            string field2 = ((ComboBox)mt_Tax.GetCellSpecific("2002", selectedRow)).Selected.Description;
            string field3 = ((ComboBox)mt_Tax.GetCellSpecific("2003", selectedRow)).Selected.Description;
            string field4 = ((ComboBox)mt_Tax.GetCellSpecific("1320002007", selectedRow)).Selected.Description;
            string field5 = ((ComboBox)mt_Tax.GetCellSpecific("1320002008", selectedRow)).Selected.Description;
            string description = ((EditText)mt_Tax.GetCellSpecific("2005", selectedRow)).Value;

            FrmRemoveTax frmRemoveTax = new FrmRemoveTax();
            frmRemoveTax.Show(selectedRow, field1, field2, field3, field4, field5, description);
        }
    }
}
