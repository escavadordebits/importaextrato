using ExcelDataReader;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ImportarExtrato
{
    [FormAttribute("385", "SystemForm1.b1f")]
    class SystemForm1 : SystemFormBase
    {


        public SystemForm1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("18").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_0").Specific));          
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Widget Widget;
        private SAPbouiCOM.Form form;
        private SAPbouiCOM.EditText EditText0;

        private void OnCustomInitialize()
        {

        }
        [STAThread]

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        
        {
            BubbleEvent = true;

//            SAPbobsCOM.Company oCompany = (SAPbobsCOM.Company)SAPbouiCOM.Framework.Application.SBO_Application.Company.GetDICompany();
  
            SAPbobsCOM.BankPages oBankPages = (SAPbobsCOM.BankPages)Program.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBankPages);
            string sConta = Convert.ToString(EditText0.Value);

            Thread t = new Thread(() =>
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                DialogResult dr = openFileDialog.ShowDialog();

                if (dr == DialogResult.OK)
                {
                    string fileName = openFileDialog.FileName;

                    using (var stream = System.IO.File.Open(fileName, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            var result = reader.AsDataSet();

                            if (result != null)
                            {
                                var table = result.Tables[0];

                                if (table != null)
                                {
                                    int i = 0;
                                    double vlcred = 0;
                                    foreach (DataRow row in table.Rows)
                                    {
                                        i++;
                                        if (i >= 5)
                                        {
                                            var arrayCells = row.ItemArray;

                                            oBankPages.AccountCode = sConta;
                                            oBankPages.ExternalCode = Convert.ToString(arrayCells[3]);
                                            oBankPages.Memo= Convert.ToString(arrayCells[2]);
                                            oBankPages.StatementNumber = Convert.ToInt32(arrayCells[3]);

                                            vlcred = Convert.ToDouble(arrayCells[4]);
                                            if (vlcred < 0)
                                            {

                                                oBankPages.CreditAmount = (Convert.ToDouble(arrayCells[4]) * -1);
                                                  oBankPages.DebitAmount = 0;

                                            }
                                            else
                                            {
                                                oBankPages.DebitAmount = Convert.ToDouble(arrayCells[4]) ;
                                                oBankPages.CreditAmount = 0;
                                            }

                                            //oBankPages.DebitAmount = 0;
                                            oBankPages.DueDate = Convert.ToDateTime(arrayCells[0]);

                                           // oBankPages.StatementNumber = 168;

                                            int oReturn = oBankPages.Add();
                                           
                                            if (oReturn != 0)
                                            {
                                                string msgerro = Program.Company.GetLastErrorDescription();

                                                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(msgerro);

                                            }
                                            else
                                            {
                                                // SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Extrato Importado");
                                               SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Linha  do extrato importada", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                                            }

                                        }

                                    }
                                    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Extrato Importado", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                                }
                            }
                        }
                    }

                }
            });
            t.IsBackground = true;
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
       
        }




    }
}
