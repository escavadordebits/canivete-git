using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using DIAPI = SAPbobsCOM;

namespace Canivete
{


    public partial class Form1 : Form
    {
        public static DIAPI.Company oCompany = new DIAPI.Company();


        public Form1()
        {
            InitializeComponent();
        }

        private void BtnConectar_Click(object sender, EventArgs e)
        {
            ConectarSAP();
        }

        public bool ConectarSAP()
        {

            oCompany.SLDServer = "https://srvevoazsap01:40000";
            oCompany.Server = "SRVEVOAZDB01";
            oCompany.language = DIAPI.BoSuppLangs.ln_Portuguese_Br;
            oCompany.DbServerType = DIAPI.BoDataServerTypes.dst_MSSQL2017;
            oCompany.DbUserName = "svc_dbsap ";
            oCompany.DbPassword = "svc_8@#rskT@3!";
           // oCompany.CompanyDB = "SBO_DESENV_EVO";
            oCompany.CompanyDB = "Lessoc";
            oCompany.UserName = "manager";
            oCompany.Password = "Evo@09";



            int retornosap = oCompany.Connect();

            if (retornosap == 0)
            {
                MessageBox.Show("Conectado com Sucesso");
                return oCompany.Connected;

            }
            else
            {

                string  ErrMsg = oCompany.GetLastErrorDescription();
                MessageBox.Show(ErrMsg);
                return false;
            }
            
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            

            if (oCompany.Connected)

            {

                DIAPI.BusinessPartners oBP = oCompany.GetBusinessObject(DIAPI.BoObjectTypes.oBusinessPartners);

                
                oBP.CardCode = "TestePN2";     // Convert.ToString(txtCardCode.Text);
                oBP.CardName = "TEstes Nome";                //Convert.ToString(txtCardName.Text);
                oBP.UserFields.Fields.Item("U_LocColEntg").Value = "Teste";
                
                int resp = oBP.Add();
                string msg;
                if (resp == 0)
                {
                    MessageBox.Show("Cadastrado Com Sucesso");
                }
                else
                {
                    msg = oCompany.GetLastErrorDescription();
                    MessageBox.Show(msg);

                }


            }else
            {
                MessageBox.Show("Favor Conectar");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (oCompany.Connected)
            {

                DIAPI.BusinessPartners oBP = oCompany.GetBusinessObject(DIAPI.BoObjectTypes.oBusinessPartners);

               // oBP.CardCode = Convert.ToString(txtCardCode.Text);
                
                oBP.GetByKey("TestePN2");

               // oBP.GetByKey(Convert.ToString(txtCardCode.Text));

                int resp = oBP.Remove();

                string msg;

                if (resp == 0)
                {
                    MessageBox.Show("Removido Com Sucesso");
                }
                else
                {
                    msg = oCompany.GetLastErrorDescription();
                    MessageBox.Show(msg);

                }


            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ConectarSAP();

            if (oCompany.Connected)
            {

                DIAPI.BusinessPartners oBP = oCompany.GetBusinessObject(DIAPI.BoObjectTypes.oBusinessPartners);

                oBP.CardCode = Convert.ToString(txtCardCode.Text);

                // oBP.GetByKey(Convert.ToString(txtCardCode.Text));

                oBP.GetByKey("TestePN2");

                oBP.Valid = DIAPI.BoYesNoEnum.tNO;
                oBP.Frozen = DIAPI.BoYesNoEnum.tYES;

                

                int resp = oBP.Update();

                string msg;

                if (resp == 0)
                {
                    MessageBox.Show("Inativo Com Sucesso");
                }
                else
                {
                    msg = oCompany.GetLastErrorDescription();
                    MessageBox.Show(msg);

                }


            }

        }

        private void button4_Click(object sender, EventArgs e)
        {


            if (oCompany.Connected)
            {

                DIAPI.Items OItem = oCompany.GetBusinessObject(DIAPI.BoObjectTypes.oItems);

                

                OItem.ItemCode = Convert.ToString(txtItem.Text);
                OItem.ItemName = Convert.ToString(txtItemName.Text);
                int resp = OItem.Add();

                if (resp == 0)
                {
                    MessageBox.Show("Item cadatrado com Sucesso");
                    oCompany.Disconnect();

                }
                else
                {
                    string msg = oCompany.GetLastErrorDescription();
                    MessageBox.Show(msg);
                }

            }
            else
            {
                string msg = oCompany.GetLastErrorDescription();
                MessageBox.Show("Desconectado - Clique em conectar" + msg);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {

            if (oCompany.Connected)
            {

                DIAPI.Items OItem = oCompany.GetBusinessObject(DIAPI.BoObjectTypes.oItems);
                OItem.ItemCode = Convert.ToString(txtItem.Text);
                OItem.GetByKey(Convert.ToString(txtItem.Text));
                int resp = OItem.Remove();

                if (resp == 0)
                {
                    MessageBox.Show("Item removido com Sucesso");
                    oCompany.Disconnect();
                }
                else
                {
                    string msg = oCompany.GetLastErrorDescription();
                    MessageBox.Show(msg);
                }
            }
            else
            {
                string msg = oCompany.GetLastErrorDescription();
                MessageBox.Show("Desconectado - Clique em conectar");
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {

            if (oCompany.Connected)
            {

                DIAPI.Items OItem = oCompany.GetBusinessObject(DIAPI.BoObjectTypes.oItems);
                OItem.ItemCode = Convert.ToString(txtItem.Text);
                OItem.GetByKey(Convert.ToString(txtItem.Text));

                OItem.Valid = DIAPI.BoYesNoEnum.tNO;
                OItem.Frozen = DIAPI.BoYesNoEnum.tYES;


                int resp = OItem.Update();

                if (resp == 0)
                {
                    MessageBox.Show("Item Inativo com Sucesso");
                    oCompany.Disconnect();
                }
                else
                {
                    string msg = oCompany.GetLastErrorDescription();
                    MessageBox.Show(msg);
                }
            }
            else
            {
                string msg = oCompany.GetLastErrorDescription();
                MessageBox.Show("Desconectado - Clique em conectar" + msg);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {



        }

        private void button7_Click(object sender, EventArgs e)
        {



            if (oCompany.Connected)
            {

                DIAPI.Documents oPC = oCompany.GetBusinessObject(DIAPI.BoObjectTypes.oPurchaseInvoices);
                
                oPC.CardCode = Convert.ToString(txtPnComp.Text);
                oPC.DocDate = DateTime.Now;
                
              

                oPC.Lines.ItemCode = Convert.ToString(txtItemcodcomp.Text);
                oPC.Lines.Quantity = Convert.ToDouble(txtqtde.Text);
                oPC.Lines.Price = Convert.ToDouble(txtprice.Text);
                oPC.Lines.Add();
                oPC.Lines.ItemCode = Convert.ToString(txtItemcodcomp.Text);
                oPC.Lines.Quantity = Convert.ToDouble(txtqtde.Text);
                oPC.Lines.Price = Convert.ToDouble(txtprice.Text);

                int resp = oPC.Add();


                if (resp == 0)
                {
                    MessageBox.Show("Pedido de compra add com sucesso");

                }
                else
                {
                    string msg = oCompany.GetLastErrorDescription();
                    MessageBox.Show(msg);
                }


            }
            else
            {
                string msg = oCompany.GetLastErrorDescription();
                MessageBox.Show("Desconectado - Clique em conectar" + msg);

            }


        }


        public void  button8_Click_1(object sender, EventArgs e)
        {
            

            DIAPI.Recordset oRecSetBuscarPed = oCompany.GetBusinessObject(DIAPI.BoObjectTypes.BoRecordset);

            DIAPI.Documents oPedido = oCompany.GetBusinessObject(DIAPI.BoObjectTypes.oPurchaseOrders);

            string docnum = Convert.ToString(txtNumdoc.Text);

            string query = "select cardcode,cardname from opor where docentry= '" + docnum + "'";

           
            oRecSetBuscarPed.DoQuery(query);

            oPedido.Browser.Recordset = oRecSetBuscarPed;

            if (oRecSetBuscarPed.RecordCount > 0)
            {
                txtforn.Text = Convert.ToString(oPedido.CardCode);
                txtraz.Text = Convert.ToString(oPedido.CardName);
                         
            }

 
        }
       
        private void button9_Click_1(object sender, EventArgs e)
        {

            if (oCompany.Connected)
            {
                DIAPI.UserTablesMD OUserTable = oCompany.GetBusinessObject(DIAPI.BoObjectTypes.oUserTables);

                OUserTable.TableName = Convert.ToString(txtNomeTabela.Text);
                OUserTable.TableDescription = Convert.ToString(txtDescTabela.Text);
                OUserTable.TableType = DIAPI.BoUTBTableType.bott_MasterData;

                int resp = OUserTable.Add();


                if (resp == 0)
                {
                    MessageBox.Show("tabela Add com sucess");


                }
                else
                {
                    string msg = oCompany.GetLastErrorDescription();
                    MessageBox.Show(msg);
                }


            }
            else
            {
                MessageBox.Show("Conectar na base e tente novamente!");
            }


        }

        private void tabPage5_Click(object sender, EventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e)
        {

            DIAPI.UserFieldsMD oUserField = oCompany.GetBusinessObject(DIAPI.BoObjectTypes.oUserFields);

            oUserField.TableName = Convert.ToString(txtNomeTabela.Text);
            oUserField.Name = Convert.ToString(txtNomeCampo.Text);
            oUserField.Description = Convert.ToString(txtDescCampo.Text);


            int resp = oUserField.Add();


            if (resp == 0)
            {
                MessageBox.Show("Campo Add com sucesso");


            }
            else
            {
                string msg = oCompany.GetLastErrorDescription();
                MessageBox.Show(msg);
            }



        }

        private void button11_Click(object sender, EventArgs e)
        {


            DIAPI.GeneralService oGeneralService;
            var oCompService = oCompany.GetCompanyService();

            DIAPI.GeneralData oGeneralData;
            oCompany.StartTransaction();
            oGeneralService = oCompService.GetGeneralService("SDK1");
            oGeneralData =(DIAPI.GeneralData)oGeneralService.GetDataInterface(DIAPI.GeneralServiceDataInterfaces.gsGeneralData);
            
            oGeneralData.SetProperty("Code", Convert.ToString(txtCode.Text));

            oGeneralData.SetProperty("Name", Convert.ToString(txtUserField.Text));
            oGeneralData.SetProperty("U_SDK1C", Convert.ToString(txtUserField.Text));


            oGeneralService.Add(oGeneralData);


            if ( oCompany.InTransaction == true)
            {
                MessageBox.Show("Registro Add com sucesso");
                oCompany.EndTransaction(DIAPI.BoWfTransOpt.wf_Commit);
                
            }
            else
            {
                string msg = oCompany.GetLastErrorDescription();
                MessageBox.Show(msg);
            }



        }

        private void button12_Click(object sender, EventArgs e)
        {
            DIAPI.UserObjectsMD oUDO = oCompany.GetBusinessObject(DIAPI.BoObjectTypes.oUserObjectsMD);

            oUDO.Code = "SDK02";
            oUDO.TableName = "SDK02";
            oUDO.CanCancel = DIAPI.BoYesNoEnum.tYES;
            oUDO.CanFind = DIAPI.BoYesNoEnum.tYES;
            oUDO.ObjectType = DIAPI.BoUDOObjType.boud_MasterData;

            int resp = oUDO.Add();

            if (resp == 0)
            {
                MessageBox.Show("Registrado com sucesso");

            }else
            {
                string msg = oCompany.GetLastErrorDescription();
                MessageBox.Show(msg);

            }




        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void btnProjeto_Click(object sender, EventArgs e)
        {
            ConectarSAP();

            if (oCompany.Connected)
            {
                oCompany.StartTransaction();

                DIAPI.CompanyService cs = oCompany.GetCompanyService();
                
                DIAPI.ProjectsService ServicoProjeto = (DIAPI.ProjectsService)cs.GetBusinessService(DIAPI.ServiceTypes.ProjectsService);
                
                //dados = Project

                DIAPI.Project dadosdoProjeto = (DIAPI.Project)ServicoProjeto.GetDataInterface(DIAPI.ProjectsServiceDataInterfaces.psProject);
                
                DIAPI.ProjectParams chavedoprojeto = (DIAPI.ProjectParams)ServicoProjeto.GetDataInterface(DIAPI.ProjectsServiceDataInterfaces.psProjectParams);
                
                dadosdoProjeto.Code = Convert.ToString(txtCodigoPrj.Text);
                dadosdoProjeto.Name = Convert.ToString(txtDescProj.Text);

                try
                {
                    chavedoprojeto = ServicoProjeto.AddProject(dadosdoProjeto);

                    if (oCompany.InTransaction)
                    {
                        MessageBox.Show("Registro Add com sucesso");
                        oCompany.EndTransaction(DIAPI.BoWfTransOpt.wf_Commit);

                    }
                    else
                    {
                        string msg = oCompany.GetLastErrorDescription();
                        MessageBox.Show(msg);
                    }
                }
                catch (Exception)
                {

                    string msg = oCompany.GetLastErrorDescription();
                    MessageBox.Show(msg);
                }

         

            }


        }

        private void button13_Click(object sender, EventArgs e)
        {
            ConectarSAP();

            if (oCompany.Connected)
            {

               

            }



        }

        private void button14_Click(object sender, EventArgs e)
        {
            ConectarSAP();

            if (oCompany.Connected)
            {
                DIAPI.ServiceCalls oSC = oCompany.GetBusinessObject(DIAPI.BoObjectTypes.oServiceCalls);
                
                oSC.CustomerCode = "C0001";
                oSC.Subject = "teste";

                int resp = oSC.Add();
                string docentry = oCompany.GetNewObjectKey();

                if ( resp == 0)
                {
                    MessageBox.Show("Chamado add Com Sucesso N :" + Convert.ToString(docentry));
                    imprimir(docentry);

                }

            }




        }

        private void imprimir(string docentry)
        {
            DIAPI.CompanyService cs = oCompany.GetCompanyService();
            DIAPI.ReportLayoutsService oReportLayoutService = cs.GetBusinessService(DIAPI.ServiceTypes.ReportLayoutsService);
            DIAPI.ReportLayoutPrintParams oReportPrintParams = oReportLayoutService.GetDataInterface(DIAPI.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutPrintParams);

            oReportPrintParams.LayoutCode = "SCL10004";
            oReportPrintParams.DocEntry = Convert.ToInt32(docentry);

            oReportLayoutService.Print(oReportPrintParams);
        }

        private void button15_Click(object sender, EventArgs e)
        {

            
            DIAPI.CompanyService cs = oCompany.GetCompanyService();
            DIAPI.ApprovalRequestsService approvalSrv = cs.GetBusinessService( DIAPI.ServiceTypes.ApprovalRequestsService);
            DIAPI.ApprovalRequestParams oParams = approvalSrv.GetDataInterface(DIAPI.ApprovalRequestsServiceDataInterfaces.arsApprovalRequestParams);

           approvalSrv.GetApprovalRequestList();
           




            oParams.Code = 2;
            DIAPI.ApprovalRequest oData = approvalSrv.GetApprovalRequest(oParams);

            //Add an approval decision 
            oData.ApprovalRequestDecisions.Add();
            //oData.ApprovalRequestDecisions.Item(0).ApproverUserName = "manager";
            //oData.ApprovalRequestDecisions.Item(0).ApproverPassword = "Evo@09";
            oData.ApprovalRequestDecisions.Item(0).Status =DIAPI.BoApprovalRequestDecisionEnum.ardApproved;
            oData.ApprovalRequestDecisions.Item(0).Remarks = "ok";

            //Update the approval request 
            approvalSrv.UpdateRequest(oData);



        }

        private void button16_Click(object sender, EventArgs e)
        {

            DIAPI.Payments oPayments = oCompany.GetBusinessObject(DIAPI.BoObjectTypes.oVendorPayments);



            oPayments.CardCode = "F0011";
            oPayments.BoeAccount = "2.1.1.01.01"; //oRecSetBolConta.Fields.Item("AcctCode").Value.ToString();
            oPayments.BillofExchangeStatus = 0;
            oPayments.DocDate = DateTime.Now;
            oPayments.BillOfExchange.BillOfExchangeDueDate = DateTime.Now;
            oPayments.BillOfExchange.BillOfExchangeNo = "9990"; //oRecSetBol.Fields.Item("numbol").Value.ToString();
            oPayments.DocDate = DateTime.Now;
            oPayments.BillOfExchange.BillOfExchangeDueDate = DateTime.Now;


//            oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_JournalEntry;

            oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice;
            oPayments.Invoices.DocEntry = 132;
            oPayments.Invoices.InstallmentId = 1; //Convert.ToInt32(((SAPbouiCOM.EditText)Matrix0.Columns.Item("ColParc").Cells.Item(i).Specific).Value);
            oPayments.BillOfExchangeAmount = 100.00;
            oPayments.BillOfExchange.PaymentMethodCode = "Movimento - Bol"; // Convert.ToString(EditText2.Value);  /*"Movimento - Bol";*/ //Corrigir parao novo campo

            if(oPayments.Add() != 0)
            {
                string msg = oCompany.GetLastErrorDescription();
                MessageBox.Show(msg);
            }

        }

        private void button17_Click(object sender, EventArgs e)
        {

            DIAPI.Payments oPayments = oCompany.GetBusinessObject(DIAPI.BoObjectTypes.oVendorPayments);



            oPayments.CardCode = "F0050";
            oPayments.BoeAccount = "2.1.1.01.01"; //oRecSetBolConta.Fields.Item("AcctCode").Value.ToString();
            oPayments.BillofExchangeStatus = 0;
            oPayments.DocDate = DateTime.Now;
            oPayments.BillOfExchange.BillOfExchangeDueDate = DateTime.Now;
            oPayments.BillOfExchange.BillOfExchangeNo = "9996"; //oRecSetBol.Fields.Item("numbol").Value.ToString();
            oPayments.DocDate = DateTime.Now;
            oPayments.BillOfExchange.BillOfExchangeDueDate = DateTime.Now;

            oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseDownPayment;
            oPayments.Invoices.DocEntry = 50;
            oPayments.Invoices.InstallmentId = 1; //Convert.ToInt32(((SAPbouiCOM.EditText)Matrix0.Columns.Item("ColParc").Cells.Item(i).Specific).Value);
            oPayments.BillOfExchangeAmount = 100.00;
            oPayments.BillOfExchange.PaymentMethodCode = "Movimento - Bol"; // Convert.ToString(EditText2.Value);  /*"Movimento - Bol";*/ //Corrigir parao novo campo

            if (oPayments.Add() != 0)
            {
                string msg = oCompany.GetLastErrorDescription();
                MessageBox.Show(msg);
            }
            try
            {

            }
            catch (Exception)
            {

                throw;
            }

        }

        private void button18_Click(object sender, EventArgs e)
        {

            SAPbobsCOM.BillOfExchangeTransaction oBillOfExchangeTransaction = (SAPbobsCOM.BillOfExchangeTransaction)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBillOfExchangeTransactions);

            SAPbobsCOM.Payments oPayments = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);



            try
            {

                //oPayments.GetByKey(240);
                //oPayments.BillofExchangeStatus = DIAPI.BoBoeStatus.boes_Created;
                //oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice;
                oPayments.GetByKey(242);
                oPayments.Invoices.DocEntry = 145;
                oPayments.Invoices.InstallmentId = 1;
                oBillOfExchangeTransaction.Lines.BillOfExchangeType = SAPbobsCOM.BoBOETypes.bobt_Outgoing;
                oBillOfExchangeTransaction.Lines.BillOfExchangeNo = 206101;
                oBillOfExchangeTransaction.StatusFrom = SAPbobsCOM.BoBOTFromStatus.btfs_Generated;
                oBillOfExchangeTransaction.StatusTo = SAPbobsCOM.BoBOTToStatus.btts_Paid;

                oBillOfExchangeTransaction.Lines.Add();
                oBillOfExchangeTransaction.Lines.SetCurrentLine(1);
                
               
                int restbol = oBillOfExchangeTransaction.Add();
                //int restbol = oPayments.Update();

                if (restbol == 0)
                {
                    MessageBox.Show("Baixa com  sucesso");
                }
                else
                {

                    string msgerro = oCompany.GetLastErrorDescription();
                    MessageBox.Show(msgerro);
                }

            }
            catch (Exception t)
            {
                MessageBox.Show(t.Message);
            }
            finally
            {
                if(oBillOfExchangeTransaction != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oBillOfExchangeTransaction);
                }
            }
            

        }

        private void button19_Click(object sender, EventArgs e)
        {
            
            string sXmlFileName = System.IO.Path.GetFileName("config.xml");

            var xmlconfigfile = XDocument.Load(sXmlFileName);

            var xmlconfig = from d in xmlconfigfile.Root.Descendants("config")
                            select new
                            {
                                Licenseserver = d.Element("Licenseserver").Value,
                                ServerSAP = d.Element("ServerSAP").Value,
                                userDB = d.Element("userDB").Value ,
                                passwordDB = d.Element("passwordDB").Value,
                                tiposerver = d.Element("tiposerver").Value,
                                Empresa =d.Element("Empresa").Value,
                                UserSAP = d.Element("UserSAP").Value, 
                                SenhaSAP = d.Element("SenhaSAP").Value
                                
                            };



            foreach (var DadosConfig in xmlconfig)
            {



                oCompany.SLDServer = DadosConfig.Licenseserver;
                oCompany.Server = DadosConfig.ServerSAP;
                oCompany.language = DIAPI.BoSuppLangs.ln_Portuguese_Br;
                oCompany.DbServerType = DIAPI.BoDataServerTypes.dst_MSSQL2017;
                oCompany.DbUserName = DadosConfig.userDB;
                oCompany.DbPassword = DadosConfig.passwordDB;
                oCompany.CompanyDB = DadosConfig.Empresa;
                oCompany.UserName = DadosConfig.UserSAP;
                oCompany.Password = DadosConfig.SenhaSAP;
                /////

                int RetValSAP;

                RetValSAP = oCompany.Connect();


                if (RetValSAP != 0)
                {
                    string ErrMsg = oCompany.GetLastErrorDescription();
                    MessageBox.Show(ErrMsg);

                }
                else
                {
                    MessageBox.Show("Item Adicionado com Sucesso!");
                }





            }




        }
    }
}
