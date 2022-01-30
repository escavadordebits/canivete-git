using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DIAPI = SAPbobsCOM;

namespace Canivete
{
    
    public partial class Config : Form
    {
        public Config()
        {
            InitializeComponent();
        }

        public static DIAPI.Company oCompany = new DIAPI.Company();

        private void button1_Click(object sender, EventArgs e)
        {

           
            if (oCompany.Connected)
            {
                DIAPI.UserTablesMD OUserTable = oCompany.GetBusinessObject(DIAPI.BoObjectTypes.oUserTables);

                OUserTable.TableName = "Reembolso" ;
                OUserTable.TableDescription = "ReembolsoCab";
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

                OUserTable.TableName = "ReembolsoLinha";
                OUserTable.TableDescription = "ReembolsoLin";
                OUserTable.TableType = DIAPI.BoUTBTableType.bott_MasterDataLines;

                int resplinha = OUserTable.Add();


                if (resplinha == 0)
                {
                    MessageBox.Show("tabela linha Add com sucess");


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
    }
}
