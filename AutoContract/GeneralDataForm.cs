using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AutoContract
{
    public partial class GeneralDataForm : Form
    {
        
        private string contrType = "";
        WordMethods _wm;
        public GeneralDataForm()
        {
            InitializeComponent();
        }

        public GeneralDataForm(string _contrType,string _pathToPatternFile,WordMethods wm)
        {
            InitializeComponent();
            contrType = _contrType;
            _wm = wm;
        }

        private void nextButton_Click(object sender, EventArgs e)
        {
            string[] date = Contract.getPartsOfDate(dateTextBox.Text);
            //_wm.replace("<DAY>", date[0]);
            //_wm.replace("<MONTHWORD>", date[1]); ;
            //_wm.replace("<YEAR>", date[2]);

            Contract.replaceString(_wm, "<DAY>", date[0]);
            Contract.replaceString(_wm, "<MONTHWORD>", date[1]);
            Contract.replaceString(_wm, "<YEAR>", date[2]);
            Contract.replaceString(_wm, "<NUMBER>", contrNumbTextBox.Text);
            Contract.replaceString(_wm, "<DATE>", dateTextBox.Text);
            Contract.replaceString(_wm, "<COMPLDATE>", complDateTextBox.Text);
            Contract.replaceString(_wm, "<FULLFIRMNAME>", fullFirmNameTextBox.Text);
            Contract.replaceString(_wm, "<SHORTFIRMNAME>", shortNameFirmTextBox.Text);
            Contract.replaceString(_wm, "<FULLDIRNAME>", directorNameTextBox.Text);
            Contract.replaceString(_wm, "<SHORTDIRNAME>", Contract.getShortDirName(directorNameTextBox.Text));
            Contract.replaceString(_wm, "<CUSTADDR>", customerAddresstextBox.Text.Equals("") ? "" : "Юр. адрес: " + customerAddresstextBox.Text + ", ");
            Contract.replaceString(_wm, "<PHYSADDR>", physicalCustomerAddressTextBox.Text.Equals("") ? "" : "факт. адрес: " + physicalCustomerAddressTextBox.Text + ", ");
            Contract.replaceString(_wm, "<COST>", costTextBox.Text);
            Contract.replaceString(_wm, "<INN>", INNtextBox.Text.Equals("") ? "" : "ИНН " + INNtextBox.Text + ", ");
            Contract.replaceString(_wm, "<OGRN>", OGRNTextBox.Text.Equals("") ? "" : "ОГРН " + OGRNTextBox.Text + ", ");
            Contract.replaceString(_wm, "<KPP>", KPPTextBox.Text.Equals("") ? "" : "КПП " + KPPTextBox.Text + ", ");
            Contract.replaceString(_wm, "<CORACC>", corAcTextBox.Text.Equals("") ? "" : "кор/сч " + corAcTextBox.Text + ", ");
            Contract.replaceString(_wm, "<BIK>", "БИК " + BIKTextBox.Text + ".");
            Contract.replaceString(_wm, "<PHONE>", telephoneTextBox.Text.Equals("") ? "" : "тел. " + telephoneTextBox.Text + ", ");
            Contract.replaceString(_wm, "<FAX>", faxTextBox.Text.Equals("") ? "" : "факс " + faxTextBox.Text + ", ");
            Contract.replaceString(_wm, "<EMAIL>", emailTextBox.Text.Equals("") ? "" : "email: " + emailTextBox.Text + ", ");
            Contract.replaceString(_wm, "<CHACC>", chAcTextBox.Text.Equals("") ? "" : "р/сч " + chAcTextBox.Text + ", ");
            Contract.replaceString(_wm, "<WHERECHACC>", whereChAcTextBox.Text.Equals("") ? "" : whereChAcTextBox.Text);


           

            //_wm.replace("<NUMBER>", contrNumbTextBox.Text);
        
            //_wm.replace("<DATE>", dateTextBox.Text);
            //_wm.replace("<COMPLDATE>", complDateTextBox.Text);
  
            //_wm.replace("<FULLFIRMNAME>", fullFirmNameTextBox.Text);
            //_wm.replace("<SHORTFIRMNAME>", shortNameFirmTextBox.Text);
            //_wm.replace("<FULLDIRNAME>", directorNameTextBox.Text);
            //_wm.replace("<SHORTDIRNAME>", Contract.getShortDirName(directorNameTextBox.Text));
            //_wm.replace("<CUSTADDR>", customerAddresstextBox.Text.Equals("") ? "" : "Юр. адрес: " +  customerAddresstextBox.Text + ", ");
            //_wm.replace("<PHYSADDR>", physicalCustomerAddressTextBox.Text.Equals("") ? "" : "факт. адрес: " + physicalCustomerAddressTextBox.Text + ", ");
            //_wm.replace("<COST>", costTextBox.Text);
            //_wm.replace("<INN>", INNtextBox.Text.Equals("") ? "" : "ИНН " + INNtextBox.Text + ", ");
            //_wm.replace("<OGRN>", OGRNTextBox.Text.Equals("") ? "" : "ОГРН " + OGRNTextBox.Text + ", ");
            //_wm.replace("<KPP>", KPPTextBox.Text.Equals("") ? "" : "КПП " + KPPTextBox.Text + ", ");
            //_wm.replace("<CORACC>", corAcTextBox.Text.Equals("") ? "" : "кор/сч " + corAcTextBox.Text + ", ");
            //_wm.replace("<BIK>", "БИК " + BIKTextBox.Text + ".");
            //_wm.replace("<PHONE>", telephoneTextBox.Text.Equals("") ? "" : "тел. " + telephoneTextBox.Text + ", ");
            //_wm.replace("<FAX>", faxTextBox.Text.Equals("") ? "" : "факс " + faxTextBox.Text + ", ");
            //_wm.replace("<EMAIL>", emailTextBox.Text.Equals("") ? "" :"email: " +  emailTextBox.Text + ", ");
            //_wm.replace("<CHACC>", chAcTextBox.Text.Equals("") ? "" : "р/сч " + chAcTextBox.Text + ", ");
            //_wm.replace("<WHERECHACC>", whereChAcTextBox.Text.Equals("") ? "" : whereChAcTextBox.Text);

            if (avanceTextBox.Text.Equals(""))
                Contract.HASADVANCE = false;
            else
                Contract.HASADVANCE = true;

            WorkSchedulesForm wsf = new WorkSchedulesForm(_wm);
            wsf.Visible = true;
            
        }

        private void exitButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

       

    }
}
