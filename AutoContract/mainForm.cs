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
    //Может придётся изменять autosize  
    public partial class mainForm : Form
    {

        private string pathToPatternFile { get { return Application.StartupPath + "\\patterns\\" + "contract_pattern.doc"; } }
        private  static string contractType;
        

        public static string  getContractType()
        {
            return contractType;
        }


        
        WordMethods _wm;
        public mainForm()
        {
            InitializeComponent();
            contrTypeComboBox.Items.Add("Определение комнонентного состава для N отходов");
            contrTypeComboBox.Items.Add("Разработка и техническое сопровождение согласования проекта предельно допустимых выбросов загрязняющих веществ (ПДВ) в атмосферный воздух от источников");
            contrTypeComboBox.Items.Add("Разработка и сопровождение согласования проекта нормативов образования отходов и лимитов на их размещение (ПНООЛР) для");
            contrTypeComboBox.Items.Add("");
            Contract.PATHTOFOLDERWITHFILES = Application.StartupPath + "\\patterns\\";
        }

        private void nextButton_Click(object sender, EventArgs e)
        {
            try
            {
                 contractType = contrTypeComboBox.SelectedItem.ToString();
                _wm = new WordMethods(pathToPatternFile, true);
                GeneralDataForm gdf = new GeneralDataForm(contractType,pathToPatternFile,_wm);
                gdf.Visible = true;        
            }
            catch
            {
                MessageBox.Show("Выберите тип договора");
            }            
        }

        private void exitButtonF_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        
    }
}
