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
    public partial class ChooseEstimatedCostForm : Form
    {
        WordMethods _wm;

        public ChooseEstimatedCostForm(WordMethods wm)
        {
            InitializeComponent();
            chooseComboBox.Items.Add("человеко-часы");
            chooseComboBox.Items.Add("другое");
            _wm = wm;
        }

        private void nextButton_Click(object sender, EventArgs e)
        {
            EstimatedCostForm esf = new EstimatedCostForm(chooseComboBox.SelectedItem.ToString() , _wm);
            esf.Visible = true;
        }

        private void exitButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
