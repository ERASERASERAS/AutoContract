using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace AutoContract
{
    public partial class WorkSchedulesForm : Form
    {

        WordMethods _wm;

        private string[] headers = { "№ п/п", "Наименование этапов", "Сроки выполнения", "Договорная цена, руб.", "Форма представления результатов" };
        public WorkSchedulesForm()
        {
            InitializeComponent();
        }

        public WorkSchedulesForm(WordMethods wm)
        {
            InitializeComponent();
            _wm = wm;
        }



        private void nextButton_Click(object sender, EventArgs e)
        {
            try
            {
                WordMethods.FillShowTemplate(insertTable, _wm);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Что-то не так со вставкой календарного плана");
            }
            ChooseEstimatedCostForm cecf = new ChooseEstimatedCostForm(_wm);
            cecf.Visible = true;

            //insert table
        }

        private void WorkSchedulesForm_Load(object sender, EventArgs e)
        {
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            var numberColumn = new DataGridViewColumn();
            numberColumn.HeaderText = "№\nп/п";
            numberColumn.Width = 50;           
            numberColumn.ReadOnly = false;
            numberColumn.Frozen = true;
            numberColumn.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(numberColumn);
            var etapsNameColumn = new DataGridViewColumn();
            etapsNameColumn.HeaderText = "Наименование \n\n\nэтапов";
            etapsNameColumn.Width = 600;
            etapsNameColumn.ReadOnly = false;
            etapsNameColumn.Frozen = true;
            etapsNameColumn.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(etapsNameColumn);
            var complDateColumn = new DataGridViewColumn();
            complDateColumn.HeaderText = "Сроки выполнения";
            complDateColumn.Width = 100;
            complDateColumn.ReadOnly = false;
            complDateColumn.Frozen = true;
            complDateColumn.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(complDateColumn);
            var costColumn = new DataGridViewColumn();
            costColumn.HeaderText = "Договорная цена, руб.";
            costColumn.Width = 100;
            costColumn.ReadOnly = false;
            costColumn.Frozen = true;
            costColumn.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(costColumn);
            var typeOfImageResColumn = new DataGridViewColumn();
            typeOfImageResColumn.HeaderText = "Форма представления результатов";
            typeOfImageResColumn.Width = 100;
            typeOfImageResColumn.ReadOnly = false;
            typeOfImageResColumn.Frozen = true;
            typeOfImageResColumn.CellTemplate = new DataGridViewTextBoxCell();
            dataGridView1.Columns.Add(typeOfImageResColumn);
            
        }

        private void exitButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void insertTable(WordMethods _wm)
        {
            Contract.insertTable(_wm, "worksch", dataGridView1.RowCount + 1, dataGridView1.ColumnCount);

            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                Contract.insertToCell(_wm, headers[i], 1, i + 1);
                Contract.insertToCell(_wm, (i + 1).ToString(), 2, i + 1);
            }

            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    Contract.insertToCell(_wm, (string)dataGridView1[j, i].Value, i + 3, j + 1);
                }
            }

            Contract.setAttributesForWorkSchedulesTable(_wm, dataGridView1.RowCount, dataGridView1.ColumnCount);
 
            Contract.setDocumentsVisible(_wm, true);
        }
        
    }
}
