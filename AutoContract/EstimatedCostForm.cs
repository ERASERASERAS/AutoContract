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
    public partial class EstimatedCostForm : Form
    {
        WordMethods wm;

        bool flag;

        public EstimatedCostForm()
        {
            InitializeComponent();
        }

        

        public EstimatedCostForm(string type, WordMethods wmVal)
        {
            InitializeComponent();
            wm = wmVal;
            string contractType = mainForm.getContractType();

            if (type == "человеко-часы")
            {
                var numberColumn = new DataGridViewColumn();
                numberColumn.HeaderText = "№\nп/п";
                numberColumn.Width = 50;
                numberColumn.ReadOnly = false;
                numberColumn.Frozen = true;
                numberColumn.CellTemplate = new DataGridViewTextBoxCell();
                dataGridView1.Columns.Add(numberColumn);
                Contract.addToHeadersOfEstimatedCostForm("№\nп/п");
                var execCountColumn = new DataGridViewColumn();
                execCountColumn.HeaderText = "Исполнители/количество";
                execCountColumn.Width = 150;
                execCountColumn.ReadOnly = false;
                execCountColumn.Frozen = true;
                execCountColumn.CellTemplate = new DataGridViewTextBoxCell();
                dataGridView1.Columns.Add(execCountColumn);
                Contract.addToHeadersOfEstimatedCostForm("Исполнители/количество");
                var execPostColumn = new DataGridViewColumn();
                execPostColumn.HeaderText = "Исполнители/Должность";
                execPostColumn.Width = 150;
                execPostColumn.ReadOnly = false;
                execPostColumn.Frozen = true;
                execPostColumn.CellTemplate = new DataGridViewTextBoxCell();
                dataGridView1.Columns.Add(execPostColumn);
                Contract.addToHeadersOfEstimatedCostForm("Исполнители/Должность");
                var manDaysColumn = new DataGridViewColumn();
                manDaysColumn.HeaderText = "Количество человеко/дней";
                manDaysColumn.Width = 150;
                manDaysColumn.ReadOnly = false;
                manDaysColumn.Frozen = true;
                manDaysColumn.CellTemplate = new DataGridViewTextBoxCell();
                dataGridView1.Columns.Add(manDaysColumn);
                Contract.addToHeadersOfEstimatedCostForm("Количество человеко/дней");
                var avSalaryColumn = new DataGridViewColumn();
                avSalaryColumn.HeaderText = "Средняя заработная плата за 1 день, руб";
                avSalaryColumn.Width = 150;
                avSalaryColumn.ReadOnly = false;
                avSalaryColumn.Frozen = true;
                avSalaryColumn.CellTemplate = new DataGridViewTextBoxCell();
                dataGridView1.Columns.Add(avSalaryColumn);
                Contract.addToHeadersOfEstimatedCostForm("Средняя заработная плата за 1 день, руб");
                var basicSalaryColumn = new DataGridViewColumn();
                basicSalaryColumn.HeaderText = "Основная заработная плата, руб.";
                basicSalaryColumn.Width = 150;
                basicSalaryColumn.ReadOnly = false;
                basicSalaryColumn.Frozen = true;
                basicSalaryColumn.CellTemplate = new DataGridViewTextBoxCell();
                dataGridView1.Columns.Add(basicSalaryColumn);
                Contract.addToHeadersOfEstimatedCostForm("Основная заработная плата, руб.");
                flag = false;
            }
            else if (type == "другое")  // Возможно,будет несколько типов
            {
                flag = true;
                switch (mainForm.getContractType())
                {
                    case "Определение комнонентного состава для N отходов":
                    {
                            var numberColumn = new DataGridViewColumn();
                            numberColumn.HeaderText = "№\nп/п";
                            numberColumn.Width = 50;
                            numberColumn.ReadOnly = false;
                            numberColumn.Frozen = true;
                            numberColumn.CellTemplate = new DataGridViewTextBoxCell();
                            dataGridView1.Columns.Add(numberColumn);
                            Contract.addToHeadersOfEstimatedCostForm("№\nп/п");
                            var nameEjColumn = new DataGridViewColumn();
                            nameEjColumn.HeaderText = "Наименование отхода";
                            nameEjColumn.Width = 437;
                            nameEjColumn.ReadOnly = false;
                            nameEjColumn.Frozen = true;
                            nameEjColumn.CellTemplate = new DataGridViewTextBoxCell();
                            dataGridView1.Columns.Add(nameEjColumn);
                            Contract.addToHeadersOfEstimatedCostForm("Наименование отхода");
                            var costColumn = new DataGridViewColumn();
                            costColumn.HeaderText = "Стоимость определения компонентного состава,руб";
                            costColumn.Width = 437;
                            costColumn.ReadOnly = false;
                            costColumn.Frozen = true;
                            costColumn.CellTemplate = new DataGridViewTextBoxCell();
                            dataGridView1.Columns.Add(costColumn);
                            Contract.addToHeadersOfEstimatedCostForm("Стоимость определения компонентного состава,руб");
                            break;
                    }

               
                }
            }
        }

        private void exitButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void insertTable(WordMethods wm)
        {
            
            Contract.insertTable(wm, "estcost", dataGridView1.RowCount + 1, dataGridView1.ColumnCount);

            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                Contract.insertToCell(wm, Contract.getElementOfHeadersOfEstimatedCostForm(i + 1), 1, i + 1);
            }



            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    Contract.insertToCell(wm, (string)dataGridView1[j, i].Value, i + 2, j + 1);
                }
            }

            Contract.insertToCell(wm, "Итого:" + "<COST> (<COSTWORD>) рублей", dataGridView1.RowCount + 1, 1);
            Contract.setDocumentsVisible(wm, true);   
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            try
            {
                WordMethods.FillShowTemplate(insertTable, wm);
                Contract.insertFileCalculateEstCost(wm);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Что-то пошло не так");
            }
        }
    }
}
