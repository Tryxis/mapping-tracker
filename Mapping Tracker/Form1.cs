using System;
using System.Drawing;
using System.Windows.Forms;

namespace Mapping_Tracker
{
    public partial class Form1 : Form
    {
        Bitmap bitmap;

        public Form1()
        {
            InitializeComponent();

        }


        // Add New Button functionality
        private void addButton_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Add(txtSupplierName.Text, txtMappingName.Text, txtSupplierVATNumber.Text, txtAssignee.Text, txtChannelPartner.Text, txtCustomer.Text, dateTimePicker.Text);
        }

        // Exit Button functionality
        private void exitButton_Click(object sender, EventArgs e)
        {
            programExit();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            programExit();
        }

        private void programExit()
        {
            DialogResult iExit;

            iExit = MessageBox.Show("Exit?", "Exit Mapping Tracker", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (iExit == DialogResult.Yes)
            {
                Application.Exit();
            }

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        // Delete rows from the Data Grid
        private void deleteRows()
        {
            foreach (DataGridViewRow item in this.dataGridView1.SelectedRows)
            {
                dataGridView1.Rows.RemoveAt(item.Index);
            }
        }

        // Delete Button functionality
        private void btnDelete(object sender, EventArgs e)
        {
            deleteRows();
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            deleteRows();
        }

        private void clearRows()
        {

            DialogResult iClear;

            iClear = MessageBox.Show("Clear all data?", "Clear Mapping Tracker", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (iClear == DialogResult.Yes)
            {
                //===== Clears the text boxes=====//
                foreach (var c in this.Controls)
                {
                    if (c is TextBox)
                    {
                        ((TextBox)c).Text = String.Empty;
                    }
                }

                //===== Clears the entire DataGrid=====//
                int numRows = dataGridView1.Rows.Count;
                for (int i = 0; i < numRows; i++)
                {
                    try
                    {
                        int max = dataGridView1.Rows.Count - 1;
                        dataGridView1.Rows.Remove(dataGridView1.Rows[max]);
                    }
                    catch (Exception exe)
                    {
                        MessageBox.Show("All rows are to be deleted " + exe, "DataGridView Delete",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }

            }

        }

        private void btnClear(object sender, EventArgs e)
        {
            clearRows();
        }

        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            clearRows();
        }

        private void btnPrint(object sender, EventArgs e)
        {
            int height = dataGridView1.Height;
            dataGridView1.Height = dataGridView1.RowCount * dataGridView1.RowTemplate.Height * 2;
            bitmap = new Bitmap(dataGridView1.Width, dataGridView1.Height);

            dataGridView1.DrawToBitmap(bitmap, new Rectangle(0, 0, dataGridView1.Width, dataGridView1.Height));
            printPreviewDialog1.PrintPreviewControl.Zoom = 1;
            printPreviewDialog1.ShowDialog();
            dataGridView1.Height = height;

        }

        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int height = dataGridView1.Height;
            dataGridView1.Height = dataGridView1.RowCount * dataGridView1.RowTemplate.Height * 2;
            bitmap = new Bitmap(dataGridView1.Width, dataGridView1.Height);

            dataGridView1.DrawToBitmap(bitmap, new Rectangle(0, 0, dataGridView1.Width, dataGridView1.Height));
            printPreviewDialog1.PrintPreviewControl.Zoom = 1;
            printPreviewDialog1.ShowDialog();
            dataGridView1.Height = height;
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.DrawImage(bitmap, 0, 0);
        }

        private void btnSave(object sender, EventArgs e)
        {
            saveFile();
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFile();
        }

        private void saveFile()
        {
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            app.Visible = true;
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            worksheet.Name = "Exported from Mapping Tracker";

            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            }

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }
        }
    }
}
