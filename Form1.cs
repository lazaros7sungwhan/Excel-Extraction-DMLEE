using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Reflection;

/*
 * Excel Component를 이용하기 위해서는 "Microsoft.Office.Interop.Excel"을 추가해야 함
 * 추가 방법은 "References" -> "Manage NuGet Packages..." -> "Browse..." -> "Microsoft.Office.Interop.Excel"
 * 설치 하면 사용 준비 끝
 */
 using Excel = Microsoft.Office.Interop.Excel;

namespace DataGridViewToExcel
{
    public partial class Form1 : Form
    {
        private Random _Random = null;
        private SaveFileDialog saveFileDialog = null;

        public Form1()
        {
            InitializeComponent();

            _Random = new Random(unchecked((int)DateTime.Now.Ticks));

            saveFileDialog = new SaveFileDialog();

        }

        private void Button1_Click(object sender, EventArgs e)
        {
            int rowID = dataGridView1.Rows.Add();

            DataGridViewRow dr = dataGridView1.Rows[rowID];
            dr.Cells[0].Value = rowID + 1;
            dr.Cells[1].Value = _Random.Next(0, 256);
            dr.Cells[2].Value = _Random.Next(0, 256);

            // Scroll to Bottom
            dataGridView1.FirstDisplayedScrollingRowIndex = rowID;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            // Clear Datagridview
            dataGridView1.Rows.Clear();
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            // 엑셀 창을 열어서 데이터를 옮기는 방법
            copyAlltoClipboard();
            Microsoft.Office.Interop.Excel.Application xlexcel;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Excel.Application();
            xlexcel.Visible = true;
            xlWorkBook = xlexcel.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
            CR.Select();
            xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
        }


        private void Button3_Click(object sender, EventArgs e)
        {
            // 파일로 저장하는 방법
            // https://mastmanban.tistory.com/235 참고
            ExportExcel(true, dataGridView1);
        }

        private void copyAlltoClipboard()
        {
            //to remove the first blank column from datagridview
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void ExportExcel(bool captions, DataGridView dataGridView)
        {
            this.saveFileDialog.FileName = "TempName";
            this.saveFileDialog.DefaultExt = "xls";
            this.saveFileDialog.Filter = "Excel files (*.xls)|*.xls";
            this.saveFileDialog.InitialDirectory = "c:\\";

            DialogResult result = saveFileDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                int num = 0;
                object missingType = Type.Missing;

                Excel.Application objApp;
                Excel._Workbook objBook;
                Excel.Workbooks objBooks;
                Excel.Sheets objSheets;
                Excel._Worksheet objSheet;
                Excel.Range range;

                string[] columns = new string[dataGridView.ColumnCount];

                for (int c = 0; c < dataGridView.ColumnCount; c++)
                {
                    num = c + 65;
                    columns[c] = Convert.ToString((char)num);
                }

                try
                {
                    objApp = new Excel.Application();
                    objBooks = objApp.Workbooks;
                    objBook = objBooks.Add(Missing.Value);
                    objSheets = objBook.Worksheets;
                    objSheet = (Excel._Worksheet)objSheets.get_Item(1);

                    if (captions)
                    {
                        for (int c = 0; c < dataGridView.ColumnCount; c++)
                        {
                            num = c + 65;
                            range = objSheet.get_Range(columns[c] + "1", Missing.Value);
                            range.set_Value(Missing.Value, dataGridView1.Columns[c].HeaderText);
                        }
                    }

                    for (int i = 0; i < dataGridView.RowCount - 1; i++)
                    {
                        for (int j = 0; j < dataGridView.ColumnCount; j++)
                        {
                            range = objSheet.get_Range(columns[j] + Convert.ToString(i + 2),
                                                                   Missing.Value);
                            range.set_Value(Missing.Value,
                                                  dataGridView.Rows[i].Cells[j].Value.ToString());
                        }
                    }

                    objApp.Visible = false;
                    objApp.UserControl = false;

                    objBook.SaveAs(@saveFileDialog.FileName,
                              Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
                              missingType, missingType, missingType, missingType,
                              Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                              missingType, missingType, missingType, missingType, missingType);
                    objBook.Close(false, missingType, missingType);

                    Cursor.Current = Cursors.Default;

                    MessageBox.Show("Save Success!!!");
                }
                catch (Exception theException)
                {
                    String errorMessage;
                    errorMessage = "Error: ";
                    errorMessage = String.Concat(errorMessage, theException.Message);
                    errorMessage = String.Concat(errorMessage, " Line: ");
                    errorMessage = String.Concat(errorMessage, theException.Source);

                    MessageBox.Show(errorMessage, "Error");
                }
            }
        }

    }
}
