namespace AgbinaDiffParser
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    using Microsoft.Office.Interop.Excel;
    using App = Microsoft.Office.Interop.Excel.Application;
    using Application = System.Windows.Forms.Application;
    using Range = Microsoft.Office.Interop.Excel.Range;

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            openFileDialog1.Title = "Select old file";
            Dictionary<string, (int, string)> oldCells;
            try
            {
                oldCells = ReadExcelFile();
            }
            catch(Exception ex)
            {
                MessageBox.Show($"Couldn't old read file. Error: {ex.Message}");
                this.Close();
                Application.Exit();
                return;
            }

            openFileDialog1.Title = "Select new file";
            Dictionary<string, (int, string)> newCells;
            try
            {
                newCells = ReadExcelFile();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Couldn't read new file. Error: {ex.Message}");
                this.Close();
                Application.Exit();
                return;
            }

            var removedKeys = oldCells.Keys.Except(newCells.Keys);
            var addedKeys = newCells.Keys.Except(oldCells.Keys);
            var removedRecords = oldCells.Where(o => removedKeys.Contains(o.Key)).Select(o => $"{o.Key} | {o.Value.Item2} | row: {o.Value.Item1}").ToArray();
            var addedRecords = newCells.Where(n => addedKeys.Contains(n.Key)).Select(n => $"{n.Key} | {n.Value.Item2} | row: {n.Value.Item1}").ToArray();
        }

        private Dictionary<string, (int, string)> ReadExcelFile()
        {

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var result = new Dictionary<string, (int, string)>();
                var xlApp = new App();
                var xlWorkBook = xlApp.Workbooks.Open(openFileDialog1.FileName);
                var xlWorkSheet = (Worksheet)xlWorkBook.Worksheets[1];
                for (int i = 1; i <= xlWorkSheet.Rows.Count; i++)
                {
                    var value = ((Range)xlWorkSheet.Cells[i, 1]).Value?.ToString();
                    if (value == null) break;
                    int j = 2;
                    while (((Range)xlWorkSheet.Cells[i, j]).Value == null) j++;
                    result.Add(value, (i, ((Range)xlWorkSheet.Cells[i, j]).Value?.ToString()));
                }

                Marshal.ReleaseComObject(xlWorkSheet);

                //close and release
                xlWorkBook.Close();
                Marshal.ReleaseComObject(xlWorkBook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                return result;
            }
            else
            {
                throw new FileNotFoundException();
            }
        }

    }
}
