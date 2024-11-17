namespace AgbinaDiffParser
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Windows.Forms;
    using ExcelDataReader;

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private Dictionary<string, (int, string)> OpenFile(string openDialogTitle)
        {
            try
            {
                openFileDialog1.Title = openDialogTitle;
                return ReadExcelFile();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Couldn't old read file. Error: {ex.Message}");
                this.Close();
                Application.Exit();
                return null;
            }
        }

        private Dictionary<string, (int, string)> ReadExcelFile()
        {

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var result = new Dictionary<string, (int, string)>();

                using var stream = File.Open(openFileDialog1.FileName, FileMode.Open, FileAccess.Read);
                using var reader = ExcelReaderFactory.CreateReader(stream);

                reader.Read();
                int i = 2;
                while (reader.Read())
                {
                    var value = reader.GetString(0);
                    if (string.IsNullOrEmpty(value)) break;
                    int j = 2;
                    while (string.IsNullOrEmpty(reader.GetString(j))) j++;
                    result.Add(value, (i, reader.GetString(j)));
                    i++;
                }

                return result;
            }
            else
            {
                throw new FileNotFoundException();
            }
        }

        public void ReadFiles()
        {
            Dictionary<string, (int, string)> oldCells = OpenFile("Select old file");
            if (oldCells != null)
            {
                Dictionary<string, (int, string)> newCells = OpenFile("Select new file");
                var removedKeys = oldCells.Keys.Except(newCells.Keys);
                var addedKeys = newCells.Keys.Except(oldCells.Keys);
                var removedRecords = oldCells.Where(o => removedKeys.Contains(o.Key)).Select(o => (o.Value.Item1, o.Key, o.Value.Item2));
                var addedRecords = newCells.Where(n => addedKeys.Contains(n.Key)).Select(n => (n.Value.Item1, n.Key, n.Value.Item2));

                foreach ((int, string, string) v in addedRecords)
                {
                    dataGridView1.Rows.Add(v.Item1, v.Item2, v.Item3);
                }

                foreach ((int, string, string) v in removedRecords)
                {
                    dataGridView2.Rows.Add(v.Item1, v.Item2, v.Item3);
                }
            }
        }

    }
}
