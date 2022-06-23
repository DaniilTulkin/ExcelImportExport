using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using Controls = System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Controls;

namespace ExcelImportExport
{
    public class Excel
    {
        private static string writeFilePath = string.Empty;
        private static string readFilePath = string.Empty;
        public static void GetFilePathToWrite()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                OverwritePrompt = true,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = ".xlsx files (*.xlsx)|*.xlsx"
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                writeFilePath = saveFileDialog.FileName;
            }
        }

        public static void WriteExcel(Controls.DataGrid dataGrid)
        {
            if (dataGrid != null)
            {
                GetFilePathToWrite();
                if (!string.IsNullOrEmpty(writeFilePath))
                {
                    using (ExcelPackage excelPackage = new ExcelPackage())
                    {
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Лист1");

                        int columnsCount = dataGrid.Columns.Count;
                        int rowsCount = dataGrid.Items.Count;
                        for (int i = 0; i < columnsCount; i++)
                        {
                            foreach (DataGridColumn column in dataGrid.Columns)
                            {
                                if (column.DisplayIndex == i)
                                {
                                    string key = column.Header.ToString();
                                    worksheet.Cells[1, i + 1].Value = key;
                                    for (int j = 0; j < rowsCount;)
                                    {
                                        foreach (DynamicDictionaryWrapper dict in dataGrid.Items)
                                        {
                                            string value = dict[key]?.ToString();
                                            worksheet.Cells[j + 2, i + 1].Value = value;
                                            j++;
                                        }
                                    }
                                    break;
                                }
                            }
                        }

                        FileInfo fi = new FileInfo($"{writeFilePath}");
                        try
                        {
                            excelPackage.SaveAs(fi);
                        }
                        catch (Exception)
                        {
                            MessageBox.Show(
                                "Файл для записи или чтения используется другим процессом. \nЗакройте файл перед записью или чтением.",
                                "Ошибка записи или чтения файла",
                                MessageBoxButtons.OK);
                            return;
                        }
                    }
                }
            }
        }

        public static void GetFilePathToRead()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = ".xlsx files (*.xlsx)|*.xlsx"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                readFilePath = openFileDialog.FileName;
            }
        }

        public static List<DynamicDictionaryWrapper> ReadExcel()
        {
            GetFilePathToRead();
            if (!string.IsNullOrEmpty(readFilePath))
            {
                List<DynamicDictionaryWrapper> dictList = new List<DynamicDictionaryWrapper>();
                byte[] bin = null;
                try
                {
                    bin = File.ReadAllBytes(readFilePath);
                }
                catch (Exception)
                {
                    MessageBox.Show(
                                "Файл для записи или чтения используется другим процессом. \nЗакройте файл перед записью или чтением.",
                                "Ошибка записи или чтения файла",
                                MessageBoxButtons.OK);
                    return null;
                }

                using (MemoryStream stream = new MemoryStream(bin))
                using (ExcelPackage excelPackage = new ExcelPackage(stream))
                {
                    ExcelWorksheet firstWorksheet = excelPackage.Workbook.Worksheets[1];
                    int columnsCount = firstWorksheet.Dimension.Columns;
                    int rowsCount = firstWorksheet.Dimension.Rows;
                    for (int i = 2; i < rowsCount+1; i++)
                    {
                        Dictionary<string, object> dict = new Dictionary<string, object>();
                        for (int j = 1; j < columnsCount+1; j++)
                        {
                            string key = firstWorksheet.Cells[1, j].Value?.ToString();
                            object value = firstWorksheet.Cells[i, j].Value;
                            dict[key] = value;
                        }
                        dictList.Add(new DynamicDictionaryWrapper(dict));
                    }
                }
                return dictList;
            }
            return null;
        }
    }
}
