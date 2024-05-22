using System.Windows;

using OfficeOpenXml;

namespace ExcelMerge
{
    public class ExcelHelper
    {
        public static string[] GetSheetNames(ExcelPackage excelPackage)
        {
            try
            {
                return excelPackage.Workbook.Worksheets.Select(s => s.Name).ToArray();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return Array.Empty<string>();
            }
        }

        public static string[] GetColumnNames(ExcelWorksheet excelWorksheet)
        {

            try
            {
                int numberOfColumns = excelWorksheet.Columns.Count();

                string[] columnNames = new string[numberOfColumns];
                for (int i = 1; i < numberOfColumns; i++)
                {
                    columnNames[i - 1] = excelWorksheet.Cells[1, i].Text;   // 1 - номер строки,
                                                                            // а так как название столбцов необязательно будет на первой строке,
                                                                            // то нужно сделать проверки на пустоту, пока не найдётся не пустая строка
                                                                            // или в доп. параметрах добавить, чтобы пользователь мог указать номер строки, где названия столбцов
                }

                return columnNames;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return Array.Empty<string>();
            }

        }

        public static async Task<int> CopyingAsync(ExcelWorksheet excelWorksheet1, ExcelWorksheet excelWorksheet2, int column1, int column2, int columnCopy, int columnPaste, CopyOptions copyOptions)
        {
            int counter = 0;

            await Task.Run(() =>
            {
                for (int i = 1; i <= excelWorksheet1.Dimension.End.Row; i++)
                {
                    string cell1Text = excelWorksheet1.Cells[i, column1].Text;
                    if (string.IsNullOrWhiteSpace(cell1Text)) { continue; }

                    for (int j = 1; j <= excelWorksheet2.Dimension.End.Row; j++)
                    {
                        string cell2Text = excelWorksheet2.Cells[j, column2].Text;
                        if (string.IsNullOrWhiteSpace(cell2Text)) { continue; }

                        string[] preparedCells = ExcelHelper.PrepareCellsForComparison(cell1Text, cell2Text, copyOptions);

                        if (preparedCells[0] == preparedCells[1])
                        {
                            excelWorksheet2.Cells[j, columnPaste].Value = excelWorksheet1.Cells[i, columnCopy].Value;

                            if (copyOptions.YellowBackground == true)
                            {
                                excelWorksheet2.Cells[j, columnPaste].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                excelWorksheet2.Cells[j, columnPaste].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                            }

                            counter++;
                        }
                    }
                }
            });

            return counter;
        }

        private static string[] PrepareCellsForComparison(string cell1, string cell2, CopyOptions copyOptions)
        {
            if (copyOptions.IgnoreCase == true)
            {
                cell1 = cell1.ToLower();
                cell2 = cell2.ToLower();
            }
            if (copyOptions.IgnoreSpace == true)
            {
                cell1 = cell1.Trim();
                cell2 = cell2.Trim();
            }

            return new string[2] { cell1, cell2 };
        }
    }
}
