using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NPOI.SS.UserModel;

namespace ExcelMerge
{
    public class ComparisonHelper
    {
        public event EventHandler<ProgressEventArgs>? ProgressChanged;

        protected virtual void OnProgressChanged(string message, double progress)
        {
            this.ProgressChanged?.Invoke(this, new ProgressEventArgs(message, progress));
        }

        public async Task ComparingAsync(ExcelManager excelManager1, ExcelManager excelManager2, int columnNum1, int columnNum2, int columnNumCopy, int columnNumPaste, ComparsionOptions copyOptions)
        {
            await Task.Run(() =>
            {
                this.OnProgressChanged($"НАЧАЛО ({DateTime.Now.ToLongTimeString()})", 0);

                for (int i = 1; i <= excelManager1.Sheet.LastRowNum; i++)
                {
                    IRow row1 = excelManager1.Sheet.GetRow(i);
                    if (row1 == null) { continue; }
                    ICell cell1 = row1.GetCell(columnNum1);
                    if (cell1 == null) { continue; }

                    if (this.ShouldSkipCell(excelManager1, cell1, copyOptions)) { continue; }

                    for (int j = 1; j <= excelManager2.Sheet.LastRowNum; j++)
                    {
                        IRow row2 = excelManager2.Sheet.GetRow(j);
                        if (row2 == null) { continue; }
                        ICell cell2 = row2.GetCell(columnNum2);
                        if (cell2 == null) { continue; }

                        if (this.ComparePart(cell1, cell2, copyOptions))
                        {
                            this.CopyPart(excelManager1, excelManager2, i, j, columnNumCopy, columnNumPaste, copyOptions);

                            this.OnProgressChanged($"[{i}][{columnNum1}] \"{cell1.StringCellValue}\" = [{j}][{columnNum2}] \"{cell2.StringCellValue}\"", (double) i / excelManager1.Sheet.LastRowNum);
                        }
                    }
                }

                this.OnProgressChanged($"Cохранение изменений в файле..", 1);
                excelManager1.SaveExcelFile();
                this.OnProgressChanged($"Cохранения изменений - успешно.", 1);

                this.OnProgressChanged($"КОНЕЦ ({DateTime.Now.ToLongTimeString()})", 1);
            });
        }

        private bool ComparePart(ICell cell1, ICell cell2, ComparsionOptions copyOptions)
        {
            string[] stringCellValues = this.PrepareCellsForComparison(cell1.StringCellValue, cell2.StringCellValue, copyOptions);

            string stringCellValue1 = stringCellValues[0];
            string stringCellValue2 = stringCellValues[1];

            return stringCellValue1 == stringCellValue2;
        }

        private void CopyPart(ExcelManager excelHelper1, ExcelManager excelHelper2, int columnNum1, int columnNum2, int columnNumCopy, int columnNumPaste, ComparsionOptions copyOptions)
        {
            IRow rowPaste = excelHelper1.Sheet.GetRow(columnNum1) ?? excelHelper1.Sheet.CreateRow(columnNum1);
            ICell cellPaste = rowPaste.GetCell(columnNumPaste) ?? rowPaste.CreateCell(columnNumPaste);

            IRow rowCopy = excelHelper2.Sheet.GetRow(columnNum2);
            ICell cellCopy = rowCopy.GetCell(columnNumCopy);

            if (cellCopy != null)
            {
                if (copyOptions.CopyCellsFormat) // copyOptions 1
                {
                    excelHelper1.CopyCellStyle(cellPaste, cellCopy);
                }
                if (copyOptions.GreenBackground) // copyOptions 1
                {
                    excelHelper1.SetCellStyle(cellCopy, IndexedColors.Green);
                }
                cellPaste.SetCellValue(cellCopy.StringCellValue);
            }
        }

        private bool ShouldSkipCell(ExcelManager excelHelper, ICell cell, ComparsionOptions copyOptions) // copyOptions 2
        {
            if (copyOptions.YellowBackground && string.IsNullOrWhiteSpace(cell.StringCellValue)) 
            { 
                excelHelper.SetCellStyle(cell, IndexedColors.Yellow);
            }

            return copyOptions.IgnoreEmptyCells && string.IsNullOrWhiteSpace(cell.StringCellValue);
        }

        private string[] PrepareCellsForComparison(string cell1, string cell2, ComparsionOptions copyOptions) // copyOptions 2
        {
            if (copyOptions.IgnoreCase)
            {
                cell1 = cell1.ToLower();
                cell2 = cell2.ToLower();
            }
            if (copyOptions.IgnoreSpace)
            {
                cell1 = cell1.Trim();
                cell2 = cell2.Trim();
            }

            return new string[2] { cell1, cell2 };
        }
    }
}
