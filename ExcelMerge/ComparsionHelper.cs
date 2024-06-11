using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

using System.Windows.Media;
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

        protected virtual void OnProgressChanged(string message, double progress, SolidColorBrush color)
        {
            this.ProgressChanged?.Invoke(this, new ProgressEventArgs(message, progress, color));
        }


		public async Task ComparingAsync(ExcelManager excelManager1, ExcelManager excelManager2, int columnNum1, int columnNum2, int columnNumCopy, int columnNumPaste, ComparisonOptions comparisonOptions)
        {
            await Task.Run(() =>
            {
				this.OnProgressChanged("------------------------------------", 0, Brushes.Green);
				this.OnProgressChanged($"СТАРТ ({DateTime.Now})\n", 0, Brushes.Green);

				ComparisonCounts comparisonCounts = new ComparisonCounts();

				for (int i = 1; i <= excelManager1.Sheet.LastRowNum; i++)
                {
					comparisonCounts.CellCount++;

					IRow row1 = excelManager1.Sheet.GetRow(i);
                    if (row1 == null) { continue; }
                    ICell cell1 = row1.GetCell(columnNum1);
                    if (cell1 == null) { comparisonCounts.EmptyCellsCount++; continue; }

                    if (this.ShouldSkipCell(excelManager1, cell1, comparisonOptions)) { comparisonCounts.EmptyCellsCount++; continue; }

                    for (int j = 1; j <= excelManager2.Sheet.LastRowNum; j++)
                    {
						IRow row2 = excelManager2.Sheet.GetRow(j);
                        if (row2 == null) { continue; }
                        ICell cell2 = row2.GetCell(columnNum2);
                        if (cell2 == null) { continue; }

                        if (this.ComparePart(cell1, cell2, comparisonOptions))
                        {
							comparisonCounts.MatchCount++;

                            this.CopyPart(excelManager1, excelManager2, i, j, columnNumCopy, columnNumPaste, comparisonOptions);
							this.OnProgressChanged($"[{i}][{columnNum1}] \"{cell1.StringCellValue}\" = [{j}][{columnNum2}] \"{cell2.StringCellValue}\"", (double) i / excelManager1.Sheet.LastRowNum);
                            
                            if (comparisonOptions.SkipFurtherMatches) { break; }
                        }
                    }
                }

                this.OnProgressChanged(comparisonCounts.ToString(), 1, Brushes.DeepSkyBlue);

                this.OnProgressChanged($"Cохранение изменений в файл..", 1);
                excelManager1.SaveExcelFile();
                this.OnProgressChanged($"Cохранение изменений - успешно.", 1);

				this.OnProgressChanged($"\nКОНЕЦ ({DateTime.Now.ToString()})", 1, Brushes.Green);
				this.OnProgressChanged("------------------------------------\n", 0, Brushes.Green);
            });
        }

        private bool ComparePart(ICell cell1, ICell cell2, ComparisonOptions comparisonOptions)
        {
            string[] stringCellValues = this.PrepareCellsForComparison(cell1.StringCellValue, cell2.StringCellValue, comparisonOptions);

            string stringCellValue1 = stringCellValues[0];
            string stringCellValue2 = stringCellValues[1];

            return stringCellValue1 == stringCellValue2;
        }

        private void CopyPart(ExcelManager excelHelper1, ExcelManager excelHelper2, int columnNum1, int columnNum2, int columnNumCopy, int columnNumPaste, ComparisonOptions comparisonOptions)
        {
            IRow rowPaste = excelHelper1.Sheet.GetRow(columnNum1) ?? excelHelper1.Sheet.CreateRow(columnNum1);
            ICell cellPaste = rowPaste.GetCell(columnNumPaste) ?? rowPaste.CreateCell(columnNumPaste);

            IRow rowCopy = excelHelper2.Sheet.GetRow(columnNum2);
            ICell cellCopy = rowCopy.GetCell(columnNumCopy);

            if (cellCopy != null)
            {
                if (comparisonOptions.CopyCellsFormat) // copyOptions 1
                {
                    excelHelper1.CopyCellStyle(cellPaste, cellCopy);
                }
                if (comparisonOptions.GreenBackground) // copyOptions 2
                {
                    excelHelper1.SetCellStyle(cellPaste, IndexedColors.Green);
                    
                }
                cellPaste.SetCellValue(cellCopy.StringCellValue);
            }
        }

        private bool ShouldSkipCell(ExcelManager excelHelper, ICell cell, ComparisonOptions comparisonOptions) // copyOptions 3
        {
            if (comparisonOptions.YellowBackground && string.IsNullOrWhiteSpace(cell.StringCellValue)) 
            { 
                excelHelper.SetCellStyle(cell, IndexedColors.Yellow);
            }

            return comparisonOptions.IgnoreEmptyCells && string.IsNullOrWhiteSpace(cell.StringCellValue);
        }

        private string[] PrepareCellsForComparison(string cell1, string cell2, ComparisonOptions comparisonOptions) // copyOptions 4
        {
            if (comparisonOptions.IgnoreCase)
            {
                cell1 = cell1.ToLower();
                cell2 = cell2.ToLower();
            }
            if (comparisonOptions.IgnoreSpace)
            {
                cell1 = cell1.Trim();
                cell2 = cell2.Trim();
            }

            return new string[2] { cell1, cell2 };
        }
    }
}
