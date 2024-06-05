using System.IO;
using System.Windows;

using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;


namespace ExcelMerge
{
    public class ExcelManager
    {
        private IWorkbook? workBook;
        private FileInfo fileInfo;

        public ISheet? Sheet { get; private set; }


        public void OpenExcelFile(FileInfo fileInfo)
        {
            using (FileStream fileStream = new FileStream(fileInfo.FullName, FileMode.Open, FileAccess.ReadWrite))
            {
                if (fileInfo.Extension == ".xls")
                {
                    this.workBook = new HSSFWorkbook(fileStream);
                }
                else if (fileInfo.Extension == ".xlsx" || fileInfo.Extension == ".xlsm")
                {
                    this.workBook = new XSSFWorkbook(fileStream);
                }
                else
                {
                    throw new ArgumentException($"Неподдерживаемый формат файла ({fileInfo.Extension})");
                }

                this.fileInfo = fileInfo;
            }
        }

        public void SaveExcelFile()
        {
            using (FileStream fileStream = new FileStream(this.fileInfo.FullName, FileMode.Create, FileAccess.Write))
            {
                this.workBook?.Write(fileStream);
            }
        }

        public void Close()
        {
            this.workBook?.Close();
        }



        public List<string> GetSheetNames()
        {
            int numberOfSheet = this.workBook.NumberOfSheets;
            List<string> sheetNames = new List<string>(numberOfSheet);

            for (int i = 0; i < numberOfSheet; i++)
            {
                sheetNames.Add(this.workBook.GetSheetName(i));
            }

            return sheetNames;
        }

        public int[] GetSheetInfo(string sheetName)
        {
            ISheet sheet = this.workBook.GetSheet(sheetName);
            if (sheet == null)
            {
                throw new InvalidOperationException($"Выбранный лист ({sheetName}) не существует!");
            }

            this.Sheet = sheet;

            int rowCount = sheet.LastRowNum + 1;

            int columnCount = 0;
            for (int i = 0; i < sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    int currentRowColumnCount = row.LastCellNum;
                    columnCount = columnCount < currentRowColumnCount ? currentRowColumnCount : columnCount;
                }
            }

            return new int[2] { rowCount, columnCount };
        }


        public void CreateColumnNames()
        {
            IRow row = this.Sheet.CreateRow(0);

            for (int i = 0; i < 100; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue($"Столбец {i + 1}");
            }
        }

        public List<string> GetColumnNames(double? rowNum)
        {
            IRow row = this.Sheet.GetRow((int) rowNum);
            if (row == null)
            {
                throw new NullReferenceException();
            }

            int columnCount = row.LastCellNum;
            if (columnCount < 0)
            {
                throw new NullReferenceException();
            }

            List<string> columnNames = new List<string>(columnCount);
            for (int i = 0; i < columnCount; i++)
            {
                ICell cell = row.GetCell(i, MissingCellPolicy.RETURN_BLANK_AS_NULL);
                string columnName = $"[{i + 1}] " + (cell != null ? cell.ToString() : "Пустота");
                columnNames.Add(columnName);
            }

            return columnNames;
        }


        public void SetCellStyle(ICell pasteCell, IndexedColors indexedColor)
        {
            //ICellStyle newCellStyle = this.workBook.CreateCellStyle();
            ICellStyle newCellStyle = pasteCell.Sheet.Workbook.CreateCellStyle();

            newCellStyle.FillForegroundColor = indexedColor.Index;
            newCellStyle.FillPattern = FillPattern.SolidForeground;

            pasteCell.CellStyle = newCellStyle;
        }

        public void CopyCellStyle(ICell pasteCell, ICell copyCell)
        {
            ICellStyle copyStyle = copyCell.CellStyle;  // Получаем стиль из ячейки, откуда копируем
            ICellStyle newCellStyle = pasteCell.Sheet.Workbook.CreateCellStyle();

            // Копирование свойств стиля
            newCellStyle.CloneStyleFrom(copyStyle);

            // Копирование шрифта
            IFont sourceFont = copyCell.Sheet.Workbook.GetFontAt(copyStyle.FontIndex);
            IFont newFont = pasteCell.Sheet.Workbook.CreateFont();

            newFont.Boldweight = sourceFont.Boldweight;
            newFont.Color = sourceFont.Color;
            newFont.FontHeightInPoints = sourceFont.FontHeightInPoints;
            newFont.FontName = sourceFont.FontName;
            newFont.IsItalic = sourceFont.IsItalic;
            newFont.IsStrikeout = sourceFont.IsStrikeout;
            newFont.TypeOffset = sourceFont.TypeOffset;
            newFont.Underline = sourceFont.Underline;
            newCellStyle.SetFont(newFont);

            pasteCell.CellStyle = newCellStyle;  // Применяем новый стиль к ячейке pasteCell
        }
    }
}
