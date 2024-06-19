using System.Data;
using System.IO;

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


namespace ExcelMerge
{
    public class ExcelManager
    {
        private IWorkbook? workBook;
        private FileInfo? fileInfo;

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

        public void ReloadExcelFile()
        {
            this.OpenExcelFile(this.fileInfo ?? throw new ArgumentNullException());
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

        public async Task<DataTable> GetTable(string sheetName)
        {
            DataTable dataTable = new DataTable();

            if (this.workBook == null)
            {
                throw new NullReferenceException();
            }

            await Task.Run(() =>
            {
                using (FileStream fileStream = new FileStream(this.fileInfo.FullName, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = new XSSFWorkbook(fileStream);
                    ISheet sheet = workbook.GetSheetAt(0);

                    IRow firstRow = sheet.GetRow(0);
                    int columnCount = this.GetSheetColumnCount(sheetName);

                    dataTable.Columns.Add("№");
                    for (int i = 0; i < columnCount; i++)
                    {
                        dataTable.Columns.Add(GetColumnNum(i + 1));
                    }

                    for (int i = 0; i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) { continue; }

                        DataRow dr = dataTable.NewRow();

                        dr[0] = $"{i + 1}";

                        for (int j = 0; j < columnCount; j++)
                        {
                            dr[j + 1] = row?.GetCell(j)?.ToString() ?? string.Empty;
                        }

                        dataTable.Rows.Add(dr);
                    }
                }
            });

            return dataTable;
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

        public int GetSheetColumnCount(string sheetName)
        {
            ISheet sheet = this.workBook.GetSheet(sheetName);
            if (sheet == null)
            {
                throw new InvalidOperationException($"Выбранный лист ({sheetName}) не существует!");
            }

            this.Sheet = sheet;

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

            return columnCount;
        }

        public int GetSheetRowCount(string sheetName)
        {
            ISheet sheet = this.workBook.GetSheet(sheetName);
            if (sheet == null)
            {
                throw new InvalidOperationException($"Выбранный лист ({sheetName}) не существует!");
            }
            this.Sheet = sheet;

            int rowCount = sheet.LastRowNum + 1;

            return rowCount;
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

                string columnNum = $"№{i + 1} [{GetColumnNum(i + 1)}] ";
                string columnName = columnNum + (cell?.ToString() ?? "Пустота");

                columnNames.Add(columnName);
            }

            return columnNames;
        }

        private string GetColumnNum(int num)
        {
            string str = "";

            while (num > 0)
            {
                num--;
                int remainder = num % 26;
                str = (char)(remainder + 65) + str;
                num /= 26;
            }

            return str;
        }


        public void SetCellStyle(ICell pasteCell, IndexedColors indexedColor)
        {
            ICellStyle newCellStyle = pasteCell.Sheet.Workbook.CreateCellStyle();

            newCellStyle.FillForegroundColor = indexedColor.Index;
            newCellStyle.FillPattern = FillPattern.SolidForeground;

            pasteCell.CellStyle = newCellStyle;
        }

        public void CopyCellStyle(ICell pasteCell, ICell copyCell)
        {
            ICellStyle copyStyle = copyCell.CellStyle;
            ICellStyle newCellStyle = pasteCell.Sheet.Workbook.CreateCellStyle();
            newCellStyle.CloneStyleFrom(copyStyle);

            IFont copyFont = copyCell.Sheet.Workbook.GetFontAt(copyStyle.FontIndex);
            IFont newFont = pasteCell.Sheet.Workbook.CreateFont();
			newFont.CloneStyleFrom(copyFont);

            pasteCell.CellStyle = newCellStyle;
        }
    }
}
