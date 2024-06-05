using Microsoft.Win32;
using System.Windows;
using System.Windows.Controls;
using System.IO;

using System.Reflection;
using NPOI.POIFS.Storage;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using ICSharpCode.SharpZipLib.Core;

using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;


namespace ExcelMerge
{
    public partial class MainWindow : MetroWindow
    {
        private ExcelManager excelManager1;
        private ExcelManager excelManager2;

        private ComparisonHelper comparisonHelper;

        public MainWindow()
        {
            InitializeComponent();

            this.excelManager1 = new ExcelManager();
            this.excelManager2 = new ExcelManager();

            this.comparisonHelper = new ComparisonHelper();
            this.comparisonHelper.ProgressChanged += ComparsionHelper_ProgressChanged;

            this.ShowMessageAsync("Внимание!", "Перед использованием программы рекомендуется сделать копии файлов, с которыми будете работать!");
        }

        private void BtnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                Filter = "Файлы Excel (.xls, .xlsx, .xlsm)|*.xls;*.xlsx;*.xlsm",
            };

            if (openFileDialog.ShowDialog() == true)
            {
                if (sender is Button btnOpenFile)
                {
                    FileInfo fileInfo = new FileInfo(openFileDialog.FileName);

                    int fileNum = Convert.ToInt32(btnOpenFile.Tag);
                    if (fileNum == 1)
                    {
                        this.ResetComponent(this.txtFile1, this.cmbSheet1, this.cmbColumn1, this.cmbPasteColumn1, this.txtRowColumnCountSheet1, this.nupRowWithColumnNames2);
                        this.LoadFile(this.excelManager1, fileInfo, this.txtFile1, this.cmbSheet1);
                    }
                    else if (fileNum == 2)
                    {
                        this.ResetComponent(this.txtFile2, this.cmbSheet2, this.cmbColumn2, this.cmbCopyColumn2, this.txtRowColumnCountSheet2, this.nupRowWithColumnNames2);
                        this.LoadFile(this.excelManager2, fileInfo, this.txtFile2, this.cmbSheet2);
                    }
                }
            }
        }

        private async void LoadFile(ExcelManager excelHelper, FileInfo fileInfo, TextBlock txtFile, ComboBox cmbSheet)
        {
            try
            {
                excelHelper.OpenExcelFile(fileInfo);
                txtFile.Text = $"Выбранный файл: {fileInfo.Name} ({fileInfo.FullName})";

                cmbSheet.ItemsSource = excelHelper.GetSheetNames();
            }
            catch (IOException)
            {
                await this.ShowMessageAsync("Внимание!", $"Файл {fileInfo.Name} уже где-то открыт. Закройте его и попробуйте ещё раз.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка при открытии Excel файла", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        
        private void CmbSheet_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox cmbSheet && cmbSheet.SelectedIndex != -1)
            {
                string sheetName = cmbSheet.SelectedItem.ToString();

                int fileNum = Convert.ToInt32(cmbSheet.Tag);
                if (fileNum == 1)
                {
                    this.LoadSheetInfo(this.excelManager1, sheetName, this.txtRowColumnCountSheet1);
                }
                else if (fileNum == 2)
                {
                    this.LoadSheetInfo(this.excelManager2, sheetName, this.txtRowColumnCountSheet2);
                }
            }
        }

        private void LoadSheetInfo(ExcelManager excelHelper, string sheetName, TextBlock txtRowColumnCountSheet)
        {
            try
            {
                int[] rowAndColumnCounts = excelHelper.GetSheetInfo(sheetName);

                txtRowColumnCountSheet.Text = $"Строк: {rowAndColumnCounts[0]}\tСтолбцов: {rowAndColumnCounts[1]}";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, $"Ошибка при получении информации о листе ({sheetName})", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


		private void BtnCreateColumnNames_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btnCreateColumnNames)
            {
                int fileNum = Convert.ToInt32(btnCreateColumnNames.Tag);

                if (fileNum == 1)
                {
                    this.CreateColumnNames(this.excelManager1);
                }
                else if (fileNum == 2)
                {
                    this.CreateColumnNames(this.excelManager2);
                }
            }
        }

		private async void CreateColumnNames(ExcelManager excelManager)
		{
			try
			{
				excelManager.CreateColumnNames();
				excelManager.SaveExcelFile();
			}
			catch (IOException)
			{
				await this.ShowMessageAsync("Внимание!", $"Файл уже где-то открыт. Закройте его и попробуйте ещё раз.");
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "Ошибка при создании наименования столбцов в Excel файле", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}



		private void CmbColumn_DropDownOpened(object sender, EventArgs e)
		{
			if (sender is ComboBox сmbColumn)
			{
				int fileNum = Convert.ToInt32(сmbColumn.Tag);

				if (fileNum == 1)
				{
					this.LoadColumns(this.excelManager1, this.cmbColumn1, this.cmbPasteColumn1, this.nupRowWithColumnNames1);
				}
				else if (fileNum == 2)
				{
			        this.LoadColumns(this.excelManager2, this.cmbColumn2, this.cmbCopyColumn2, this.nupRowWithColumnNames2);
				}
			}
		}

        private async void LoadColumns(ExcelManager excelHelper, ComboBox cmbColumn, ComboBox cmbPasteOrCopyColumn, NumericUpDown nupRowWithColumnNames)
        {
            List<string> columnNames = new List<string>();

            try
            {
                double? rowNum = nupRowWithColumnNames.Value - 1;
				columnNames = excelHelper.GetColumnNames(rowNum);
            }
            catch (NullReferenceException)
            {
                await this.ShowMessageAsync("Внимание!", $"Выбранная строка для загрузки имени столбцов ({nupRowWithColumnNames.Value}) пуста. Выберите другую или создайте наименование столбцов и попробуйте ещё раз.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.GetType()}" + ex.Message, "Ошибка при получении списка столбцов", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                cmbColumn.ItemsSource = columnNames;
                cmbPasteOrCopyColumn.ItemsSource = columnNames;
            }
        }


        private async void BtnCopy_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateInputs())
            {
                await this.ShowMessageAsync("Внимание!", "Пожалуйста, выберите все необходимые параметры.");

                return;
            }

            try
            {
                int column1 = this.cmbColumn1.SelectedIndex;
                int column2 = this.cmbColumn2.SelectedIndex;

                int columnCopy = this.cmbCopyColumn2.SelectedIndex;
                int columnPaste = this.cmbPasteColumn1.SelectedIndex;

                ComparsionOptions copyOptions = new ComparsionOptions
                    ((bool) this.cbIgnoreEmptyCells.IsChecked, 
                    (bool) this.cbYellowBackground.IsChecked, 
                    (bool) this.cbGreenBackground.IsChecked, 
                    (bool) this.cbCopyCellsFormat.IsChecked, 
                    (bool) this.cbIngoreCase.IsChecked, 
                    (bool) this.cbIngoreSpace.IsChecked);

                await this.comparisonHelper.ComparingAsync(this.excelManager1, this.excelManager2, column1, column2, columnCopy, columnPaste, copyOptions);
            }
            catch (IOException)
            {
                await this.ShowMessageAsync("Внимание!", $"Файл уже где-то открыт. Закройте его и попробуйте ещё раз.");
            }
            catch (Exception ex)
            {
				MessageBox.Show($"{ex.GetType()}" + ex.Message, "Ошибка при процессе сравнения.", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

        private bool ValidateInputs()
        {
            return this.excelManager1 != null && this.excelManager2 != null &&
                   this.cmbSheet1.SelectedItem != null && this.cmbSheet2.SelectedItem != null &&
                   this.cmbColumn1.SelectedIndex != -1 && this.cmbColumn2.SelectedIndex != -1 &&
                   this.cmbPasteColumn1.SelectedIndex != -1 && this.cmbCopyColumn2.SelectedIndex != -1;
        }

        private void ComparsionHelper_ProgressChanged(object sender, ProgressEventArgs e)
        {
            this.Dispatcher.Invoke(() =>
            {
                this.rtbProgress.AppendText(e.Message + "\n");
                this.progressBar.Value = e.Progress * 100;
            });
        }


		private void ResetComponent(TextBlock txtFile, ComboBox cmbSheet, ComboBox cmbColumn, ComboBox cmbPasteOrCopyColumn, TextBlock txtRowColumnCountSheet, NumericUpDown nupRowWithColumnNames)
		{
			txtFile.Text = "Файл не выбран";

			cmbSheet.ItemsSource = null;
			cmbColumn.ItemsSource = null;
			cmbPasteOrCopyColumn.ItemsSource = null;

			txtRowColumnCountSheet.Text = "Лист не выбран";

			nupRowWithColumnNames.Value = 1;

            this.cbIgnoreEmptyCells.IsChecked = true;
            this.cbYellowBackground.IsChecked = false;

            this.cbGreenBackground.IsChecked = false;
            this.cbCopyCellsFormat.IsChecked = false;

            this.cbIngoreCase.IsChecked = false;
            this.cbIngoreSpace.IsChecked = true;
		}


		private void MainWindow_Closed(object sender, EventArgs e)
        {
            this.excelManager1.Close();
            this.excelManager2.Close();
		}

		private void NupRowWithColumnNames_ValueChanged(object sender, RoutedEventArgs e)
		{
			if (sender is NumericUpDown nupRowWithColumnNames)
			{
				nupRowWithColumnNames.Value ??= 1;
			}
		}
	}
}