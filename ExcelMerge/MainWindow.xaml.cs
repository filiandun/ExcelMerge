using Microsoft.Win32;
using System.Windows;
using System.Windows.Controls;
using System.IO;

using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using System.Windows.Documents;
using NPOI.OpenXmlFormats.Spreadsheet;


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
        }


        private async void MetroWindow_Loaded(object sender, RoutedEventArgs e)
        {
            if (!Properties.Settings.Default.DoNotShowWarning)
            {
                var messageDialogResult = await this.ShowMessageAsync("Внимание!", "Перед использованием программы рекомендуется создать копии файлов, с которыми будете работать!", MessageDialogStyle.AffirmativeAndNegative, new MetroDialogSettings() { NegativeButtonText = "Больше не показывать" });
                if (messageDialogResult == MessageDialogResult.Negative)
                {
                    Properties.Settings.Default.DoNotShowWarning = true;
                    Properties.Settings.Default.Save();
                }
            }
        }

        private async void BtnOpenFile_Click(object sender, RoutedEventArgs e)
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
                        this.ResetAllComponent(this.txtFile1, this.cmbSheet1, this.cmbColumn1, this.cmbPasteColumn1, this.txtRowColumnCountSheet1, this.nupRowWithColumnNames1);
                        await this.LoadFile(this.excelManager1, fileInfo, this.txtFile1, this.cmbSheet1);
                    }
                    else if (fileNum == 2)
                    {
                        this.ResetAllComponent(this.txtFile2, this.cmbSheet2, this.cmbColumn2, this.cmbCopyColumn2, this.txtRowColumnCountSheet2, this.nupRowWithColumnNames2);
                        await this.LoadFile(this.excelManager2, fileInfo, this.txtFile2, this.cmbSheet2);
                    }
                }
            }
        }

        private async void BtnReloadFile_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btnReloadFile)
            {
                int fileNum = Convert.ToInt32(btnReloadFile.Tag);
                if (fileNum == 1)
                {
                    await this.ReloadFile(this.excelManager1, this.cmbSheet1, this.cmbColumn1, this.cmbPasteColumn1, this.txtRowColumnCountSheet1, this.nupRowWithColumnNames1);
                }
                else if (fileNum == 2)
                {
                    await this.ReloadFile(this.excelManager2, this.cmbSheet2, this.cmbColumn2, this.cmbCopyColumn2, this.txtRowColumnCountSheet2, this.nupRowWithColumnNames2);
                }
            }
        }

        public async Task ReloadFile(ExcelManager excelManager, ComboBox cmbSheet, ComboBox cmbColumn, ComboBox cmbPasteOrCopyColumn, TextBlock txtRowColumnCountSheet, NumericUpDown nupRowWithColumnNames)
        {
            try
            {
                // ПЕРЕОТКРЫТИЕ ФАЙЛА
                excelManager.ReloadExcelFile();

                string? selectedColumnName = cmbColumn.SelectedItem as string ?? null; // нужно взять заранее, иначе в перезагрузке листа сбросится выбранный столбец
                string? selectedColumnName2 = cmbPasteOrCopyColumn.SelectedItem as string ?? null; // нужно взять заранее, иначе в перезагрузке листа сбросится выбранный столбец

                // ПЕРЕЗАГРУЗКА ЛИСТА
                string? selectedSheetName = cmbSheet.SelectedItem as string ?? null;
                if (selectedSheetName != null)
                {
                    txtRowColumnCountSheet.Text = "Лист не выбран";

                    cmbSheet.ItemsSource = excelManager.GetSheetNames();

                    cmbSheet.SelectedItem = null; // чтобы вызвать selectionChanged
                    cmbSheet.SelectedItem = selectedSheetName;
                }

                // ПЕРЕЗАГРУЗКА СТОЛБЦОВ
                if (selectedColumnName != null)
                {
                    await LoadColumns(excelManager, cmbColumn, cmbPasteOrCopyColumn, nupRowWithColumnNames);

                    cmbColumn.SelectedItem = selectedColumnName;
                    cmbPasteOrCopyColumn.SelectedItem = selectedColumnName2;
                }
            }
            catch (ArgumentNullException)
            {
                await this.ShowMessageAsync("Внимание!", "Выберите файл, прежде чем его обновить!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка при просмотре Excel файла", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void BtnShowFile_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btnShowFile)
            {
                int fileNum = Convert.ToInt32(btnShowFile.Tag);

                if (fileNum == 1)
                {
                    await this.ShowTableWindow(this.excelManager1);
                }
                else if (fileNum == 2)
                {
                    await this.ShowTableWindow(this.excelManager2);
                }
            }
        }

        private async Task ShowTableWindow(ExcelManager excelManager)
        {
            try
            {
                ShowTableWindow showTableWindow = new ShowTableWindow(excelManager);
                showTableWindow.ShowDialog();
            }
            catch (NullReferenceException)
            {
                await this.ShowMessageAsync("Внимание!", "Выберите файл, прежде чем его просмотреть!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка при просмотре Excel файла", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private async Task LoadFile(ExcelManager excelManager, FileInfo fileInfo, TextBlock txtFile, ComboBox cmbSheet)
        {
            try
            {
                excelManager.OpenExcelFile(fileInfo);
                txtFile.Text = $"Выбранный файл: {fileInfo.Name} ({fileInfo.FullName})";

                cmbSheet.ItemsSource = excelManager.GetSheetNames();
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
					this.ResetColumnsComponent(this.cmbColumn1, cmbPasteColumn1);
				}
				else if (fileNum == 2)
                {
                    this.LoadSheetInfo(this.excelManager2, sheetName, this.txtRowColumnCountSheet2);
                    this.ResetColumnsComponent(this.cmbColumn2, cmbCopyColumn2);
				}
            }
        }

        private void LoadSheetInfo(ExcelManager excelManager, string sheetName, TextBlock txtRowColumnCountSheet)
        {
            try
            {
                int columnCount = excelManager.GetSheetColumnCount(sheetName);
                int rowCount = excelManager.GetSheetRowCount(sheetName);

                txtRowColumnCountSheet.Text = $"Строк: {rowCount}\tСтолбцов: {columnCount}";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, $"Ошибка при получении информации о листе ({sheetName})", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

		private void NupRowWithColumnNames_ValueChanged(object sender, RoutedEventArgs e)
		{
			if (sender is NumericUpDown nupRowWithColumnNames)
			{
				nupRowWithColumnNames.Value ??= 1;
			}
		}

		private async void BtnCreateColumnNames_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btnCreateColumnNames)
            {
                int fileNum = Convert.ToInt32(btnCreateColumnNames.Tag);

                if (fileNum == 1)
                {
                    await this.CreateColumnNames(this.excelManager1);
                }
                else if (fileNum == 2)
                {
                    await this.CreateColumnNames(this.excelManager2);
                }
            }
        }

		private async Task CreateColumnNames(ExcelManager excelManager)
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



		private async void CmbColumn_DropDownOpened(object sender, EventArgs e)
		{
			if (sender is ComboBox сmbColumn)
			{
				int fileNum = Convert.ToInt32(сmbColumn.Tag);

				if (fileNum == 1)
				{
					await this.LoadColumns(this.excelManager1, this.cmbColumn1, this.cmbPasteColumn1, this.nupRowWithColumnNames1);
				}
				else if (fileNum == 2)
				{
                    await this.LoadColumns(this.excelManager2, this.cmbColumn2, this.cmbCopyColumn2, this.nupRowWithColumnNames2);
				}
			}
		}

        private async Task LoadColumns(ExcelManager excelManager, ComboBox cmbColumn, ComboBox cmbPasteOrCopyColumn, NumericUpDown nupRowWithColumnNames)
        {
            List<string> columnNames = new List<string>();

            try
            {
                double? rowNum = nupRowWithColumnNames.Value - 1;
				columnNames = excelManager.GetColumnNames(rowNum);
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
                cmbPasteOrCopyColumn.ItemsSource = columnNames.Append("[вставить в новый столбец]");
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

                ComparisonOptions copyOptions = new ComparisonOptions
                (
                    (bool) this.cbIgnoreEmptyCells.IsChecked!, 
                    (bool) this.cbYellowBackground.IsChecked!, 
                    (bool) this.cbGreenBackground.IsChecked!, 
                    (bool) this.cbCopyCellsFormat.IsChecked!,
				    (bool) this.cbSkipFurtherMatches.IsChecked!,
				    (bool) this.cbIngoreCase.IsChecked!, 
                    (bool) this.cbIngoreSpace.IsChecked!
                );

                await this.comparisonHelper.ComparingAsync(this.excelManager1, this.excelManager2, column1, column2, columnCopy, columnPaste, copyOptions);
            }
            catch (IOException)
            {
                await this.ShowMessageAsync("Внимание!", $"Файл не удалось сохранить, так как он уже где-то открыт. Закройте его и попробуйте ещё раз.");
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
                TextRange textRange = new TextRange(this.rtbProgress.Document.ContentEnd, this.rtbProgress.Document.ContentEnd);
                textRange.Text = e.Message + "\n";
				textRange.ApplyPropertyValue(TextElement.ForegroundProperty, e.Color);
				textRange.ApplyPropertyValue(Paragraph.MarginProperty, new Thickness(0));

				this.progressBar.Value = e.Progress * 100;
            });
        }


		private void ResetAllComponent(TextBlock txtFile, ComboBox cmbSheet, ComboBox cmbColumn, ComboBox cmbPasteOrCopyColumn, TextBlock txtRowColumnCountSheet, NumericUpDown nupRowWithColumnNames)
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
            this.cbSkipFurtherMatches.IsChecked = false;

            this.cbIngoreCase.IsChecked = false;
            this.cbIngoreSpace.IsChecked = true;
		}

        private void ResetColumnsComponent(ComboBox cmbColumn, ComboBox cmbPasteOrCopyColumn)
        {
			cmbColumn.ItemsSource = null;
			cmbPasteOrCopyColumn.ItemsSource = null;
		}


		private void MainWindow_Closed(object sender, EventArgs e)
        {
            this.excelManager1.Close();
            this.excelManager2.Close();
		}
    }
}