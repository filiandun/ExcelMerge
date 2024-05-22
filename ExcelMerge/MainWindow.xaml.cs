using Microsoft.Win32;
using System.Windows;
using System.Windows.Controls;
using System.IO;

using OfficeOpenXml;


namespace ExcelMerge
{
    public partial class MainWindow : Window
    {
        private ExcelPackage excelPackage1;
        private ExcelPackage excelPackage2;

        public MainWindow()
        {
            InitializeComponent();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // иначе Exception
        }


        private void BtnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                Filter = "Excel Files (*.xlsx, *.xlsm)|*.xlsx;*.xlsm",
            };

            if (openFileDialog.ShowDialog() == true)
            {
                FileInfo fileInfo = new FileInfo(openFileDialog.FileName);

                Button button = sender as Button;
                if (button == null) { return; };

                switch (button.Name)
                {
                    case "btnOpenFile1":
                        this.OpenExcelFile(ref this.excelPackage1, fileInfo, this.txtFile1, this.cmbSheet1);
                        break;

                    case "btnOpenFile2":
                        this.OpenExcelFile(ref this.excelPackage2, fileInfo, this.txtFile2, this.cmbSheet2);
                        break;
                }
            }
        }

        private void OpenExcelFile(ref ExcelPackage excelPackage, FileInfo fileInfo, TextBlock textBox, ComboBox comboBox)
        {
            excelPackage = new ExcelPackage(fileInfo);

            textBox.Text = $"{fileInfo.Name} ({fileInfo.FullName})";
            comboBox.ItemsSource = ExcelHelper.GetSheetNames(excelPackage);
        }




        private void CmbSheet_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox comboBox = sender as ComboBox;
            if (comboBox == null || comboBox.SelectedIndex == -1) return;

            switch (comboBox.Name)
            {
                case "cmbSheet1":
                    LoadSheetColumns(this.excelPackage1, comboBox, this.cmbColumn1, this.cmbCopyColumn1);
                    break;

                case "cmbSheet2":
                    LoadSheetColumns(this.excelPackage2, comboBox, this.cmbColumn2, this.cmbPasteColumn2);
                    break;
            }
        }

        private void LoadSheetColumns(ExcelPackage excelPackage, ComboBox sheetComboBox, ComboBox columnComboBox1, ComboBox columnComboBox2)
        {
            if (excelPackage == null) return;

            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[sheetComboBox.SelectedItem.ToString()];

            columnComboBox1.ItemsSource = ExcelHelper.GetColumnNames(worksheet);
            columnComboBox2.ItemsSource = ExcelHelper.GetColumnNames(worksheet);
        }




        private async void btnCopy_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateInputs())
            {
                MessageBox.Show("Пожалуйста, выберите все необходимые параметры.");
            }

            var sheet1 = this.excelPackage1.Workbook.Worksheets[this.cmbSheet1.SelectedItem.ToString()];
            var sheet2 = this.excelPackage2.Workbook.Worksheets[this.cmbSheet2.SelectedItem.ToString()];

            int column1 = this.cmbColumn1.SelectedIndex + 1;
            int column2 = this.cmbColumn2.SelectedIndex + 1;

            int columnCopy = this.cmbCopyColumn1.SelectedIndex + 1;
            int columnPaste = this.cmbPasteColumn2.SelectedIndex + 1;

            CopyOptions copyOptions = new CopyOptions(this.cbIngoreCase.IsChecked, this.cbIngoreSpace.IsChecked, this.cbYellowBackground.IsChecked);

            int counter = await ExcelHelper.CopyingAsync(sheet1, sheet2, column1, column2, columnCopy, columnPaste, copyOptions);
            this.excelPackage2.Save();

            MessageBox.Show($"Копирование завершено {counter}"); // counter не совсем корретно работает, так как содержимое ячеек может повторяться
        }

        private bool ValidateInputs()
        {
            return this.excelPackage1 != null && this.excelPackage2 != null &&
                   this.cmbSheet1.SelectedItem != null && this.cmbSheet2.SelectedItem != null &&
                   this.cmbColumn1.SelectedIndex != -1 && this.cmbColumn2.SelectedIndex != -1 &&
                   this.cmbCopyColumn1.SelectedIndex != -1 && this.cmbPasteColumn2.SelectedIndex != -1;
        }
    }


    public struct CopyOptions
    {
        public bool? IgnoreCase { get; set; }
        public bool? IgnoreSpace { get; set; }
        public bool? YellowBackground { get; set; }

        public CopyOptions(bool? ignoreCase, bool? ignoreSpace, bool? yellowbackground) 
        {
            this.IgnoreCase = ignoreCase;
            this.IgnoreSpace = ignoreSpace;
            this.YellowBackground = yellowbackground;
        }
    }
}