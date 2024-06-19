using System.Windows;
using System.Windows.Controls;

namespace ExcelMerge
{
    public partial class ShowTableWindow
    {
        private ExcelManager _excelManager;

        public ShowTableWindow(ExcelManager excelManager)
        {
            InitializeComponent();

            this._excelManager = excelManager;
            
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.LoadSheetNames();
            await this.LoadDataTable();
            this.LoadFirstColumnStyle();
        }

        private async void cmdSheet_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            await this.LoadDataTable();
            this.LoadFirstColumnStyle();
        }

        private void LoadSheetNames()
        {
            try
            {
                this.cmdSheet.ItemsSource = this._excelManager.GetSheetNames();
                this.cmdSheet.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка при чтении листа", MessageBoxButton.OK, MessageBoxImage.Error);
                this.Close();
            }
        }

        private async Task LoadDataTable()
        {
            try
            {
                string? selectedSheetName = this.cmdSheet.SelectedItem as string ?? null;
                if (selectedSheetName != null)
                {
                    this.dataGrid.ItemsSource = (await this._excelManager.GetTable(selectedSheetName)).DefaultView;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка при чтении листа", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadFirstColumnStyle()
        {
            var column = this.dataGrid.Columns.FirstOrDefault(c => c.Header.ToString() == "№");
            if (column != null)
            {
                column.CellStyle = (Style)FindResource("FirstColumn");
            }
        }
    }
}
