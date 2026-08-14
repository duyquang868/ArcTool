using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace ArcTool.UI
{
    public partial class CreateVoidModeToolbar : Window
    {
        public bool? IsBulkMode { get; private set; }
        public FamilySymbol SelectedSymbol { get; private set; }

        public CreateVoidModeToolbar(List<FamilySymbol> symbols)
        {
            InitializeComponent();

            FamilySymbolComboBox.ItemsSource = symbols
                .Select(symbol => new FamilySymbolOption(symbol))
                .ToList();

            if (FamilySymbolComboBox.Items.Count > 0)
            {
                FamilySymbolComboBox.SelectedIndex = 0;
            }
        }

        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            FamilySymbolOption option = FamilySymbolComboBox.SelectedItem as FamilySymbolOption;
            if (option == null || option.Symbol == null)
            {
                System.Windows.MessageBox.Show(this, "Vui lòng chọn Void Family.", "Create Void", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SelectedSymbol = option.Symbol;
            IsBulkMode = CreateFromLinkRadio.IsChecked == true;
            DialogResult = true;
            Close();
        }

        protected override void OnClosed(System.EventArgs e)
        {
            if (!IsBulkMode.HasValue)
            {
                DialogResult = false;
            }

            base.OnClosed(e);
        }

        private class FamilySymbolOption
        {
            public FamilySymbolOption(FamilySymbol symbol)
            {
                Symbol = symbol;
                DisplayName = $"{symbol.FamilyName} : {symbol.Name}";
            }

            public FamilySymbol Symbol { get; }
            public string DisplayName { get; }
        }
    }
}
