using System.Collections.Generic;
using System.Linq;
using System.Windows;
using WpfComboBox = System.Windows.Controls.ComboBox;
using WpfComboBoxItem = System.Windows.Controls.ComboBoxItem;
using WpfWindow = System.Windows.Window;

namespace ArcTool.UI
{
    /// <summary>
    /// Compact coordinate settings dialog used by RegisterCoordParamsCommand.
    /// The dialog is pure WPF UI and contains no Revit API interaction.
    /// </summary>
    public partial class CoordSettingsDialog : WpfWindow
    {
        private readonly List<(string UserLabel, string CodeKey)> _axisMappingOptions = new List<(string UserLabel, string CodeKey)>
        {
            ("Standard  (EW→X, NS→Y)", "Standard"),
            ("VN-2000   (NS→X, EW→Y)", "VN-2000")
        };

        private readonly List<(string UserLabel, string CodeKey)> _outputUnitOptions = new List<(string UserLabel, string CodeKey)>
        {
            ("Meters", "Meters"),
            ("Millimeters", "Millimeters")
        };

        private readonly List<(string UserLabel, string CodeKey)> _triggerFilterOptions = new List<(string UserLabel, string CodeKey)>
        {
            ("Structural Columns", "StructuralColumns"),
            ("Structural Foundations", "StructuralFoundations")
        };

        /// <summary>
        /// Initializes the coordinate settings dialog and pre-selects the current project settings.
        /// </summary>
        /// <param name="currentAxisMapping">Current persisted axis mapping key.</param>
        /// <param name="currentOutputUnit">Current persisted output unit key.</param>
        /// <param name="currentTriggerFilter">Current supported trigger filter key.</param>
        public CoordSettingsDialog(
            string currentAxisMapping,
            string currentOutputUnit,
            string currentTriggerFilter)
        {
            InitializeComponent();
            PopulateDropdowns();
            SetCurrentValues(currentAxisMapping, currentOutputUnit, currentTriggerFilter);
        }

        /// <summary>
        /// Gets the selected axis mapping key after the dialog is accepted.
        /// </summary>
        public string SelectedAxisMappingKey { get; private set; } = string.Empty;

        /// <summary>
        /// Gets the selected output unit key after the dialog is accepted.
        /// </summary>
        public string SelectedOutputUnitKey { get; private set; } = string.Empty;

        /// <summary>
        /// Gets the selected trigger filter key after the dialog is accepted.
        /// </summary>
        public string SelectedTriggerFilterKey { get; private set; } = string.Empty;

        private void PopulateDropdowns()
        {
            PopulateComboBox(AxisMappingComboBox, _axisMappingOptions);
            PopulateComboBox(OutputUnitComboBox, _outputUnitOptions);
            PopulateComboBox(TriggerFilterComboBox, _triggerFilterOptions);
        }

        private static void PopulateComboBox(
            WpfComboBox comboBox,
            IEnumerable<(string UserLabel, string CodeKey)> options)
        {
            comboBox.Items.Clear();

            foreach ((string userLabel, string codeKey) in options)
            {
                comboBox.Items.Add(new WpfComboBoxItem
                {
                    Content = userLabel,
                    Tag = codeKey
                });
            }
        }

        private void SetCurrentValues(
            string currentAxisMapping,
            string currentOutputUnit,
            string currentTriggerFilter)
        {
            SelectByCodeKey(AxisMappingComboBox, currentAxisMapping);
            SelectByCodeKey(OutputUnitComboBox, currentOutputUnit);
            SelectByCodeKey(TriggerFilterComboBox, currentTriggerFilter);
        }

        private static void SelectByCodeKey(WpfComboBox comboBox, string codeKey)
        {
            WpfComboBoxItem item = comboBox.Items
                .OfType<WpfComboBoxItem>()
                .FirstOrDefault(i => string.Equals(i.Tag as string, codeKey, System.StringComparison.OrdinalIgnoreCase));

            comboBox.SelectedItem = item ?? comboBox.Items.OfType<WpfComboBoxItem>().FirstOrDefault();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            SelectedAxisMappingKey = GetSelectedCodeKey(AxisMappingComboBox);
            SelectedOutputUnitKey = GetSelectedCodeKey(OutputUnitComboBox);
            SelectedTriggerFilterKey = GetSelectedCodeKey(TriggerFilterComboBox);
            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private string GetSelectedCodeKey(WpfComboBox comboBox)
        {
            if (comboBox?.SelectedItem is WpfComboBoxItem item && item.Tag is string codeKey)
            {
                return codeKey;
            }

            return string.Empty;
        }
    }
}
