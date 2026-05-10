using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using Brush = System.Windows.Media.Brush;
using Brushes = System.Windows.Media.Brushes;
using ArcTool.Core.Models;
using ArcTool.Core.Services;
using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;
using Win32OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using Autodesk.Revit.DB;

namespace ArcTool.UI
{
    // ══════════════════════════════════════════════════════════════════════════
    //  HELPER TYPES
    // ══════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Đại diện một lựa chọn trong Region ComboBox.
    /// Gộp DisplayName (hiển thị) + RegionName (giá trị gửi đi) + RegionType (enum).
    /// </summary>
    public sealed class RegionOption
    {
        public string DisplayName { get; }
        public string RegionName { get; }
        public ExcelRegionType RegionType { get; }

        public RegionOption(string displayName, string regionName, ExcelRegionType regionType)
        {
            DisplayName = displayName;
            RegionName  = regionName;
            RegionType  = regionType;
        }

        // Static singletons cho các lựa chọn không thay đổi
        public static RegionOption PrintArea { get; } =
            new RegionOption("Print Area", null, ExcelRegionType.PrintArea);

        public static RegionOption UsedRange { get; } =
            new RegionOption("Used Range", null, ExcelRegionType.UsedRange);

        public static RegionOption NamedRange(string name) =>
            new RegionOption(name, name, ExcelRegionType.NamedRange);
    }

    /// <summary>
    /// Đại diện một lựa chọn trong View Type ComboBox.
    /// </summary>
    public sealed class ViewTypeOption
    {
        public string DisplayName { get; }
        public ExcelViewType Value { get; }

        public ViewTypeOption(string displayName, ExcelViewType value)
        {
            DisplayName = displayName;
            Value       = value;
        }
    }

    // ══════════════════════════════════════════════════════════════════════════
    //  ROW VIEW MODEL
    // ══════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// ViewModel cho mỗi dòng trong DataGrid.
    /// Wrap ExcelMapping, expose các computed property phục vụ binding WPF.
    ///
    /// Nguyên tắc mutation:
    ///   - Các property như WorkSheet, ViewType, AutoSync write-through xuống _mapping
    ///   - FilePath write-through + trigger reload SheetNames/RegionOptions
    ///   - ViewName là computed (read-only), cập nhật qua UpdateViewName()
    /// </summary>
    public sealed class ExcelMappingRowViewModel : INotifyPropertyChanged
    {
        private readonly ExcelMapping _mapping;
        private bool      _isSelected;
        private bool      _fileExists;
        private bool      _hasChanges;
        private RegionOption _selectedRegionOption;

        public ExcelMappingRowViewModel(ExcelMapping mapping)
        {
            _mapping = mapping ?? throw new ArgumentNullException(nameof(mapping));
        }

        public ExcelMapping Mapping => _mapping;
        public string Id => _mapping.Id;

        // ── SELECT ────────────────────────────────────────────────────────────

        public bool IsSelected
        {
            get => _isSelected;
            set { if (_isSelected == value) return; _isSelected = value; OnPropertyChanged(); }
        }

        // ── STATUS ────────────────────────────────────────────────────────────

        public bool FileExists => _fileExists;
        public bool HasChanges => _hasChanges;

        public Brush DotBrush =>
            !_fileExists ? Brushes.Goldenrod          // file bị move/xóa
            : _hasChanges ? Brushes.IndianRed          // có thay đổi chưa sync
                          : Brushes.MediumSeaGreen;    // đã sync

        public Brush UpdateBrush =>
            !CanUpdate   ? Brushes.LightGray
            : _hasChanges ? Brushes.IndianRed
                          : Brushes.MediumSeaGreen;

        public string StatusTooltip =>
            !_fileExists ? "File không tìm thấy. Click để chọn lại đường dẫn."
            : _hasChanges ? "Excel file có thay đổi. Click để update."
                          : "Excel file đã sync.";

        public bool CanUpdate => !AutoSync && _fileExists;

        // ── REVIT TARGET ─────────────────────────────────────────────────────

        public string ViewName => _mapping.ViewName;

        public ExcelViewType ViewType
        {
            get => _mapping.ViewType;
            set
            {
                if (_mapping.ViewType == value) return;
                _mapping.ViewType = value;
                OnPropertyChanged();
            }
        }

        // ── SYNC CONTROL ─────────────────────────────────────────────────────

        public bool AutoSync
        {
            get => _mapping.AutoSync;
            set
            {
                if (_mapping.AutoSync == value) return;
                _mapping.AutoSync = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanUpdate));
                OnPropertyChanged(nameof(UpdateBrush));
            }
        }

        public string LastModifiedText =>
            _mapping.LastModified == DateTime.MinValue
                ? "—"
                : _mapping.LastModified.ToString("dd/MM/yyyy HH:mm");

        // ── EXCEL SOURCE ─────────────────────────────────────────────────────

        public string WorkSheet
        {
            get => _mapping.WorkSheet;
            set
            {
                value ??= string.Empty;
                if (_mapping.WorkSheet == value) return;
                _mapping.WorkSheet = value;
                UpdateViewName();
                OnPropertyChanged();
            }
        }

        public string FilePath
        {
            get => _mapping.FilePath;
            set
            {
                value ??= string.Empty;
                if (_mapping.FilePath == value) return;
                _mapping.FilePath = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(FileDisplayText));
                OnPropertyChanged(nameof(FileTooltip));
            }
        }

        public string FileDisplayText =>
            string.IsNullOrWhiteSpace(_mapping.FilePath)
                ? "(chưa chọn file)"
                : Path.GetFileName(_mapping.FilePath);

        public string FileTooltip =>
            string.IsNullOrWhiteSpace(_mapping.FilePath)
                ? "Chưa chọn file"
                : _mapping.FilePath;

        // ── DROPDOWNS ─────────────────────────────────────────────────────────

        public ObservableCollection<string>       SheetNames    { get; } = new ObservableCollection<string>();
        public ObservableCollection<RegionOption> RegionOptions { get; } = new ObservableCollection<RegionOption>();

        public RegionOption SelectedRegionOption
        {
            get => _selectedRegionOption;
            set
            {
                if (ReferenceEquals(_selectedRegionOption, value)) return;
                _selectedRegionOption = value;

                if (value != null)
                {
                    _mapping.RegionType = value.RegionType;
                    _mapping.Region     = value.RegionName;
                    UpdateViewName();
                }

                OnPropertyChanged();
            }
        }

        // ── PUBLIC MUTATION API (dùng từ code-behind) ─────────────────────────

        /// <summary>Cập nhật trạng thái FileExists/HasChanges và notify tất cả dependent properties.</summary>
        public void SetStatus(bool fileExists, bool hasChanges)
        {
            _fileExists = fileExists;
            _hasChanges = hasChanges;

            OnPropertyChanged(nameof(FileExists));
            OnPropertyChanged(nameof(HasChanges));
            OnPropertyChanged(nameof(DotBrush));
            OnPropertyChanged(nameof(StatusTooltip));
            OnPropertyChanged(nameof(CanUpdate));
            OnPropertyChanged(nameof(UpdateBrush));
        }

        /// <summary>Notify UI refresh sau khi _mapping fields bị mutate bởi ExcelSyncEngine.</summary>
        public void RefreshFromMapping()
        {
            OnPropertyChanged(nameof(ViewName));
            OnPropertyChanged(nameof(LastModifiedText));
            OnPropertyChanged(nameof(FileDisplayText));
            OnPropertyChanged(nameof(FileTooltip));
            OnPropertyChanged(nameof(CanUpdate));
            OnPropertyChanged(nameof(UpdateBrush));
        }

        public void ReplaceSheetNames(IEnumerable<string> sheetNames)
        {
            SheetNames.Clear();
            foreach (string name in sheetNames ?? Enumerable.Empty<string>())
                if (!string.IsNullOrWhiteSpace(name))
                    SheetNames.Add(name);
        }

        public void ReplaceRegionOptions(IEnumerable<RegionOption> regionOptions)
        {
            RegionOptions.Clear();
            foreach (RegionOption opt in regionOptions ?? Enumerable.Empty<RegionOption>())
                if (opt != null)
                    RegionOptions.Add(opt);
        }

        /// <summary>
        /// Đồng bộ SelectedRegionOption với _mapping.RegionType + _mapping.Region sau khi RegionOptions thay đổi.
        /// Ưu tiên: NamedRange match by name → UsedRange → PrintArea (default).
        /// </summary>
        public void SyncSelectedRegionOption()
        {
            RegionOption selected = null;

            if (_mapping.RegionType == ExcelRegionType.NamedRange
                && !string.IsNullOrWhiteSpace(_mapping.Region))
            {
                selected = RegionOptions.FirstOrDefault(o =>
                    o.RegionType == ExcelRegionType.NamedRange &&
                    string.Equals(o.RegionName, _mapping.Region, StringComparison.OrdinalIgnoreCase));
            }
            else if (_mapping.RegionType == ExcelRegionType.UsedRange)
            {
                selected = RegionOptions.FirstOrDefault(o => o.RegionType == ExcelRegionType.UsedRange);
            }

            // Fallback: PrintArea — luôn là phần tử đầu tiên của list
            if (selected == null)
                selected = RegionOptions.FirstOrDefault(o => o.RegionType == ExcelRegionType.PrintArea);

            SelectedRegionOption = selected;
        }

        // ── PRIVATE ───────────────────────────────────────────────────────────

        private void UpdateViewName()
        {
            string viewName = _mapping.BuildViewName();
            if (_mapping.ViewName == viewName)
            {
                OnPropertyChanged(nameof(ViewName));
                return;
            }
            _mapping.ViewName = viewName;
            OnPropertyChanged(nameof(ViewName));
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    // ══════════════════════════════════════════════════════════════════════════
    //  WINDOW CODE-BEHIND
    // ══════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// WPF modal dialog cho Excel to Revit V3.0.
    ///
    /// DataContext = this (set trong constructor).
    /// DataGrid bind vào Rows (ObservableCollection).
    /// ViewTypeOptions bind vào ComboBox của cột "View Type".
    ///
    /// Event flow:
    ///   Window_Loaded → LoadMappingsIntoRows → RefreshAllStatuses → RunAutoSyncRows → RefreshAllStatuses
    ///   User thay đổi FilePath → BrowseForRow → LoadLookupData (1 lần duy nhất) → PersistMappings
    ///   User thay đổi WorkSheet → Row_PropertyChanged → LoadRegionOptionsForRow → PersistMappings
    ///   User nhấn Update → TryUpdateRow → ExcelSyncEngine.ExecuteUpdate → RefreshAllStatuses
    /// </summary>
    public partial class ExcelToRevitWindow : Window, INotifyPropertyChanged
    {
        private readonly Document                                    _doc;
        private readonly List<ExcelMapping>                          _mappings     = new List<ExcelMapping>();
        private readonly ObservableCollection<ExcelMappingRowViewModel> _rows      = new ObservableCollection<ExcelMappingRowViewModel>();
        private readonly IReadOnlyList<ViewTypeOption>               _viewTypeOptions = new[]
        {
            new ViewTypeOption("Drafting View", ExcelViewType.DraftingView),
            new ViewTypeOption("Legend View",   ExcelViewType.LegendView)
        };

        // Guards chống cascade events
        private bool _suppressRowEvents;
        private bool _isLoading;

        // ── CONSTRUCTOR ───────────────────────────────────────────────────────

        // Parameterless overload cho XAML designer — không dùng trong production
        public ExcelToRevitWindow() : this(null) { }

        public ExcelToRevitWindow(Document doc)
        {
            _doc = doc;
            InitializeComponent();
            DataContext = this;
        }

        // ── BINDING PROPERTIES ────────────────────────────────────────────────

        public ObservableCollection<ExcelMappingRowViewModel> Rows            => _rows;
        public IReadOnlyList<ViewTypeOption>                   ViewTypeOptions => _viewTypeOptions;

        // ── WINDOW LIFECYCLE ─────────────────────────────────────────────────

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            if (_isLoading) return;
            _isLoading = true;

            try
            {
                LoadMappingsIntoRows();
                RefreshAllStatuses();
                RunAutoSyncRows();
                RefreshAllStatuses(); // cập nhật sau AutoSync
            }
            catch (Exception ex)
            {
                RevitTaskDialog.Show("ArcTool Error", ex.Message);
                Close();
            }
            finally
            {
                _isLoading = false;
            }
        }

        // ── LOAD DATA ─────────────────────────────────────────────────────────

        private void LoadMappingsIntoRows()
        {
            _mappings.Clear();
            _mappings.AddRange(ArcToolSettingsService.LoadMappings(_doc));

            _rows.Clear();
            foreach (ExcelMapping mapping in _mappings)
            {
                var row = new ExcelMappingRowViewModel(mapping);
                row.PropertyChanged += Row_PropertyChanged;
                _rows.Add(row);
                LoadLookupData(row, defaultToFirstSheet: false);
            }
        }

        /// <summary>
        /// Load SheetNames + RegionOptions cho một row từ file Excel.
        /// Nếu file không tồn tại hoặc không mở được: giữ nguyên (không crash).
        ///
        /// QUAN TRỌNG: Luôn set _suppressRowEvents = true TRƯỚC khi gọi —
        /// vì method này set WorkSheet (→ trigger Row_PropertyChanged → vòng lặp vô tận).
        /// </summary>
        private void LoadLookupData(ExcelMappingRowViewModel row, bool defaultToFirstSheet)
        {
            _suppressRowEvents = true;
            try
            {
                row.ReplaceSheetNames(Array.Empty<string>());
                row.ReplaceRegionOptions(new[] { RegionOption.PrintArea });
                row.SyncSelectedRegionOption();

                if (!ArcToolSettingsService.FileExists(row.Mapping))
                    return;

                using (var excelService = new ExcelInteropService())
                {
                    if (!excelService.OpenFile(row.FilePath))
                        return;

                    List<string> sheetNames = excelService.GetSheetNames();
                    row.ReplaceSheetNames(sheetNames);

                    // Chọn sheet: default first nếu chưa có hoặc không tìm thấy sheet cũ
                    bool shouldSelectFirst = defaultToFirstSheet
                        || string.IsNullOrWhiteSpace(row.WorkSheet)
                        || !sheetNames.Any(s => string.Equals(s, row.WorkSheet, StringComparison.OrdinalIgnoreCase));

                    if (shouldSelectFirst && sheetNames.Count > 0)
                        row.WorkSheet = sheetNames[0];

                    // Load region options cho sheet được chọn
                    string sheetName = row.WorkSheet;
                    List<string> namedRanges = string.IsNullOrWhiteSpace(sheetName)
                        ? new List<string>()
                        : excelService.GetNamedRanges(sheetName);

                    bool includeUsedRange = row.Mapping.RegionType == ExcelRegionType.UsedRange;
                    row.ReplaceRegionOptions(BuildRegionOptions(namedRanges, includeUsedRange));
                    row.SyncSelectedRegionOption();
                }
            }
            finally
            {
                _suppressRowEvents = false;
            }
        }

        /// <summary>
        /// Load lại chỉ RegionOptions (không load lại SheetNames) khi user đổi WorkSheet.
        /// Gọi trong Row_PropertyChanged, lúc đó _suppressRowEvents đã = true.
        /// </summary>
        private void LoadRegionOptionsForRow(ExcelMappingRowViewModel row)
        {
            row.ReplaceRegionOptions(new[] { RegionOption.PrintArea });
            row.SyncSelectedRegionOption();

            if (!ArcToolSettingsService.FileExists(row.Mapping))
                return;

            using (var excelService = new ExcelInteropService())
            {
                if (!excelService.OpenFile(row.FilePath))
                    return;

                List<string> namedRanges = string.IsNullOrWhiteSpace(row.WorkSheet)
                    ? new List<string>()
                    : excelService.GetNamedRanges(row.WorkSheet);

                bool includeUsedRange = row.Mapping.RegionType == ExcelRegionType.UsedRange;
                row.ReplaceRegionOptions(BuildRegionOptions(namedRanges, includeUsedRange));
                row.SyncSelectedRegionOption();
            }
        }

        /// <summary>
        /// Xây dựng danh sách RegionOption: PrintArea đầu tiên, rồi NamedRanges, tuỳ chọn UsedRange cuối.
        /// UsedRange không hiện mặc định (theo spec UI: Print Areas + Named Ranges).
        /// Chỉ thêm UsedRange nếu mapping đang dùng UsedRange — để giữ nguyên lựa chọn hiện tại.
        /// </summary>
        private static List<RegionOption> BuildRegionOptions(IEnumerable<string> namedRanges, bool includeUsedRange)
        {
            var options = new List<RegionOption> { RegionOption.PrintArea };

            foreach (string name in namedRanges ?? Enumerable.Empty<string>())
                if (!string.IsNullOrWhiteSpace(name))
                    options.Add(RegionOption.NamedRange(name));

            if (includeUsedRange)
                options.Add(RegionOption.UsedRange);

            return options;
        }

        // ── STATUS MANAGEMENT ─────────────────────────────────────────────────

        private void RefreshAllStatuses()
        {
            IReadOnlyDictionary<string, MappingSyncStatus> statuses =
                ExcelSyncEngine.CheckForChanges(_mappings);

            foreach (ExcelMappingRowViewModel row in _rows)
            {
                if (statuses.TryGetValue(row.Id, out MappingSyncStatus status))
                {
                    row.SetStatus(status.FileExists, status.HasChanges);
                }
                else
                {
                    row.SetStatus(false, false);
                }
                row.RefreshFromMapping();
            }
        }

        /// <summary>
        /// Tự động update tất cả rows có AutoSync = true và HasChanges = true.
        /// Gọi một lần khi Window_Loaded, trước lần RefreshAllStatuses() cuối.
        /// </summary>
        private void RunAutoSyncRows()
        {
            List<ExcelMappingRowViewModel> pendingRows = _rows
                .Where(row => row.AutoSync && row.FileExists && row.HasChanges)
                .ToList();

            foreach (ExcelMappingRowViewModel row in pendingRows)
            {
                try
                {
                    if (ExcelSyncEngine.ExecuteUpdate(row.Mapping, _doc, _mappings))
                        row.RefreshFromMapping();
                }
                catch (Exception ex)
                {
                    RevitTaskDialog.Show("ArcTool Error", ex.Message);
                }
            }
        }

        // ── PERSIST ───────────────────────────────────────────────────────────

        private void PersistMappings()
        {
            if (_doc == null) return;
            try
            {
                ArcToolSettingsService.SaveMappings(_doc, _mappings);
            }
            catch (Exception ex)
            {
                RevitTaskDialog.Show("ArcTool Error", ex.Message);
            }
        }

        // ── ROW PROPERTY CHANGE ────────────────────────────────────────────────

        /// <summary>
        /// Lắng nghe thay đổi từ ViewModel rows để reload dữ liệu và persist.
        /// Guard _suppressRowEvents tránh vòng lặp cascade khi code set property.
        /// </summary>
        private void Row_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (_isLoading || _suppressRowEvents) return;
            if (sender is not ExcelMappingRowViewModel row) return;

            switch (e.PropertyName)
            {
                case nameof(ExcelMappingRowViewModel.WorkSheet):
                    // FilePath case đã được xử lý trực tiếp trong BrowseForRow (không qua event này)
                    // Chỉ reload RegionOptions — SheetNames đã có rồi
                    _suppressRowEvents = true;
                    try   { LoadRegionOptionsForRow(row); }
                    finally { _suppressRowEvents = false; }
                    PersistMappings();
                    break;

                case nameof(ExcelMappingRowViewModel.SelectedRegionOption):
                case nameof(ExcelMappingRowViewModel.ViewType):
                case nameof(ExcelMappingRowViewModel.AutoSync):
                    PersistMappings();
                    break;
            }
        }

        // ── TOOLBAR BUTTONS ───────────────────────────────────────────────────

        private void AddRow_Click(object sender, RoutedEventArgs e)
        {
            var mapping = new ExcelMapping();
            var row     = new ExcelMappingRowViewModel(mapping);
            row.PropertyChanged += Row_PropertyChanged;

            _mappings.Add(mapping);
            _rows.Add(row);

            _suppressRowEvents = true;
            try
            {
                row.ReplaceSheetNames(Array.Empty<string>());
                row.ReplaceRegionOptions(new[] { RegionOption.PrintArea });
                row.SyncSelectedRegionOption();
                row.SetStatus(false, false);
            }
            finally
            {
                _suppressRowEvents = false;
            }

            PersistMappings();
        }

        private void RemoveRows_Click(object sender, RoutedEventArgs e)
        {
            List<ExcelMappingRowViewModel> selectedRows =
                _rows.Where(row => row.IsSelected).ToList();

            if (selectedRows.Count == 0) return;

            foreach (ExcelMappingRowViewModel row in selectedRows)
            {
                row.PropertyChanged -= Row_PropertyChanged;
                _rows.Remove(row);
                _mappings.Remove(row.Mapping);
            }

            PersistMappings();
            RefreshAllStatuses();
        }

        private void UpdateAll_Click(object sender, RoutedEventArgs e)
        {
            List<ExcelMappingRowViewModel> pendingRows =
                _rows.Where(row => row.FileExists && row.HasChanges).ToList();

            foreach (ExcelMappingRowViewModel row in pendingRows)
                TryUpdateRow(row);

            RefreshAllStatuses();
        }

        // ── DATAGRID CELL BUTTONS ─────────────────────────────────────────────

        private void UpdateRow_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement el && el.Tag is ExcelMappingRowViewModel row)
            {
                TryUpdateRow(row);
                RefreshAllStatuses();
            }
        }

        private void BrowseFile_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not FrameworkElement el || el.Tag is not ExcelMappingRowViewModel row)
                return;

            BrowseForRow(row);
        }

        private void StatusDot_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not FrameworkElement el || el.Tag is not ExcelMappingRowViewModel row)
                return;

            // Chỉ mở Browse khi file không tìm thấy (dot màu vàng)
            if (!row.FileExists)
                BrowseForRow(row);
        }

        // ── HELPER METHODS ────────────────────────────────────────────────────

        /// <summary>
        /// Mở OpenFileDialog cho user chọn file Excel, sau đó load SheetNames + RegionOptions.
        ///
        /// BUG-P3-01 FIX:
        ///   _suppressRowEvents = true phải set TRƯỚC khi gán row.FilePath để ngăn
        ///   Row_PropertyChanged fire và gọi LoadLookupData lần thứ hai (double-call).
        ///   Thứ tự đúng: suppress → set FilePath → LoadLookupData (1 lần) → persist.
        /// </summary>
        private void BrowseForRow(ExcelMappingRowViewModel row)
        {
            var dialog = new Win32OpenFileDialog
            {
                Title  = "Chọn file Excel",
                Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*"
            };

            // Mặc định mở thư mục của file đang chọn (nếu có)
            if (!string.IsNullOrWhiteSpace(row.FilePath))
            {
                try   { dialog.InitialDirectory = Path.GetDirectoryName(row.FilePath); }
                catch { dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); }
            }
            else
            {
                dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }

            if (dialog.ShowDialog(this) != true) return;

            // BUG-P3-01 FIX: suppress trước khi gán FilePath
            // → Row_PropertyChanged bị chặn → LoadLookupData sẽ chỉ được gọi 1 lần bên dưới
            _suppressRowEvents = true;
            try
            {
                row.FilePath = dialog.FileName;       // property change bị suppress
                LoadLookupData(row, defaultToFirstSheet: true);  // gọi đúng 1 lần
                RefreshAllStatuses();
            }
            finally
            {
                _suppressRowEvents = false;
            }

            PersistMappings();
        }

        private bool TryUpdateRow(ExcelMappingRowViewModel row)
        {
            try
            {
                bool success = ExcelSyncEngine.ExecuteUpdate(row.Mapping, _doc, _mappings);
                row.RefreshFromMapping();
                return success;
            }
            catch (InvalidOperationException ex)
            {
                // Cấu hình sai: ViewName rỗng, không có Legend template, v.v.
                RevitTaskDialog.Show("ArcTool Error", ex.Message);
                return false;
            }
            catch (IOException ex)
            {
                // Không lưu được JSON
                RevitTaskDialog.Show("ArcTool Error", $"Không thể lưu settings:\n{ex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                RevitTaskDialog.Show("ArcTool Error", ex.Message);
                return false;
            }
        }

        // ── INOTIFYPROPERTYCHANGED ────────────────────────────────────────────

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
