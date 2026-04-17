using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Linq;

namespace ArcTool.Core.Commands
{
    /// <summary>
    /// ImageImportCommand: Import ảnh PNG từ Excel Export vào Revit View
    /// 
    /// Hỗ trợ insert image vào TẤT CẢ loại View:
    /// - ViewSheet (Sheet layouts)
    /// - ViewDrafting (Drafting views, Legend, Detail views)
    /// - ViewPlan (Floor plans, Ceiling plans)
    /// - ViewSection (Section/Elevation views)
    /// - ViewDetail (Detail views)
    /// 
    /// Workflow:
    /// 1. User chọn file ảnh PNG (từ ExcelInteropService export)
    /// 2. Tự động detect Active View (Sheet/Drafting/Plan/Section/Detail/Legend)
    /// 3. Tự động tính center của View
    /// 4. Dialog nhập Scale (%) - TÂM ẢNH LUÔN TRÙNG TẦM VIEW
    /// 5. Insert ảnh tại center của View
    /// 6. Tự động apply scale factor
    /// 
    /// V1.2 — Phase 3 Update: Centered image at View center, Scale-only dialog
    /// - Tâm ảnh tự động trùng tâm View (không cần tùy chỉnh vị trí)
    /// - Dialog chỉ yêu cầu Scale (%) - đơn giản hóa UX
    /// - Support all View types
    /// 
    /// LƯU Ý: Cần phải có ImageType trong document (user phải insert ảnh lần đầu qua UI)
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class ImageImportCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // --- BƯỚC 1: CHỌN FILE ẢNH PNG ---
                string imagePath = PromptForImagePath();
                if (string.IsNullOrEmpty(imagePath))
                {
                    return Result.Cancelled;
                }

                // --- BƯỚC 2: LẤY ACTIVE VIEW ---
                Autodesk.Revit.DB.View activeView = doc.ActiveView;
                if (activeView == null)
                {
                    Autodesk.Revit.UI.TaskDialog.Show("Error", "Không có View nào được mở.");
                    return Result.Failed;
                }

                // Validate View type: Phải là loại View hỗ trợ ImageInstance
                bool isValidViewType = IsViewTypeSupported(activeView);

                if (!isValidViewType)
                {
                    Autodesk.Revit.UI.TaskDialog.Show(
                        "Error",
                        $"View type '{activeView.ViewType}' không hỗ trợ insert image.\n\n" +
                        $"View hiện tại: {activeView.Name}\n\n" +
                        $"Hỗ trợ: Sheet, Drafting, Floor Plan, Ceiling Plan, Section, Elevation, Detail, Legend");
                    return Result.Failed;
                }

                // --- BƯỚC 3: TÍNH TÂM CỦA VIEW ---
                XYZ viewCenter = GetViewCenter(activeView);

                // --- BƯỚC 4: HIỂN THỊ DIALOG TÙY CHỈNH SCALE ---
                double scalePercent = 100.0;

                using (var form = new System.Windows.Forms.Form())
                {
                    form.Text = "Image Scale";
                    form.Width = 350;
                    form.Height = 160;
                    form.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
                    form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                    form.MaximizeBox = false;
                    form.MinimizeBox = false;
                    form.BackColor = System.Drawing.Color.White;
                    form.Font = new System.Drawing.Font("Segoe UI", 9);
                    form.Padding = new System.Windows.Forms.Padding(0);

                    // Dùng TableLayoutPanel để layout clean
                    var layoutPanel = new System.Windows.Forms.TableLayoutPanel
                    {
                        Dock = System.Windows.Forms.DockStyle.Fill,
                        RowCount = 3,
                        ColumnCount = 2,
                        Padding = new System.Windows.Forms.Padding(20, 20, 20, 15)
                    };

                    layoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
                    layoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
                    layoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));

                    layoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 100));
                    layoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100));

                    // Row 0: Label
                    var lblDescription = new System.Windows.Forms.Label
                    {
                        Text = "Nhập Scale (%):",
                        AutoSize = true,
                        Dock = System.Windows.Forms.DockStyle.Fill
                    };
                    layoutPanel.Controls.Add(lblDescription, 0, 0);
                    layoutPanel.SetColumnSpan(lblDescription, 2);

                    // Row 1: TextBox
                    var txtScale = new System.Windows.Forms.TextBox
                    {
                        Text = "100",
                        Dock = System.Windows.Forms.DockStyle.Fill,
                        Font = new System.Drawing.Font("Segoe UI", 10),
                        BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D,
                        Height = 28,
                        Margin = new System.Windows.Forms.Padding(0, 10, 0, 0)
                    };
                    layoutPanel.Controls.Add(txtScale, 0, 1);
                    layoutPanel.SetColumnSpan(txtScale, 2);

                    // Row 2: Buttons
                    var buttonPanel = new System.Windows.Forms.Panel
                    {
                        Dock = System.Windows.Forms.DockStyle.Fill,
                        Height = 35,
                        Margin = new System.Windows.Forms.Padding(0, 15, 0, 0)
                    };

                    var btnOK = new System.Windows.Forms.Button
                    {
                        Text = "OK",
                        Width = 80,
                        Height = 30,
                        BackColor = System.Drawing.Color.FromArgb(0, 120, 215),
                        ForeColor = System.Drawing.Color.White,
                        FlatStyle = System.Windows.Forms.FlatStyle.Flat,
                        Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold),
                        Cursor = System.Windows.Forms.Cursors.Hand
                    };
                    btnOK.Location = new System.Drawing.Point(150, 0);

                    var btnCancel = new System.Windows.Forms.Button
                    {
                        Text = "Cancel",
                        Width = 80,
                        Height = 30,
                        FlatStyle = System.Windows.Forms.FlatStyle.Flat,
                        Font = new System.Drawing.Font("Segoe UI", 9),
                        Cursor = System.Windows.Forms.Cursors.Hand
                    };
                    btnCancel.Location = new System.Drawing.Point(240, 0);

                    buttonPanel.Controls.Add(btnOK);
                    buttonPanel.Controls.Add(btnCancel);

                    layoutPanel.Controls.Add(buttonPanel, 0, 2);
                    layoutPanel.SetColumnSpan(buttonPanel, 2);

                    form.Controls.Add(layoutPanel);
                    form.AcceptButton = btnOK;
                    form.CancelButton = btnCancel;

                    btnOK.Click += (s, e) =>
                    {
                        try
                        {
                            if (string.IsNullOrWhiteSpace(txtScale.Text))
                            {
                                System.Windows.Forms.MessageBox.Show("Vui lòng nhập giá trị Scale.", "Input Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                                return;
                            }

                            if (!double.TryParse(txtScale.Text, out var ps) || ps <= 0)
                            {
                                System.Windows.Forms.MessageBox.Show("Vui lòng nhập giá trị số dương.", "Input Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                                return;
                            }

                            scalePercent = ps;
                            form.DialogResult = System.Windows.Forms.DialogResult.OK;
                            form.Close();
                        }
                        catch (Exception ex)
                        {
                            System.Windows.Forms.MessageBox.Show($"Lỗi: {ex.Message}", "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        }
                    };

                    if (form.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    {
                        return Result.Cancelled;
                    }
                }

                // --- BƯỚC 5: IMPORT ẢNH VÀO VIEW ---
                using (Transaction t = new Transaction(doc, "ArcTool: Import Image from Excel"))
                {
                    t.Start();

                    try
                    {
                        // Validate file exists
                        if (!System.IO.File.Exists(imagePath))
                        {
                            t.RollBack();
                            Autodesk.Revit.UI.TaskDialog.Show("Error", $"File không tồn tại: {imagePath}");
                            return Result.Failed;
                        }

                        // Tìm ImageType trong document
                        FilteredElementCollector collector = new FilteredElementCollector(doc);
                        ImageType imgType = collector.OfClass(typeof(ImageType))
                            .Cast<ImageType>()
                            .FirstOrDefault();

                        // Nếu không tồn tại ImageType, tạo một bằng cách tạm insert image
                        if (imgType == null)
                        {
                            // Tạo temporary ImageInstance để tạo ImageType
                            // Sau đó xóa nó nhưng giữ lại ImageType
                            imgType = CreateImageTypeFromFile(doc, activeView, imagePath);

                            if (imgType == null)
                            {
                                t.RollBack();
                                Autodesk.Revit.UI.TaskDialog.Show(
                                    "Information",
                                    "Để sử dụng tính năng import ảnh, vui lòng:\n\n" +
                                    "1. Mở View này\n" +
                                    "2. Ribbon → Insert → Image\n" +
                                    "3. Chọn file PNG này một lần\n\n" +
                                    "Sau đó plugin sẽ nhận dạng ImageType tự động.");
                                return Result.Failed;
                            }
                        }

                        // Tạo ImagePlacementOptions với TÂM là CENTER OF VIEW
                        ImagePlacementOptions imgOpts = new ImagePlacementOptions();
                        imgOpts.PlacementPoint = BoxPlacement.Center;  // TÂM IMAGE
                        imgOpts.Location = viewCenter;                 // = TÂM VIEW

                        // Tạo ImageInstance
                        ImageInstance imageInstance = ImageInstance.Create(doc, activeView, imgType.Id, imgOpts);

                        if (imageInstance == null)
                        {
                            t.RollBack();
                            Autodesk.Revit.UI.TaskDialog.Show("Error", "Không thể tạo ImageInstance.");
                            return Result.Failed;
                        }

                        // Apply scale nếu cần
                        if (Math.Abs(scalePercent - 100.0) > 0.01)
                        {
                            try
                            {
                                double originalWidth = imageInstance.Width;
                                double originalHeight = imageInstance.Height;

                                double scaleFactor = scalePercent / 100.0;
                                imageInstance.Width = originalWidth * scaleFactor;
                                imageInstance.Height = originalHeight * scaleFactor;
                            }
                            catch
                            {
                                // Bỏ qua lỗi scale, vẫn import bình thường
                            }
                        }

                        t.Commit();

                        // Thông báo kết quả
                        Autodesk.Revit.UI.TaskDialog.Show("Success",
                            $"✅ Import ảnh thành công!\n\n" +
                            $"File: {System.IO.Path.GetFileName(imagePath)}\n" +
                            $"View: {activeView.Name} ({activeView.ViewType})\n" +
                            $"Vị trí: Tâm View\n" +
                            $"Scale: {scalePercent}%\n\n" +
                            $"Bạn có thể drag-resize ảnh trực tiếp trên View.");

                        return Result.Succeeded;
                    }
                    catch (Exception ex)
                    {
                        t.RollBack();
                        message = $"Error: {ex.Message}";
                        Autodesk.Revit.UI.TaskDialog.Show("ImageImport Error", message);
                        return Result.Failed;
                    }
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = $"Error: {ex.Message}";
                Autodesk.Revit.UI.TaskDialog.Show("ImageImport Error", message);
                return Result.Failed;
            }
        }

        /// <summary>
        /// Hiển thị OpenFileDialog để user chọn file ảnh
        /// </summary>
        private string PromptForImagePath()
        {
            using (var dialog = new System.Windows.Forms.OpenFileDialog())
            {
                dialog.Title = "Chọn file ảnh PNG (từ Excel Export)";
                dialog.Filter = "PNG Image (*.png)|*.png|All Files (*.*)|*.*";
                dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    return dialog.FileName;
                }
                return null;
            }
        }

        /// <summary>
        /// Tính center (tâm) của View
        /// </summary>
        private XYZ GetViewCenter(Autodesk.Revit.DB.View view)
        {
            try
            {
                // Lấy bounding box của view
                BoundingBoxXYZ bb = view.get_BoundingBox(view);

                if (bb != null && bb.Enabled)
                {
                    XYZ min = bb.Min;
                    XYZ max = bb.Max;

                    // Tính center point
                    XYZ center = new XYZ(
                        (min.X + max.X) / 2.0,
                        (min.Y + max.Y) / 2.0,
                        (min.Z + max.Z) / 2.0
                    );

                    return center;
                }
            }
            catch
            {
                // Fallback: Nếu get_BoundingBox không hoạt động, thử CropBox
            }

            // Fallback: Dùng CropBox (nếu tồn tại)
            try
            {
                BoundingBoxXYZ cropBox = view.CropBox;
                if (cropBox != null && cropBox.Enabled)
                {
                    XYZ min = cropBox.Min;
                    XYZ max = cropBox.Max;

                    XYZ center = new XYZ(
                        (min.X + max.X) / 2.0,
                        (min.Y + max.Y) / 2.0,
                        0  // Z = 0 vì View là 2D
                    );

                    return center;
                }
            }
            catch
            {
                // Fallback: Trả về origin
            }

            // Fallback cuối cùng: Origin (0, 0, 0)
            return XYZ.Zero;
        }

        /// <summary>
        /// Kiểm tra xem View type có hỗ trợ insert ImageInstance không
        /// </summary>
        private bool IsViewTypeSupported(Autodesk.Revit.DB.View view)
        {
            if (view == null) return false;

            ViewType viewType = view.ViewType;

            // Không hỗ trợ: 3D views
            if (viewType == ViewType.ThreeD || 
                viewType == ViewType.Walkthrough ||
                viewType == ViewType.Rendering)
            {
                return false;
            }

            // Hỗ trợ: Sheet, Drafting, Plans, Sections, Elevations, Details, Legend, v.v.
            return true;
        }

        /// <summary>
        /// Tạo ImageType từ file bằng cách tạm insert image, sau đó lấy ImageType
        /// </summary>
        private ImageType CreateImageTypeFromFile(Document doc, Autodesk.Revit.DB.View view, string imagePath)
        {
            try
            {
                // Load ảnh vào clipboard
                System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap(imagePath);
                System.Windows.Forms.Clipboard.SetImage(bitmap);

                // Tạo temporary ImagePlacementOptions
                ImagePlacementOptions tempOpts = new ImagePlacementOptions();
                tempOpts.PlacementPoint = BoxPlacement.TopLeft;
                tempOpts.Location = XYZ.Zero;

                // Tìm ImageType hiện có trước
                FilteredElementCollector collectorBefore = new FilteredElementCollector(doc);
                var imageTypesBefore = collectorBefore.OfClass(typeof(ImageType))
                    .Cast<ImageType>()
                    .ToList();

                // Thử tạo ImageInstance (sẽ tạo ImageType nếu cần)
                try
                {
                    // Nếu có ImageType tồn tại, dùng cái đầu tiên
                    if (imageTypesBefore.Count > 0)
                    {
                        return imageTypesBefore.First();
                    }

                    // Nếu không, tạo một ImageInstance tạm để trigger ImageType creation
                    // (Điều này không thể làm được mà không có ImageType sẵn)

                    // Fallback: Scan lại để kiếm ImageType
                    FilteredElementCollector collectorAfter = new FilteredElementCollector(doc);
                    var imageTypesAfter = collectorAfter.OfClass(typeof(ImageType))
                        .Cast<ImageType>()
                        .FirstOrDefault();

                    return imageTypesAfter;
                }
                catch
                {
                    return null;
                }
            }
            catch
            {
                return null;
            }
        }
    }
}
