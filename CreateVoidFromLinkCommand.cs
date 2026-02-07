using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace ArcTool.Core.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class CreateVoidFromLinkCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // --- BƯỚC 1: HIỆN FORM CHỌN FAMILY VOID (GIỮ NGUYÊN) ---
                var genericModelSymbols = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .OfCategory(BuiltInCategory.OST_GenericModel)
                    .Cast<FamilySymbol>()
                    .OrderBy(x => x.FamilyName)
                    .ToList();

                if (genericModelSymbols.Count == 0)
                {
                    Autodesk.Revit.UI.TaskDialog.Show("ArcTool Error", "Không tìm thấy Family Generic Model nào.");
                    return Result.Failed;
                }

                FamilySymbol selectedVoidSymbol = null;
                using (var form = new FamilySelectionForm(genericModelSymbols))
                {
                    if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK) selectedVoidSymbol = form.SelectedSymbol;
                    else return Result.Cancelled;
                }

                using (Transaction t = new Transaction(doc, "Activate Symbol"))
                {
                    t.Start();
                    if (!selectedVoidSymbol.IsActive) selectedVoidSymbol.Activate();
                    t.Commit();
                }

                // --- BƯỚC 2: CHỌN FILE LINK (GIỮ NGUYÊN) ---
                Reference linkRef = uidoc.Selection.PickObject(ObjectType.Element, new LinkSelectionFilter(), "Bước 1: Chọn File Link chứa dầm");
                RevitLinkInstance linkInstance = doc.GetElement(linkRef) as RevitLinkInstance;
                Document linkDoc = linkInstance.GetLinkDocument();
                Transform linkTransform = linkInstance.GetTotalTransform();

                // --- BƯỚC 3: LẤY TẤT CẢ DẦM TRONG LINK (TỰ ĐỘNG HÓA HOÀN TOÀN) ---
                // Không cần quét chọn PickBox nữa
                List<Element> linkedBeams = new FilteredElementCollector(linkDoc)
                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .WhereElementIsNotElementType()
                    .ToElements()
                    .ToList();

                if (linkedBeams.Count == 0)
                {
                    Autodesk.Revit.UI.TaskDialog.Show("Info", "File Link này không chứa dầm nào.");
                    return Result.Cancelled;
                }

                // --- BƯỚC 4: TẠO VOID HÀNG LOẠT ---
                using (Transaction t = new Transaction(doc, "ArcTool: Create All Voids"))
                {
                    t.Start();

                    int successCount = 0;
                    int totalBeams = linkedBeams.Count;

                    // Duyệt qua toàn bộ dầm trong file Link
                    for (int i = 0; i < totalBeams; i++)
                    {
                        Element linkedBeam = linkedBeams[i];

                        // Cập nhật status bar để biết tiến độ
                        uidoc.Application.Application.WriteJournalComment($"Generating Void {i + 1}/{totalBeams}", true);

                        try
                        {
                            // A. Lấy thông số
                            FamilyInstance beamInstance = linkedBeam as FamilyInstance;
                            if (beamInstance == null) continue;
                            FamilySymbol beamSymbol = beamInstance.Symbol;

                            double beamWidth = GetParamValueFromSymbol(beamSymbol, new[] { "b", "Width", "B", "Rộng" });
                            double beamHeight = GetParamValueFromSymbol(beamSymbol, new[] { "h", "Height", "H", "Depth", "Cao" });
                            double beamCutLength = GetParamValue(linkedBeam, new[] { "Cut Length", "Length", "Chiều dài" });

                            if (beamWidth == 0 || beamHeight == 0 || beamCutLength == 0) continue;

                            // B. Tính vị trí (LOGIC CHUẨN: LocationCurve)
                            LocationCurve locCurve = linkedBeam.Location as LocationCurve;
                            if (locCurve == null) continue;

                            // Chuyển tọa độ từ Link sang Host
                            XYZ startPoint = linkTransform.OfPoint(locCurve.Curve.GetEndPoint(0));
                            XYZ endPoint = linkTransform.OfPoint(locCurve.Curve.GetEndPoint(1));

                            // Tính trung điểm và hướng
                            XYZ midPoint = (startPoint + endPoint) / 2.0;
                            XYZ beamDir = (endPoint - startPoint).Normalize();

                            // C. Tạo Void
                            FamilyInstance voidInst = doc.Create.NewFamilyInstance(midPoint, selectedVoidSymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                            RotateInstanceToVector(doc, voidInst, beamDir, midPoint);

                            // D. Gán tham số (Height ÂM để đảo chiều như bạn yêu cầu)
                            SetParam(voidInst, "Width", beamWidth);
                            SetParam(voidInst, "Height", -beamHeight);
                            SetParam(voidInst, "Length", beamCutLength);

                            successCount++;
                        }
                        catch
                        {
                            // Bỏ qua dầm lỗi, chạy tiếp dầm sau
                            continue;
                        }
                    }

                    doc.Regenerate();
                    t.Commit();

                    Autodesk.Revit.UI.TaskDialog.Show("Success",
                        $"Đã hoàn tất!\n" +
                        $"- Tổng số dầm trong Link: {totalBeams}\n" +
                        $"- Số Void đã tạo thành công: {successCount}\n\n" +
                        $"Tiếp theo: Hãy dùng lệnh 'Multi-Cut Walls' để cắt tường.");
                }

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }


        // ====================================================
        //              CÁC HÀM HỖ TRỢ (GIỮ NGUYÊN)
        // ====================================================

        private double GetParamValueFromSymbol(FamilySymbol symbol, string[] paramNames)
        {
            if (symbol == null) return 0;
            try
            {
                foreach (Parameter param in symbol.Parameters)
                {
                    foreach (string searchName in paramNames)
                    {
                        if (param.Definition.Name.Equals(searchName, StringComparison.OrdinalIgnoreCase))
                        {
                            if (param.StorageType == StorageType.Double && param.HasValue) return param.AsDouble();
                        }
                    }
                }
            }
            catch { }
            return 0;
        }

        private double GetParamValue(Element elem, string[] paramNames)
        {
            if (elem == null) return 0;
            try
            {
                foreach (Parameter param in elem.Parameters)
                {
                    foreach (string searchName in paramNames)
                    {
                        if (param.Definition.Name.Equals(searchName, StringComparison.OrdinalIgnoreCase))
                        {
                            if (param.StorageType == StorageType.Double && param.HasValue) return param.AsDouble();
                        }
                    }
                }
            }
            catch { }
            return 0;
        }

        private void SetParam(Element elem, string paramName, double value)
        {
            Parameter p = elem.LookupParameter(paramName);
            if (p != null && !p.IsReadOnly) p.Set(value);
        }

        private void RotateInstanceToVector(Document doc, FamilyInstance inst, XYZ targetDir, XYZ center)
        {
            XYZ currentDir = inst.HandOrientation;
            double angle = currentDir.AngleTo(targetDir);
            if (Math.Abs(angle) > 0.001)
            {
                XYZ cross = currentDir.CrossProduct(targetDir);
                if (cross.IsZeroLength()) cross = XYZ.BasisZ;
                if (cross.Z < 0) angle = -angle;
                Line axis = Line.CreateBound(center, center + XYZ.BasisZ);
                ElementTransformUtils.RotateElement(doc, inst.Id, axis, angle);
            }
        }

        public class LinkSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => elem is RevitLinkInstance;
            public bool AllowReference(Reference reference, XYZ position) => false;
        }

        public class FamilySelectionForm : System.Windows.Forms.Form
        {
            public FamilySymbol SelectedSymbol { get; private set; }
            private System.Windows.Forms.ComboBox cmbFamilies;
            private System.Windows.Forms.Button btnOk;
            private System.Windows.Forms.Button btnCancel;
            private List<FamilySymbol> _symbols;

            public FamilySelectionForm(List<FamilySymbol> symbols)
            {
                _symbols = symbols;
                InitializeComponent();
            }

            private void InitializeComponent()
            {
                this.Text = "ArcTool - Chọn Void Family";
                this.Size = new System.Drawing.Size(350, 180);
                this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;

                System.Windows.Forms.Label lbl = new System.Windows.Forms.Label() { Text = "Vui lòng chọn Family Void:", Location = new System.Drawing.Point(20, 20), AutoSize = true };
                cmbFamilies = new System.Windows.Forms.ComboBox() { Location = new System.Drawing.Point(20, 50), Size = new System.Drawing.Size(290, 30), DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList };
                foreach (var sym in _symbols) cmbFamilies.Items.Add($"{sym.FamilyName} : {sym.Name}");
                if (cmbFamilies.Items.Count > 0) cmbFamilies.SelectedIndex = 0;

                btnOk = new System.Windows.Forms.Button() { Text = "OK", Location = new System.Drawing.Point(130, 100), DialogResult = System.Windows.Forms.DialogResult.OK };
                btnOk.Click += (s, e) => { if (cmbFamilies.SelectedIndex >= 0) SelectedSymbol = _symbols[cmbFamilies.SelectedIndex]; };
                btnCancel = new System.Windows.Forms.Button() { Text = "Cancel", Location = new System.Drawing.Point(220, 100), DialogResult = System.Windows.Forms.DialogResult.Cancel };

                this.Controls.Add(lbl); this.Controls.Add(cmbFamilies); this.Controls.Add(btnOk); this.Controls.Add(btnCancel);
                this.AcceptButton = btnOk; this.CancelButton = btnCancel;
            }
        }
    }
}
