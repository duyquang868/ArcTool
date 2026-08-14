using ArcTool.UI;
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
                // --- BƯỚC 1: HIỆN FORM CHỌN FAMILY VOID ---
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

                CreateVoidModeToolbar toolbar = new CreateVoidModeToolbar(genericModelSymbols);
                var helper = new System.Windows.Interop.WindowInteropHelper(toolbar);
                helper.Owner = Autodesk.Windows.ComponentManager.ApplicationWindow;

                bool? toolbarResult = toolbar.ShowDialog();
                if (toolbarResult != true || toolbar.SelectedSymbol == null || !toolbar.IsBulkMode.HasValue)
                {
                    return Result.Cancelled;
                }

                FamilySymbol selectedVoidSymbol = toolbar.SelectedSymbol;

                using (Transaction t = new Transaction(doc, "Activate Symbol"))
                {
                    t.Start();
                    if (!selectedVoidSymbol.IsActive) selectedVoidSymbol.Activate();
                    t.Commit();
                }

                // --- BƯỚC 2 & 3: CHỌN LINK HOẶC DẦM CỤ THỂ TRONG LINK ---
                LinkedBeamSelection selection = toolbar.IsBulkMode.Value
                    ? PromptForLinkBulkMode(uidoc, doc)
                    : PromptForLinkedBeamsMode(uidoc, doc);

                if (selection == null) return Result.Cancelled;

                RevitLinkInstance linkInstance = selection.LinkInstance;
                Document linkDoc = selection.LinkDocument;
                Transform linkTransform = selection.LinkTransform;
                List<Element> linkedBeams = selection.LinkedBeams;

                if (linkedBeams.Count == 0)
                {
                    string infoMessage = selection.IsSingleBeamMode
                        ? "Không lấy được dầm đã chọn trong file Link."
                        : "File Link này không chứa dầm nào.";
                    Autodesk.Revit.UI.TaskDialog.Show("Info", infoMessage);
                    return Result.Cancelled;
                }

                // --- BƯỚC 4: TẠO VOID ---
                using (Transaction t = new Transaction(doc, "ArcTool: Create Voids"))
                {
                    t.Start();

                    int successCount = 0;
                    int errorCount = 0;
                    int totalBeams = linkedBeams.Count;

                    // Duyệt qua toàn bộ dầm trong file Link
                    for (int i = 0; i < totalBeams; i++)
                    {
                        Element linkedBeam = linkedBeams[i];

                        // Cập nhật status bar để biết tiến độ
                        uidoc.Application.Application.WriteJournalComment($"Generating Void {i + 1}/{totalBeams}", true);

                        try
                        {
                            FamilyInstance beamInstance = linkedBeam as FamilyInstance;
                            if (beamInstance == null) continue;

                            LocationCurve locCurve = linkedBeam.Location as LocationCurve;
                            if (locCurve == null || locCurve.Curve == null) continue;

                            // A. Xử lý Tọa độ Không gian
                            XYZ startPoint = linkTransform.OfPoint(locCurve.Curve.GetEndPoint(0));
                            XYZ endPoint = linkTransform.OfPoint(locCurve.Curve.GetEndPoint(1));
                            XYZ midPoint = (startPoint + endPoint) / 2.0;
                            XYZ beamDir = (endPoint - startPoint).Normalize();

                            // B. Lấy thông số kích thước
                            double beamCutLength = startPoint.DistanceTo(endPoint);
                            // Try Instance parameters first, then fallback to Symbol parameters
                            double beamWidth = GetParamValue(beamInstance, new[] { "b", "Width", "B", "Rộng" });
                            if (beamWidth <= 0)
                                beamWidth = GetParamValue(beamInstance.Symbol, new[] { "b", "Width", "B", "Rộng" });

                            double beamHeight = GetParamValue(beamInstance, new[] { "h", "Height", "H", "Depth", "Cao" });
                            if (beamHeight <= 0)
                                beamHeight = GetParamValue(beamInstance.Symbol, new[] { "h", "Height", "H", "Depth", "Cao" });

                            if (beamWidth <= 0 || beamHeight <= 0 || beamCutLength <= 0)
                            {
                                errorCount++;
                                continue;
                            }

                            // C. TRÍCH XUẤT MẶT PHẲNG TỪ DẦM LINK (FACE-BASED APPROACH)
                            Options geomOptions = new Options { ComputeReferences = true };
                            GeometryElement geomElem = linkedBeam.get_Geometry(geomOptions);
                            PlanarFace targetFace = GetTopOrBottomFace(geomElem);

                            if (targetFace == null || targetFace.Reference == null)
                            {
                                errorCount++;
                                continue;
                            }

                            // D. CHUYỂN ĐỔI REFERENCE SANG LINKED REFERENCE
                            Reference linkedFaceRef = targetFace.Reference.CreateLinkReference(linkInstance);

                            // E. TẠO VOID DỰA TRÊN MẶT PHẲNG (FACE-BASED)
                            FamilyInstance voidInst = doc.Create.NewFamilyInstance(
                                linkedFaceRef,
                                midPoint,
                                beamDir,
                                selectedVoidSymbol
                            );

                            // F. Gán tham số kích thước
                            SetParam(voidInst, "Width", beamWidth);
                            SetParam(voidInst, "Height", -beamHeight);
                            SetParam(voidInst, "Length", beamCutLength);

                            successCount++;
                        }
                        catch
                        {
                            // Bỏ qua dầm lỗi, chạy tiếp dầm sau
                            errorCount++;
                            continue;
                        }
                    }

                    t.Commit();

                    Autodesk.Revit.UI.TaskDialog.Show("Success",
                        $"Đã hoàn tất!\n" +
                        $"- Chế độ: {selection.ModeLabel}\n" +
                        $"- Số dầm được xử lý: {totalBeams}\n" +
                        $"- Số Void đã tạo thành công: {successCount}\n" +
                        $"- Lỗi/Bỏ qua: {errorCount}\n\n" +
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
        //              CÁC HÀM HỖ TRỢ (ĐÃ TỐI ƯU HÓA)
        // ====================================================

        /// <summary>
        /// Kết quả phân giải nguồn dầm: hoặc toàn bộ dầm của File Link,
        /// hoặc chỉ các dầm cụ thể mà người dùng đã Tab-chọn trong Link.
        /// </summary>
        private class LinkedBeamSelection
        {
            public RevitLinkInstance LinkInstance { get; set; }
            public Document LinkDocument { get; set; }
            public Transform LinkTransform { get; set; }
            public List<Element> LinkedBeams { get; set; }
            public bool IsSingleBeamMode { get; set; }

            public string ModeLabel => IsSingleBeamMode
                ? "Dầm chỉ định trong Link (Tab)"
                : "Toàn bộ dầm trong Link";
        }


        /// <summary>
        /// Logic cũ: chọn File Link và lấy toàn bộ dầm trong đó.
        /// </summary>
        private LinkedBeamSelection PromptForLinkBulkMode(UIDocument uidoc, Document doc)
        {
            Reference linkRef;
            try
            {
                linkRef = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    new LinkSelectionFilter(doc),
                    "Bước 1: Chọn File Link chứa dầm");
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return null;
            }

            RevitLinkInstance linkInstance = doc.GetElement(linkRef) as RevitLinkInstance;
            Document linkDoc = linkInstance?.GetLinkDocument();
            if (linkDoc == null)
            {
                Autodesk.Revit.UI.TaskDialog.Show("ArcTool Error", "Không đọc được nội dung File Link đã chọn.");
                return null;
            }

            return new LinkedBeamSelection
            {
                LinkInstance = linkInstance,
                LinkDocument = linkDoc,
                LinkTransform = linkInstance.GetTotalTransform(),
                IsSingleBeamMode = false,
                LinkedBeams = new FilteredElementCollector(linkDoc)
                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .WhereElementIsNotElementType()
                    .ToElements()
                    .ToList()
            };
        }

        /// <summary>
        /// Logic mới: dùng LinkedElement để pick dầm cụ thể trong File Link.
        /// </summary>
        private LinkedBeamSelection PromptForLinkedBeamsMode(UIDocument uidoc, Document doc)
        {
            IList<Reference> pickedRefs;
            try
            {
                pickedRefs = uidoc.Selection.PickObjects(
                    ObjectType.LinkedElement,
                    new LinkedBeamSelectionFilter(doc),
                    "Bước 1: Chọn các dầm cụ thể trong File Link, rồi bấm Finish");
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return null;
            }

            if (pickedRefs == null || pickedRefs.Count == 0) return null;

            RevitLinkInstance targetLink = null;
            Document targetLinkDoc = null;
            List<Element> beams = new List<Element>();
            HashSet<long> seenBeamIds = new HashSet<long>();

            foreach (Reference r in pickedRefs)
            {
                RevitLinkInstance li = doc.GetElement(r.ElementId) as RevitLinkInstance;
                Document ld = li?.GetLinkDocument();
                if (li == null || ld == null) continue;

                if (targetLink == null)
                {
                    targetLink = li;
                    targetLinkDoc = ld;
                }
                else if (targetLink.Id.Value != li.Id.Value)
                {
                    continue;
                }

                Element beam = ld.GetElement(r.LinkedElementId);
                if (!IsStructuralFramingInstance(beam)) continue;
                if (seenBeamIds.Add(beam.Id.Value)) beams.Add(beam);
            }

            if (targetLink == null || targetLinkDoc == null)
            {
                Autodesk.Revit.UI.TaskDialog.Show("ArcTool Error", "Không xác định được File Link của dầm đã chọn.");
                return null;
            }

            return new LinkedBeamSelection
            {
                LinkInstance = targetLink,
                LinkDocument = targetLinkDoc,
                LinkTransform = targetLink.GetTotalTransform(),
                IsSingleBeamMode = true,
                LinkedBeams = beams
            };
        }

        /// <summary>
        /// Chỉ chấp nhận dầm (Structural Framing dạng FamilyInstance).
        /// </summary>
        private static bool IsStructuralFramingInstance(Element elem)
        {
            if (!(elem is FamilyInstance)) return false;

            Category cat = elem.Category;
            return cat != null && cat.Id.Value == (long)BuiltInCategory.OST_StructuralFraming;
        }

        /// <summary>
        /// Lấy giá trị tham số bằng LookupParameter (tối ưu hơn duyệt tất cả tham số)
        /// </summary>
        private double GetParamValue(Element elem, string[] paramNames)
        {
            if (elem == null) return 0;

            foreach (string name in paramNames)
            {
                Parameter p = elem.LookupParameter(name);
                if (p != null && p.StorageType == StorageType.Double && p.HasValue)
                {
                    return p.AsDouble();
                }
            }
            return 0;
        }

        private void SetParam(Element elem, string paramName, double value)
        {
            Parameter p = elem.LookupParameter(paramName);
            if (p != null && !p.IsReadOnly) p.Set(value);
        }

        /// <summary>
        /// Trích xuất mặt phẳng Top hoặc Bottom từ geometry của dầm
        /// </summary>
        private PlanarFace GetTopOrBottomFace(GeometryElement geomElem)
        {
            foreach (GeometryObject geomObj in geomElem)
            {
                Solid solid = geomObj as Solid;

                // Nếu dầm được bọc trong GeometryInstance (vì là Family)
                if (geomObj is GeometryInstance geomInst)
                {
                    solid = GetSolidFromInstance(geomInst);
                }

                if (solid != null && solid.Faces.Size > 0)
                {
                    foreach (Face face in solid.Faces)
                    {
                        if (face is PlanarFace pFace)
                        {
                            // Lấy mặt phẳng có pháp tuyến hướng lên hoặc xuống (Top/Bottom Face)
                            // Chấp nhận sai số cho dầm xiên (sloped beams)
                            if (Math.Abs(pFace.FaceNormal.Z) > 0.5)
                            {
                                return pFace;
                            }
                        }
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Trích xuất Solid từ GeometryInstance
        /// </summary>
        private Solid GetSolidFromInstance(GeometryInstance geomInst)
        {
            foreach (GeometryObject instObj in geomInst.GetInstanceGeometry())
            {
                if (instObj is Solid solid && solid.Volume > 0)
                {
                    return solid;
                }
            }
            return null;
        }

        /// <summary>
        /// Chỉ cho phép chọn thẳng Revit Link Instance cho chế độ bulk.
        /// </summary>
        public class LinkSelectionFilter : ISelectionFilter
        {
            public LinkSelectionFilter(Document hostDoc)
            {
            }

            public bool AllowElement(Element elem) => elem is RevitLinkInstance;
            public bool AllowReference(Reference reference, XYZ position) => false;
        }

        /// <summary>
        /// Cho phép chọn dầm cụ thể bên trong File Link bằng ObjectType.LinkedElement.
        /// </summary>
        public class LinkedBeamSelectionFilter : ISelectionFilter
        {
            private readonly Document _hostDoc;

            public LinkedBeamSelectionFilter(Document hostDoc)
            {
                _hostDoc = hostDoc;
            }

            public bool AllowElement(Element elem) => elem is RevitLinkInstance;

            public bool AllowReference(Reference reference, XYZ position)
            {
                if (reference == null || _hostDoc == null) return false;
                if (reference.LinkedElementId == ElementId.InvalidElementId) return false;

                RevitLinkInstance linkInstance = _hostDoc.GetElement(reference.ElementId) as RevitLinkInstance;
                Document linkDoc = linkInstance?.GetLinkDocument();
                if (linkDoc == null) return false;

                Element linkedElem = linkDoc.GetElement(reference.LinkedElementId);
                return IsStructuralFramingInstance(linkedElem);
            }
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