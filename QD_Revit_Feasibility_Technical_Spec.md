# Báo cáo nghiên cứu tính năng Quick Dimension cho Revit trong ArcTool

Ngày lập: 2026-05-26  
Phạm vi: phân tích tài liệu PDF do người dùng cung cấp, video mô phỏng Lisp QD/QD2 trong AutoCAD, đối chiếu codebase ArcTool hiện tại, xác minh các API Revit 2026 liên quan.  
Ràng buộc phiên làm việc: chỉ đọc và nghiên cứu code; không sửa file nguồn ArcTool.

---

## 1. Kết luận điều hành

Tính năng Quick Dimension kiểu Lisp QD của AutoCAD có thể triển khai trong Revit, nhưng không thể bê nguyên tư duy AutoCAD sang một cách cơ học. AutoCAD QD đo theo giao điểm hình học 2D giữa một đường cắt và các đường/thực thể CAD. Revit Dimension không đo giữa các điểm ảo độc lập; nó cần `Reference` hợp lệ tới hình học, datum, hoặc element có khả năng dimension. Vì vậy tính năng “y hệt trải nghiệm” là khả thi, còn “y hệt thuật toán CAD 2D” là không đúng hướng.

Kết luận đề xuất là làm tính năng theo lộ trình 3 cấp. V1 nên triển khai QD cho mặt bằng với hai điểm người dùng pick trực tiếp bằng `Selection.PickPoint`, tự động tìm references của Grid và Wall/Column trong active Plan View, tạo một chain dimension chính và tùy chọn một total dimension. V2 mở rộng sang linked models, nhiều category hơn, tùy chọn DimensionType và offset. V3 mới xét trải nghiệm rubberband/dynamic preview nếu thật sự cần, vì Revit API không thân thiện với kiểu preview như AutoCAD command line.

Video mô phỏng cho thấy hành vi thực tế quan trọng hơn PDF ở một điểm: Lisp QD2 không chỉ tạo một chain dimension; nó tạo các segment liên tiếp phía trên và có một dimension tổng phía dưới. Các segment trong video gồm các khoảng như `300`, `2230`, `300`, `3530`, `2420`, `300`, `3900`, `300`, và tổng `13280`. Người dùng thao tác bằng lệnh `QD2`, chọn/định tuyến một đường ngang dưới hoặc qua cụm tường/trục, sau đó kết quả dimension nằm theo đường người dùng chỉ định. Đây là hành vi nên mô phỏng trong Revit.

Tôi phản bác một phần báo cáo PDF ở các điểm sau. Một là đề xuất dùng `ReferenceIntersector` như xương sống duy nhất cho tất cả wall/grid là quá nặng cho V1 và có rủi ro phụ thuộc 3D view, visibility, section box, link transform, face layer. Hai là đề xuất bắt người dùng vẽ Detail Line rồi tool xóa sau không cần thiết cho V1; Revit 2026 có `Selection.PickPoint(ObjectSnapTypes, string)` đủ để lấy hai điểm, dù không có rubberband đẹp như AutoCAD. Ba là cơ chế auto-group bằng `NewGroup` không nên là mặc định, vì Revit Group mang ý nghĩa model/annotation group nghiêm ngặt, có thể gây phiền hơn lợi cho dimension. Bốn là phần “DirectContext3D rubberband” nên xem là R&D, không nằm trong MVP.

Tận dụng code hiện có trong ArcTool được, nhưng chủ yếu là tận dụng pattern và một số thành phần nhỏ, không có module nào tái sử dụng trực tiếp thành QD engine. `ArrangeDimensionCommand.cs` có kinh nghiệm về linear dimension, `TransactionGroup`, vector perpendicular với `ViewDirection`, selection filter cho `Dimension`. `App.cs` đã có `Annotation Tools` panel phù hợp để thêm nút QD sau này. `CoordinateBatchService` và `CoordinateExtractionService` là mẫu kiến trúc tốt để tách command/service/model/result/logging. `CoordinateLogService` có thể tái sử dụng hoặc nhân bản thành log cho QD. Không nên gắn QD vào Filter Manager hoặc Excel stack.

---

## 2. Những gì video mô phỏng thể hiện

Video dài khoảng 22.77 giây, độ phân giải 1920x1080, nội dung là AutoCAD với overlay “QD2 — Dim nhanh mặt bằng”. Các frame chính cho thấy người dùng gõ hoặc chọn lệnh `QD2` trong danh sách lệnh gồm `QD`, `QDT`, `QD3`, `QD1`, `QD2`, `QDM`, `QDIM`. Lệnh yêu cầu người dùng chỉ định điểm bắt đầu/điểm kết thúc hoặc vị trí đặt dimension bằng dynamic input của AutoCAD.

Kết quả sau khi chạy lệnh là một hàng dimension segment nằm phía trên đường dimension, các đoạn đo lần lượt theo các vị trí giao cắt trên mặt bằng. Dãy số hiển thị rõ gồm nhiều đoạn tường/khoảng trống: `300`, `2230`, `300`, `3230` hoặc về sau `3530`, `2420`, `300`, `3900`, `300`. Ở trạng thái cuối, có thêm một dimension tổng phía dưới với giá trị `13280`, spanning toàn bộ chiều dài từ mốc trái đến mốc phải. Điều này cho thấy QD2 không chỉ tạo dimension từng cặp riêng rẽ; nó tạo chuỗi dimension liên tục và một tổng thể, rất gần với nhu cầu “dim nhanh mặt bằng” trong triển khai kiến trúc.

Điểm UX quan trọng là vị trí dimension không phải do thuật toán tự đoán hoàn toàn; người dùng vẫn điều khiển đường đặt dimension bằng thao tác pick/drag. Đây là chìa khóa khi chuyển sang Revit: thay vì cố làm preview giống AutoCAD, ta có thể yêu cầu người dùng pick hai điểm của dimension line. Với Revit, dimension line chính là `Line` truyền vào `doc.Create.NewDimension(view, line, referenceArray, dimensionType)`, nên hai điểm pick đã đủ quyết định vị trí, hướng và offset của kết quả.

---

## 3. Tóm tắt và phản biện báo cáo PDF

Báo cáo PDF đã nêu đúng bản chất cốt lõi: Revit Dimension cần `Reference`, không phải chỉ cần điểm tọa độ. Đây là điểm phân biệt quan trọng nhất giữa CAD 2D và BIM database. Khi `NewDimension` nhận `ReferenceArray`, các reference phải hợp lệ và cùng logic hình học với dimension line, nếu không Revit sẽ trả lỗi hoặc tạo kết quả không ổn định.

Báo cáo cũng đúng khi nhấn mạnh cần sort references theo khoảng cách dọc đường cắt. Nếu thứ tự reference không tăng dần theo hướng dimension line, chain dimension dễ bị đảo đoạn, segment bằng 0 hoặc lỗi invalid references. Đây nên là invariant trong QD engine.

Báo cáo cũng đúng khi cảnh báo về Grid. Grid không nên lấy reference theo cách tùy tiện từ `Grid.Curve.Reference`. Thực tế triển khai dimension tới datum trong Revit thường cần thử nghiệm rất kỹ. Candidate tốt là `new Reference(grid)`, vì nó biểu diễn reference tới datum element thay vì surface/curve reference không phù hợp. Tuy nhiên phần này phải được Revit-test trong file mẫu, không nên coi là chắc chắn tuyệt đối chỉ từ lý thuyết.

Điểm cần phản biện đầu tiên là báo cáo đề xuất `ReferenceIntersector` làm trung tâm raycasting. API này yêu cầu `View3D`, trả về references theo visibility/geometry trong 3D view và có thể tìm references trong linked models nếu bật `FindReferencesInRevitLinks`. Nó mạnh, nhưng V1 Quick Dimension mặt bằng không nhất thiết cần bắt đầu bằng raycast 3D. Đối với walls trong active plan, cách ổn định hơn là thu thập candidate elements trong view, lấy location/geometry với `Options.ComputeReferences = true`, tìm face/reference phù hợp theo hướng dimension line. Với walls thẳng, ta còn có thể lấy references dựa trên wall side faces qua geometry hoặc utility API nếu phù hợp. Raycasting nên là một strategy, không phải kiến trúc duy nhất.

Điểm cần phản biện thứ hai là “Detail Line rồi xóa” không phải UX tốt nhất cho ArcTool V1. Revit có `Selection.PickPoint(ObjectSnapTypes, string)` và `PickPoint` trả về `XYZ` trên active work plane. Dùng hai lần PickPoint là đủ để mô phỏng “chỉ một đường dimension”. Cách này ít side effect hơn, không tạo/xóa element tạm và không cần giao dịch chỉ để thu input.

Điểm cần phản biện thứ ba là DirectContext3D. Rubberband preview đẹp nhưng không đáng đưa vào MVP. Revit không có command-loop linh hoạt như AutoCAD. Preview theo `DirectContext3D` hoặc Idling có chi phí bảo trì, rủi ro flicker, rủi ro model lớn, và tạo thêm surface area debug. Với ArcTool hiện tại, ưu tiên đúng là engine ổn định trước, preview sau.

Điểm cần phản biện thứ tư là Group. Trong AutoCAD, group dimension là UX nhẹ. Trong Revit, `Group` là một entity mạnh, ảnh hưởng chọn/sửa, có behavior riêng khi copy, mirror, edit group, view-specific annotation. Nếu group mặc định, người dùng có thể khó chỉnh từng dimension. Đề xuất: V1 không group mặc định; nếu cần, lưu created ids trong result/log hoặc tạo tùy chọn “Create annotation group” sau khi đã test.

---

## 4. Feasibility trong Revit

### 4.1 Có thể làm được gì ngay trong V1

Có thể tạo linear dimension trong active view bằng `doc.Create.NewDimension(activeView, dimensionLine, referenceArray, dimensionType)` hoặc overload không truyền `DimensionType`. API Revit 2026 xác nhận overload `NewDimension(View, Line, ReferenceArray, DimensionType)` tồn tại và tạo linear dimension theo style chỉ định.

Có thể lấy hai điểm người dùng pick bằng `Selection.PickPoint(ObjectSnapTypes, string)`. Điểm này giúp thay thế AutoCAD rubberband ở MVP. Ta không có preview live đẹp như AutoCAD, nhưng flow hai điểm là tự nhiên với Revit.

Có thể thu thập Grid trong active view bằng `FilteredElementCollector(doc, activeView.Id).OfClass(typeof(Grid))` và kiểm tra giao điểm 2D giữa `Grid.Curve` với dimension line. Grid reference candidate nên là `new Reference(grid)` và phải test với `NewDimension`.

Có thể thu thập Wall/Column trong active view bằng collector theo category và view id. Với wall thẳng, có thể tìm hai mặt chính hoặc centerline tùy mode. Với column/family instance, khả thi nhưng khó hơn vì geometry family có nhiều face và reference có thể không ổn định nếu family không bật reference planes phù hợp. V1 nên ưu tiên `Grid + Wall`, sau đó mới `StructuralColumn/Column`.

Có thể tạo chain dimension nhiều references trong một `ReferenceArray`. Revit sẽ tạo dimension với nhiều segments nếu references hợp lệ và cùng một dimension line.

Có thể tạo total dimension riêng bằng một `ReferenceArray` chỉ gồm reference đầu và cuối, đặt trên line offset song song với chain dimension. Đây là cách mô phỏng kết quả cuối video có tổng `13280`.

### 4.2 Không nên hứa “y hệt AutoCAD” ở cấp thuật toán

AutoCAD entities là 2D linework, điểm giao là đủ. Revit dimension là annotation có liên kết tới model reference. Điều này tạo ra khác biệt căn bản. Nếu một wall không expose face reference phù hợp, một family column không có reference plane mạnh, hoặc geometry nằm trong link không cho reference usable trong host dimension, tool phải báo unsupported rõ ràng thay vì cố tạo dimension sai.

Revit Plan View có crop, view range, phase, design option, linked view visibility, discipline, detail level. Candidate nhìn thấy trong view chưa chắc geometry reference sẽ hợp lệ cho dimension trong view. Engine phải có validation và diagnostic.

### 4.3 Kết luận feasibility theo module

`Grid + Wall straight` là khả thi cao và nên là MVP. `Grid + Wall + rectangular Column` khả thi trung bình, nên vào V1.1 hoặc V2 sau khi test family references. `Linked model walls/grids` khả thi nhưng rủi ro cao hơn vì host dimension tới linked geometry có nhiều hạn chế; nên để V2. `Rubberband preview` khả thi nghiên cứu nhưng không nên làm trong V1. `Auto-group` khả thi về API nhưng không nên mặc định.

---

## 5. Đối chiếu ArcTool codebase hiện tại

### 5.1 File có thể tận dụng trực tiếp hoặc gần trực tiếp

`ArcTool.Core/Commands/ArrangeDimensionCommand.cs` là file gần nhất với tính năng QD. Nó đã có command manual transaction, chọn linear dimension bằng `ISelectionFilter`, dùng `TransactionGroup` để gộp nhiều thao tác thành một undo, đọc `Dimension.Curve as Line`, tính vector vuông góc bằng `baseDirection.CrossProduct(activeView.ViewDirection).Normalize()`, và move dimension theo offset. Logic này không tạo dimension mới, nhưng có nhiều kinh nghiệm hình học 2D trong view có thể reuse về cách tính vector, cách guard linear dimension, và cách tổ chức undo.

Điểm cần sửa nếu tái sử dụng pattern từ `ArrangeDimensionCommand`: hiện tại command thiếu guard `activeView.Scale == 0`, đang nằm trong BUG-06 của CLAUDE.md. QD engine cũng phải guard active view không hợp lệ, view không phải plan/section/elevation phù hợp, view direction null/không ổn định, và dimension line quá ngắn.

`ArcTool.Core/Utilities/SelectionFilters.cs` có `LinearDimensionSelectionFilter`. QD sẽ cần thêm filter mới, ví dụ `DetailLineSelectionFilter` nếu sau này hỗ trợ chọn line có sẵn, hoặc `QdCandidateSelectionFilter` nếu user chọn phạm vi element. File này là chỗ hợp lý để đặt selection filter nhỏ, nhưng không nên nhồi engine logic vào đây.

`ArcTool.Core/App.cs` đã có ribbon panel `Annotation Tools` với nút `Arrange Dimensions`. Đây là nơi tự nhiên để thêm `Quick Dimension` sau khi được phê duyệt code. Không nên tạo panel mới nếu không cần. ToolTip và LongDescription trong file này đang theo phong cách đơn giản, dễ mở rộng.

`ArcTool.Core/Services/CoordinateBatchService.cs` là pattern tốt cho QD: service trả về result object, command chịu trách nhiệm transaction/UI, service chịu trách nhiệm pipeline, per-element failure không làm chết toàn bộ batch. QD nên học mô hình này thay vì viết tất cả trong command.

`ArcTool.Core/Services/CoordinateExtractionService.cs` là pattern tốt về phân loại unsupported explicit. QD engine nên trả `QdReferenceCandidate` và `QdBuildSummary`, trong đó mỗi candidate có `Outcome`, `DiagnosticMessage`, `ElementId`, `SourceKind`, `DistanceAlongLine`.

`ArcTool.Core/Services/CoordinateLogService.cs` có thể tái sử dụng như mẫu logging. Nếu QD có nhiều failure do references, log cực kỳ quan trọng. Không nên silent catch như `CreateVoidFromLinkCommand` và `MultiCutCommand` đang có.

### 5.2 File chỉ nên dùng làm bài học, không reuse trực tiếp

`MultiCutCommand.cs` có anti-pattern đúng như PDF nói: nó chuyển selected refs thành `List<Element>` rồi lọc bằng LINQ `targetElements.Where(e => bbFilter.PassesFilter(e))`. Với QD, không nên gom tất cả elements rồi scan nếu có thể dùng collector theo active view/category trước. Tuy nhiên vì QD thường chạy trên active view và số wall/grid có thể vừa phải, giải pháp V1 có thể vẫn collector một lần theo category rồi lọc hình học trong memory, miễn là giới hạn theo view id và quick filters trước.

`CreateVoidFromLinkCommand.cs` có điểm mạnh là dùng `linkInstance.GetTotalTransform()` và `linkTransform.OfPoint(...)`. Đây là pattern bắt buộc nếu V2 hỗ trợ linked models. Nhưng command cũng có nhiều điểm không nên copy: WinForms nội tuyến trong command, lookup parameter theo string, catch trống, geometry helper nằm trong command. QD phải tách service ngay từ đầu.

`FilterWindow.xaml.cs` đang là skeleton UI và không liên quan trực tiếp đến QD. Không nên tận dụng Filter Manager cho QD. Nếu QD cần UI settings, hãy học pattern từ `CoordSettingsDialog.xaml` hoặc Excel WPF hơn là skeleton FilterWindow.

### 5.3 Khoảng trống cần thêm mới

Codebase hiện không có `ReferenceIntersector`, không có `ReferenceArray`, không có `NewDimension`, không có service dimension creation. Vì vậy QD là một feature mới, không phải mở rộng nhỏ của ArrangeDimension. Cần thêm command, model, service, selection filters, có thể thêm settings dialog nhẹ.

---

## 6. Kiến trúc đề xuất

### 6.1 Nguyên tắc kiến trúc

Command chỉ làm UI boundary: lấy `UIDocument`, guard active view, pick points, mở transaction/transaction group, gọi service, hiện summary. Command không chứa geometry scanning phức tạp.

Service tách thành các lớp nhỏ. `QuickDimensionInputService` hoặc command thu input. `QuickDimensionCandidateCollector` thu grid/wall/column candidates từ document/view. `QuickDimensionReferenceResolver` chuyển element/geometry thành usable `Reference`. `QuickDimensionGeometryService` xử lý project 2D, intersect, sort, tolerance. `QuickDimensionCreationService` tạo dimension. `QuickDimensionLogService` ghi diagnostics.

Model phải immutable hoặc ít nhất rõ trạng thái. Nên có record như `QdPickLine`, `QdReferenceCandidate`, `QdReferenceSource`, `QdBuildOptions`, `QdBuildSummary`, `QdFailureReason`. Naming theo project nên prefix domain, ví dụ `QdReferenceSource`, tránh enum chung kiểu `ReferenceType` dễ đụng namespace Revit.

Transaction boundary phải rõ. Input picking không cần transaction. Candidate collection không cần transaction nếu chỉ đọc. Tạo dimension cần transaction. Nếu tạo 2 dimension chain + total, dùng một `TransactionGroup` và một transaction hoặc hai transaction có rollback group. Nếu failure ở total dimension nhưng chain dimension thành công, V1 nên coi total là soft failure nếu option total bật, còn chain là primary output; nhưng quyết định này cần được chốt trước khi code.

### 6.2 File layout đề xuất

Đề xuất thêm các file sau khi bắt đầu triển khai, chưa sửa trong phiên này:

```text
ArcTool.Core/
├── Commands/
│   └── QuickDimensionCommand.cs
├── Models/
│   └── QuickDimensionContract.cs
├── Services/
│   ├── QuickDimensionGeometryService.cs
│   ├── QuickDimensionCandidateCollector.cs
│   ├── QuickDimensionReferenceResolver.cs
│   ├── QuickDimensionCreationService.cs
│   └── QuickDimensionLogService.cs
└── Utilities/
    └── QuickDimensionSelectionFilters.cs
```

Nếu có settings UI sau V1, thêm:

```text
ArcTool.Core/UI/
├── QuickDimensionSettingsDialog.xaml
└── QuickDimensionSettingsDialog.xaml.cs
```

Tôi không đề xuất WPF settings cho MVP nếu mục tiêu là prove feasibility nhanh. MVP có thể dùng constants hoặc một TaskDialog đơn giản, sau đó mới thêm settings.

### 6.3 Command flow đề xuất cho V1

`QuickDimensionCommand.Execute()` guard `uidoc`, `doc`, `activeView`. Chỉ cho chạy trong `ViewPlan` trước. Nếu active view là template, sheet, 3D, schedule, legend, drafting view không có model references, trả lỗi rõ.

Command gọi `PickPoint(ObjectSnapTypes.Endpoints | ObjectSnapTypes.Intersections | ObjectSnapTypes.Nearest, "Chọn điểm đầu đường dim QD")` và pick point thứ hai. Nếu user ESC, return `Result.Cancelled`. Nếu khoảng cách giữa hai điểm nhỏ hơn tolerance, return cancelled hoặc failed với message rõ.

Từ hai điểm tạo `Line dimensionLine = Line.CreateBound(p1, p2)`. Tính `direction = (p2 - p1).Normalize()`. Tính `normalInView = activeView.ViewDirection`. Tính `offsetDirection = direction.CrossProduct(normalInView).Normalize()` để đặt total line nếu cần.

Gọi collector lấy candidates trong active view. V1 nên gồm Grid và Wall. Tùy chọn column có thể bật sau khi wall ổn.

Gọi resolver tạo danh sách reference candidates. Mỗi candidate có distance along line: `(candidatePoint - p1).DotProduct(direction)`. Chỉ nhận distance trong `[0, lineLength]` với tolerance.

Sort theo distance, dedupe theo tolerance. Tolerance nên phân biệt: geometric tolerance nội bộ feet rất nhỏ, business tolerance để bỏ references quá sát nhau có thể 1 mm hoặc 2 mm tùy thực tế. Không hardcode 0.001 ft nếu chưa cân nhắc vì 0.001 ft khoảng 0.3048 mm, có thể quá nhỏ với wall layer hoặc import CAD; đề xuất V1 dùng 1 mm quy đổi sang internal feet.

Tạo `ReferenceArray` cho chain. Nếu ít hơn 2 references, báo không đủ reference. Gọi `doc.Create.NewDimension(activeView, dimensionLine, refArray, dimType)`.

Nếu option total bật, tạo line offset song song. Total references là first và last sau dedupe. Gọi `NewDimension` lần thứ hai. Không group mặc định.

### 6.4 Strategy thu Grid

Grid thu bằng collector trong view. Với mỗi `Grid`, lấy `grid.Curve`. Vì grid có thể là line hoặc arc, V1 chỉ hỗ trợ grid line thẳng; arc grid trả unsupported diagnostic. Project curve xuống mặt phẳng view hoặc ít nhất bỏ Z trong plan. Intersect giữa grid curve và dimension line để tìm điểm giao. Nếu có intersection trong segment dimension line, tạo candidate với `Reference = new Reference(grid)`, source `Grid`, distance tính theo point giao.

Điểm rủi ro: `new Reference(grid)` phải được test trong Revit 2026 với `NewDimension`. Báo cáo PDF nêu đây là trick đúng hướng, nhưng cần fixture test. Nếu fail, phải thử alternative như `grid.Curve.Reference` hoặc `grid.GetCurvesInView(...)` nếu cần, nhưng không được đoán trước khi test.

### 6.5 Strategy thu Wall

Wall thu bằng collector trong active view: `new FilteredElementCollector(doc, activeView.Id).OfClass(typeof(Wall)).OfCategory(BuiltInCategory.OST_Walls)`. Quick filter trước, sau đó LINQ geometry.

Có hai hướng giải quyết wall reference.

Hướng A: dùng geometry faces. Lấy geometry với `Options { ComputeReferences = true, IncludeNonVisibleObjects = false, View = activeView }` nếu phù hợp. Duyệt `Solid.Faces`, nhận `PlanarFace`. Tính `faceNormal`. Với dimension line direction, chỉ nhận face có `Abs(faceNormal.DotProduct(direction))` gần 1 nếu ta cần mặt vuông góc với hướng đo. Sau đó lấy face reference và intersection point/projection. Đây là general hơn nhưng tốn chi phí và phải xử lý compound wall nhiều lớp.

Hướng B: dùng wall location curve và side face references. Nếu API/utility có cách lấy references tới exterior/interior side faces ổn định, đây là hướng tốt hơn cho wall thẳng vì ít phải duyệt geometry. Tuy nhiên cần xác minh cụ thể trước code. Nếu không chắc, dùng geometry faces cho V1 nhưng giới hạn `Wall` thẳng và face planar.

Với video QD2, tool dường như đo theo các đường biên tường/đường trục mặt bằng. Trong Revit, cần quyết định mode: `Wall Finish Faces`, `Wall Core Faces`, hay `Wall Centerline`. MVP nên làm `Finish Faces` vì giống CAD linework nhất. Sau đó thêm setting `WallReferenceMode` nếu cần.

### 6.6 Strategy thu Column

Column khó hơn Wall vì family geometry và references không nhất quán giữa family. Với rectangular structural column, có thể lấy planar faces có normal song song dimension direction và compute intersection. Nhưng nếu family không có geometry reference hợp lệ hoặc column tròn, kết quả không như mong muốn. Đề xuất V1 chỉ thu walls + grids; V1.1 thêm rectangular columns sau khi test.

### 6.7 Strategy linked models

Linked model support không nên nằm trong V1. Nếu triển khai V2, phải xử lý `RevitLinkInstance`, `GetTotalTransform()`, linked document, linked element references. `ReferenceIntersector.FindReferencesInRevitLinks = true` là một route. Nhưng dimension trong host tới linked reference cần test nghiêm ngặt. Pattern từ `CreateVoidFromLinkCommand.cs` về `linkTransform.OfPoint(...)` là thứ có thể học lại.

### 6.8 Strategy ReferenceIntersector

`ReferenceIntersector` là công cụ mạnh cho V2 hoặc fallback. API Revit 2026 có constructor nhận `ElementFilter`, `FindReferenceTarget`, `View3D`, và `Find(origin, direction)` trả `IList<ReferenceWithContext>`. `ReferenceWithContext` có `Proximity`, `GetReference()` và `GetInstanceTransform()`. Nếu dùng, phải có service `QdRaycastService` tìm hoặc tạo 3D view riêng. Tạo 3D view cần transaction qua `View3D.CreateIsometric(Document, viewFamilyTypeId)` và `viewFamilyTypeId` phải là `ViewFamilyType` có `ViewFamily.ThreeDimensional`.

Nhưng raycast phụ thuộc View3D visibility. Nếu view 3D ẩn walls/grids hoặc có section box không bao phủ, kết quả thiếu. Nếu tự tạo view 3D, lại phát sinh element mới trong model. Vì vậy không nên là V1 mặc định.

---

## 7. Spec triển khai V1 chi tiết

### 7.1 Use case

Người dùng mở Plan View, bấm `ArcTool > Annotation Tools > Quick Dimension`. Tool yêu cầu chọn điểm đầu và điểm cuối của đường dimension. Tool tự tìm Grid và Wall giao với đoạn thẳng này trong active view. Tool tạo một chain dimension theo đoạn line người dùng chỉ định. Nếu bật total, tool tạo thêm một total dimension ở phía offset song song. Nếu không đủ reference, tool không tạo gì và báo lý do.

### 7.2 Input

Input bắt buộc là active document, active view, hai điểm `XYZ` trên active work plane, dimension type mặc định hoặc selected `DimensionType`, tùy chọn categories và wall reference mode. V1 có thể hardcode category gồm Grid và Wall, dimension type dùng default của Revit hoặc lấy type hiện hành nếu API cho phép. Nếu không có default dimension type hợp lệ, báo lỗi.

### 7.3 Output

Output là một `QdBuildSummary` gồm tổng số candidates tìm thấy, số references hợp lệ, số references bị loại do duplicate/tolerance, số unsupported per category, id của chain dimension, id của total dimension nếu có, và danh sách diagnostics ngắn.

### 7.4 Invariants

Dimension line phải nằm trong active view plane. Direction phải normalized. Reference candidates phải có distance tăng dần. Không tạo dimension nếu số reference sau dedupe nhỏ hơn 2. Không silently swallow exception; mọi failure per element phải đi vào diagnostics. Không giữ reference tới Revit element lâu hơn phạm vi API call/transaction. Không mở transaction trong service đọc dữ liệu. Chỉ service tạo dimension hoặc command mở transaction.

### 7.5 Tolerance

`EpsilonFeet` cho phép toán vector có thể là `1e-9` đến `1e-6` feet tùy phép toán. `DuplicateDistanceToleranceFeet` nên tương đương 1 mm bằng `UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters)`. `ParallelDotTolerance` nên nhận `Abs(dot) > 0.985` cho face gần song song với hướng đo, tương đương khoảng lệch dưới 10 độ, hoặc chặt hơn `0.999` nếu model sạch. V1 nên dùng chặt để tránh lấy nhầm face.

### 7.6 Error handling

User cancel ở bất kỳ PickPoint nào phải return `Result.Cancelled`. Active view không hỗ trợ phải return `Result.Failed` với TaskDialog. Không tìm thấy references phải return `Result.Cancelled` hoặc `Succeeded` kèm dialog “không tạo dimension” tùy UX; tôi đề xuất `Succeeded` nếu command chạy đúng nhưng không có output, để không làm Revit coi là lỗi add-in.

Exception trong transaction tạo dimension phải rollback transaction. Nếu dùng `TransactionGroup`, exception phải rollback group. Per-element geometry failure không được throw ra command; ghi vào summary.

### 7.7 Performance

Collector phải giới hạn theo active view id. Dùng quick filters: `OfClass(typeof(Wall))`, `OfCategory(BuiltInCategory.OST_Walls)`, `OfClass(typeof(Grid))`. Không dùng collector toàn document cho mỗi category nếu không cần. Không gọi `doc.GetElement()` lặp lại từ ElementId nếu collector đã trả element. Không gọi geometry extraction cho mọi element trong model; trước khi extract geometry, lọc nhanh bằng bounding box hoặc location curve giao với expanded outline của dimension line.

### 7.8 Logging

QD nên có log riêng `ArcTool_QD.log` cạnh `.rvt`, tương tự `CoordinateLogService`. Log cần ghi mode, active view id/name, line length, total candidates, success ids, failures. Logging failure không được chặn command.

---

## 8. API Revit 2026 đã xác minh

`NewDimension(View, Line, ReferenceArray, DimensionType)` tồn tại trong Revit API 2026 và tạo linear dimension bằng dimension style chỉ định. Đây là API tạo output chính.

`ReferenceIntersector.Find(XYZ origin, XYZ direction)` tồn tại và trả danh sách `ReferenceWithContext`. `ReferenceIntersector` có property `FindReferencesInRevitLinks`. Constructor quan trọng nhận `ElementFilter`, `FindReferenceTarget`, `View3D`.

`ReferenceWithContext` có `GetReference()`, `Proximity`, `GetInstanceTransform()`, `Dispose()`. Nếu dùng raycast, nên xử lý vòng đời cẩn thận.

`Grid.Curve` tồn tại để lấy hình học grid line. `Curve.Reference` tồn tại nhưng không chắc là reference phù hợp nhất cho dimension grid; cần test `new Reference(grid)` so với `grid.Curve.Reference`.

`Selection.PickPoint(ObjectSnapTypes, string)` tồn tại và trả `XYZ`, dùng active work plane và snap settings. Đây là route input V1 tốt hơn vẽ Detail Line tạm.

`ElementMulticategoryFilter(ICollection<BuiltInCategory>)` tồn tại và là quick filter theo category, hữu ích cho `ReferenceIntersector` hoặc collector đa category.

`Face.ComputeNormal(UV)` và `PlanarFace.FaceNormal` tồn tại để xác định hướng face khi chọn wall/column faces.

`View3D.CreateIsometric(Document, ElementId)` tồn tại, cần open transaction và `ViewFamilyType` loại ThreeDimensional. Dùng cho V2 nếu bắt buộc có raycast view riêng.

Tôi không tìm được trang RevitAPIDocs 2026 trực tiếp cho `Document.Create.NewGroup(ICollection<ElementId>)` trong search; kết quả gần nhất là trang 2024. Vì vậy không nên đưa group vào thiết kế V1 dựa trên giả định chưa xác minh 2026.

---

## 9. Rủi ro kỹ thuật và edge cases

Active view không phải plan view là rủi ro đầu tiên. QD mặt bằng nên chỉ chạy trong `ViewPlan`. Section/elevation có thể hỗ trợ sau nhưng geometry plane khác, direction/view normal khác, reference rules khác.

Work plane không set có thể làm `PickPoint` lỗi hoặc không như mong muốn. Trong Plan View thường ổn, nhưng vẫn cần catch `OperationCanceledException` và API exception.

Dimension line quá ngắn hoặc gần song song với view normal không hợp lệ phải reject sớm.

Wall joined, stacked wall, curtain wall, wall có nhiều layer có thể tạo nhiều face sát nhau. Dedupe tolerance phải loại face rác nhưng không được mất kích thước tường 100/200/300 mm.

Wall arc/curve không nên hỗ trợ trong V1. Nếu gặp, ghi unsupported. Arc grid cũng tương tự.

Linked elements có transform và reference khác host. Đưa vào V1 sẽ làm tăng rủi ro lớn.

Design options, phase, view template, crop region, hidden category, section box của 3D raycast có thể làm candidate thiếu hoặc sai. V1 collector theo active view ít nhất tôn trọng visibility của view hiện tại.

DimensionType có snap/format riêng. Tool không nên can thiệp unit format ở V1. Revit project units và dimension type sẽ quyết định hiển thị. Nếu sau này cần output mm fixed, phải xử lý qua dimension type/settings, không qua value raw.

User muốn dim theo face nào là câu hỏi nghiệp vụ lớn. Finish face, core face, centerline, grid center, column edge đều khác nhau. Video AutoCAD đang đo linework mặt bằng, nên V1 nên mặc định finish/visible edge, nhưng cần ghi rõ giới hạn.

---

## 10. Test plan đề xuất

Test đầu tiên là một file Revit đơn giản có 4 walls thẳng song song/chéo nhẹ và 5 grids thẳng. Chạy QD horizontal, kỳ vọng chain dimension có đúng số segment, thứ tự trái sang phải, không duplicate.

Test thứ hai dùng layout giống video: nhiều tường/khoảng có kích thước 300, 2230, 300, 3530, 2420, 300, 3900, 300. So sánh segment values và total. Đây là acceptance test chính.

Test thứ ba kiểm tra line direction ngược phải vẫn sort đúng. Người dùng pick phải sang trái thì dimension vẫn hợp lệ, không đảo loạn segment.

Test thứ tư kiểm tra không có đủ reference. Tool phải báo “found 0/1 usable references” và không crash.

Test thứ năm kiểm tra wall joined/layered. Tool không được tạo segment rác vài mm nếu mode không yêu cầu layer.

Test thứ sáu kiểm tra active view không hỗ trợ: 3D view, sheet, drafting view, schedule. Tool phải reject sạch.

Test thứ bảy kiểm tra performance với mặt bằng lớn khoảng 1000 walls/grids. Collector phải chạy trong thời gian chấp nhận được, không freeze dài, log summary rõ.

Test thứ tám sau V1.1 kiểm tra columns: rectangular, round, family custom thiếu references. Unsupported phải có diagnostics.

Test thứ chín sau V2 kiểm tra linked model với transform rotate/translate. Segment phải đúng trong host coordinates hoặc feature phải báo không hỗ trợ nếu Revit không chấp nhận linked reference.

---

## 11. Roadmap đề xuất

Giai đoạn 0 là spike Revit API trong một branch riêng, không đụng production feature khác. Mục tiêu duy nhất: tạo được dimension giữa hai grids bằng `new Reference(grid)` và giữa hai wall faces bằng references thật. Nếu bước này thất bại thì mọi UI/architecture đều vô nghĩa.

Giai đoạn 1 là MVP `QuickDimensionCommand` cho active Plan View, Grid + Wall, hai PickPoint, tạo chain dimension. Không UI settings, không linked model, không rubberband, không group.

Giai đoạn 2 thêm total dimension giống video QD2, offset song song. Đây nên làm ngay sau khi chain dimension ổn vì video thể hiện rõ total là phần giá trị của tính năng.

Giai đoạn 3 thêm settings tối thiểu: categories, wall reference mode, create total yes/no, duplicate tolerance, dimension type. Nếu dùng UI, theo WPF nhỏ như `CoordSettingsDialog`, không tạo window lớn.

Giai đoạn 4 thêm columns và linked models. Mỗi category mới phải có extraction/reference rules khóa rõ giống coordinate feature trước đây.

Giai đoạn 5 mới xét preview/rubberband hoặc chọn Detail Line có sẵn. Đây là UX polish, không phải core.

---

## 12. Quyết định khuyến nghị trước khi code

Tôi khuyến nghị chốt V1 như sau: chỉ active `ViewPlan`; input bằng hai `PickPoint`; output gồm chain dimension và total dimension tùy chọn mặc định bật; categories gồm Grid và Wall; không group; không linked model; không DirectContext3D; không WPF settings ở lần đầu. Mục tiêu là chứng minh Revit references ổn định trước.

Nếu sau test thực tế wall face references quá nhiễu, fallback V1 có thể giới hạn chỉ Grid dimension để chứng minh pipeline, sau đó mở Wall bằng một strategy riêng. Tuy nhiên với nhu cầu mô phỏng QD2 mặt bằng, Wall là giá trị chính nên spike wall face phải được làm sớm.

---

## 13. Nguồn và bằng chứng đã dùng

Nguồn nội bộ từ file người dùng cung cấp: `Mô phỏng Lisp QD trong Revit.pdf`, video `Lisp dim 3 lớp QD3 ... mp4`. Video được trích frame tại các mốc 2 giây để nhận diện flow QD2, chain segments và total dimension.

Nguồn codebase ArcTool đã đọc: `ArcTool.Core/Commands/ArrangeDimensionCommand.cs`, `ArcTool.Core/Utilities/SelectionFilters.cs`, `ArcTool.Core/App.cs`, `ArcTool.Core/Commands/MultiCutCommand.cs`, `ArcTool.Core/Commands/CreateVoidFromLinkCommand.cs`, `ArcTool.Core/UI/FilterWindow.xaml.cs`, `ArcTool.Core/Services/CoordinateBatchService.cs`, `ArcTool.Core/Services/CoordinateExtractionService.cs`, `ArcTool.Core/Services/CoordinateLogService.cs`, `ArcTool.Core/ArcTool.Core.csproj`.

Nguồn Revit API 2026: `NewDimension(View, Line, ReferenceArray, DimensionType)`, `ReferenceIntersector`, `ReferenceWithContext`, `Grid.Curve`, `Curve.Reference`, `Selection.PickPoint`, `View3D.CreateIsometric`, `ElementMulticategoryFilter`, `Face.ComputeNormal`, `PlanarFace.FaceNormal`. Các URL cụ thể được liệt kê trong phần Sources của phản hồi bàn giao.
