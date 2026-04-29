# ARCTOOL — COPILOT BUG FIX INSTRUCTIONS

## VAI TRÒ CỦA COPILOT TRONG DỰ ÁN NÀY
Copilot CHỈ được dùng để fix bug trong file được chỉ định cụ thể.
Copilot KHÔNG tạo file mới, KHÔNG refactor, KHÔNG "cải thiện" code không liên quan.

---

## FILE ĐƯỢC PHÉP SỬA — CHỈ KHI ĐƯỢC YÊU CẦU RÕ RÀNG

Chỉ sửa file nào được nhắc tên trực tiếp trong prompt của user.
Nếu prompt không nhắc tên file → hỏi lại, không tự quyết định.

---

## FILE TUYỆT ĐỐI KHÔNG ĐƯỢC ĐỤNG VÀO

Các file sau đã được Claude (Anthropic) viết và verify.
Dù phát hiện vấn đề gì trong các file này → CHỈ báo cáo, không tự sửa:

## PHÂN QUYỀN VỚI 3 FILE KHUNG SƯỜN

CONTEXT.md, SKILL.md, ArcTool_Instructions.md:
- ✅ Copilot ĐƯỢC ĐỌC để hiểu kiến trúc và coding patterns
- ✅ Copilot ĐƯỢC THAM CHIẾU khi fix bug
- ❌ Copilot KHÔNG ĐƯỢC SỬA dù user yêu cầu
- ❌ Nếu user yêu cầu cập nhật 3 file này → trả lời:
     "Việc này cần thực hiện trên Claude web, không phải Copilot."

### Commands/
- CreateVoidFromLinkCommand.cs
- MultiCutCommand.cs  
- ArrangeDimensionCommand.cs
- FilterManagerCommand.cs

### Services/
- ExcelInteropService.cs

### UI/
- FilterWindow.xaml
- FilterWindow.xaml.cs

### Root/
- App.cs
- Utilities/SelectionFilters.cs
- Properties/Resources.resx
- Properties/Resources.Designer.cs

---

## QUY TẮC FIX BUG BẮT BUỘC

1. Chỉ sửa đúng dòng gây ra lỗi — không sửa thêm bất cứ thứ gì xung quanh
2. Không đổi tên biến, method, class dù thấy tên "không đẹp"
3. Không thêm using directive mới trừ khi bắt buộc để fix lỗi
4. Không chuyển đổi code style (var → explicit type, hoặc ngược lại)
5. Nếu fix 1 chỗ kéo theo phải sửa file khác → báo cho user quyết định

---

## KHI GẶP LỖI LIÊN QUAN ĐẾN CÁC RULE SAU — KHÔNG TỰ SỬA
Các pattern này có lý do kỹ thuật đặc biệt, sửa sai sẽ gây crash:

- `(long)elem.Category.Id.Value` → KHÔNG đổi thành (int)
- Không thêm `Marshal.ReleaseComObject()` sau `chartObj.Delete()`
- Không null biến trong `ReleaseObject()` — phải null ở caller
- `view.Duplicate(ViewDuplicateOption.WithDetailing)` cho Legend View
  → KHÔNG thay bằng bất kỳ Create() nào khác