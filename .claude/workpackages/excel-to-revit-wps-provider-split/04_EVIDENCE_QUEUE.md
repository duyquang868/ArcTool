# EXCEL TO REVIT — WPS PROVIDER SPLIT — OPERATOR EVIDENCE QUEUE

The master owns this file. Agents never write here.

Rules:
- Only the master asks the human for evidence. Agents raise `blocker:` instead.
- When evidence arrives, forward only the relevant routing excerpt to the analysis task.
- Heavy artifacts (PDF, PNG, screenshots, journals) are worker-read by path. The master records
  paths and routes them; it does not open them.
- Each request names: the runbook, what to run, and exactly what to return.

---

## Environment constraint driving this queue

Re-verified 2026-08-09 after the user installed WPS on this machine: `Excel.Application` resolves
(CLSID `00024500-0000-0000-c000-000000000046`); `KET.Application` now also resolves
(CLSID `45540001-5750-5300-4b49-4e47534f4655`, per-user/HKCU registration, launches via
`wps.exe /prometheus /et /Automation`); `ET.Application` and `Kingsoft.ET.Application` still resolve
to null. `et.exe` is present under `...\WPS Office\12.1.0.28032\office6\`.

Both engines are exercisable on this one machine now. `EV-1`, `EV-2`, and `EV-3` all run here — the
"operator's separate WPS machine" framing below is stale and no longer required. The earlier
"WPS absent, cannot be exercised here" note is superseded and void.

Runtime is still operator-owned (R1): the master/workers do not launch WPS, Excel, or Revit
themselves. The user runs the runbooks below on this machine and reports results back.

---

## Live queue

### EV-1 — Phase 5 — RUNBOOK READY
- runbook: `results/T4.2_result.md`
- needed for: `T5.1`
- asked on: —
- what the operator must run: on **this machine**, execute the WPS ProgID + late-bound member probe below. No Revit involved.
- operator runbook:
  1. Đóng toàn bộ cửa sổ WPS Writer/Spreadsheet/PDF đang mở. Expected result: không còn cửa sổ WPS nào trên màn hình. What to record: ảnh chụp màn hình desktop sạch hoặc ghi rõ "all WPS windows closed".
  2. Mở Task Manager hoặc Command Prompt và kiểm tra còn `et.exe` hay `wps.exe` nào đang chạy không bằng `tasklist | findstr /I "et.exe wps.exe"`. Expected result: không còn tiến trình WPS liên quan, hoặc nếu có thì bạn ghi lại đúng tên tiến trình. What to record: toàn bộ output của lệnh.
  3. Mở WPS Spreadsheet, vào Help/About hoặc Account/About, ghi lại đúng version/build number rồi đóng WPS lại. Expected result: lấy được chuỗi version đầy đủ, ví dụ `12.1.0.28032`. What to record: ảnh chụp cửa sổ About hoặc chuỗi version nguyên văn.
  4. Mở PowerShell và chạy nguyên khối script sau:
     ```powershell
     $progIds = @('KET.Application','ET.Application','Kingsoft.ET.Application','KWPS.Application')
     foreach ($id in $progIds) {
       try {
         $t = [Type]::GetTypeFromProgID($id, $false)
         if ($null -eq $t) {
           [pscustomobject]@{ ProgID = $id; Resolved = $false; TypeName = $null; CLSID = $null; Error = $null }
         }
         else {
           [pscustomobject]@{ ProgID = $id; Resolved = $true; TypeName = $t.FullName; CLSID = $t.GUID.Guid; Error = $null }
         }
       }
       catch {
         [pscustomobject]@{ ProgID = $id; Resolved = 'ERROR'; TypeName = $null; CLSID = $null; Error = $_.Exception.Message }
       }
     } | Format-Table -AutoSize
     ```
     Expected result: nhận được bảng 4 dòng cho `KET.Application`, `ET.Application`, `Kingsoft.ET.Application`, `KWPS.Application`. What to record: copy/paste toàn bộ output PowerShell.
  5. Trong cùng cửa sổ PowerShell, chạy nguyên khối script late-bound member probe sau:
     ```powershell
     $progId = 'KET.Application'
     $type = [Type]::GetTypeFromProgID($progId, $false)
     if ($null -eq $type) { throw "ProgID not registered: $progId" }
     $app = $null
     $books = $null
     $book = $null
     $sheets = $null
     $sheet = $null
     $names = $null
     $nameObj = $null
     $range = $null
     $pageSetup = $null
     $flagsGet = [Reflection.BindingFlags]'Public,Instance,GetProperty'
     $flagsSet = [Reflection.BindingFlags]'Public,Instance,SetProperty'
     $flagsCall = [Reflection.BindingFlags]'Public,Instance,InvokeMethod'
     $results = New-Object System.Collections.Generic.List[object]
     function Add-Result($name, $present, $detail) {
       $results.Add([pscustomobject]@{ Member = $name; Present = $present; Detail = $detail })
     }
     try {
       $app = [Activator]::CreateInstance($type)
       Add-Result 'CreateInstance' 'present' $null
       try { $type.InvokeMember('Visible', $flagsSet, $null, $app, @($false)); Add-Result 'Visible(set)' 'present' $null } catch { Add-Result 'Visible(set)' 'absent/error' $_.Exception.Message }
       try { $type.InvokeMember('DisplayAlerts', $flagsSet, $null, $app, @($false)); Add-Result 'DisplayAlerts(set)' 'present' $null } catch { Add-Result 'DisplayAlerts(set)' 'absent/error' $_.Exception.Message }
       try { $books = $type.InvokeMember('Workbooks', $flagsGet, $null, $app, $null); Add-Result 'Workbooks' 'present' $null } catch { Add-Result 'Workbooks' 'absent/error' $_.Exception.Message }
       if ($books) {
         $bookType = $books.GetType()
         $testPath = Read-Host 'Nhap duong dan day du toi file .xlsx thu nghiem'
         try { $book = $bookType.InvokeMember('Open', $flagsCall, $null, $books, @($testPath)); Add-Result 'Workbooks.Open' 'present' $null } catch { Add-Result 'Workbooks.Open' 'absent/error' $_.Exception.Message }
       }
       if ($book) {
         $bookType = $book.GetType()
         try { $sheets = $bookType.InvokeMember('Worksheets', $flagsGet, $null, $book, $null); Add-Result 'Worksheets' 'present' $null } catch { Add-Result 'Worksheets' 'absent/error' $_.Exception.Message }
         try { $names = $bookType.InvokeMember('Names', $flagsGet, $null, $book, $null); Add-Result 'Names' 'present' $null } catch { Add-Result 'Names' 'absent/error' $_.Exception.Message }
       }
       if ($sheets) {
         $sheetsType = $sheets.GetType()
         try { $sheet = $sheetsType.InvokeMember('Item', $flagsGet, $null, $sheets, @(1)); Add-Result 'Worksheets.Item(1)' 'present' $null } catch { Add-Result 'Worksheets.Item(1)' 'absent/error' $_.Exception.Message }
       }
       if ($sheet) {
         $sheetType = $sheet.GetType()
         try { $sheetType.InvokeMember('Name', $flagsGet, $null, $sheet, $null) | Out-Null; Add-Result 'Worksheet.Name' 'present' $null } catch { Add-Result 'Worksheet.Name' 'absent/error' $_.Exception.Message }
         try { $range = $sheetType.InvokeMember('UsedRange', $flagsGet, $null, $sheet, $null); Add-Result 'UsedRange' 'present' $null } catch { Add-Result 'UsedRange' 'absent/error' $_.Exception.Message }
         try { $pageSetup = $sheetType.InvokeMember('PageSetup', $flagsGet, $null, $sheet, $null); Add-Result 'PageSetup' 'present' $null } catch { Add-Result 'PageSetup' 'absent/error' $_.Exception.Message }
         try { $sheetType.InvokeMember('Range', $flagsGet, $null, $sheet, @('A1')) | Out-Null; Add-Result 'Range("A1")' 'present' $null } catch { Add-Result 'Range("A1")' 'absent/error' $_.Exception.Message }
       }
       if ($names) {
         $namesType = $names.GetType()
         try { $nameObj = $namesType.InvokeMember('Item', $flagsGet, $null, $names, @(1)); Add-Result 'Names.Item(1)' 'present' $null } catch { Add-Result 'Names.Item(1)' 'absent/error' $_.Exception.Message }
       }
       if ($nameObj) {
         $nameType = $nameObj.GetType()
         try { $nameType.InvokeMember('Name', $flagsGet, $null, $nameObj, $null) | Out-Null; Add-Result 'Name.Name' 'present' $null } catch { Add-Result 'Name.Name' 'absent/error' $_.Exception.Message }
         try { $range = $nameType.InvokeMember('RefersToRange', $flagsGet, $null, $nameObj, $null); Add-Result 'RefersToRange' 'present' $null } catch { Add-Result 'RefersToRange' 'absent/error' $_.Exception.Message }
       }
       if ($range) {
         $rangeType = $range.GetType()
         try { $rangeType.InvokeMember('Worksheet', $flagsGet, $null, $range, $null) | Out-Null; Add-Result 'Range.Worksheet' 'present' $null } catch { Add-Result 'Range.Worksheet' 'absent/error' $_.Exception.Message }
         try { $rangeType.InvokeMember('Address', $flagsGet, $null, $range, @($false,$false)) | Out-Null; Add-Result 'Range.Address(false,false)' 'present' $null } catch { Add-Result 'Range.Address(false,false)' 'absent/error' $_.Exception.Message }
       }
       if ($pageSetup) {
         $psType = $pageSetup.GetType()
         foreach ($member in 'PrintArea','Zoom','FitToPagesWide','FitToPagesTall','TopMargin','BottomMargin','LeftMargin','RightMargin','PaperSize') {
           try { $psType.InvokeMember($member, $flagsGet, $null, $pageSetup, $null) | Out-Null; Add-Result "PageSetup.$member(get)" 'present' $null } catch { Add-Result "PageSetup.$member(get)" 'absent/error' $_.Exception.Message }
         }
       }
       if ($sheet) {
         $sheetType = $sheet.GetType()
         $tempPdf = Join-Path $env:TEMP ('ArcTool_EV1_' + [guid]::NewGuid().ToString('N') + '.pdf')
         try {
           $sheetType.InvokeMember('ExportAsFixedFormat', $flagsCall, $null, $sheet, @(0, $tempPdf, 0, $false, $false, 1, 1, $false))
           Add-Result 'ExportAsFixedFormat' 'present' $tempPdf
         } catch {
           Add-Result 'ExportAsFixedFormat' 'absent/error' $_.Exception.Message
         }
       }
     }
     finally {
       if ($book) { try { $book.GetType().InvokeMember('Close', $flagsCall, $null, $book, @($false)); Add-Result 'Close' 'present' $null } catch { Add-Result 'Close' 'absent/error' $_.Exception.Message } }
       if ($app) { try { $type.InvokeMember('Quit', $flagsCall, $null, $app, $null); Add-Result 'Quit' 'present' $null } catch { Add-Result 'Quit' 'absent/error' $_.Exception.Message } }
     }
     $results | Format-Table -Wrap -AutoSize
     ```
     Expected result: nhận được bảng `Member / Present / Detail` cho toàn bộ member giả định của nhánh WPS. What to record: copy/paste toàn bộ output PowerShell và đường dẫn file `.xlsx` bạn đã nhập.
  6. Nếu script ở bước 5 tạo ra một file PDF trong `%TEMP%`, mở đúng file đó để xác nhận nó tồn tại rồi đóng lại. Expected result: file PDF tồn tại hoặc script báo lỗi rõ ràng. What to record: đường dẫn file PDF nếu có, hoặc lỗi nguyên văn nếu không có.
  7. Đóng PowerShell. Kiểm tra lại tiến trình bằng `tasklist | findstr /I "et.exe wps.exe"`. Expected result: không còn `et.exe`/`wps.exe` bị kẹt sau probe. What to record: toàn bộ output của lệnh.
  8. Gửi lại đầy đủ 4 nhóm bằng chứng: version/build WPS, bảng ProgID, bảng member, và mọi error text nguyên văn. Expected result: bộ bằng chứng đủ để phân tích mà không cần chạy lại. What to record: chính nội dung bạn gửi lại.
- what to return:
  - [ ] WPS version string và build number
  - [ ] toàn bộ bảng ProgID probe
  - [ ] toàn bộ bảng late-bound member probe
  - [ ] đường dẫn file `.xlsx` đã dùng và đường dẫn PDF tạm nếu được tạo
  - [ ] mọi error text nguyên văn
  - [ ] output `tasklist` trước và sau khi probe
- supplied on: —
- forwarded to: —

### EV-2 — Phase 5 — RUNBOOK READY
- runbook: `results/T4.2_result.md`
- needed for: `T5.1`
- asked on: —
- what the operator must run: trên **chính máy này**, chạy end-to-end WPS export và so sánh fidelity với MS Excel. Vì coordinator ưu tiên `Excel.Application`, phải tạm thời làm cho `Excel.Application` không resolve được trước khi chạy nhánh WPS.
- operator runbook:
  1. Chuẩn bị **scratch workbook** `.xlsx` chỉ dùng cho test, không dùng file production. Workbook phải có đủ 3 case: một named range, một sheet có print area, và một sheet không có named range lẫn print area để buộc UsedRange fallback. Expected result: có 1 file test duy nhất đáp ứng đủ 3 case. What to record: đường dẫn đầy đủ tới file `.xlsx`.
  2. Xác nhận build refactor đã được deploy lên máy này và bạn biết cách mở đúng lệnh Excel-to-Revit trong Revit test context. Expected result: sẵn sàng chạy UI test trên build mới. What to record: đường dẫn/thư mục build đang dùng hoặc ảnh chụp About/version của add-in nếu có.
  3. Trước khi force WPS, chạy lại PowerShell probe ngắn sau để xác nhận trạng thái hiện tại của hai ProgID chính:
     ```powershell
     'Excel.Application','KET.Application' | ForEach-Object {
       $t = [Type]::GetTypeFromProgID($_, $false)
       [pscustomobject]@{ ProgID = $_; Resolved = ($null -ne $t) }
     } | Format-Table -AutoSize
     ```
     Expected result: cả `Excel.Application` và `KET.Application` đều đang resolve. What to record: toàn bộ output.
  4. Tạm thời làm cho `Excel.Application` **không resolve được** trong thời gian test WPS, bằng cách unregister hoặc vô hiệu hóa COM registration của Microsoft Excel theo quy trình an toàn mà bạn kiểm soát trên máy này. Không uninstall Office. Đây là bước chỉ được làm trên môi trường scratch/test, không làm trên production machine nếu không có rollback rõ ràng. Expected result: Excel COM detection tạm thời bị vô hiệu hóa nhưng có thể khôi phục ngay sau test. What to record: chính xác bạn đã làm cách nào để disable và cách bạn sẽ restore.
  5. Chạy lại probe ở bước 3. Expected result: `Excel.Application` = `Resolved False`, `KET.Application` = `Resolved True`. Nếu không đạt đúng điều này thì **dừng EV-2** vì nhánh WPS chưa thực sự được force. What to record: toàn bộ output PowerShell.
  6. Đóng mọi cửa sổ WPS/Excel/Revit còn mở, rồi kiểm tra `tasklist | findstr /I "excel.exe et.exe wps.exe revit.exe"`. Expected result: môi trường sạch trước khi chạy. What to record: toàn bộ output.
  7. Mở Revit test context, mở Excel-to-Revit window, chọn file scratch workbook ở bước 1. Expected result: cửa sổ mở bình thường và chấp nhận file. What to record: ảnh chụp cửa sổ sau khi browse file, gồm đường dẫn file.
  8. Kiểm tra dropdown `WorkSheet`. Expected result: dropdown được populate từ workbook qua nhánh WPS. What to record: ảnh chụp dropdown mở ra và danh sách sheet hiển thị.
  9. Kiểm tra dropdown `Region` cho từng case: sheet có named range, sheet có print area, và sheet chỉ còn UsedRange fallback. Expected result: named range hiển thị đúng trên sheet tương ứng; sheet print-area hiển thị `PrintArea`; sheet fallback vẫn export được bằng UsedRange dù không có named range. What to record: 3 ảnh chụp hoặc 1 video ngắn cho cả 3 case, kèm ghi chú sheet nào ứng với case nào.
  10. Thực hiện update/import để tạo image trong Revit cho từng case cần chứng minh, ít nhất 1 lần cho named range và 1 lần cho non-named-range path. Expected result: image được đặt vào view, không xuất hiện dialog lỗi. What to record: ảnh chụp view trong Revit sau khi image xuất hiện, hoặc dialog text nguyên văn nếu lỗi.
  11. Thu thập artifact của nhánh WPS: intermediate PDF, cropped PNG, và nếu có log/debug output nào chỉ ra provider đang dùng là WPS thì chụp lại. Nếu UI không hiện provider name, dùng bằng chứng gián tiếp của bước 5 (`Excel.Application` false, `KET.Application` true) để chứng minh run này chỉ có thể đi vào nhánh WPS. Expected result: có đủ PDF + PNG của nhánh WPS. What to record: đường dẫn đầy đủ đến PDF và PNG, cùng mọi text/log liên quan.
  12. Khôi phục `Excel.Application` về trạng thái resolve được như ban đầu, rồi chạy lại cùng workbook/cùng region qua MS path để tạo bộ chứng cứ đối chiếu. Expected result: `Excel.Application` resolve lại, export MS chạy thành công. What to record: output probe sau restore, đường dẫn PDF và PNG của MS path.
  13. So sánh WPS vs MS trên cùng workbook/cùng region theo đúng 7 tiêu chí: print-area bounds, scaling/fit, margins, font substitution, merged cells, line weights, crop result. Sau đó kiểm tra cleanup bằng `tasklist | findstr /I "excel.exe et.exe wps.exe"` và liệt kê `%TEMP%\ArcTool_ExcelSync_*.pdf`. Expected result: có kết luận pass/fail cho từng tiêu chí, không còn orphan process, không còn temp PDF thừa sau khi app đóng. What to record: nhận xét pass/fail từng tiêu chí, ảnh/PDF/PNG của cả hai path, output `tasklist`, và output liệt kê file temp.
- what to return:
  - [ ] output probe trước khi disable Excel COM
  - [ ] mô tả chính xác cách đã force `Excel.Application` thành unresolved và cách restore
  - [ ] output probe sau khi force WPS
  - [ ] ảnh chụp sheet dropdown và region dropdown cho các case
  - [ ] kết quả Revit-side: image placed / not placed, kèm dialog text nguyên văn nếu lỗi
  - [ ] đường dẫn PDF + PNG của WPS path
  - [ ] đường dẫn PDF + PNG của MS path sau khi restore Excel
  - [ ] pass/fail cho 7 tiêu chí fidelity
  - [ ] output `tasklist` và listing `%TEMP%\ArcTool_ExcelSync_*.pdf`
- supplied on: —
- forwarded to: —

### EV-3 — Phase 5 — RUNBOOK READY
- runbook: `results/T4.2_result.md`
- needed for: `T5.2`
- asked on: —
- what the operator must run: trên **máy này với MS Excel hoạt động bình thường**, chạy non-regression matrix để chứng minh refactor không đổi hành vi cũ.
- operator runbook:
  1. Khôi phục trạng thái bình thường của Microsoft Excel trước khi bắt đầu: `Excel.Application` phải resolve được. Chạy probe ngắn sau:
     ```powershell
     'Excel.Application','KET.Application' | ForEach-Object {
       $t = [Type]::GetTypeFromProgID($_, $false)
       [pscustomobject]@{ ProgID = $_; Resolved = ($null -ne $t) }
     } | Format-Table -AutoSize
     ```
     Expected result: `Excel.Application` = `Resolved True`. What to record: toàn bộ output.
  2. Chuẩn bị một **scratch copy** của workbook test để chạy update/import. Nếu có pre-refactor baseline PNG/PDF của cùng workbook/cùng region thì đặt sẵn để đối chiếu. Expected result: có file scratch riêng và có baseline nếu sẵn có. What to record: đường dẫn workbook scratch và đường dẫn baseline PNG/PDF.
  3. Đóng mọi cửa sổ Excel đang mở, sau đó chạy `tasklist | findstr /I "excel.exe"`. Expected result: không còn `EXCEL.EXE` trước khi test. What to record: toàn bộ output.
  4. Mở Revit test context, mở Excel-to-Revit window, browse tới workbook scratch. Expected result: cửa sổ load file bình thường. What to record: ảnh chụp cửa sổ sau khi browse file.
  5. Kiểm tra `WorkSheet` dropdown khi load lần đầu. Expected result: dropdown populate đúng như trước refactor. What to record: ảnh chụp dropdown mở ra.
  6. Chọn một sheet có named range và kiểm tra `Region` dropdown. Expected result: dropdown region populate đúng các named range và vẫn có `PrintArea` khi phù hợp. What to record: ảnh chụp dropdown region.
  7. Đóng và mở lại row/mapping đã lưu trước đó để kiểm tra restore behavior. Expected result: lựa chọn sheet/region trước đó được restore đúng; nếu không có giá trị đã lưu thì logic default-to-first-sheet vẫn hoạt động như cũ. What to record: ảnh chụp trước/sau reload chứng minh selection được giữ hoặc default đúng.
  8. Chạy export/import cho 3 đường dẫn nội dung tối thiểu: named range, print area, và neither/UsedRange fallback. Expected result: cả 3 path đều tạo được image như trước refactor. What to record: pass/fail từng path, kèm ảnh chụp view sau mỗi lần import.
  9. Kiểm tra existing mapping update path. Expected result: update thay image cũ và giữ Smart Scale size như trước. What to record: ảnh chụp trước update và sau update, kèm ghi chú kích thước không đổi nếu nhìn thấy được.
  10. Kiểm tra first-time import path. Expected result: image mới lần đầu vẫn vào với default width 2000 mm. What to record: ảnh chụp hoặc thông số thể hiện width mặc định nếu UI/Revit property hiển thị được.
  11. Chạy 3 failure paths riêng biệt bằng dữ liệu scratch: (a) missing file, (b) wrong sheet name, (c) workbook đang bị Excel mở/lock. Expected result: mỗi case đều thất bại theo cách cũ và hiện đúng thông điệp `InvalidOperationException` tiếng Việt như trước refactor. What to record: ảnh chụp từng dialog lỗi và chép lại nguyên văn text tiếng Việt.
  12. So sánh PNG sau refactor với baseline pre-refactor của cùng region để bắt crop/DPI regression. Expected result: không có khác biệt nhìn thấy được về crop hoặc độ nét; nếu có khác biệt thì ghi rõ. What to record: cả hai file PNG/PDF đối chiếu và ghi chú pass/fail.
  13. Kết thúc test, đóng dialog/app liên quan rồi chạy `tasklist | findstr /I "excel.exe"` và liệt kê `%TEMP%\ArcTool_ExcelSync_*.pdf`. Expected result: không còn orphan `EXCEL.EXE` và không còn temp PDF thừa. What to record: toàn bộ output hai lệnh.
- what to return:
  - [ ] output probe xác nhận `Excel.Application` resolve
  - [ ] đường dẫn workbook scratch và baseline PNG/PDF
  - [ ] ảnh chụp WorkSheet dropdown, Region dropdown, và row reload/restore behavior
  - [ ] pass/fail cho named range, print area, UsedRange fallback
  - [ ] bằng chứng update existing mapping giữ Smart Scale
  - [ ] bằng chứng first-time import giữ default width 2000 mm
  - [ ] ảnh chụp + text nguyên văn của 3 failure dialogs tiếng Việt
  - [ ] so sánh PNG/PDF pre-refactor vs post-refactor
  - [ ] output `tasklist` và listing `%TEMP%\ArcTool_ExcelSync_*.pdf`
- supplied on: —
- forwarded to: —
