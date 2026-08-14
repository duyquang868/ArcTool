---
name: project_qd_wall_spike_handoff
description: Session handoff cho Quick Dimension Wall Spike — trạng thái sau nhiều lượt smoke, tồn đọng cần xử lý ở phiên mới.
type: project
---

Wall Spike Phase 1.2 đã được reset và rewrite hoàn toàn theo yêu cầu người dùng "test từng logic trước, tránh càng sửa càng sai" (2026-07-14). Đây là bản ghi trạng thái sau nhiều lượt smoke test, để phiên mới bắt đúng vị trí đang dang dở.

**Command/service/model liên quan:**
- `ArcTool.Core/Commands/QuickDimensionWallReferenceSpikeCommand.cs`
- `ArcTool.Core/Services/QuickDimensionWallReferenceProbeService.cs`
- `ArcTool.Core/Services/QuickDimensionWallSpikeXmlLogService.cs`
- `ArcTool.Core/Models/QuickDimensionWallReferenceProbe.cs`
- Ribbon binding trong `App.cs` giữ nguyên tên class + assembly path.

**Hiện trạng logic Wall Spike (bản 2026-07-17 — SIDE-FACE BOUNDARY + DIRECTIONAL FULL-HEIGHT RESOLVER, 100% SMOKE PASS):**
- Input: pick 1 straight non-curtain host `Wall` + pick 1 điểm side Left/Right; không tạo dimension.
- Map side pick → shell layer: `Wall.Orientation` (planar) dot `targetSideNormal` > 0 → `ShellLayerType.Exterior`, ngược lại `Interior`.
- `TryCollectSideFaceBoundary(selectedWall, ..., includeBothShells:false)` lấy side planes qua `HostObjectUtils.GetSideFaces(wall, shellLayer)` và resolve bằng `wall.GetGeometryObjectFromReference`.
- Duyệt `wall.get_Geometry(Options{ComputeReferences=true, IncludeNonVisibleObjects=true})`, giữ boundary candidate thuộc side planes: vertical edge midpoint (`source=vertical-edge`, có `Edge.Reference`) + endpoint của horizontal side-face edge song song wall axis (`source=horizontal-endpoint`, không có vertical ref). `EdgeTouchesSideFace` kiểm tra `edge.GetFace(0/1)` bằng face identity, fallback bằng midpoint-to-plane ≤ 1mm + normal match.
- Chọn `mainRun` = horizontal side-face segment song song wall axis có `Length` lớn nhất. Base Start/Finish = hai endpoint của `mainRun` (qua `SelectBestCandidateAtPoint`, ưu tiên candidate có reference nếu cùng station). Lý do: L-joint có vertical tip/solid overshoot vượt quá góc dimension, nên không được lấy min/max station mù quáng.
- Tự động lấy joined walls bằng `LocationCurve.get_ElementsAtJoin(0/1)` + `JoinGeometryUtils.GetJoinedElements` fallback. `CollectJoinedWallBoundaryCandidates` gom boundary candidates từ cả Exterior/Interior side faces của joined walls, rồi lọc điểm nằm trên cùng side-line của selected wall (`DistanceToSideLine <= 5mm`). `ResolveEndCandidate` dùng directional full-height resolver: Interior chọn full-height vertical `Reference` gần nhất theo hướng inward vào span từ selected+joined candidates; Exterior chọn joined full-height vertical `Reference` theo hướng outward nếu có, nếu không giữ base.
- Critical bug fix đã khóa: full-height threshold phải tính trên `referenceCandidates` (`Reference != null`, tức vertical-edge thật), KHÔNG tính trên horizontal endpoints ở z=8000; nếu không Interior sẽ rơi về raw main-run endpoint t=0/length.
- Diagnostics message mới: `Side-face boundary model (vertical + horizontal + joined-wall outward candidates)`, kèm số selected candidates, joined walls, joined candidates on side line, source của Start/Finish, và mapped refs.
- 2026-07-17 smoke PASS 100% trên bộ wall 379467, 379468, 379469, 379470, 379933, 380187 Left/Right (12 XML logs). Đây là spike model đã được chứng minh cho L-joint/T-joint wall-end anchors; chưa port sang production collector.

**Trạng thái smoke test tính đến 2026-07-14:**
- Left side đã trả về đúng tọa độ so với khảo sát thủ công.
- Right side smoke gần nhất (trước bước top-face footprint) vẫn lệch: dùng largest side face vertical-edge min/max cho Start ≈ (6935.83, 3443.5), Finish ≈ (12572.94, 6965.95) — vượt qua gốc giao thật khoảng vài trăm mm vì solid tường ôm phần join extension.
- Sau bước switch sang top-face footprint corners + map vertical edge (áp dụng cuối phiên), smoke báo `Wall footprint corner could not be mapped to a vertical Edge.Reference in wall geometry` vì plan-visible join corner không còn là vertical edge của selected wall.
- Lịch sử bị supersede: top-face footprint/intersection/near-far từng pass L-joint nhưng T-joint smoke bác bỏ. `JoinGeometryUtils.AreElementsJoined`/`IsCuttingElementInJoin` trả role=unknown cho WALL END-JOIN ở góc; rule EXTERIOR→far / INTERIOR→near chỉ là heuristic đúng một số L-joint, sai T-joint.
- Người dùng cung cấp node Dynamo/Python `Wall Edges References` của Genius Loci: `HostObjectUtils.GetSideFaces(Exterior/Interior)` + `edge.GetFace(0/1)` lọc vertical edges theo side face, trả đúng các cặp góc xanh/đỏ của T-joint. Đây là bằng chứng chuyển model sang SIDE-FACE VERTICAL EDGE.
- Code hiện tại đã gỡ toàn bộ helper top-face/intersection/near-far/join-role khỏi `QuickDimensionWallReferenceProbeService.cs`. T-joint smoke PASS với side-face VERTICAL edges, nhưng L-joint smoke FAIL với vertical-only.
- 2026-07-15 (L-joint): node `Wall Edges References` chứng minh vertical-only chưa đủ: một số góc L-joint đúng nằm ở endpoint của horizontal side-face edge; một góc exterior đầu bị cắt (ví dụ 10341.836,18569.557) KHÔNG nằm trong selected-wall edge list mà thuộc joined wall. Vì vậy bản cuối phiên thử mô hình: selected-wall main horizontal side-run endpoints + joined-wall outward candidates trên cùng side-line.
- Bằng chứng số liệu wall 379468: Right/Interior Start đúng = selected vertical (10307.842,18312.479); Right/Interior Finish đúng = selected horizontal endpoint (14301.973,14019.466), không phải selected vertical tip (14452.252,13857.942). Left/Exterior Finish đúng = selected vertical (14662.117,13925.988). Left/Exterior Start đúng = joined-wall outer point (10341.836,18569.557), không có trong selected-wall edge data. Bản thử nghiệm mới được viết để bắt 4 tình huống này.
- `QuickDimensionWallReferenceProbeService.cs` bản cuối phiên ~681 dòng, brace balance OK, grep xác nhận có `WallSpikeBoundaryCandidate`, `WallSpikeHorizontalSegmentCandidate`, `CollectJoinedWalls`, `CollectJoinedWallBoundaryCandidates`, `ResolveEndCandidate`, `SelectBestCandidateAtPoint`; không còn `PLACEHOLDER`. Chưa build/smoke trong Revit sau rewrite cuối.

**Smoke test 2026-07-16 (bản joined-wall-extension) + XML log utility mới:**
- Kết quả smoke: TOÀN BỘ tọa độ phía Left ĐÚNG; TOÀN BỘ tọa độ phía Right SAI (xanh=đúng, đỏ=sai trên bản khảo sát). T-joint cũng bị cùng triệu chứng: Left đúng, Right sai. Đây là bằng chứng lỗi có tính hệ thống theo SIDE, không phải theo loại joint.
- Thêm tiện ích XML log cho nút Wall Spike: `QuickDimensionWallSpikeXmlLogService.WriteWallSpikeLog(doc, wall, sidePickPoint, result)`, gọi trong `TryWriteXmlLog` của command (bọc try/catch — log fail KHÔNG làm hỏng smoke summary). File ghi cạnh `.rvt`, tên `ArcTool_QD_WallSpike_{wallId}_{yyyyMMdd_HHmmss}.xml`.
- Nội dung log: ProbeResult (side, shell, message, Start/Finish anchor), SelectedWall (id, type, LocationCurve start/end, toàn bộ boundary corner candidates), JoinedWalls (mỗi joined wall: id, type, toàn bộ corner candidates). Corner = ĐÚNG boundary candidate mà logic spike đang dùng (vertical-edge XY midpoint + horizontal-endpoint) qua `CollectBoundaryCornerPointsForLog` (dùng lại `TryCollectSideFaceBoundary`), joined walls dùng lại `CollectJoinedWalls` — nên log phản ánh chính xác logic hiện tại, phục vụ rà soát.
- Tọa độ trong log: Survey N/E qua `CoordinateConversionService.ToSharedMm` (`ProjectLocation.GetProjectPosition`). Ghi cả mét (`n`,`e`,`elevation` — khớp nhãn ảnh smoke) lẫn mm (`n_mm`,`e_mm`,`elevation_mm`). Anchor có thêm `stationMm` = station dọc wall axis.
- CHƯA build/smoke XML log trong Revit (dotnet không có trong Linux workspace). Việc đầu tiên phiên sau: build Windows/Revit, chạy Wall Spike, mở XML log, đối chiếu corner candidates của cả selected + joined walls để tìm vì sao Right luôn sai.

**ROOT CAUSE xác định 2026-07-16 từ 2 XML log wall 379467 (Left=Exterior, Right=Interior):**
- Trên tường này `side→shell`: Left→Exterior, Right→Interior. Nên "Left đúng/Right sai" = "Exterior đúng/Interior sai". Lỗi theo SHELL, KHÔNG theo joint type.
- Bằng chứng message: Left `joined candidates on side line: 0` (extension KHÔNG chạy → anchor = vertical-edge của chính tường → đúng). Right `joined candidates on side line: 4` (extension CHẠY và ghi đè cả 2 anchor → sai).
- Overshoot đo được (tường dày 200mm): Right Start đúng phải là góc trong tường 379467 st≈78 (15059.098/5101.342) nhưng code trả st −128.076 (14949.852/4926.513) = đúng góc JoinedWall 379470 idx18, lố ~206mm. Right Finish đúng phải st≈6217 (18312.479/10307.842) nhưng code trả st 6421.168 (18420.423/10480.588) = đúng góc JoinedWall 379468 idx17, lố ~204mm.
- Hình học: ở góc lồi, Exterior là cạnh NGOÀI của miter (góc tường được chọn đã là điểm ngoài cùng → không candidate joint nào lọt → đúng). Interior là cạnh TRONG (góc tường joint thò ra ngoài góc trong tường được chọn). `ResolveEndCandidate` chọn candidate station xa nhất trong `JoinExtensionMargin=500mm` → trên Interior giả định "góc ngoài thuộc joint & xa hơn" bị NGƯỢC DẤU → overshoot đúng 1 bề dày.
- Kết luận: bước joined-wall outward extension chỉ đúng cho shell tạo cạnh NGOÀI của góc; áp lên shell trong luôn kéo lố qua đầu solid tường thật. Logic không phân biệt shell nào là "ngoài" nên phía nào trigger extension là phía đó sai.

**Hướng sửa đề xuất (chưa code, chờ user xác nhận):**
1. Clamp: anchor sau resolve KHÔNG được vượt min/max station của chính side-face tường được chọn (shell đó). Extension chỉ điền góc thiếu, không đẩy xa hơn đầu solid tường được chọn.
2. Chỉ mở rộng khi base endpoint tường được chọn CHƯA phải vertical-edge có Reference. Nếu tường đã có góc riêng (đúng trường hợp Interior) thì từ chối candidate joint.
3. (Dài hạn) Thay gom vertical-edge joint theo margin lỏng bằng GIAO mặt bên đã chọn của tường được chọn với mặt tường joint (đường miter), suy anchor từ side-plane tường được chọn → bỏ được cả `JoinExtensionMargin` lẫn `DistanceToSideLine`.
- Claude nghiêng về làm 1+2 trước (bước nhanh), rồi cân nhắc 3. Chưa đụng probe logic.

**FIX 1+2 đã áp dụng 2026-07-16 vào `ResolveEndCandidate` (chưa build/smoke Windows):**
- Thêm hằng `JoinCoincidenceTolerance = 1mm`.
- Guard chính (option 2, mạnh nhất): nếu `baseCandidate.Reference != null` → return baseCandidate ngay, KHÔNG extend. Lý do: base có Edge.Reference nghĩa là chính selected wall đã có góc vertical-edge của nó tại đầu đó (đúng trường hợp Interior của log 379467 khi mainRun endpoint rơi vào vertical-edge z=4000). Extension chỉ chạy khi base là horizontal/main-run endpoint không có reference (đầu bị cắt, góc thật thuộc joined wall — trường hợp Exterior của wall 379468 trước đây).
- Guard phụ (option 1 dạng mềm): khi vẫn extend, bỏ qua joined candidate nào trùng station với candidate mà selected wall đã sở hữu (`SelectedWallAlreadyOwnsStation`, tol 1mm) — đó là geometry join-cleanup của chính selected wall, không phải góc thiếu; extend tới nó sẽ overshoot.
- `ResolveEndCandidate` đổi chữ ký: thêm tham số `selectedCandidates`; hai call site trong `RunWallReferenceProbe` đã cập nhật truyền `selectedCandidates`.
- Smoke mới sau fix 1+2: wall 379467 (selected wall cắt tường khác) Left/Right đều PASS; wall 379470 (selected wall bị tường khác cắt) Left/Right đều FAIL; wall 379933 T-joint Left PASS, Right FAIL. 6 XML logs chứng minh rule cần theo HƯỚNG đầu tường, không chỉ theo shell.
- Final fix hiện tại trong `ResolveEndCandidate`: nhận `shellLayer`; dùng `SelectDirectionalFullHeightReference` + `IsCandidateInRequestedDirection`. Với `Interior`, chọn full-height vertical-edge có Reference gần nhất theo hướng INWARD vào span (Start: station > base; Finish: station < base), lấy từ selected + joined candidates. Với `Exterior`, chọn full-height vertical-edge có Reference theo hướng OUTWARD ngoài span từ joined candidates; nếu không có thì giữ base. `FullHeightReferenceTolerance=10mm`, `JoinExtensionMargin=500mm`.
- Mô phỏng theo logs: wall 379467 giữ PASS; wall 379470 kỳ vọng sửa thành Start/Finish đúng cho cả hai side (Interior dùng inward st=100/3926; Exterior dùng outward from joined side-line); wall 379933 Right kỳ vọng sửa từ raw st=0/4992 sang inward selected st=113/4879. Vẫn cần build/smoke Windows/Revit vì dotnet không có trong Linux workspace.

**BUG NGƯỠNG FULL-HEIGHT — sửa 2026-07-16 (bản mới nhất):**
- Triệu chứng smoke 6 log 16:01-16:08: 379467 Left/Right PASS; 379470 (bị cắt) Left/Right FAIL; 379933 (T-joint) Left PASS/Right FAIL.
- Root cause: trong `SelectDirectionalFullHeightReference`, `maxMidpointZ` tính trên TẤT CẢ candidate gồm cả horizontal-endpoint ở z=8000 → threshold ~7990 → vertical-edge (midpoint z=4000, mới là edge mang Reference) luôn bị loại → nhánh Interior luôn rơi về base = raw main-run endpoint (t=0/length) → sai.
- Fix: tính `referenceCandidates` = candidate có `Reference != null` (chính là vertical edges) TRƯỚC; `maxMidpointZ` chỉ tính trên tập này; directional filter cũng chạy trên tập này. Guard rỗng trả null → giữ base.
- Mô phỏng lại các case Interior joined_on_side_line=0 (tái hiện 100% từ selected candidates): 379470 Right → Start st=100 (11418.969/6281.605), Finish st=3926.69 (15059.098/5101.342); 379933 Right → Start st=113.029 (17034/8261.847), Finish st=4879.436 (12714.169/10276.218). Hai giá trị finish/start này khớp các ô đỏ trong ảnh smoke (điểm ĐÚNG mong đợi).
- LƯU Ý mô phỏng: XML log dump TOÀN BỘ joined candidates (không lọc side-line), còn code thật chỉ dùng tập đã lọc `DistanceToSideLine<=5mm` = số "joined candidates on side line" trong message. Vì vậy mô phỏng Exterior/Interior có joined>0 (379467 Right=4, 379470 Left=4) chỉ đúng xu hướng, KHÔNG khớp 1-1; phải smoke Windows/Revit để chốt.
- Cần smoke lại toàn bộ: 379467 L/R (không regress), 379470 L/R (kỳ vọng chuyển sang PASS), 379933 L/R (Right kỳ vọng PASS). File-tool đọc file đầy đủ 804 dòng kết thúc đúng; bash mount trong phiên này bị stale/truncate, KHÔNG dùng bash xác thực file service.

**Closure 2026-07-17 — Wall Spike isolated smoke PASS 100%:**
- User smoke result: PASS 100% on walls 379467, 379468, 379469, 379470, 379933, 380187, both Left/Right. 12 XML logs confirm `Succeeded=true` and expected survey N/E coordinates.
- Representative passed anchors: 379467 L=(-78.079/6382.53), R=(78.079/6217.47); 379470 L=(-100/4082.848), R=(100/3926.69); 379933 L=(88.473/4904.119), R=(113.029/4879.436); 380187 L=(70.067/4284.344), R=(142.721/4386.249).
- Final resolver model: base anchors come from longest selected-wall side-run endpoints. `Interior` resolves to nearest full-height vertical `Reference` in the inward direction into the selected wall span from selected+joined candidates. `Exterior` resolves to outward joined full-height vertical `Reference` if one exists on the side line; otherwise keeps selected-wall base.
- Final implementation detail that must not regress: `SelectDirectionalFullHeightReference` first builds `referenceCandidates = candidates.Where(c => c.Reference != null && c.Point != null)` and computes `maxMidpointZ` only from that set. Horizontal endpoints at top elevation must never define the full-height threshold.
- XML logger (`QuickDimensionWallSpikeXmlLogService`) proved essential and should remain until production collector is ported and smoked.

**Closure 2026-07-20 — refined mid-run classifier re-smoke PASS; production port allowed:**
- Four real Wall Spike re-smoke sets passed after the accepted-mid-run classifier fix: selected wall 380815 accepted true mid-run wall 381185 only on Right/Interior; 379467 accepted 379933 only on Right/Interior; 379469 accepted 379933 only on Right/Interior; 379470 accepted 380187 only on Right/Interior.
- Opposite/clean shells reported `mid-run crossings: 0`; proximity-only candidates stayed ignored; endpoint join artifacts stayed `EndJoinOnly` with `acceptedMidRunStationCount=0` even when raw `referenceHitCount > 0`.
- This cleared Wall Spike logic for production collector + read-only aggregator port. It did NOT clear Phase 3 by itself.

**Closure 2026-07-22 — production read-only re-smoke PASS; NewDimension still gated:**
- Production `QuickDimensionReadOnlySummaryCommand` passed the four real sets after BUG-09: 380815, 379467, 379469, 379470, both shells. `FinalCandidate.elementId` now matches the stable-reference owner and `hostElementId` preserves the selected wall.
- Classifier stayed PASS: true mid-run walls accepted only on Interior/Right with `acceptedMidRunStationCount=2`; opposite shells stayed clean; end joins stayed `EndJoinOnly`; visible survey labels matched XML coordinates.
- Remaining gates before Phase 3 `NewDimension`: construct dimension line from resolved final candidate span instead of raw `0..axisLength`; fix read-only XML `includeGrids` metadata mismatch; verify close-opening FamilyInstance Left/Right reference semantics from test 379470 before trusting generated `ReferenceArray` output.

**Tồn đọng cần xử lý ở phiên mới (ưu tiên theo thứ tự):**
1. Inspect and fix `NewDimension` prerequisites only; do not rewrite the proven collector/classifier.
2. Ensure future dimension-line construction derives from final ordered candidates/resolved anchors and covers outside-span exterior anchors.
3. Fix `QuickDimensionReadOnlyXmlLogService` options serialization so Grid disabled state is not logged as `includeGrids="true"` in wall-axis mode.
4. Investigate test 379470 close-opening interleave: Window 379479, Door 379482, Window 379478 produce plausible station spacing but owner/side labels interleave; verify with live references before creating production dimensions.
5. Then implement the smallest Phase 3 `NewDimension` smoke path using the already ordered final candidates and live `Reference` objects; smoke 379470 first, then 380815/379467/379469.

**Ràng buộc không được phá:**
- Product behavior is manual and reviewable: one operator-selected straight wall at a time, producing one dimension chain on that wall axis. Never introduce bulk/automatic multi-wall dimension creation; a high-volume batch prevents reliable human validation of every generated dimension.
- The one-wall chain may contain a mixed sequence of L-joint and T-joint stations. Per-joint left/right correctness is a prerequisite, not evidence that aggregation/ordering/deduplication is correct.
- Do not merge Wall + Door/Window + Grid into the Read-Only Summary command before each spike isolated pass.
- Wall Spike phải giữ format `Placement side` + `Selected side face` + `Vertical edges on side face` + Start/Finish anchor mm để so sánh với các smoke report cũ.
- Không tự thay đổi vai trò: Claude là Chief Architect, Gemma 4 chỉ generate code qua MCP; mọi output Gemma phải review + reject nếu sai contract trước khi apply.
- Đơn vị hiển thị trong dialog phải là millimeters (`UnitUtils.ConvertFromInternalUnits(..., UnitTypeId.Millimeters)`); internal math giữ Revit internal feet.
- `dotnet` không có trong shell của cả hai phiên (Linux workspace); mọi build/smoke phải chạy trong Windows/Revit developer environment.
