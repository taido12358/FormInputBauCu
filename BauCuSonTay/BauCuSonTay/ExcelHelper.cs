using System;
using System.Collections.Generic;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace BauCuSonTay
{
    public class SheetData
    {
        public string SheetName { get; set; }
        public List<RowData> Rows { get; set; } = new List<RowData>();
        public List<string> CandidateNames { get; set; } = new List<string>();
    }

    public class RowData
    {
        public int RowIndex { get; set; }
        public string KhuVuc { get; set; }
        public string ToDP { get; set; }
        public double TongSoKV { get; set; }
        public double TongSoCuTri { get; set; }
    }

    public static class ExcelHelper
    {
        public static string FilePath { get; set; }

        /// <summary>
        /// Doc toan bo file vao byte[], NPOI lam viec tren MemoryStream.
        /// Khong giu file handle -> tranh IOException "used by another process".
        /// </summary>
        private static HSSFWorkbook OpenWorkbook()
        {
            // FileShare.ReadWrite: doc duoc ke ca khi Excel dang mo file
            byte[] fileBytes;
            using (var fs = new FileStream(FilePath, FileMode.Open,
                                           FileAccess.Read, FileShare.ReadWrite))
            {
                fileBytes = new byte[fs.Length];
                int totalRead = 0;
                while (totalRead < fileBytes.Length)
                {
                    int read = fs.Read(fileBytes, totalRead, fileBytes.Length - totalRead);
                    if (read == 0) break;
                    totalRead += read;
                }
            }
            // NPOI doc tu MemoryStream, file lock da duoc giai phong
            using (var ms = new MemoryStream(fileBytes))
            {
                return new HSSFWorkbook(ms);
            }
        }

        public static List<string> GetSheetNames()
        {
            var names = new List<string>();
            var wb = OpenWorkbook();
            for (int i = 0; i < wb.NumberOfSheets; i++)
                names.Add(wb.GetSheetName(i));
            return names;
        }

        public static SheetData GetSheetData(string sheetName, int y = 5)
        {
            var data = new SheetData { SheetName = sheetName };
            var wb = OpenWorkbook();
            var sheet = wb.GetSheet(sheetName);
            if (sheet == null) return data;

            // Ten ung vien: hang 5 (index 5)
            // y=5 -> cot T(19) den het; y=3 -> cot Q(16) den U(20)
            var headerRow = sheet.GetRow(5);
            if (headerRow != null)
            {
                // y=3: ứng viên từ col 16 (sau PB1-3 ở 13-15)
                // y=5: ứng viên từ col 19 (sau PB1-5 ở 14-18)
                int colStart = (y == 3) ? 16 : 19;
                int colEnd   = headerRow.LastCellNum; // đọc hết đến cuối
                for (int c = colStart; c < colEnd; c++)
                {
                    string val = GetCellString(headerRow.GetCell(c));
                    if (!string.IsNullOrWhiteSpace(val))
                        data.CandidateNames.Add(val);
                }
            }

            // Du lieu: bat dau tu hang 7 (index 7)
            for (int r = 7; r <= sheet.LastRowNum; r++)
            {
                var row = sheet.GetRow(r);
                if (row == null) continue;
                string bVal = GetCellString(row.GetCell(1));
                if (string.IsNullOrWhiteSpace(bVal)) continue;
                if (bVal.StartsWith("TONG") || bVal.StartsWith("TỔNG") ||
                    bVal.StartsWith("Ghi") || bVal.StartsWith("-")) continue;

                data.Rows.Add(new RowData
                {
                    RowIndex    = r,
                    KhuVuc      = bVal,
                    ToDP        = GetCellString(row.GetCell(2)),
                    TongSoKV    = GetCellDouble(row.GetCell(3)),
                    TongSoCuTri = GetCellDouble(row.GetCell(4))
                });
            }

            return data;
        }

        public static void WriteRowData(string sheetName, int rowIndex, SaveData sd)
        {
            // Doc workbook vao memory (giai phong lock ngay)
            var wb = OpenWorkbook();

            var sheet = wb.GetSheet(sheetName);
            if (sheet == null) return;

            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);

            // Cột bắt đầu phụ thuộc vào y (số đại biểu được bầu)
            // y=5: cấu trúc cũ — TongSoCuTri bắt đầu từ col 4, PB1-5 ở col 14-18, ứng viên từ col 19
            // y=3: cấu trúc mới — TongSoCuTri bắt đầu từ col 3, PB1-3 ở col 13-15, ứng viên từ col 16
            int baseCol     = (sd.Y == 3) ? 3  : 4;   // Tổng số cử tri
            int pb1Col      = (sd.Y == 3) ? 13 : 14;  // Phiếu bầu 1 ĐB
            int candColStart= (sd.Y == 3) ? 16 : 19;  // Ứng viên bắt đầu

            SetCellNum(row, baseCol,     sd.TongSoCuTri);
            SetCellNum(row, baseCol + 1, sd.SoLuongCuTriThamGia);
            SetCellPct(row, baseCol + 2, sd.TyLeCuTri     / 100.0, wb);  // % cử tri
            SetCellNum(row, baseCol + 3, sd.SoPhieuPhatRa);
            SetCellNum(row, baseCol + 4, sd.SoPhieuThuVao);
            SetCellPct(row, baseCol + 5, sd.TyLeThuVao    / 100.0, wb);  // % thu vào
            SetCellNum(row, baseCol + 6, sd.SoPhieuHopLe);
            SetCellPct(row, baseCol + 7, sd.TyLeHopLe     / 100.0, wb);  // % hợp lệ
            SetCellNum(row, baseCol + 8, sd.SoPhieuKhongHopLe);
            SetCellPct(row, baseCol + 9, sd.TyLeKhongHopLe / 100.0, wb); // % không hợp lệ

            SetCell(row, pb1Col,     sd.PhieuBau1DB);
            SetCell(row, pb1Col + 1, sd.PhieuBau2DB);
            SetCell(row, pb1Col + 2, sd.PhieuBau3DB);
            // Chỉ ghi PB4, PB5 khi y=5
            if (sd.Y == 5)
            {
                SetCell(row, pb1Col + 3, sd.PhieuBau4DB);
                SetCell(row, pb1Col + 4, sd.PhieuBau5DB);
            }

            for (int i = 0; i < sd.CandidateVotes.Count; i++)
                SetCell(row, candColStart + i, sd.CandidateVotes[i]);

            // Sau khi ghi dữ liệu hàng, cập nhật hàng TỔNG CỘNG của sheet
            TinhTongHangKhuVuc(sheet, sd.Y);

            // Ghi ra file voi FileShare.ReadWrite
            using (var fs = new FileStream(FilePath, FileMode.Create,
                                           FileAccess.Write, FileShare.ReadWrite))
            {
                wb.Write(fs);
            }
        }

        // ── Tính tổng và ghi vào hàng TỔNG CỘNG KV BỎ PHIẾU ───────
        // Tìm hàng TỔNG CỘNG, cộng tất cả hàng data từ row 7 đến (tongRow-1)
        // Bỏ trống các cột %
        private static void TinhTongHangKhuVuc(ISheet sheet, int y)
        {
            // 1. Tìm hàng TỔNG CỘNG
            int tongRow = -1;
            for (int r = 0; r <= sheet.LastRowNum; r++)
            {
                var row = sheet.GetRow(r);
                if (row == null) continue;
                string bVal = GetCellStringFull(row.GetCell(1));
                if (bVal.Contains("TỔNG CỘNG") || bVal.Contains("TONG CONG"))
                { tongRow = r; break; }
            }
            if (tongRow < 0) return;

            // 2. Xác định các cột cần tính tổng và các cột % cần bỏ trống
            // y=5: baseCol=4, pctCols={6,9,11,13}, sumCols=D(3),E(4),F(5),H(7),I(8),K(10),M(12),O-S(14-18),T+(19..)
            // y=3: baseCol=3, pctCols={5,8,10,12}, sumCols=C(2),D(3),E(4),G(6),H(7),J(9),L(11),N-P(13-15),Q+(16..)
            HashSet<int> pctCols;
            int firstDataCol; // cột đầu tiên có dữ liệu số (bỏ STT, tên, tổ DP)

            if (y == 5)
            {
                pctCols     = new HashSet<int> { 6, 9, 11, 13 };
                firstDataCol = 3;   // bắt đầu từ col D(3)
            }
            else // y == 3
            {
                pctCols     = new HashSet<int> { 5, 8, 10, 12 };
                firstDataCol = 2;   // bắt đầu từ col C(2) — tổng số KV
            }

            // Xác định cột cuối dựa vào header row 5
            var headerRow5 = sheet.GetRow(5);
            int lastDataCol = 0;
            if (headerRow5 != null)
                lastDataCol = headerRow5.LastCellNum - 1;

            // 3. Tính tổng từng cột (data rows: 7 → tongRow-1)
            var tongRowObj = sheet.GetRow(tongRow) ?? sheet.CreateRow(tongRow);

            for (int c = firstDataCol; c <= lastDataCol; c++)
            {
                // Bỏ qua cột % — để trống
                if (pctCols.Contains(c))
                {
                    var pctCell = tongRowObj.GetCell(c);
                    if (pctCell != null)
                    {
                        if (pctCell.CellType == CellType.Formula)
                            pctCell.SetCellType(CellType.Blank);
                        else if (pctCell.CellType == CellType.Numeric)
                            pctCell.SetCellValue(0.0);
                    }
                    continue;
                }

                // Bỏ qua col C(2) khi y=5 (là "Thuộc tổ dân phố" — text)
                if (y == 5 && c == 2) continue;

                // Cộng tổng từ row 7 đến tongRow-1
                double sum = 0;
                for (int r = 7; r < tongRow; r++)
                {
                    var dataRow = sheet.GetRow(r);
                    if (dataRow == null) continue;
                    sum += SafeGetNum(dataRow, c);
                }

                // Ghi tổng vào hàng TỔNG CỘNG
                SetCell(tongRowObj, c, sum);
            }
        }

        private static void SetCell(IRow row, int col, double value)
        {
            var cell = row.GetCell(col) ?? row.CreateCell(col);
            // Xóa formula nếu có, chuyển sang Numeric để ghi đè giá trị thật
            if (cell.CellType == CellType.Formula)
                cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(value);
        }

        // Ghi ô % với format "0.00%" — giá trị truyền vào là tỷ lệ (0..1)
        // Ví dụ: 0.9875 → hiển thị "98.75%" trong Excel
        private static void SetCellPct(IRow row, int col, double ratio,
                                        NPOI.SS.UserModel.IWorkbook wb)
        {
            var cell = row.GetCell(col) ?? row.CreateCell(col);
            // Xóa formula nếu có
            if (cell.CellType == CellType.Formula)
                cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(ratio);

            var style = wb.CreateCellStyle();
            var fmt   = wb.CreateDataFormat();
            style.DataFormat = fmt.GetFormat("0.00%");
            cell.CellStyle   = style;
        }

        // Ghi ô số thông thường — xóa formula nếu có
        private static void SetCellNum(IRow row, int col, double value)
        {
            var cell = row.GetCell(col) ?? row.CreateCell(col);
            if (cell.CellType == CellType.Formula)
                cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(value);
        }

        private static string GetCellString(ICell cell)
        {
            if (cell == null) return "";
            switch (cell.CellType)
            {
                case CellType.String:  return cell.StringCellValue?.Trim() ?? "";
                case CellType.Numeric: return cell.NumericCellValue.ToString();
                case CellType.Formula:
                    try   { return cell.StringCellValue?.Trim() ?? ""; }
                    catch { return ""; }
                default: return "";
            }
        }

        private static double GetCellDouble(ICell cell)
        {
            if (cell == null) return 0;
            if (cell.CellType == CellType.Numeric) return cell.NumericCellValue;
            return 0;
        }

        // ── Xuất tổng hợp vào sheet TH ──────────────────────────────
        // Cột ĐV nguồn: D(3)=KV, E(4)=CuTri, F(5)=SoLuong,
        //               H(7)=PhatRa, I(8)=ThuVao, K(10)=HopLe, M(12)=KhongHopLe
        //               T(19)..AA(26) = Ứng viên
        // Cột TH đích:  C(2)=KV, D(3)=CuTri, E(4)=SoLuong, F(5)=%CuTri
        //               G(6)=PhatRa, H(7)=ThuVao, I(8)=%ThuVao
        //               J(9)=HopLe, K(10)=%HopLe, L(11)=KhongHopLe, M(12)=%KhongHopLe
        public static string XuatTongHop(string filePath)
        {
            HSSFWorkbook wb;
            try { wb = OpenWorkbookFrom(filePath); }
            catch (Exception ex) { return "Lỗi đọc file: " + ex.Message; }

            var shTH = wb.GetSheet("TH");
            if (shTH == null)
            {
                // Thử tìm sheet tên gần giống
                string sheetList = "";
                for (int i = 0; i < wb.NumberOfSheets; i++)
                    sheetList += wb.GetSheetName(i) + ", ";
                return $"Không tìm thấy sheet TH. Các sheet hiện có: {sheetList}";
            }

            // (dvSheet, thRow, thCandColStart, candCount)
            var dvMap = new (string, int, int, int)[]
            {
                ("ĐV1", 7,  18, 8),
                ("ĐV2", 8,  26, 8),
                ("ĐV3", 9,  34, 8),
                ("ĐV4", 10, 42, 8),
                ("ĐV5", 11, 50, 7),
            };

            var errors = new System.Text.StringBuilder();
            int processed = 0;

            // Tạo evaluator — chỉ dùng EvaluateInCell từng ô cụ thể
            // KHÔNG dùng EvaluateAll() vì file có thể có external reference gây lỗi
            var evaluator = new NPOI.HSSF.UserModel.HSSFFormulaEvaluator(wb);

            foreach (var (dvName, thRow, thCandStart, candCount) in dvMap)
            {
                var shDV = wb.GetSheet(dvName);
                if (shDV == null) { errors.AppendLine($"Không tìm thấy sheet {dvName}"); continue; }

                // Tìm hàng TỔNG CỘNG (cột B = index 1)
                int srcRow = -1;
                for (int r = 0; r <= shDV.LastRowNum; r++)
                {
                    var row = shDV.GetRow(r);
                    if (row == null) continue;
                    var cell = row.GetCell(1);
                    if (cell == null) continue;
                    string bVal = GetCellStringFull(cell);
                    if (bVal.Contains("TỔNG CỘNG") || bVal.Contains("TONG CONG"))
                    { srcRow = r; break; }
                }

                if (srcRow < 0)
                { errors.AppendLine($"{dvName}: không tìm thấy hàng TỔNG CỘNG"); continue; }

                var dvRow    = shDV.GetRow(srcRow);
                var thRowObj = shTH.GetRow(thRow) ?? shTH.CreateRow(thRow);

                // Reset từng ô formula → Numeric bằng cách evaluate tại chỗ
                // Bọc try-catch để bỏ qua external reference hoặc lỗi formula phức tạp
                int[] readCols = { 3, 4, 5, 7, 8, 10, 12, 19, 20, 21, 22, 23, 24, 25, 26 };
                foreach (int rc in readCols)
                {
                    var c = dvRow.GetCell(rc);
                    if (c == null || c.CellType != CellType.Formula) continue;
                    try
                    {
                        evaluator.EvaluateInCell(c);  // Formula → Numeric với giá trị thật
                    }
                    catch
                    {
                        // Nếu không evaluate được (external ref...), giữ nguyên cached value
                    }
                }

                // Đọc giá trị đã evaluate và ghi vào TH
                SetCell(thRowObj, 2,  SafeGetNum(dvRow, 3));   // D(3)  → TH C(2)
                SetCell(thRowObj, 3,  SafeGetNum(dvRow, 4));   // E(4)  → TH D(3)
                SetCell(thRowObj, 4,  SafeGetNum(dvRow, 5));   // F(5)  → TH E(4)
                SetCell(thRowObj, 6,  SafeGetNum(dvRow, 7));   // H(7)  → TH G(6)
                SetCell(thRowObj, 7,  SafeGetNum(dvRow, 8));   // I(8)  → TH H(7)
                SetCell(thRowObj, 9,  SafeGetNum(dvRow, 10));  // K(10) → TH J(9)
                SetCell(thRowObj, 11, SafeGetNum(dvRow, 12));  // M(12) → TH L(11)

                // Tính % từ giá trị đã đọc
                double cuTri     = SafeGetNum(dvRow, 4);
                double soLuong   = SafeGetNum(dvRow, 5);
                double phatRa    = SafeGetNum(dvRow, 7);
                double thuVao    = SafeGetNum(dvRow, 8);
                double hopLe     = SafeGetNum(dvRow, 10);
                double khongHopLe= SafeGetNum(dvRow, 12);

                SetCellPct(thRowObj, 5,  cuTri   > 0 ? Math.Round(soLuong    / cuTri  , 4) : 0, wb);
                SetCellPct(thRowObj, 8,  phatRa  > 0 ? Math.Round(thuVao     / phatRa , 4) : 0, wb);
                SetCellPct(thRowObj, 10, thuVao  > 0 ? Math.Round(hopLe      / thuVao , 4) : 0, wb);
                SetCellPct(thRowObj, 12, thuVao  > 0 ? Math.Round(khongHopLe / thuVao , 4) : 0, wb);

                // Ghi ứng viên T(19)..AA(26) → TH thCandStart..
                for (int i = 0; i < candCount; i++)
                    SetCell(thRowObj, thCandStart + i, SafeGetNum(dvRow, 19 + i));

                processed++;
            }

            // Hàng TỔNG CỘNG TH (row 12) — cộng rows 7..11
            var totalRow = shTH.GetRow(12) ?? shTH.CreateRow(12);
            int[] sumCols = { 2, 3, 4, 6, 7, 9, 11 };
            foreach (int tc in sumCols)
            {
                double sum = 0;
                for (int r = 7; r <= 11; r++)
                {
                    var rr = shTH.GetRow(r);
                    if (rr != null) sum += SafeGetNum(rr, tc);
                }
                SetCell(totalRow, tc, sum);
            }
            double ttCuTri   = SafeGetNum(totalRow, 3);
            double ttSoLuong = SafeGetNum(totalRow, 4);
            double ttPhatRa  = SafeGetNum(totalRow, 6);
            double ttThuVao  = SafeGetNum(totalRow, 7);
            double ttHopLe   = SafeGetNum(totalRow, 9);
            double ttKhongHL = SafeGetNum(totalRow, 11);
            SetCellPct(totalRow, 5,  ttCuTri  > 0 ? Math.Round(ttSoLuong / ttCuTri , 4) : 0, wb);
            SetCellPct(totalRow, 8,  ttPhatRa > 0 ? Math.Round(ttThuVao  / ttPhatRa, 4) : 0, wb);
            SetCellPct(totalRow, 10, ttThuVao > 0 ? Math.Round(ttHopLe   / ttThuVao, 4) : 0, wb);
            SetCellPct(totalRow, 12, ttThuVao > 0 ? Math.Round(ttKhongHL / ttThuVao, 4) : 0, wb);

            // Ghi file
            using (var fs = new System.IO.FileStream(filePath, System.IO.FileMode.Create,
                                                      System.IO.FileAccess.Write,
                                                      System.IO.FileShare.ReadWrite))
            { wb.Write(fs); }

            string msg = $"OK:{processed}";
            if (errors.Length > 0) msg += "|WARN:" + errors.ToString();
            return msg;
        }

        // Đọc cell bất kỳ type → double (xử lý cả Numeric, Formula, String)
        private static double SafeGetNum(IRow row, int col)
        {
            if (row == null) return 0;
            var cell = row.GetCell(col);
            if (cell == null) return 0;
            switch (cell.CellType)
            {
                case CellType.Numeric: return cell.NumericCellValue;
                case CellType.Formula:
                    try { return cell.NumericCellValue; } catch { return 0; }
                case CellType.String:
                    return double.TryParse(cell.StringCellValue,
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out double v) ? v : 0;
                default: return 0;
            }
        }

        // Đọc cell string, xử lý multiline (\n trong Excel)
        private static string GetCellStringFull(ICell cell)
        {
            if (cell == null) return "";
            string raw = "";
            switch (cell.CellType)
            {
                case CellType.String:  raw = cell.StringCellValue ?? ""; break;
                case CellType.Formula:
                    try { raw = cell.StringCellValue ?? ""; } catch { raw = ""; } break;
                default: return "";
            }
            // Chuẩn hóa: xóa newline, uppercase để so sánh
            return raw.Replace("\n", " ").Replace("\r", " ").ToUpper();

        }

        // OpenWorkbook từ path bất kỳ (không phải FilePath static)
        private static HSSFWorkbook OpenWorkbookFrom(string path)
        {
            byte[] bytes;
            using (var fs = new System.IO.FileStream(path, System.IO.FileMode.Open,
                                                      System.IO.FileAccess.Read,
                                                      System.IO.FileShare.ReadWrite))
            {
                bytes = new byte[fs.Length];
                int total = 0;
                while (total < bytes.Length)
                {
                    int n = fs.Read(bytes, total, bytes.Length - total);
                    if (n == 0) break;
                    total += n;
                }
            }
            using (var ms = new System.IO.MemoryStream(bytes))
                return new HSSFWorkbook(ms);
        }
    }

    public class SaveData
    {
        public double TongSoCuTri         { get; set; }
        public double SoLuongCuTriThamGia { get; set; }
        public double TyLeCuTri           { get; set; }
        public double SoPhieuPhatRa       { get; set; }
        public double SoPhieuThuVao       { get; set; }
        public double TyLeThuVao          { get; set; }
        public double SoPhieuHopLe        { get; set; }
        public double TyLeHopLe           { get; set; }
        public double SoPhieuKhongHopLe   { get; set; }
        public double TyLeKhongHopLe      { get; set; }
        public double PhieuBau1DB         { get; set; }
        public double PhieuBau2DB         { get; set; }
        public double PhieuBau3DB         { get; set; }
        public double PhieuBau4DB         { get; set; }
        public double PhieuBau5DB         { get; set; }
        public List<double> CandidateVotes { get; set; } = new List<double>();
        public int Y { get; set; } = 5;  // so dai bieu duoc bau (3 hoac 5)
    }
}
