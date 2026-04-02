using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace BauCuSonTay
{
    public partial class MainForm : Form
    {
        private SheetData _currentSheet;
        private RowData   _selectedRow;
        private bool      _updating = false;

        public MainForm() { InitializeComponent(); }

        // ══════════════════════════════════════════════════════════
        // FILE
        // ══════════════════════════════════════════════════════════
        private void BtnOpenFile_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog {
                Title  = "Chọn file Excel bầu cử",
                Filter = "Excel 97-2003 (*.xls)|*.xls|Tất cả file|*.*"
            })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;
                ExcelHelper.FilePath = dlg.FileName;
                lblFileInfo.Text      = dlg.FileName;
                lblFileInfo.ForeColor = Color.FromArgb(30,130,76);
                lblFileInfo.Font      = new Font("Segoe UI", 9f);

                cboSheet.Items.Clear();
                cboKhuVuc.Items.Clear();
                foreach (var s in ExcelHelper.GetSheetNames()) cboSheet.Items.Add(s);
                cboSheet.Enabled = true;
                if (cboSheet.Items.Count > 0) cboSheet.SelectedIndex = 0;
            }
        }

        // ══════════════════════════════════════════════════════════
        // SHEET CHANGED
        // ══════════════════════════════════════════════════════════
        private void CboSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboSheet.SelectedItem == null) return;
            string selected = cboSheet.SelectedItem.ToString();

            if (selected == "TH")
            {
                SetModeTH(true);
                return;
            }

            SetModeTH(false);

            _updating = true;
            try
            {
                _currentSheet = ExcelHelper.GetSheetData(selected, GetY());
                cboKhuVuc.Items.Clear();
                foreach (var row in _currentSheet.Rows)
                    cboKhuVuc.Items.Add(row.KhuVuc);
                cboKhuVuc.Enabled = true;
                RebuildCandidates(_currentSheet.CandidateNames);
                if (cboKhuVuc.Items.Count > 0) cboKhuVuc.SelectedIndex = 0;
            }
            finally { _updating = false; }
        }

        // Ẩn/hiện toàn bộ form nhập liệu, chỉ để lại nút xuất khi ở mode TH
        private void SetModeTH(bool isTH)
        {
            pnlSec1.Visible        = !isTH;
            pnlSec2.Visible        = !isTH;
            pnlSec3.Visible        = !isTH;
            pnlSec4.Visible        = !isTH;
            pnlSec5.Visible        = !isTH;
            pnlCandidates.Visible  = !isTH;
            pnlErrors.Visible      = !isTH;
            btnSave.Visible        = !isTH;
            btnClear.Visible       = !isTH;
            pnlModeTH.Visible      = isTH;

            if (isTH)
            {
                pnlMain.Height = 420;
            }
            else
            {
                RefreshLayout();
            }
        }

        // ══════════════════════════════════════════════════════════
        // KHU VỰC CHANGED
        // ══════════════════════════════════════════════════════════
        private void CboKhuVuc_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_currentSheet == null || cboKhuVuc.SelectedIndex < 0) return;
            _selectedRow = _currentSheet.Rows[cboKhuVuc.SelectedIndex];
            txtToDP.Text        = _selectedRow.ToDP;
            txtTongSoKV.Text    = _selectedRow.TongSoKV.ToString("0");
            txtTongSoCuTri.Text = _selectedRow.TongSoCuTri > 0
                                    ? _selectedRow.TongSoCuTri.ToString("0") : "";
            ApplySoCuTriBau();
            RecalcAndValidate();
        }

        // ══════════════════════════════════════════════════════════
        // SỐ CỬ TRI BẦU (y = 3 hoặc 5)
        // ══════════════════════════════════════════════════════════
        private void CboSoCuTriBau_Changed(object sender, EventArgs e)
        {
            if (_updating) return;
            ApplySoCuTriBau();
            // Đọc lại tên ứng viên theo y mới
            if (_currentSheet != null)
            {
                var freshData = ExcelHelper.GetSheetData(_currentSheet.SheetName, GetY());
                RebuildCandidates(freshData.CandidateNames);
            }
            RecalcAndValidate();
        }

        // Cập nhật label "Có x cử tri bầu lấy y" và ẩn/hiện ô 4-5
        private void ApplySoCuTriBau()
        {
            int y = GetY();
            double x = D(txtTongSoCuTri.Text);
            string xStr = x > 0 ? x.ToString("0") : "?";
            lblSoCuTriBau.Text = $"Có {xStr} cử tri bầu lấy:";

            bool show45 = (y == 5);
            txtPhieuBau4.Visible      = show45;
            txtPhieuBau5.Visible      = show45;
            lblPhieuBau4Label.Visible = show45;
            lblPhieuBau5Label.Visible = show45;

            if (!show45)
            {
                txtPhieuBau4.Text = "0";
                txtPhieuBau5.Text = "0";
            }
        }

        private int GetY() =>
            cboSoCuTriBau.SelectedItem?.ToString() == "3" ? 3 : 5;

        // ══════════════════════════════════════════════════════════
        // DYNAMIC CANDIDATES
        // ══════════════════════════════════════════════════════════
        private void RebuildCandidates(List<string> names)
        {
            foreach (var l in lblCandidates)    pnlCandidates.Controls.Remove(l);
            foreach (var t in txtCandidateVotes) {
                t.TextChanged -= InputChanged;
                pnlCandidates.Controls.Remove(t);
            }
            foreach (var l in errCandidates) pnlCandidates.Controls.Remove(l);
            lblCandidates.Clear();
            txtCandidateVotes.Clear();
            errCandidates.Clear();

            if (names.Count == 0) {
                lblCandidatesHeader.Text = "Không có ứng viên trong sheet này.";
                pnlCandidates.Height = 70;
                RefreshLayout();
                return;
            }

            lblCandidatesHeader.Text      = $"Danh sách {names.Count} ứng viên — nhập số phiếu bầu cho từng người:";
            lblCandidatesHeader.ForeColor = Color.FromArgb(52, 73, 94);
            lblCandidatesHeader.Font      = new Font("Segoe UI", 9.5f);

            // ── Danh sách dọc: mỗi ứng viên 1 dòng [Tên] [Input] + [Error] ─
            const int CV_PAD    = 14;    // padding trái
            const int CV_START  = 46;    // margin-top dòng đầu tiên
            const int CV_INP_Y  = 8;     // offset y của input trong row
            const int CV_LBL_W  = 420;   // width label tên ứng viên (rộng hơn)
            const int CV_INP_W  = 140;   // width ô input số phiếu
            const int CV_ERR_Y  = 36;    // offset y của error label trong row
            const int CV_ROW_H  = 58;    // tổng chiều cao 1 row: label+input(26)+gap+error(16)+margin(10)

            for (int i = 0; i < names.Count; i++)
            {
                int rowY = CV_START + i * CV_ROW_H;

                string name = names[i];
                if (name.StartsWith("Số phiếu bầu cho "))
                    name = name.Substring("Số phiếu bầu cho ".Length);

                // Label tên ứng viên — canh giữa theo chiều cao input
                var lbl = new Label {
                    Text      = name + ":",
                    Location  = new Point(CV_PAD, rowY + CV_INP_Y + 4),
                    Size      = new Size(CV_LBL_W, 22),
                    ForeColor = Color.FromArgb(40, 60, 90),
                    Font      = new Font("Segoe UI", 9.5f),
                    TextAlign = ContentAlignment.MiddleLeft
                };
                pnlCandidates.Controls.Add(lbl);
                lblCandidates.Add(lbl);

                // Input số phiếu
                var txt = new TextBox {
                    Location    = new Point(CV_PAD + CV_LBL_W + 8, rowY + CV_INP_Y),
                    Width       = CV_INP_W,
                    Font        = new Font("Segoe UI", 10.5f),
                    BorderStyle = BorderStyle.FixedSingle,
                    TextAlign   = HorizontalAlignment.Center,
                    Tag         = names[i]
                };
                txt.Enter += (s, ev) => ((TextBox)s).BackColor = Color.FromArgb(235, 245, 255);
                txt.Leave += (s, ev) => ((TextBox)s).BackColor = Color.White;
                txt.TextChanged += InputChanged;
                pnlCandidates.Controls.Add(txt);
                txtCandidateVotes.Add(txt);

                // Error label — nằm dưới input, đủ khoảng cách
                var errLbl = new Label {
                    Text      = "",
                    ForeColor = Color.Crimson,
                    Font      = new Font("Segoe UI", 8.5f, FontStyle.Italic),
                    Location  = new Point(CV_PAD, rowY + CV_ERR_Y),
                    Size      = new Size(CV_LBL_W + CV_INP_W + 8, 16),
                    Visible   = false
                };
                pnlCandidates.Controls.Add(errLbl);
                errCandidates.Add(errLbl);
            }

            // Height = start + n×rowH + bottom padding
            pnlCandidates.Height = CV_START + names.Count * CV_ROW_H + 20;
            RefreshLayout();
        }

        // Đẩy pnlErrors và buttons xuống sau pnlCandidates
        private void RefreshLayout()
        {
            pnlMain.SuspendLayout();

            int newY = pnlCandidates.Location.Y + pnlCandidates.Height + 8;

            pnlErrors.Location = new Point(10, newY);
            newY += pnlErrors.Height + 12;

            btnSave.Location  = new Point(10,  newY);
            btnClear.Location = new Point(212, newY);

            int totalH = newY + 70;
            pnlMain.Height = totalH;

            pnlMain.ResumeLayout(true);
            pnlMain.Invalidate();
            pnlScroll.Invalidate();
        }

        // ══════════════════════════════════════════════════════════
        // INPUT CHANGED → recalc %
        // ══════════════════════════════════════════════════════════
        private void InputChanged(object sender, EventArgs e)
        {
            if (_updating) return;
            RecalcAndValidate();
        }

        // ══════════════════════════════════════════════════════════
        // RECALC % + VALIDATE
        // ══════════════════════════════════════════════════════════
        private void RecalcAndValidate()
        {
            // Cập nhật label "Có x cử tri bầu lấy y" khi tổng số cử tri thay đổi
            double tongSoCuTri       = D(txtTongSoCuTri.Text);
            lblSoCuTriBau.Text = "Bầu lấy:";
            double soLuongCuTri      = D(txtSoLuongCuTri.Text);
            double soPhieuPhatRa     = D(txtSoPhieuPhatRa.Text);
            double soPhieuThuVao     = D(txtSoPhieuThuVao.Text);
            double soPhieuHopLe      = D(txtSoPhieuHopLe.Text);
            double soPhieuKhongHopLe = D(txtSoPhieuKhongHopLe.Text);
            double pb1 = D(txtPhieuBau1.Text), pb2 = D(txtPhieuBau2.Text),
                   pb3 = D(txtPhieuBau3.Text), pb4 = D(txtPhieuBau4.Text),
                   pb5 = D(txtPhieuBau5.Text);

            // ── auto % ─────────────────────────────────────────────
            SetPct(txtTyLeCuTri,       soLuongCuTri,      tongSoCuTri);
            SetPct(txtTyLeThuVao,      soPhieuThuVao,     soPhieuPhatRa);
            SetPct(txtTyLeHopLe,       soPhieuHopLe,      soPhieuThuVao);
            SetPct(txtTyLeKhongHopLe,  soPhieuKhongHopLe, soPhieuThuVao);

            // ── kiểm tra công thức đại biểu ──────────────────────
            int    yVal        = GetY();
            double tongUngVien = 0;
            foreach (var t in txtCandidateVotes) tongUngVien += D(t.Text);
            // Khi y=3: pb4=pb5=0 nên công thức tự rút gọn
            double kt = 1*pb1 + 2*pb2 + 3*pb3 + 4*pb4 + 5*pb5;

            // Cập nhật label công thức kiểm tra theo y
            string ktFormula = yVal == 3
                ? "🔍  Kiểm tra:  1×P1 + 2×P2 + 3×P3  =  Tổng phiếu ứng viên"
                : "🔍  Kiểm tra:  1×P1 + 2×P2 + 3×P3 + 4×P4 + 5×P5  =  Tổng phiếu ứng viên";
            // Tìm label kiểm tra (control ngay trước lblKiemTraResult trong cùng parent)
            var ktParent = lblKiemTraResult.Parent;
            if (ktParent != null)
                foreach (Control c in ktParent.Controls)
                    if (c is Label l && l != lblKiemTraResult && l.Font.Bold && l.Text.StartsWith("🔍"))
                    { l.Text = ktFormula; break; }

            if (pb1+pb2+pb3+pb4+pb5 > 0 || tongUngVien > 0)
            {
                bool ok = Math.Abs(kt - tongUngVien) < 0.001;
                lblKiemTraResult.Text      = ok
                    ? $"✅  HỢP LỆ  →  {kt:0} = {tongUngVien:0}"
                    : $"❌  KHÔNG HỢP LỆ  →  {kt:0} ≠ {tongUngVien:0}  (lệch {Math.Abs(kt-tongUngVien):0})";
                lblKiemTraResult.ForeColor = ok ? Color.FromArgb(30,130,76) : Color.Crimson;
            }
            else
            {
                lblKiemTraResult.Text      = "—";
                lblKiemTraResult.ForeColor = Color.Gray;
            }

            // ── validate ──────────────────────────────────────────
            var errs = Validate(tongSoCuTri, soLuongCuTri, soPhieuPhatRa, soPhieuThuVao,
                                soPhieuHopLe, soPhieuKhongHopLe, pb1,pb2,pb3,pb4,pb5,
                                tongUngVien, kt, GetY());
            ShowErrors(errs);
            btnSave.Enabled = errs.Count == 0 && _selectedRow != null
                                              && ExcelHelper.FilePath != null;
        }

        private void SetPct(TextBox box, double numerator, double denominator)
        {
            if (denominator > 0)
            {
                double pct = numerator / denominator * 100;
                box.Text      = pct.ToString("F2") + "%";
                box.ForeColor = pct > 100
                    ? Color.Crimson
                    : Color.FromArgb(30,130,76);
            }
            else
            {
                box.Text      = "—";
                box.ForeColor = Color.FromArgb(150,150,150);
            }
        }

        // ══════════════════════════════════════════════════════════
        // VALIDATE
        // ══════════════════════════════════════════════════════════
        private List<string> Validate(
            double tongSoCuTri, double soLuongCuTri,
            double soPhieuPhatRa, double soPhieuThuVao,
            double soPhieuHopLe, double soPhieuKhongHopLe,
            double pb1, double pb2, double pb3, double pb4, double pb5,
            double tongUngVien, double kt, int y)
        {
            var errs = new List<string>();
            if (_selectedRow == null) return errs;

            // 1. Bắt buộc nhập
            if (Empty(txtSoLuongCuTri))      errs.Add("• Số lượng cử tri tham gia bỏ phiếu: bắt buộc nhập.");
            if (Empty(txtSoPhieuPhatRa))      errs.Add("• Số phiếu phát ra: bắt buộc nhập.");
            if (Empty(txtSoPhieuThuVao))      errs.Add("• Số phiếu thu vào: bắt buộc nhập.");
            if (Empty(txtSoPhieuHopLe))       errs.Add("• Số phiếu hợp lệ: bắt buộc nhập.");
            if (Empty(txtSoPhieuKhongHopLe))  errs.Add("• Số phiếu không hợp lệ: bắt buộc nhập.");
            if (Empty(txtPhieuBau1)) errs.Add("• Phiếu bầu 1 đại biểu: bắt buộc nhập.");
            if (Empty(txtPhieuBau2)) errs.Add("• Phiếu bầu 2 đại biểu: bắt buộc nhập.");
            if (Empty(txtPhieuBau3)) errs.Add("• Phiếu bầu 3 đại biểu: bắt buộc nhập.");
            if (y == 5 && Empty(txtPhieuBau4)) errs.Add("• Phiếu bầu 4 đại biểu: bắt buộc nhập.");
            if (y == 5 && Empty(txtPhieuBau5)) errs.Add("• Phiếu bầu 5 đại biểu: bắt buộc nhập.");
            foreach (var t in txtCandidateVotes)
                if (string.IsNullOrWhiteSpace(t.Text)) {
                    string n = (t.Tag?.ToString() ?? "").Replace("Số phiếu bầu cho ","");
                    errs.Add($"• Phiếu bầu cho {n}: bắt buộc nhập.");
                }

            if (errs.Count > 0) return errs; // dừng sớm nếu còn trống

            // 2. Tổng số cử tri phải > 0
            if (tongSoCuTri <= 0)
                errs.Add("• Tổng số cử tri = 0. Kiểm tra lại dữ liệu khu vực trong Excel.");

            // 3. % không vượt 100
            double tyLeCuTri = tongSoCuTri > 0 ? soLuongCuTri / tongSoCuTri * 100 : 0;
            if (tyLeCuTri > 100)
                errs.Add($"• Tỷ lệ cử tri tham gia = {tyLeCuTri:F2}% > 100%."
                       + $"  →  Giảm số lượng cử tri tham gia (max {tongSoCuTri:0}).");

            // 4. Cử tri tham gia <= Tổng
            if (soLuongCuTri > tongSoCuTri)
                errs.Add($"• Cử tri tham gia ({soLuongCuTri:0}) > Tổng số cử tri ({tongSoCuTri:0}).");

            // 5. Phát ra <= Cử tri tham gia
            if (soPhieuPhatRa > soLuongCuTri)
                errs.Add($"• Phiếu phát ra ({soPhieuPhatRa:0}) > Cử tri tham gia ({soLuongCuTri:0}).");

            // 6. Thu vào <= Phát ra
            if (soPhieuThuVao > soPhieuPhatRa)
                errs.Add($"• Phiếu thu vào ({soPhieuThuVao:0}) > Phiếu phát ra ({soPhieuPhatRa:0}).");

            // 7. Thu vào = Hợp lệ + Không hợp lệ
            double sumHL = soPhieuHopLe + soPhieuKhongHopLe;
            if (Math.Abs(soPhieuThuVao - sumHL) > 0.001)
                errs.Add($"• Phiếu thu vào ({soPhieuThuVao:0}) ≠ Hợp lệ + Không hợp lệ = {sumHL:0}."
                       + $"  →  Sửa để tổng = {soPhieuThuVao:0}.");

            // 8. Không hợp lệ = Thu vào - Hợp lệ
            double khl = soPhieuThuVao - soPhieuHopLe;
            if (Math.Abs(soPhieuKhongHopLe - khl) > 0.001)
                errs.Add($"• Phiếu không hợp lệ ({soPhieuKhongHopLe:0}) ≠ Thu vào − Hợp lệ = {khl:0}."
                       + $"  →  Đặt = {khl:0}.");

            // 9. Hợp lệ = tổng 1+2+3+4+5 ĐB
            double tongDB = pb1+pb2+pb3+pb4+pb5;
            if (Math.Abs(soPhieuHopLe - tongDB) > 0.001)
                errs.Add($"• Phiếu hợp lệ ({soPhieuHopLe:0}) ≠ Tổng 1-5 ĐB = {tongDB:0}."
                       + $"  →  Kiểm tra lại các ô phiếu theo đại biểu.");

            // 10. Mỗi phiếu ĐB <= Hợp lệ
            void ChkDB(double v, string n) {
                if (v > soPhieuHopLe)
                    errs.Add($"• Phiếu {n} ({v:0}) > Phiếu hợp lệ ({soPhieuHopLe:0}).");
            }
            ChkDB(pb1,"bầu 1 ĐB"); ChkDB(pb2,"bầu 2 ĐB"); ChkDB(pb3,"bầu 3 ĐB");
            if (y == 5) { ChkDB(pb4,"bầu 4 ĐB"); ChkDB(pb5,"bầu 5 ĐB"); }

            // 11. Công thức tổng ứng viên
            if (Math.Abs(kt - tongUngVien) > 0.001)
                errs.Add($"• Kiểm tra ĐB: {kt:0} ≠ Tổng phiếu ứng viên {tongUngVien:0}."
                       + $"  →  Chênh lệch {Math.Abs(kt-tongUngVien):0}. Kiểm tra lại.");

            // 12. Mỗi ứng viên <= Hợp lệ
            foreach (var t in txtCandidateVotes) {
                double v = D(t.Text);
                if (v > soPhieuHopLe) {
                    string n = (t.Tag?.ToString() ?? "").Replace("Số phiếu bầu cho ","");
                    errs.Add($"• Phiếu bầu cho {n} ({v:0}) > Phiếu hợp lệ ({soPhieuHopLe:0}).");
                }
            }

            return errs;
        }

        // ══════════════════════════════════════════════════════════
        // SHOW ERRORS
        // ══════════════════════════════════════════════════════════
        private void ShowErrors(List<string> errs)
        {
            // ── Reset tất cả inline error labels ──────────────────
            ClearInlineErrors();

            lblErrors.AutoSize = false;

            if (errs.Count == 0)
            {
                lblErrors.Text      = "✅  Tất cả điều kiện hợp lệ – có thể lưu file!";
                lblErrors.ForeColor = Color.FromArgb(30, 130, 76);
                lblErrors.Font      = new Font("Segoe UI", 9.5f, FontStyle.Bold);
                pnlErrors.BackColor = Color.FromArgb(240, 255, 245);
            }
            else
            {
                lblErrors.Text      = string.Join("\r\n", errs);
                lblErrors.ForeColor = Color.FromArgb(180, 30, 30);
                lblErrors.Font      = new Font("Segoe UI", 9f);
                pnlErrors.BackColor = Color.FromArgb(255, 245, 245);

                // ── Hiện inline error dưới từng ô tương ứng ──────
                ApplyInlineErrors(errs);
            }

            int textH = TextRenderer.MeasureText(
                lblErrors.Text,
                lblErrors.Font,
                new Size(1192, int.MaxValue),
                TextFormatFlags.WordBreak).Height;

            lblErrors.Location = new Point(14, 36);
            lblErrors.Size     = new Size(1192, textH + 4);
            pnlErrors.Height   = textH + 56;

            RefreshLayout();
        }

        private void ClearInlineErrors()
        {
            var allInline = new[] {
                errSoLuongCuTri, errSoPhieuPhatRa, errSoPhieuThuVao,
                errSoPhieuHopLe, errSoPhieuKhongHopLe,
                errPhieuBau1, errPhieuBau2, errPhieuBau3,
                errPhieuBau4, errPhieuBau5
            };
            foreach (var l in allInline)
                if (l != null) { l.Text = ""; l.Visible = false; }
            foreach (var l in errCandidates)
                if (l != null) { l.Text = ""; l.Visible = false; }

            // Reset màu viền input về bình thường
            foreach (var t in new[] {
                txtSoLuongCuTri, txtSoPhieuPhatRa, txtSoPhieuThuVao,
                txtSoPhieuHopLe, txtSoPhieuKhongHopLe,
                txtPhieuBau1, txtPhieuBau2, txtPhieuBau3, txtPhieuBau4, txtPhieuBau5
            }) if (t != null) t.BackColor = Color.White;
        }

        private void ApplyInlineErrors(List<string> errs)
        {
            // Map từng key trong error message → inline label + textbox
            var map = new (string key, Label errLbl, TextBox txt)[] {
                ("cử tri tham gia",      errSoLuongCuTri,      txtSoLuongCuTri),
                ("phiếu phát ra",        errSoPhieuPhatRa,     txtSoPhieuPhatRa),
                ("phiếu thu vào",        errSoPhieuThuVao,     txtSoPhieuThuVao),
                ("phiếu hợp lệ",         errSoPhieuHopLe,      txtSoPhieuHopLe),
                ("không hợp lệ",         errSoPhieuKhongHopLe, txtSoPhieuKhongHopLe),
                ("bầu 1 đại biểu",       errPhieuBau1,         txtPhieuBau1),
                ("bầu 2 đại biểu",       errPhieuBau2,         txtPhieuBau2),
                ("bầu 3 đại biểu",       errPhieuBau3,         txtPhieuBau3),
                ("bầu 4 đại biểu",       errPhieuBau4,         txtPhieuBau4),
                ("bầu 5 đại biểu",       errPhieuBau5,         txtPhieuBau5),
            };

            foreach (var err in errs)
            {
                string errLower = err.ToLower();
                foreach (var (key, lbl, txt) in map)
                {
                    if (errLower.Contains(key.ToLower()) && lbl != null)
                    {
                        // Lấy thông báo ngắn gọn
                        string shortMsg = err.Replace("•", "").Trim();
                        int arrowIdx = shortMsg.IndexOf("→");
                        if (arrowIdx > 0) shortMsg = shortMsg.Substring(0, arrowIdx).Trim();
                        if (shortMsg.Length > 60) shortMsg = shortMsg.Substring(0, 60) + "…";
                        lbl.Text    = shortMsg;
                        lbl.Visible = true;
                        if (txt != null) txt.BackColor = Color.FromArgb(255, 235, 235);
                        break;
                    }
                }

                // Ứng viên inline errors
                for (int i = 0; i < errCandidates.Count && i < txtCandidateVotes.Count; i++)
                {
                    string tag = txtCandidateVotes[i].Tag?.ToString() ?? "";
                    string name = tag.Replace("Số phiếu bầu cho ", "").ToLower();
                    if (err.ToLower().Contains(name.Substring(0, Math.Min(10, name.Length))))
                    {
                        errCandidates[i].Text    = "Bắt buộc nhập";
                        errCandidates[i].Visible = true;
                        txtCandidateVotes[i].BackColor = Color.FromArgb(255, 235, 235);
                    }
                }
            }
        }

        // ══════════════════════════════════════════════════════════
        // XUẤT FILE TỔNG HỢP
        // ══════════════════════════════════════════════════════════
        private void BtnXuatTH_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(ExcelHelper.FilePath))
            {
                MessageBox.Show(
                    "⚠️  Vui lòng mở file Excel trước.",
                    "Chưa mở file", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var confirm = MessageBox.Show(
                $"Xuất tổng hợp từ các sheet ĐV1–ĐV5 vào sheet TH?\n\nFile: {System.IO.Path.GetFileName(ExcelHelper.FilePath)}",
                "Xác nhận xuất tổng hợp",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes) return;

            try
            {
                string result = ExcelHelper.XuatTongHop(ExcelHelper.FilePath);
                if (result.StartsWith("OK:"))
                {
                    string[] parts = result.Split('|');
                    int count = int.Parse(parts[0].Split(':')[1]);
                    string msg = $"✅  Xuất tổng hợp thành công!\n\nĐã tổng hợp {count}/5 đơn vị bầu cử vào sheet TH.";
                    if (parts.Length > 1)
                        msg += "\n\n⚠️  Cảnh báo:\n" + parts[1].Replace("WARN:", "");
                    MessageBox.Show(msg, "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show($"❌  {result}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"❌  Lỗi khi xuất tổng hợp:\n{ex.Message}\n\n→ Đảm bảo file không đang mở trong Excel.",
                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ══════════════════════════════════════════════════════════
        // SAVE
        // ══════════════════════════════════════════════════════════
        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (_selectedRow == null) return;

            double tongSoCuTri       = D(txtTongSoCuTri.Text);
            double soLuongCuTri      = D(txtSoLuongCuTri.Text);
            double soPhieuPhatRa     = D(txtSoPhieuPhatRa.Text);
            double soPhieuThuVao     = D(txtSoPhieuThuVao.Text);
            double soPhieuHopLe      = D(txtSoPhieuHopLe.Text);
            double soPhieuKhongHopLe = D(txtSoPhieuKhongHopLe.Text);

            var sd = new SaveData {
                TongSoCuTri         = tongSoCuTri,
                SoLuongCuTriThamGia = soLuongCuTri,
                TyLeCuTri           = tongSoCuTri   > 0 ? soLuongCuTri      / tongSoCuTri   * 100 : 0,
                SoPhieuPhatRa       = soPhieuPhatRa,
                SoPhieuThuVao       = soPhieuThuVao,
                TyLeThuVao          = soPhieuPhatRa > 0 ? soPhieuThuVao     / soPhieuPhatRa * 100 : 0,
                SoPhieuHopLe        = soPhieuHopLe,
                TyLeHopLe           = soPhieuThuVao > 0 ? soPhieuHopLe      / soPhieuThuVao * 100 : 0,
                SoPhieuKhongHopLe   = soPhieuKhongHopLe,
                TyLeKhongHopLe      = soPhieuThuVao > 0 ? soPhieuKhongHopLe / soPhieuThuVao * 100 : 0,
                PhieuBau1DB         = D(txtPhieuBau1.Text),
                PhieuBau2DB         = D(txtPhieuBau2.Text),
                PhieuBau3DB         = D(txtPhieuBau3.Text),
                PhieuBau4DB         = D(txtPhieuBau4.Text),
                PhieuBau5DB         = D(txtPhieuBau5.Text)
            };
            sd.Y = GetY();
            foreach (var t in txtCandidateVotes) sd.CandidateVotes.Add(D(t.Text));

            try
            {
                ExcelHelper.WriteRowData(cboSheet.SelectedItem.ToString(), _selectedRow.RowIndex, sd);
                MessageBox.Show(
                    $"✅ Đã lưu thành công!\n\nSheet: {cboSheet.SelectedItem}\nKhu vực: {_selectedRow.KhuVuc}",
                    "Lưu thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"❌ Lỗi khi ghi file:\n{ex.Message}\n\n→ Đảm bảo file không bị mở bởi Excel.",
                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ══════════════════════════════════════════════════════════
        // CLEAR
        // ══════════════════════════════════════════════════════════
        private void BtnClear_Click(object sender, EventArgs e)
        {
            _updating = true;
            foreach (var t in new[] {
                txtSoLuongCuTri, txtSoPhieuPhatRa, txtSoPhieuThuVao,
                txtSoPhieuHopLe, txtSoPhieuKhongHopLe,
                txtPhieuBau1, txtPhieuBau2, txtPhieuBau3, txtPhieuBau4, txtPhieuBau5
            }) t.Text = "";
            foreach (var t in txtCandidateVotes) t.Text = "";
            foreach (var t in new[] {
                txtTyLeCuTri, txtTyLeThuVao, txtTyLeHopLe, txtTyLeKhongHopLe
            }) { t.Text = "—"; t.ForeColor = Color.FromArgb(150,150,150); }
            lblKiemTraResult.Text      = "—";
            lblKiemTraResult.ForeColor = Color.Gray;
            _updating = false;

            ShowErrors(new List<string>());
            btnSave.Enabled = false;
        }

        // ══════════════════════════════════════════════════════════
        // HELPERS
        // ══════════════════════════════════════════════════════════
        private double D(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return 0;
            s = s.Replace(",",".").Replace("%","").Trim();
            return double.TryParse(s,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : 0;
        }

        private bool Empty(TextBox t) => string.IsNullOrWhiteSpace(t.Text);
    }
}
