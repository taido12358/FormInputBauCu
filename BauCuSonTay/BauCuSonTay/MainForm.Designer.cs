using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace BauCuSonTay
{
    public partial class MainForm : Form
    {
        // ── selectors ──────────────────────────────────────────────
        private Label    lblFileInfo;
        private Button   btnOpenFile;
        private ComboBox cboSheet;
        private ComboBox cboKhuVuc;

        // ── readonly info ──────────────────────────────────────────
        private TextBox txtToDP, txtTongSoKV, txtTongSoCuTri;

        // ── editable inputs ────────────────────────────────────────
        private TextBox txtSoLuongCuTri;
        private TextBox txtSoPhieuPhatRa;
        private TextBox txtSoPhieuThuVao;
        private TextBox txtSoPhieuHopLe;
        private TextBox txtSoPhieuKhongHopLe;
        private TextBox txtPhieuBau1, txtPhieuBau2, txtPhieuBau3, txtPhieuBau4, txtPhieuBau5;
        private Label   lblPhieuBau4Label, lblPhieuBau5Label;

        // ── auto-calc % ──────────────────────────────────────────
        private TextBox txtTyLeCuTri;
        private TextBox txtTyLeThuVao;
        private TextBox txtTyLeHopLe;
        private TextBox txtTyLeKhongHopLe;

        // ── so cu tri bau (y) ────────────────────────────────────
        private Label    lblSoCuTriBau;
        private ComboBox cboSoCuTriBau;

        // ── inline error labels ──────────────────────────────────
        private Label errSoLuongCuTri, errSoPhieuPhatRa, errSoPhieuThuVao;
        private Label errSoPhieuHopLe, errSoPhieuKhongHopLe;
        private Label errPhieuBau1, errPhieuBau2, errPhieuBau3, errPhieuBau4, errPhieuBau5;
        private List<Label> errCandidates = new List<Label>();

        // ── kiem tra ─────────────────────────────────────────────
        private Label lblKiemTraResult;

        // ── candidates ───────────────────────────────────────────
        private Panel         pnlCandidates;
        private Label         lblCandidatesHeader;
        private List<Label>   lblCandidates     = new List<Label>();
        private List<TextBox> txtCandidateVotes = new List<TextBox>();

        // ── errors / buttons ─────────────────────────────────────
        private Button btnXuatTH;
        private Panel  pnlErrors;
        private Label  lblErrors;
        private Button btnSave, btnClear;

        // ── scroll container ─────────────────────────────────────
        private Panel pnlScroll;
        private Panel pnlMain;

        // ── section panels (để toggle ẩn/hiện) ───────────────────
        private Panel pnlSec1, pnlSec2, pnlSec3, pnlSec4, pnlSec5;
        private Panel pnlModeTH;     // hiển thị khi chọn sheet TH

        // ═══════════════════════════════════════════════════════════
        // LAYOUT CONSTANTS  (single source of truth)
        // ═══════════════════════════════════════════════════════════
        private const int FORM_W    = 1280;   // form width
        private const int INNER_W   = 1220;   // section width
        private const int PAD       = 14;     // left padding inside section
        private const int ROW_H     = 36;     // height per input row
        private const int SEC_HDR   = 30;     // section header height
        private const int INPUT_H   = 26;     // textbox height (approx, WinForms auto)

        // Column widths for the 2-column "label | input | label | input" grid
        private const int COL_LBL   = 260;    // label width (left block)
        private const int COL_INP   = 140;    // input width (left block)
        private const int COL_GAP   = 40;     // gap between left and right block
        private const int COL2_X    = PAD + COL_LBL + COL_INP + COL_GAP + 8; // right block start X
        private const int COL2_LBL  = 280;    // label width (right block = % label)
        private const int COL2_INP  = 160;    // input width (right block = % value)

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.Text          = "Nhập Kết Quả Bầu Cử – Phường Sơn Tây";
            this.Size          = new Size(FORM_W, 860);
            this.MinimumSize   = new Size(1000, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font          = new Font("Segoe UI", 9.5f);
            this.BackColor     = Color.FromArgb(235, 240, 246);

            // ── scroll wrapper ───────────────────────────────────
            pnlScroll = new Panel {
                Dock       = DockStyle.Fill,
                AutoScroll = true,
                BackColor  = Color.FromArgb(235, 240, 246)
            };
            this.Controls.Add(pnlScroll);

            pnlMain = new Panel {
                Width     = FORM_W - 20,
                Location  = new Point(0, 0),
                BackColor = Color.FromArgb(235, 240, 246)
            };
            pnlScroll.Controls.Add(pnlMain);

            int y = 12;

            // ═══════════════════════════════════════════════════════
            // TITLE
            // ═══════════════════════════════════════════════════════
            var pnlTitle = new Panel {
                Location  = new Point(10, y),
                Width     = INNER_W,
                Height    = 50,
                BackColor = Color.FromArgb(21, 67, 96)
            };
            pnlTitle.Controls.Add(new Label {
                Text      = "   📋  BIỂU NHẬP KẾT QUẢ BẦU CỬ – PHƯỜNG SƠN TÂY",
                Font      = new Font("Segoe UI", 14f, FontStyle.Bold),
                ForeColor = Color.White,
                Dock      = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            });

            btnXuatTH = new Button {
                Text      = "📊  Xuất file tổng hợp",
                Size      = new Size(210, 34),
                Location  = new Point(INNER_W - 220, 8),
                BackColor = Color.FromArgb(155, 89, 182),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Segoe UI", 10f, FontStyle.Bold),
                Cursor    = Cursors.Hand
            };
            btnXuatTH.FlatAppearance.BorderSize = 0;
            btnXuatTH.Click += BtnXuatTH_Click;
            pnlTitle.Controls.Add(btnXuatTH);

            pnlMain.Controls.Add(pnlTitle);
            y += 58;

            // ═══════════════════════════════════════════════════════
            // SEC 1 – File & khu vực
            // ═══════════════════════════════════════════════════════
            pnlSec1 = Sec("📂  File Excel & Khu Vực", ref y, 112);

            // row 1: file info + button
            SL(pnlSec1, "File đang dùng:", PAD, 38);
            lblFileInfo = new Label {
                Text      = "Chưa chọn file …",
                ForeColor = Color.FromArgb(160, 160, 160),
                Location  = new Point(160, 41),
                Size      = new Size(INNER_W - 160 - 180 - 20, 22),
                Font      = new Font("Segoe UI", 9f, FontStyle.Italic)
            };
            pnlSec1.Controls.Add(lblFileInfo);

            btnOpenFile = Btn("📂  Mở File Excel",
                new Point(INNER_W - 190, 36),
                Color.FromArgb(41, 128, 185), 180, 32);
            btnOpenFile.Click += BtnOpenFile_Click;
            pnlSec1.Controls.Add(btnOpenFile);

            // row 2: sheet + khu vực
            SL(pnlSec1, "Đơn vị bầu cử:", PAD, 78);
            cboSheet = new ComboBox {
                Location      = new Point(160, 75),
                Width         = 160,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Enabled       = false,
                Font          = new Font("Segoe UI", 10f)
            };
            cboSheet.SelectedIndexChanged += CboSheet_SelectedIndexChanged;
            pnlSec1.Controls.Add(cboSheet);

            SL(pnlSec1, "Khu vực bỏ phiếu:", 340, 78);
            cboKhuVuc = new ComboBox {
                Location      = new Point(470, 75),
                Width         = INNER_W - 490,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Enabled       = false,
                Font          = new Font("Segoe UI", 10f)
            };
            cboKhuVuc.SelectedIndexChanged += CboKhuVuc_SelectedIndexChanged;
            pnlSec1.Controls.Add(cboKhuVuc);

            // ═══════════════════════════════════════════════════════
            // SEC 2 – Thông tin khu vực (3 ô 1 hàng)
            // ═══════════════════════════════════════════════════════
            pnlSec2 = Sec("ℹ️  Thông Tin Khu Vực", ref y, 68);

            // Chia đều 3 ô trên 1 hàng
            int s2w = (INNER_W - PAD * 2 - 40) / 3;  // width mỗi nhóm label+input
            int s2y = 38;

            txtToDP = ReadBox(pnlSec2, "Thuộc tổ dân phố:",
                PAD, s2y, 160, s2w - 164);

            txtTongSoKV = ReadBox(pnlSec2, "Tổng số khu vực bỏ phiếu:",
                PAD + s2w + 20, s2y, 220, s2w - 224);

            txtTongSoCuTri = InputBox(pnlSec2, "Tổng số cử tri:",
                PAD + (s2w + 20) * 2, s2y, 140, s2w - 144);
            txtTongSoCuTri.TextChanged += InputChanged;

            // ═══════════════════════════════════════════════════════
            // SEC 3 – Cử tri
            // ═══════════════════════════════════════════════════════
            pnlSec3 = Sec("👥  Thông Tin Cử Tri", ref y, 90);
            int s3mid = INNER_W / 2;   // midpoint

            txtSoLuongCuTri = InputBox(pnlSec3,
                "Số lượng cử tri đã tham gia bỏ phiếu:",
                PAD, 38, 310, COL_INP);
            txtSoLuongCuTri.TextChanged += InputChanged;
            errSoLuongCuTri = new Label {
                Text      = "",
                ForeColor = Color.Crimson,
                Font      = new Font("Segoe UI", 8.5f, FontStyle.Italic),
                Location  = new Point(PAD, 38+26),
                Size      = new Size(380, 16),
                Visible   = false
            };
            pnlSec3.Controls.Add(errSoLuongCuTri);

            CalcBox(pnlSec3, "Tỷ lệ % cử tri tham gia / Tổng số cử tri:",
                out txtTyLeCuTri,
                s3mid + 10, 38, 330, COL2_INP);

            // ═══════════════════════════════════════════════════════
            // SEC 4 – Phiếu bầu  (2 cột cân bằng)
            // ═══════════════════════════════════════════════════════
            pnlSec4 = Sec("🗳️  Thông Tin Phiếu Bầu", ref y, 282);

            // Cột trái: label+input, Cột phải: label % + value %
            int c1x = PAD;
            int c1lw = 190;   // label width cột trái
            int c1iw = COL_INP;
            int c2x  = s3mid + 10;
            int c2lw = 280;
            int c2iw = COL2_INP;
            const int ERR_ROW = 58; // input(26) + gap(6) + error(16) + margin(10)

            int r4 = 36;

            // Phát ra – chỉ input (không có %)
            txtSoPhieuPhatRa = InputBox(pnlSec4, "Số phiếu phát ra:", c1x, r4, c1lw, c1iw);
            txtSoPhieuPhatRa.TextChanged += InputChanged;
            errSoPhieuPhatRa = new Label { Text="", ForeColor=Color.Crimson, Font=new Font("Segoe UI",8.5f,FontStyle.Italic), Location=new Point(c1x, r4+26), Size=new Size(c1lw+c1iw+4, 16), Visible=false };
            pnlSec4.Controls.Add(errSoPhieuPhatRa);
            r4 += ERR_ROW;

            // Thu vào  +  %
            txtSoPhieuThuVao = InputBox(pnlSec4, "Số phiếu thu vào:", c1x, r4, c1lw, c1iw);
            txtSoPhieuThuVao.TextChanged += InputChanged;
            errSoPhieuThuVao = new Label { Text="", ForeColor=Color.Crimson, Font=new Font("Segoe UI",8.5f,FontStyle.Italic), Location=new Point(c1x, r4+26), Size=new Size(c1lw+c1iw+4, 16), Visible=false };
            pnlSec4.Controls.Add(errSoPhieuThuVao);
            CalcBox(pnlSec4, "% số phiếu thu vào / phiếu phát ra:", out txtTyLeThuVao, c2x, r4, c2lw, c2iw);
            r4 += ERR_ROW;

            // Hợp lệ  +  %
            txtSoPhieuHopLe = InputBox(pnlSec4, "Số phiếu hợp lệ:", c1x, r4, c1lw, c1iw);
            txtSoPhieuHopLe.TextChanged += InputChanged;
            errSoPhieuHopLe = new Label { Text="", ForeColor=Color.Crimson, Font=new Font("Segoe UI",8.5f,FontStyle.Italic), Location=new Point(c1x, r4+26), Size=new Size(c1lw+c1iw+4, 16), Visible=false };
            pnlSec4.Controls.Add(errSoPhieuHopLe);
            CalcBox(pnlSec4, "% số phiếu hợp lệ / phiếu thu vào:", out txtTyLeHopLe, c2x, r4, c2lw, c2iw);
            r4 += ERR_ROW;

            // Không hợp lệ  +  %
            txtSoPhieuKhongHopLe = InputBox(pnlSec4, "Số phiếu không hợp lệ:", c1x, r4, c1lw, c1iw);
            txtSoPhieuKhongHopLe.TextChanged += InputChanged;
            errSoPhieuKhongHopLe = new Label { Text="", ForeColor=Color.Crimson, Font=new Font("Segoe UI",8.5f,FontStyle.Italic), Location=new Point(c1x, r4+26), Size=new Size(c1lw+c1iw+4, 16), Visible=false };
            pnlSec4.Controls.Add(errSoPhieuKhongHopLe);
            CalcBox(pnlSec4, "% số phiếu không hợp lệ / phiếu thu vào:", out txtTyLeKhongHopLe, c2x, r4, c2lw + 20, c2iw);
            pnlSec4.Height = r4 + ERR_ROW + 16;

            // ═══════════════════════════════════════════════════════
            // SEC 5 – Số phiếu theo đại biểu  (3 cột × 2 hàng)
            // ═══════════════════════════════════════════════════════
            pnlSec5 = Sec("📊  Số Phiếu Bầu Theo Số Đại Biểu", ref y, 200);

            // ── Row 0: Có x cử tri bầu lấy y ─────────────────────
            int db_y0 = 36;
            lblSoCuTriBau = new Label {
                Text      = "Bầu lấy:",
                Location  = new Point(PAD, db_y0 + 4),
                AutoSize  = true,
                Font      = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(31, 78, 121)
            };
            pnlSec5.Controls.Add(lblSoCuTriBau);

            cboSoCuTriBau = new ComboBox {
                Location      = new Point(PAD + 68, db_y0),
                Width         = 80,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font          = new Font("Segoe UI", 10.5f, FontStyle.Bold)
            };
            cboSoCuTriBau.Items.AddRange(new object[] { "3", "5" });
            cboSoCuTriBau.SelectedIndex = 1; // mặc định = 5
            cboSoCuTriBau.SelectedIndexChanged += CboSoCuTriBau_Changed;
            pnlSec5.Controls.Add(cboSoCuTriBau);

            pnlSec5.Controls.Add(new Label {
                Text      = "đại biểu",
                Location  = new Point(PAD + 156, db_y0 + 4),
                AutoSize  = true,
                ForeColor = Color.FromArgb(52, 73, 94)
            });

            // ── 5 ô input: hàng 1 = [1db][2db][3db], hàng 2 = [4db][5db]
            int dbCols = 3;
            int dbW    = (INNER_W - PAD * 2 - (dbCols - 1) * 20) / dbCols;
            int dbLW   = 210;
            int dbIW   = dbW - dbLW - 12;
            int db_y1  = db_y0 + ROW_H + 4;
            int db_y2  = db_y1 + 58;  // 58 = input + error gap

            txtPhieuBau1 = InputBox(pnlSec5, "Số phiếu bầu 1 đại biểu:", PAD,                    db_y1, dbLW, dbIW); txtPhieuBau1.TextChanged += InputChanged;
            txtPhieuBau2 = InputBox(pnlSec5, "Số phiếu bầu 2 đại biểu:", PAD + dbW + 20,         db_y1, dbLW, dbIW); txtPhieuBau2.TextChanged += InputChanged;
            txtPhieuBau3 = InputBox(pnlSec5, "Số phiếu bầu 3 đại biểu:", PAD + (dbW + 20) * 2,  db_y1, dbLW, dbIW); txtPhieuBau3.TextChanged += InputChanged;
            errPhieuBau1 = new Label { Text="", ForeColor=Color.Crimson, Font=new Font("Segoe UI",8.5f,FontStyle.Italic), Location=new Point(PAD, db_y1+26), Size=new Size(dbLW+dbIW+4, 16), Visible=false }; pnlSec5.Controls.Add(errPhieuBau1);
            errPhieuBau2 = new Label { Text="", ForeColor=Color.Crimson, Font=new Font("Segoe UI",8.5f,FontStyle.Italic), Location=new Point(PAD+dbW+20, db_y1+26), Size=new Size(dbLW+dbIW+4, 16), Visible=false }; pnlSec5.Controls.Add(errPhieuBau2);
            errPhieuBau3 = new Label { Text="", ForeColor=Color.Crimson, Font=new Font("Segoe UI",8.5f,FontStyle.Italic), Location=new Point(PAD+(dbW+20)*2, db_y1+26), Size=new Size(dbLW+dbIW+4, 16), Visible=false }; pnlSec5.Controls.Add(errPhieuBau3);
            txtPhieuBau4 = InputBoxWithLabelRef(pnlSec5, "Số phiếu bầu 4 đại biểu:", PAD,           db_y2, dbLW, dbIW, out lblPhieuBau4Label); txtPhieuBau4.TextChanged += InputChanged;
            txtPhieuBau5 = InputBoxWithLabelRef(pnlSec5, "Số phiếu bầu 5 đại biểu:", PAD+dbW+20, db_y2, dbLW, dbIW, out lblPhieuBau5Label); txtPhieuBau5.TextChanged += InputChanged;
            errPhieuBau4 = new Label { Text="", ForeColor=Color.Crimson, Font=new Font("Segoe UI",8.5f,FontStyle.Italic), Location=new Point(PAD, db_y2+26), Size=new Size(dbLW+dbIW+4, 16), Visible=false }; pnlSec5.Controls.Add(errPhieuBau4);
            errPhieuBau5 = new Label { Text="", ForeColor=Color.Crimson, Font=new Font("Segoe UI",8.5f,FontStyle.Italic), Location=new Point(PAD+dbW+20, db_y2+26), Size=new Size(dbLW+dbIW+4, 16), Visible=false }; pnlSec5.Controls.Add(errPhieuBau5);

            // Dòng kiểm tra
            int ktY = db_y2 + 58 + 8;  // 58 = input(26)+err(16)+gap(16)
            pnlSec5.Controls.Add(new Label {
                Text      = "🔍  Kiểm tra:  1×P1 + 2×P2 + 3×P3 + 4×P4 + 5×P5  =  Tổng phiếu ứng viên",
                Font      = new Font("Segoe UI", 9f, FontStyle.Bold),
                ForeColor = Color.FromArgb(52, 73, 94),
                Location  = new Point(PAD, ktY),
                AutoSize  = true
            });
            lblKiemTraResult = new Label {
                Text      = "—",
                Font      = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.Gray,
                Location  = new Point(PAD + 20, ktY + 24),
                Size      = new Size(INNER_W - PAD * 2 - 20, 22)
            };
            pnlSec5.Controls.Add(lblKiemTraResult);
            pnlSec5.Height = ktY + 56;

            // ═══════════════════════════════════════════════════════
            // SEC 6 – Ứng viên (dynamic)
            // ═══════════════════════════════════════════════════════
            pnlCandidates = Sec("🏅  Số Phiếu Bầu Cho Từng Ứng Viên", ref y, 68);
            lblCandidatesHeader = new Label {
                Text      = "Vui lòng chọn Sheet và Khu vực bỏ phiếu để hiển thị ứng viên",
                ForeColor = Color.FromArgb(160, 160, 160),
                Font      = new Font("Segoe UI", 9f, FontStyle.Italic),
                Location  = new Point(PAD, 40),
                AutoSize  = true
            };
            pnlCandidates.Controls.Add(lblCandidatesHeader);

            // ═══════════════════════════════════════════════════════
            // SEC 7 – Kiểm tra / lỗi
            // ═══════════════════════════════════════════════════════
            pnlErrors = Sec("⚠️  Kiểm Tra Điều Kiện", ref y, 68);
            pnlErrors.BackColor = Color.FromArgb(255, 251, 245);
            lblErrors = new Label {
                Text        = "Chưa có dữ liệu để kiểm tra.",
                ForeColor   = Color.FromArgb(160, 160, 160),
                Font        = new Font("Segoe UI", 9f, FontStyle.Italic),
                Location    = new Point(PAD, 36),
                AutoSize    = false,
                Size        = new Size(INNER_W - PAD * 2, 22)
            };
            pnlErrors.Controls.Add(lblErrors);

            // ═══════════════════════════════════════════════════════
            // BUTTONS
            // ═══════════════════════════════════════════════════════
            y += 12;
            btnSave = Btn("💾  Lưu vào Excel", new Point(10, y), Color.FromArgb(39, 174, 96), 200, 44);
            btnSave.Font    = new Font("Segoe UI", 11f, FontStyle.Bold);
            btnSave.Enabled = false;
            btnSave.Click  += BtnSave_Click;
            pnlMain.Controls.Add(btnSave);

            btnClear = Btn("🗑️  Xóa dữ liệu", new Point(220, y), Color.FromArgb(189, 54, 47), 165, 44);
            btnClear.Font  = new Font("Segoe UI", 10.5f);
            btnClear.Click += BtnClear_Click;
            pnlMain.Controls.Add(btnClear);

            y += 64;
            pnlMain.Height = y;

            // ── Panel chế độ TH (ẩn mặc định) ───────────────────
            pnlModeTH = new Panel {
                Location  = new Point(10, 70),
                Width     = INNER_W,
                Height    = 400,
                BackColor = Color.FromArgb(235, 240, 246),
                Visible   = false
            };

            var btnQuayLai = new Button {
                Text      = "← Quay lại",
                Size      = new Size(140, 36),
                Location  = new Point(PAD, 16),
                BackColor = Color.FromArgb(100, 110, 130),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Segoe UI", 10f),
                Cursor    = Cursors.Hand
            };
            btnQuayLai.FlatAppearance.BorderSize = 0;
            btnQuayLai.Click += (s, ev) => {
                cboSheet.SelectedIndex = 0;
            };
            pnlModeTH.Controls.Add(btnQuayLai);

            var btnXuatLon = new Button {
                Text      = "📊  Xuất File Tổng Hợp",
                Size      = new Size(340, 80),
                Location  = new Point((INNER_W - 340) / 2, 120),
                BackColor = Color.FromArgb(125, 60, 152),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Segoe UI", 16f, FontStyle.Bold),
                Cursor    = Cursors.Hand
            };
            btnXuatLon.FlatAppearance.BorderSize = 0;
            btnXuatLon.Click += BtnXuatTH_Click;
            pnlModeTH.Controls.Add(btnXuatLon);

            pnlModeTH.Controls.Add(new Label {
                Text      = "Tổng hợp kết quả từ sheet ĐV1, ĐV2, ĐV3, ĐV4, ĐV5 → ghi vào sheet TH",
                Font      = new Font("Segoe UI", 11f, FontStyle.Italic),
                ForeColor = Color.FromArgb(100, 100, 120),
                Location  = new Point(0, 220),
                Size      = new Size(INNER_W, 30),
                TextAlign = ContentAlignment.MiddleCenter
            });

            pnlMain.Controls.Add(pnlModeTH);

            this.ResumeLayout(false);
        }

        // ═══════════════════════════════════════════════════════════
        // HELPERS
        // ═══════════════════════════════════════════════════════════

        private Panel Sec(string title, ref int y, int height)
        {
            var pnl = new Panel {
                Location  = new Point(10, y),
                Width     = INNER_W,
                Height    = height,
                BackColor = Color.White
            };
            pnl.Paint += (s, e) => {
                using (var pen = new System.Drawing.Pen(Color.FromArgb(205, 218, 235)))
                    e.Graphics.DrawRectangle(pen, 0, 0, pnl.Width - 1, pnl.Height - 1);
            };
            var hdr = new Label {
                Text      = "  " + title,
                Font      = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(31, 78, 121),
                Location  = new Point(0, 0),
                Size      = new Size(INNER_W, SEC_HDR),
                TextAlign = ContentAlignment.MiddleLeft
            };
            pnl.Controls.Add(hdr);
            pnlMain.Controls.Add(pnl);
            y += height + 8;
            return pnl;
        }

        private void SL(Panel p, string text, int x, int y)
        {
            p.Controls.Add(new Label {
                Text      = text,
                Location  = new Point(x, y + 4),
                AutoSize  = true,
                ForeColor = Color.FromArgb(52, 73, 94)
            });
        }

        private TextBox ReadBox(Panel p, string lbl, int lx, int ly, int lw, int tw)
        {
            SL(p, lbl, lx, ly);
            var t = new TextBox {
                Location    = new Point(lx + lw + 4, ly),
                Width       = tw,
                Font        = new Font("Segoe UI", 10f),
                BackColor   = Color.FromArgb(236, 240, 241),
                ReadOnly    = true,
                BorderStyle = BorderStyle.FixedSingle
            };
            p.Controls.Add(t);
            return t;
        }

        private TextBox InputBox(Panel p, string lbl, int lx, int ly, int lw, int tw)
        {
            SL(p, lbl, lx, ly);
            var t = new TextBox {
                Location    = new Point(lx + lw + 4, ly),
                Width       = tw,
                Font        = new Font("Segoe UI", 10.5f),
                BorderStyle = BorderStyle.FixedSingle
            };
            t.Enter += (s, e) => ((TextBox)s).BackColor = Color.FromArgb(235, 245, 255);
            t.Leave += (s, e) => ((TextBox)s).BackColor = Color.White;
            p.Controls.Add(t);
            return t;
        }

        // Giống InputBox nhưng trả ra reference của label để ẩn/hiện
        private TextBox InputBoxWithLabelRef(Panel p, string lbl, int lx, int ly, int lw, int tw, out Label labelRef)
        {
            labelRef = new Label {
                Text      = lbl,
                Location  = new Point(lx, ly + 4),
                AutoSize  = true,
                ForeColor = Color.FromArgb(52, 73, 94)
            };
            p.Controls.Add(labelRef);
            var t = new TextBox {
                Location    = new Point(lx + lw + 4, ly),
                Width       = tw,
                Font        = new Font("Segoe UI", 10.5f),
                BorderStyle = BorderStyle.FixedSingle
            };
            t.Enter += (s, e) => ((TextBox)s).BackColor = Color.FromArgb(235, 245, 255);
            t.Leave += (s, e) => ((TextBox)s).BackColor = Color.White;
            p.Controls.Add(t);
            return t;
        }

        private void CalcBox(Panel p, string lbl, out TextBox t, int lx, int ly, int lw, int tw)
        {
            SL(p, lbl, lx, ly);
            t = new TextBox {
                Location    = new Point(lx + lw + 4, ly),
                Width       = tw,
                Font        = new Font("Segoe UI", 10.5f, FontStyle.Bold),
                BackColor   = Color.FromArgb(232, 248, 232),
                ForeColor   = Color.FromArgb(30, 130, 76),
                ReadOnly    = true,
                BorderStyle = BorderStyle.FixedSingle,
                TextAlign   = HorizontalAlignment.Center
            };
            p.Controls.Add(t);
        }

        private Button Btn(string text, Point loc, Color bg, int w, int h)
        {
            var b = new Button {
                Text      = text,
                Location  = loc,
                Size      = new Size(w, h),
                BackColor = bg,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor    = Cursors.Hand
            };
            b.FlatAppearance.BorderSize = 0;
            return b;
        }
    }
}
