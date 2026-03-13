using ClosedXML.Excel;
using Microsoft.EntityFrameworkCore;
using QuanLyQuanAn.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QuanLyQuanAn.Forms
{
    public partial class frmHoaDon : Form
    {
        QLQADbcontext context = new QLQADbcontext();
        int id;

        public frmHoaDon()
        {
            InitializeComponent();
        }

        private void frmHoaDon_Load(object sender, EventArgs e)
        {
            dataGridView.AutoGenerateColumns = false;

            // --- CHÌA KHÓA FIX LỖI XÓA ---
            // Xóa bộ nhớ đệm (Cache) để nó lấy dữ liệu mới nhất từ Database, không bị nhớ dai
            context.ChangeTracker.Clear();

            // Dùng .Include để yêu cầu SQL lôi luôn dữ liệu bảng liên quan lên RAM
            var dsHoaDon = context.HoaDon
                .Include(r => r.NhanVien)
                .Include(r => r.KhachHang)
                .Include(r => r.HoaDon_ChiTiet)
                .ToList();

            // Đổ dữ liệu vào danh sách hiển thị
            List<DanhSachHoaDon> hd = dsHoaDon.Select(r => new DanhSachHoaDon
            {
                ID = r.ID,
                NhanVienID = r.NhanVienID,
                HoVaTenNhanVien = r.NhanVien?.HoVaTen ?? "Chưa xác định",
                KhachHangID = r.KhachHangID,
                HoVaTenKhachHang = r.KhachHang?.HoVaTen ?? "Khách vãng lai",
                NgayLap = r.NgayLap,
                GhiChuHoaDon = r.GhiChuHoaDon,
                TongTienHoaDon = r.HoaDon_ChiTiet.Sum(ct => (double?)((int)ct.SoLuongBan * ct.DonGiaBan)) ?? 0,
                XemChiTiet = "Xem chi tiết"
            }).ToList();

            dataGridView.DataSource = hd;
        }

        private void btnLapHoaDon_Click(object sender, EventArgs e)
        {
            using (frmHoaDon_ChiTiet chiTiet = new frmHoaDon_ChiTiet())
            {
                chiTiet.ShowDialog();
            }
            frmHoaDon_Load(sender, e);
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            if (dataGridView.CurrentRow == null) return;

            id = Convert.ToInt32(dataGridView.CurrentRow.Cells["ID"].Value.ToString());
            using (frmHoaDon_ChiTiet chiTiet = new frmHoaDon_ChiTiet(id))
            {
                chiTiet.ShowDialog();
            }
            frmHoaDon_Load(sender, e);
        }

        // --- HÀM XÓA ĐÃ ĐƯỢC LÀM LẠI ĐỂ TRÁNH LỖI CONCURRENCY ---
        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (dataGridView.CurrentRow == null) return;

            if (MessageBox.Show("Bạn có chắc chắn muốn xóa hóa đơn này?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    id = Convert.ToInt32(dataGridView.CurrentRow.Cells["ID"].Value.ToString());

                    // Dùng một Context MỚI TINH để thực hiện lệnh xóa
                    using (var dbDelete = new QLQADbcontext())
                    {
                        var hd = dbDelete.HoaDon.Find(id);
                        if (hd != null)
                        {
                            // Xóa các chi tiết món ăn trước
                            var chiTiets = dbDelete.HoaDon_ChiTiet.Where(ct => ct.HoaDonID == id);
                            dbDelete.HoaDon_ChiTiet.RemoveRange(chiTiets);

                            // Xóa hóa đơn chính sau
                            dbDelete.HoaDon.Remove(hd);
                            dbDelete.SaveChanges();
                        }
                    }

                    MessageBox.Show("Đã xóa hóa đơn thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // Gọi lại hàm Load để làm mới lưới
                    frmHoaDon_Load(sender, e);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi xóa: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn thoát?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                this.Close();
            }
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            string tuKhoa = NhapTuKhoaPopUp("Nhập tên Khách hàng hoặc Nhân viên cần tìm:", "Tìm kiếm Hóa Đơn");

            if (!string.IsNullOrWhiteSpace(tuKhoa))
            {
                tuKhoa = tuKhoa.ToLower();

                var hd = context.HoaDon
                    .Where(r => r.KhachHang.HoVaTen.ToLower().Contains(tuKhoa) || r.NhanVien.HoVaTen.ToLower().Contains(tuKhoa))
                    .Select(r => new DanhSachHoaDon
                    {
                        ID = r.ID,
                        NhanVienID = r.NhanVienID,
                        HoVaTenNhanVien = r.NhanVien.HoVaTen,
                        KhachHangID = r.KhachHangID,
                        HoVaTenKhachHang = r.KhachHang.HoVaTen,
                        NgayLap = r.NgayLap,
                        GhiChuHoaDon = r.GhiChuHoaDon,
                        TongTienHoaDon = r.HoaDon_ChiTiet.Sum(ct => ct.SoLuongBan * ct.DonGiaBan),
                        XemChiTiet = "Xem chi tiết"
                    }).ToList();

                dataGridView.DataSource = hd;

                if (hd.Count == 0)
                {
                    MessageBox.Show("Không tìm thấy hóa đơn nào khớp với từ khóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    frmHoaDon_Load(sender, e);
                }
            }
            else
            {
                frmHoaDon_Load(sender, e);
            }
        }

        private string NhapTuKhoaPopUp(string loiNhan, string tieuDe)
        {
            Form prompt = new Form()
            {
                Width = 400,
                Height = 160,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = tieuDe,
                StartPosition = FormStartPosition.CenterScreen,
                MaximizeBox = false
            };
            Label lblText = new Label() { Left = 20, Top = 20, Width = 350, Text = loiNhan };
            TextBox txtInput = new TextBox() { Left = 20, Top = 50, Width = 340 };
            Button btnXacNhan = new Button() { Text = "Tìm kiếm", Left = 260, Width = 100, Top = 80, DialogResult = DialogResult.OK };

            prompt.Controls.Add(lblText); prompt.Controls.Add(txtInput); prompt.Controls.Add(btnXacNhan);
            prompt.AcceptButton = btnXacNhan;

            return prompt.ShowDialog() == DialogResult.OK ? txtInput.Text : "";
        }

        private void dataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView.Columns[e.ColumnIndex].Name == "XemChiTiet" && e.RowIndex >= 0)
            {
                btnSua_Click(sender, e);
            }
        }

        // --- NÚT XUẤT EXCEL ---
        private void btnXuat_Click(object sender, EventArgs e)
        {
            if (dataGridView.Rows.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu để xuất!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Xuất danh sách hóa đơn ra Excel";
            saveFileDialog.Filter = "Excel Files|*.xlsx";
            saveFileDialog.FileName = "DanhSachHoaDon_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    DataTable table = new DataTable();
                    table.Columns.AddRange(new DataColumn[] {
                        new DataColumn("Mã HĐ", typeof(int)),
                        new DataColumn("Nhân Viên", typeof(string)),
                        new DataColumn("Khách Hàng", typeof(string)),
                        new DataColumn("Ngày Lập", typeof(string)),
                        new DataColumn("Ghi Chú", typeof(string)),
                        new DataColumn("Tổng Tiền", typeof(double))
                    });

                    // Lấy dữ liệu MỚI NHẤT từ database để xuất
                    using (var dbExport = new QLQADbcontext())
                    {
                        var dsHoaDon = dbExport.HoaDon
                            .Include(r => r.NhanVien)
                            .Include(r => r.KhachHang)
                            .Include(r => r.HoaDon_ChiTiet)
                            .ToList();

                        foreach (var hd in dsHoaDon)
                        {
                            string tenNV = hd.NhanVien?.HoVaTen ?? "Chưa xác định";
                            string tenKH = hd.KhachHang?.HoVaTen ?? "Khách vãng lai";
                            string ngayLap = hd.NgayLap.ToString("dd/MM/yyyy HH:mm") ?? "";
                            double tongTien = hd.HoaDon_ChiTiet.Sum(ct => (double?)((int)ct.SoLuongBan * ct.DonGiaBan)) ?? 0;

                            table.Rows.Add(hd.ID, tenNV, tenKH, ngayLap, hd.GhiChuHoaDon, tongTien);
                        }
                    }

                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        var sheet = wb.Worksheets.Add(table, "DanhSachHoaDon");
                        sheet.Column(6).Style.NumberFormat.Format = "#,##0";
                        sheet.Columns().AdjustToContents();
                        wb.SaveAs(saveFileDialog.FileName);
                        MessageBox.Show("Xuất danh sách hóa đơn thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi xuất Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}