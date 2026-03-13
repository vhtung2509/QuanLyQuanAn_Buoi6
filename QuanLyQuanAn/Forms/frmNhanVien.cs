using ClosedXML.Excel;
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
using BC = BCrypt.Net.BCrypt;

namespace QuanLyQuanAn.Forms
{
    public partial class frmNhanVien : Form
    {
        QLQADbcontext context = new QLQADbcontext();
        bool XuLyThem = false;
        int id;

        public frmNhanVien()
        {
            InitializeComponent();
        }

        private void BatTatChucNang(bool giaTri)
        {
            btnLuu.Enabled = giaTri;
            btnHuyBo.Enabled = giaTri;
            txtHoVaTen.Enabled = giaTri;
            txtDienThoai.Enabled = giaTri;
            txtDiaChi.Enabled = giaTri;
            txtTenDangNhap.Enabled = giaTri;
            txtMatKhau.Enabled = giaTri;
            cboQuyenHan.Enabled = giaTri;

            btnThem.Enabled = !giaTri;
            btnSua.Enabled = !giaTri;
            btnXoa.Enabled = !giaTri;
            btnTimKiem.Enabled = !giaTri;
            btnNhap.Enabled = !giaTri;
            btnXuat.Enabled = !giaTri;
        }

        private void frmNhanVien_Load(object sender, EventArgs e)
        {
            BatTatChucNang(false);
            dataGridView.AutoGenerateColumns = false;

            // Lấy dữ liệu từ Database
            List<NhanVien> nv = context.NhanVien.ToList();
            BindingSource bindingSource = new BindingSource();
            bindingSource.DataSource = nv;

            // BINDING DỮ LIỆU LÊN TEXTBOX
            txtHoVaTen.DataBindings.Clear();
            txtHoVaTen.DataBindings.Add("Text", bindingSource, "HoVaTen", false, DataSourceUpdateMode.Never);

            txtDienThoai.DataBindings.Clear();
            txtDienThoai.DataBindings.Add("Text", bindingSource, "DienThoai", false, DataSourceUpdateMode.Never);

            // FIX LỖI 1: Sửa "DienThoai" thành "DiaChi" để hiện đúng địa chỉ
            txtDiaChi.DataBindings.Clear();
            txtDiaChi.DataBindings.Add("Text", bindingSource, "DiaChi", false, DataSourceUpdateMode.Never);

            txtTenDangNhap.DataBindings.Clear();
            txtTenDangNhap.DataBindings.Add("Text", bindingSource, "TenDangNhap", false, DataSourceUpdateMode.Never);

            // ĐỒNG BỘ COMBOBOX QUYỀN HẠN KHI CHỌN DÒNG TRÊN LƯỚI
            bindingSource.PositionChanged += (s, ev) => {
                if (bindingSource.Current is NhanVien item)
                {
                    cboQuyenHan.SelectedIndex = item.Quyen == true ? 0 : 1;
                }
            };

            // FIX LỖI 2: HIỂN THỊ CHỮ QUẢN LÝ/NHÂN VIÊN TRÊN LƯỚI (DataPropertyName = QuyenHan)
            dataGridView.CellFormatting += (s, ev) => {
                // Kiểm tra nếu đang vẽ cột Quyền hạn (dựa vào DataPropertyName bạn đặt là QuyenHan)
                if (dataGridView.Columns[ev.ColumnIndex].DataPropertyName == "QuyenHan")
                {
                    var row = dataGridView.Rows[ev.RowIndex].DataBoundItem as NhanVien;
                    if (row != null)
                    {
                        ev.Value = row.Quyen == true ? "Quản lý" : "Nhân viên";
                        ev.FormattingApplied = true;
                    }
                }
            };

            dataGridView.DataSource = bindingSource;
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            XuLyThem = true;
            BatTatChucNang(true);
            txtHoVaTen.Clear();
            txtDienThoai.Clear();
            txtDiaChi.Clear();
            txtTenDangNhap.Clear();
            txtMatKhau.Clear();
            cboQuyenHan.SelectedIndex = 1; // Mặc định là nhân viên
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            if (dataGridView.CurrentRow == null) return;
            XuLyThem = false;
            BatTatChucNang(true);
            id = Convert.ToInt32(dataGridView.CurrentRow.Cells["ID"].Value.ToString());
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (dataGridView.CurrentRow == null) return;
            if (MessageBox.Show("Xác nhận xóa nhân viên " + txtHoVaTen.Text + "?", "Xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                id = Convert.ToInt32(dataGridView.CurrentRow.Cells["ID"].Value.ToString());
                NhanVien nv = context.NhanVien.Find(id);
                if (nv != null) context.NhanVien.Remove(nv);
                context.SaveChanges();
                frmNhanVien_Load(sender, e);
            }
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtHoVaTen.Text) || string.IsNullOrWhiteSpace(txtTenDangNhap.Text))
            {
                MessageBox.Show("Vui lòng nhập đủ thông tin!", "Lỗi");
                return;
            }

            try
            {
                if (XuLyThem)
                {
                    if (string.IsNullOrWhiteSpace(txtMatKhau.Text))
                    {
                        MessageBox.Show("Vui lòng nhập mật khẩu!", "Lỗi");
                        return;
                    }
                    NhanVien nv = new NhanVien
                    {
                        HoVaTen = txtHoVaTen.Text,
                        DienThoai = txtDienThoai.Text,
                        DiaChi = txtDiaChi.Text,
                        TenDangNhap = txtTenDangNhap.Text,
                        MatKhau = BC.HashPassword(txtMatKhau.Text),
                        Quyen = cboQuyenHan.SelectedIndex == 0
                    };
                    context.NhanVien.Add(nv);
                }
                else
                {
                    NhanVien nv = context.NhanVien.Find(id);
                    if (nv != null)
                    {
                        nv.HoVaTen = txtHoVaTen.Text;
                        nv.DienThoai = txtDienThoai.Text;
                        nv.DiaChi = txtDiaChi.Text;
                        nv.TenDangNhap = txtTenDangNhap.Text;
                        nv.Quyen = cboQuyenHan.SelectedIndex == 0;
                        if (!string.IsNullOrEmpty(txtMatKhau.Text))
                            nv.MatKhau = BC.HashPassword(txtMatKhau.Text);
                        context.NhanVien.Update(nv);
                    }
                }
                context.SaveChanges();
                MessageBox.Show("Lưu thành công!");
                frmNhanVien_Load(sender, e);
            }
            catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message); }
        }

        private void btnNhap_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Nhập dữ liệu từ tập tin Excel";
            openFileDialog.Filter = "Tập tin Excel|*.xls;*.xlsx";
            openFileDialog.Multiselect = false;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    DataTable table = new DataTable();
                    using (XLWorkbook workbook = new XLWorkbook(openFileDialog.FileName))
                    {
                        IXLWorksheet worksheet = workbook.Worksheet(1);
                        bool firstRow = true;
                        string readRange = "1:1";

                        foreach (IXLRow row in worksheet.RowsUsed())
                        {
                            // Bước 1: Đọc dòng tiêu đề (dòng đầu tiên) để tạo cột cho DataTable
                            if (firstRow)
                            {
                                readRange = string.Format("{0}:{1}", 1, row.LastCellUsed().Address.ColumnNumber);
                                foreach (IXLCell cell in row.Cells(readRange))
                                    table.Columns.Add(cell.Value.ToString());
                                firstRow = false;
                            }
                            else // Bước 2: Đọc các dòng nội dung đưa vào DataTable
                            {
                                table.Rows.Add();
                                int cellIndex = 0;
                                foreach (IXLCell cell in row.Cells(readRange))
                                {
                                    table.Rows[table.Rows.Count - 1][cellIndex] = cell.Value.ToString();
                                    cellIndex++;
                                }
                            }
                        }

                        // Bước 3: Đưa dữ liệu từ DataTable vào Database
                        if (table.Rows.Count > 0)
                        {
                            foreach (DataRow r in table.Rows)
                            {
                                NhanVien nv = new NhanVien();
                                // Lưu ý: Tên trong ngoặc vuông [] phải khớp 100% với tên cột trong file Excel
                                nv.HoVaTen = r["HoVaTen"].ToString();
                                nv.DienThoai = r["DienThoai"].ToString();
                                nv.DiaChi = r["DiaChi"].ToString();

                                context.NhanVien.Add(nv);
                            }
                            context.SaveChanges();

                            MessageBox.Show("Đã nhập thành công " + table.Rows.Count + " nhân viên.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }

                        if (firstRow)
                            MessageBox.Show("Tập tin Excel rỗng.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
        }

        private void btnXuat_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Xuất danh sách nhân viên ra Excel";
            saveFileDialog.Filter = "Excel Files|*.xlsx";
            saveFileDialog.FileName = "DanhSachNhanVien_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // Tạo DataTable để chứa dữ liệu xuất
                    DataTable table = new DataTable();
                    table.Columns.AddRange(new DataColumn[] {
                    new DataColumn("ID", typeof(int)),
                    new DataColumn("HoVaTen", typeof(string)),
                    new DataColumn("DienThoai", typeof(string)),
                    new DataColumn("DiaChi", typeof(string))
                    });

                    // Lấy dữ liệu từ database
                    var danhSach = context.NhanVien.ToList();
                    foreach (var nv in danhSach)
                    {
                        // Đã loại bỏ nv.Email
                        table.Rows.Add(nv.ID, nv.HoVaTen, nv.DienThoai, nv.DiaChi);
                    }

                    // Sử dụng ClosedXML để ghi file
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        var sheet = wb.Worksheets.Add(table, "NhanVien");
                        sheet.Columns().AdjustToContents(); // Tự động căn chỉnh độ rộng cột
                        wb.SaveAs(saveFileDialog.FileName);

                        MessageBox.Show("Xuất dữ liệu thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi xuất Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnHuyBo_Click(object sender, EventArgs e)
        {
            frmNhanVien_Load(sender, e);
        }
        private void btnThoat_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có chắc chắn muốn thoát không?", "Xác nhận", MessageBoxButtons.YesNo) == DialogResult.Yes)
                this.Close();
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            string tuKhoa = txtHoVaTen.Text.Trim().ToLower();
            var ketQua = context.NhanVien
                .Where(nv => nv.HoVaTen.ToLower().Contains(tuKhoa) || nv.DienThoai.Contains(tuKhoa)).ToList();

            BindingSource bs = new BindingSource();
            bs.DataSource = ketQua;
            dataGridView.DataSource = bs;
        }
    }
}