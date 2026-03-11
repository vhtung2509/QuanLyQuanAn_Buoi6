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
using static QuanLyQuanAn.Data.MonAn;
using Microsoft.EntityFrameworkCore;
using System.ComponentModel.DataAnnotations.Schema;

namespace QuanLyQuanAn.Forms
{
    public partial class frmMonAn : Form
    {
        QLQADbcontext context = new QLQADbcontext();
        int id;
        bool xuLyThem = false;
        string imagesFolder = Application.StartupPath.Replace("bin\\Debug\\net8.0-windows", "Images");
        public frmMonAn()
        {
            InitializeComponent();
        }

        private void BatTatChucNang(bool giaTri)
        {
            btnLuu.Enabled = giaTri;
            btnHuyBo.Enabled = giaTri;

            cboLoaiMonAn.Enabled = giaTri;
            txtTenMonAn.Enabled = giaTri;
            numSoLuong.Enabled = giaTri;
            numDonGia.Enabled = giaTri;
            txtMoTa.Enabled = giaTri;
            picHinhAnh.Enabled = giaTri;

            btnThem.Enabled = !giaTri;
            btnSua.Enabled = !giaTri;
            btnXoa.Enabled = !giaTri;

            btnDoiAnh.Enabled = giaTri;
            btnXoayAnh.Enabled = giaTri;


            btnTimKiem.Enabled = !giaTri;
            btnNhap.Enabled = !giaTri;
            btnXuat.Enabled = !giaTri;
        }

        public void LayLoaiMonAnVaoComboBox()
        {
            cboLoaiMonAn.DataSource = context.LoaiMonAn.ToList();
            cboLoaiMonAn.ValueMember = "ID";
            cboLoaiMonAn.DisplayMember = "TenLoai";
        }

        private void frmMonAn_Load(object sender, EventArgs e)
        {
            BatTatChucNang(false);
            LayLoaiMonAnVaoComboBox();

            dataGridView.AutoGenerateColumns = false;

            // Load danh sách món ăn từ DB lên ViewModel
            List<DanhSachMonAn> sp = new List<DanhSachMonAn>();
            sp = context.MonAn.Select(r => new DanhSachMonAn
            {
                ID = r.ID,
                LoaiMonAnID = r.LoaiMonAnID,
                TenLoai = r.LoaiMonAn.TenLoai,
                TenMonAn = r.TenMon,
                SoLuong = r.SoLuong,
                DonGia = r.DonGia,
                MoTa = r.MoTa,
                HinhAnh = r.HinhAnh
            }).ToList();

            BindingSource bindingSource = new BindingSource();
            bindingSource.DataSource = sp;

            // Clear và Add Binding giống y hệt bài mẫu
            cboLoaiMonAn.DataBindings.Clear();
            cboLoaiMonAn.DataBindings.Add("SelectedValue", bindingSource, "LoaiMonAnID", false, DataSourceUpdateMode.Never);

            txtTenMonAn.DataBindings.Clear();
            txtTenMonAn.DataBindings.Add("Text", bindingSource, "TenMonAn", false, DataSourceUpdateMode.Never);

            numSoLuong.DataBindings.Clear();
            numSoLuong.DataBindings.Add("Value", bindingSource, "SoLuong", false, DataSourceUpdateMode.Never);

            numDonGia.DataBindings.Clear();
            numDonGia.DataBindings.Add("Value", bindingSource, "DonGia", false, DataSourceUpdateMode.Never);

            txtMoTa.DataBindings.Clear();
            txtMoTa.DataBindings.Add("Text", bindingSource, "MoTa", false, DataSourceUpdateMode.Never);

            picHinhAnh.DataBindings.Clear();
            Binding hinhAnh = new Binding("Tag", bindingSource, "HinhAnh");
            hinhAnh.Format += (s, ev) =>
            {
                if (ev.Value != null && !string.IsNullOrEmpty(ev.Value.ToString()))
                {
                    string fullPath = Path.Combine(imagesFolder, ev.Value.ToString());
                    if (File.Exists(fullPath))
                    {
                        // Dùng FileStream để không khóa file ảnh (giúp nút Xoay Ảnh hoạt động được)
                        using (FileStream fs = new FileStream(fullPath, FileMode.Open, FileAccess.Read))
                        {
                            picHinhAnh.Image = Image.FromStream(fs);
                        }
                    }
                    else picHinhAnh.Image = null;
                }
                else picHinhAnh.Image = null;
            };
            picHinhAnh.DataBindings.Add(hinhAnh);

            dataGridView.DataSource = bindingSource;
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            xuLyThem = true;
            BatTatChucNang(true); // Mở khóa các ô nhập liệu

            // Xóa trắng các ô nhập liệu để chuẩn bị thêm mới
            cboLoaiMonAn.SelectedIndex = -1;
            txtTenMonAn.Clear();
            txtMoTa.Clear();
            numSoLuong.Value = 0;
            numDonGia.Value = 0;
            picHinhAnh.Image = null;
            picHinhAnh.Tag = null; // Cực kỳ quan trọng để reset tên ảnh
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            // Kiểm tra xem người dùng đã chọn dòng nào trên lưới chưa
            if (dataGridView.CurrentRow == null) return;

            xuLyThem = false;
            BatTatChucNang(true);

            // Lấy ID của dòng đang chọn (Lưu ý: Đảm bảo cột ID trên lưới của bạn Name là "ID")
            id = Convert.ToInt32(dataGridView.CurrentRow.Cells["ID"].Value);
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (dataGridView.CurrentRow == null) return;

            if (MessageBox.Show("Xác nhận xóa món ăn: " + txtTenMonAn.Text + "?", "Cảnh báo Xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                try
                {
                    id = Convert.ToInt32(dataGridView.CurrentRow.Cells["ID"].Value);
                    MonAn sp = context.MonAn.Find(id);
                    if (sp != null)
                    {
                        context.MonAn.Remove(sp);
                        context.SaveChanges();
                    }
                    frmMonAn_Load(sender, e); // Load lại Form
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi xóa: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnHuyBo_Click(object sender, EventArgs e)
        {
            frmMonAn_Load(sender, e); // Load lại Form để reset về trạng thái ban đầu
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            // 1. Bắt lỗi người dùng nhập thiếu
            if (string.IsNullOrWhiteSpace(cboLoaiMonAn.Text))
                MessageBox.Show("Vui lòng chọn Loại món ăn.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else if (string.IsNullOrWhiteSpace(txtTenMonAn.Text))
                MessageBox.Show("Vui lòng nhập Tên món ăn.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else if (numDonGia.Value <= 0)
                MessageBox.Show("Đơn giá phải lớn hơn 0.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                try
                {
                    if (xuLyThem) // NẾU ĐANG BẤM NÚT THÊM
                    {
                        MonAn sp = new MonAn();
                        sp.TenMon = txtTenMonAn.Text;
                        sp.LoaiMonAnID = Convert.ToInt32(cboLoaiMonAn.SelectedValue);
                        sp.SoLuong = Convert.ToInt32(numSoLuong.Value);
                        sp.DonGia = Convert.ToInt32(numDonGia.Value);
                        sp.MoTa = txtMoTa.Text;

                        // Lấy tên file ảnh đã được lưu tạm trong Tag
                        if (picHinhAnh.Tag != null) sp.HinhAnh = picHinhAnh.Tag.ToString();

                        context.MonAn.Add(sp);
                    }
                    else // NẾU ĐANG BẤM NÚT SỬA
                    {
                        MonAn sp = context.MonAn.Find(id);
                        if (sp != null)
                        {
                            sp.TenMon = txtTenMonAn.Text;
                            sp.LoaiMonAnID = Convert.ToInt32(cboLoaiMonAn.SelectedValue);
                            sp.SoLuong = Convert.ToInt32(numSoLuong.Value);
                            sp.DonGia = Convert.ToInt32(numDonGia.Value);
                            sp.MoTa = txtMoTa.Text;

                            if (picHinhAnh.Tag != null) sp.HinhAnh = picHinhAnh.Tag.ToString();
                        }
                    }

                    context.SaveChanges(); // Đẩy xuống SQL Server
                    MessageBox.Show("Lưu dữ liệu thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    frmMonAn_Load(sender, e); // Load lại lưới cho đẹp
                }
                catch (Exception ex)
                {
                    string loi = ex.InnerException != null ? ex.InnerException.Message : ex.Message;
                    MessageBox.Show("Lỗi CSDL: " + loi, "Lỗi Lưu", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn thoát?", "Xác nhận Thoát", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                this.Close();
            }
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            // 1. Tạo một hộp thoại pop-up siêu nhanh bằng code
            Form prompt = new Form()
            {
                Width = 400,
                Height = 160,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Tìm kiếm món ăn",
                StartPosition = FormStartPosition.CenterScreen,
                MaximizeBox = false,
                MinimizeBox = false
            };
            Label textLabel = new Label() { Left = 20, Top = 20, Text = "Nhập tên món ăn cần tìm:" };
            TextBox textBox = new TextBox() { Left = 20, Top = 50, Width = 340 };
            Button confirmation = new Button() { Text = "Tìm kiếm", Left = 260, Width = 100, Top = 80, DialogResult = DialogResult.OK };

            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(textLabel);
            prompt.AcceptButton = confirmation; // Nhấn Enter là tự bấm nút Tìm

            // 2. Hiển thị hộp thoại và chờ người dùng nhập
            if (prompt.ShowDialog() == DialogResult.OK)
            {
                string tuKhoa = textBox.Text.Trim().ToLower();

                // 3. Lọc dữ liệu từ Database chứa từ khóa (không phân biệt hoa thường)
                List<DanhSachMonAn> sp = context.MonAn
                    .Where(m => m.TenMon.ToLower().Contains(tuKhoa))
                    .Select(r => new DanhSachMonAn
                    {
                        ID = r.ID,
                        LoaiMonAnID = r.LoaiMonAnID,
                        TenLoai = r.LoaiMonAn.TenLoai,
                        TenMonAn = r.TenMon,
                        SoLuong = r.SoLuong,
                        DonGia = r.DonGia,
                        MoTa = r.MoTa,
                        HinhAnh = r.HinhAnh
                    }).ToList();

                // 4. Đẩy kết quả tìm được lên lưới DataGridView
                BindingSource bindingSource = new BindingSource();
                bindingSource.DataSource = sp;
                dataGridView.DataSource = bindingSource;

                // 5. Nếu danh sách rỗng (không tìm thấy)
                if (sp.Count == 0)
                {
                    MessageBox.Show("Không tìm thấy món ăn nào có chứa chữ: " + textBox.Text, "Kết quả tìm kiếm", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    frmMonAn_Load(sender, e); // Load lại toàn bộ danh sách
                }
            }
        }

        private void btnDoiAnh_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Chọn hình ảnh món ăn";
            openFileDialog.Filter = "Tập tin hình ảnh|*.jpg;*.jpeg;*.png;*.gif;*.bmp";
            openFileDialog.Multiselect = false;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string sourcePath = openFileDialog.FileName;
                string fileName = Path.GetFileNameWithoutExtension(sourcePath);
                string ext = Path.GetExtension(sourcePath);

                string newFileName = fileName.GenerateSlug() + ext;
                string fileSavePath = Path.Combine(imagesFolder, newFileName);

                if (!Directory.Exists(imagesFolder)) Directory.CreateDirectory(imagesFolder);

                try
                {
                    // 1. MỞ KHÓA: Hủy ảnh cũ đang hiển thị trên PictureBox (nếu có)
                    if (picHinhAnh.Image != null)
                    {
                        picHinhAnh.Image.Dispose();
                        picHinhAnh.Image = null;
                    }
                    // Dọn rác bộ nhớ ngay lập tức để chắc chắn nhả khóa file
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    // 2. CHỐNG LỖI CHỌN TRÙNG: Chỉ copy nếu file nguồn và đích khác nhau
                    if (!sourcePath.Equals(fileSavePath, StringComparison.OrdinalIgnoreCase))
                    {
                        File.Copy(sourcePath, fileSavePath, true);
                    }

                    // 3. LOAD ẢNH SIÊU AN TOÀN (Tạo bản sao lưu trên RAM)
                    using (FileStream fs = new FileStream(fileSavePath, FileMode.Open, FileAccess.Read))
                    {
                        using (Image tempImg = Image.FromStream(fs))
                        {
                            picHinhAnh.Image = new Bitmap(tempImg); // Cắt đứt hoàn toàn với file gốc
                        }
                    }

                    picHinhAnh.Tag = newFileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi đổi ảnh: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnXoayAnh_Click(object sender, EventArgs e)
        {
            if (picHinhAnh.Image != null && picHinhAnh.Tag != null)
            {
                // 1. Xoay hình 90 độ theo chiều kim đồng hồ ngay trên màn hình
                picHinhAnh.Image.RotateFlip(RotateFlipType.Rotate90FlipNone);
                picHinhAnh.Refresh();

                // 2. Lưu lại file ảnh thật trong thư mục để nó nhớ góc xoay
                try
                {
                    string filePath = Path.Combine(imagesFolder, picHinhAnh.Tag.ToString());

                    // Nếu file tồn tại, ta xóa nó đi rồi lưu bản đã xoay xuống
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }
                    picHinhAnh.Image.Save(filePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi lưu ảnh xoay: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void dataGridView_CellFormatting_1(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (dataGridView.Columns[e.ColumnIndex].Name == "HinhAnh" && e.Value != null)
            {
                try
                {
                    string path = Path.Combine(imagesFolder, e.Value.ToString());
                    if (File.Exists(path))
                    {
                        // LOAD ẢNH LÊN LƯỚI KHÔNG KHÓA FILE
                        using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read))
                        {
                            using (Image tempImg = Image.FromStream(fs))
                            {
                                // Thu nhỏ ảnh và nhân bản vào RAM
                                e.Value = new Bitmap(tempImg, 40, 40);
                            }
                        }
                    }
                    else e.Value = null;
                }
                catch { e.Value = null; }
            }
        }
    }
    public static class StringExtensions
    {
        // 1. Hàm tạo Slug (Dùng cho tên file ảnh)
        public static string GenerateSlug(this string phrase)
        {
            if (string.IsNullOrEmpty(phrase)) return "";

            string str = phrase.ToLower();

            // Xử lý riêng chữ đ/Đ
            str = str.Replace("đ", "d").Replace("Đ", "d");

            // Bỏ dấu tiếng việt
            str = str.Normalize(System.Text.NormalizationForm.FormD);
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            foreach (char c in str)
            {
                if (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c) != System.Globalization.UnicodeCategory.NonSpacingMark)
                    sb.Append(c);
            }
            str = sb.ToString().Normalize(System.Text.NormalizationForm.FormC);

            // Thay thế các ký tự đặc biệt, giữ lại chữ cái và số
            str = System.Text.RegularExpressions.Regex.Replace(str, @"[^a-z0-9\s-]", "");

            // Thay nhiều khoảng trắng thành 1 khoảng trắng
            str = System.Text.RegularExpressions.Regex.Replace(str, @"\s+", " ").Trim();

            // Đổi khoảng trắng thành dấu gạch ngang
            str = str.Replace(" ", "-");

            return str;
        }

        // 2. Hàm bỏ dấu Tiếng Việt (Dùng cho nút Lưu Tên món ăn nếu bạn cần)
        public static string ChuyenTiengVietKhongDau(this string str)
        {
            if (string.IsNullOrEmpty(str)) return str;
            str = str.Replace("đ", "d").Replace("Đ", "D");
            str = str.Normalize(System.Text.NormalizationForm.FormD);
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            foreach (char c in str)
            {
                if (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c) != System.Globalization.UnicodeCategory.NonSpacingMark)
                    sb.Append(c);
            }
            return sb.ToString().Normalize(System.Text.NormalizationForm.FormC);
        }
    }
}