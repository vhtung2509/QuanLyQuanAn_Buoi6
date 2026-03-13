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
using ClosedXML.Excel;

namespace QuanLyQuanAn.Forms
{
    public partial class frmLoaiMonAn : Form
    {
        QLQADbcontext context = new QLQADbcontext(); // Khởi tạo biến ngữ cảnh CSDL
        bool xuLyThem = false; // Kiểm tra có nhấn vào nút Thêm hay không?
        int id; // Lấy mã loại món ăn (dùng cho Sửa và Xóa)
        public frmLoaiMonAn()
        {
            InitializeComponent();
        }
        private void BatTatChucNang(bool giaTri)
        {
            btnLuu.Enabled = giaTri;
            btnHuyBo.Enabled = giaTri;
            txtTenLoai.Enabled = giaTri;

            btnThem.Enabled = !giaTri;
            btnSua.Enabled = !giaTri;
            btnXoa.Enabled = !giaTri;
        }

        private void frmLoaiMonAn_Load(object sender, EventArgs e)
        {
            BatTatChucNang(false);

            List<LoaiMonAn> lma = new List<LoaiMonAn>();
            lma = context.LoaiMonAn.ToList();

            BindingSource bindingSource = new BindingSource();
            bindingSource.DataSource = lma;

            txtTenLoai.DataBindings.Clear();
            // Tên thuộc tính vẫn là "TenLoai" giống trong class LoaiMonAn của bạn
            txtTenLoai.DataBindings.Add("Text", bindingSource, "TenLoai", false, DataSourceUpdateMode.Never);

            dataGridView.DataSource = bindingSource;
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            xuLyThem = true;
            BatTatChucNang(true);
            txtTenLoai.Clear();
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            xuLyThem = false;
            BatTatChucNang(true);
            id = Convert.ToInt32(dataGridView.CurrentRow.Cells["ID"].Value.ToString());
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Xác nhận xóa loại món ăn này?", "Xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                id = Convert.ToInt32(dataGridView.CurrentRow.Cells["ID"].Value.ToString());

                LoaiMonAn lma = context.LoaiMonAn.Find(id);

                if (lma != null)
                {
                    context.LoaiMonAn.Remove(lma);
                }
                context.SaveChanges();

                frmLoaiMonAn_Load(sender, e);
            }
        }

        private void btnHuyBo_Click(object sender, EventArgs e)
        {
            frmLoaiMonAn_Load(sender, e);
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn thoát không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                this.Close();
            }
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtTenLoai.Text))
                MessageBox.Show("Vui lòng nhập tên loại món ăn?", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                if (xuLyThem)
                {
                    LoaiMonAn lma = new LoaiMonAn();
                    lma.TenLoai = txtTenLoai.Text;
                    context.LoaiMonAn.Add(lma);

                    context.SaveChanges();
                }
                else
                {
                    LoaiMonAn lma = context.LoaiMonAn.Find(id);
                    if (lma != null)
                    {
                        lma.TenLoai = txtTenLoai.Text;
                        context.LoaiMonAn.Update(lma);

                        context.SaveChanges();
                    }
                }
                frmLoaiMonAn_Load(sender, e);
            }
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
                            if (firstRow)
                            {
                                readRange = string.Format("{0}:{1}", 1, row.LastCellUsed().Address.ColumnNumber);
                                foreach (IXLCell cell in row.Cells(readRange))
                                    table.Columns.Add(cell.Value.ToString().Trim());
                                firstRow = false;
                            }
                            else
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

                        if (table.Rows.Count > 0)
                        {
                            int dem = 0;
                            foreach (DataRow r in table.Rows)
                            {
                                LoaiMonAn lma = new LoaiMonAn();
                                // Lấy dữ liệu từ cột có chữ "TenLoai" trong Excel
                                lma.TenLoai = table.Columns.Contains("TenLoai") ? r["TenLoai"].ToString() : "";

                                // Chỉ lưu nếu tên loại không bị rỗng
                                if (!string.IsNullOrWhiteSpace(lma.TenLoai))
                                {
                                    context.LoaiMonAn.Add(lma);
                                    dem++;
                                }
                            }
                            context.SaveChanges();

                            MessageBox.Show("Đã nhập thành công " + dem + " loại món ăn.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            frmLoaiMonAn_Load(sender, e);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi nhập Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
        }

        private void btnXuat_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Xuất danh sách loại món ăn ra Excel";
            saveFileDialog.Filter = "Excel Files|*.xlsx";
            saveFileDialog.FileName = "DanhSachLoaiMonAn_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // Tạo DataTable để chứa dữ liệu xuất
                    DataTable table = new DataTable();
                    table.Columns.AddRange(new DataColumn[] {
                        new DataColumn("ID", typeof(int)),
                        new DataColumn("TenLoai", typeof(string))
                    });

                    // Lấy dữ liệu từ database
                    var danhSach = context.LoaiMonAn.ToList();
                    foreach (var lma in danhSach)
                    {
                        table.Rows.Add(lma.ID, lma.TenLoai);
                    }

                    // Sử dụng ClosedXML để ghi file
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        var sheet = wb.Worksheets.Add(table, "LoaiMonAn");
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
    }
}