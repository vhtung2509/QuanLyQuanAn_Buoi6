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
    public partial class frmHoaDon_ChiTiet : Form
    {
        QLQADbcontext context = new QLQADbcontext();
        int id;
        BindingList<DanhSachHoaDon_ChiTiet> hoaDonChiTiet = new BindingList<DanhSachHoaDon_ChiTiet>();
        public frmHoaDon_ChiTiet(int maHoaDon = 0)
        {
            InitializeComponent();
            id = maHoaDon;
        }

        public void LayNhanVienVaoComboBox()
        {
            cboNhanVien.DataSource = context.NhanVien.ToList();
            cboNhanVien.ValueMember = "ID";
            cboNhanVien.DisplayMember = "HoVaTen";
            cboNhanVien.SelectedIndex = -1;
        }

        public void LayKhachHangVaoComboBox()
        {
            cboKhachHang.DataSource = context.KhachHang.ToList();
            cboKhachHang.ValueMember = "ID";
            cboKhachHang.DisplayMember = "HoVaTen";
            cboKhachHang.SelectedIndex = -1;
        }

        public void LayMonAnVaoComboBox()
        {
            cboMonAn.DataSource = context.MonAn.ToList();
            cboMonAn.ValueMember = "ID";
            cboMonAn.DisplayMember = "TenMon";
            cboMonAn.SelectedIndex = -1;
        }

        public void BatTatChucNang()
        {
            btnLuuHoaDon.Enabled = hoaDonChiTiet.Count > 0;
            btnXoa.Enabled = dataGridView.CurrentRow != null;
        }

        private void frmHoaDon_ChiTiet_Load(object sender, EventArgs e)
        {
            LayNhanVienVaoComboBox();
            LayKhachHangVaoComboBox();
            LayMonAnVaoComboBox();

            dataGridView.AutoGenerateColumns = false;
            if (id != 0) // Đã tồn tại chi tiết 
            {
                var hoaDon = context.HoaDon.Where(r => r.ID == id).SingleOrDefault();
                cboNhanVien.SelectedValue = hoaDon.NhanVienID;
                cboKhachHang.SelectedValue = hoaDon.KhachHangID;
                txtGhiChuHoaDon.Text = hoaDon.GhiChuHoaDon;

                // CHÍNH LÀ ĐOẠN NÀY ĐÃ FIX TRIỆT ĐỂ LỖI INT16 BẰNG Convert.ToInt32
                var ct = context.HoaDon_ChiTiet.Where(r => r.HoaDonID == id).Select(r => new DanhSachHoaDon_ChiTiet
                {
                    ID = r.ID,
                    HoaDonID = r.HoaDonID,
                    MonAnID = r.MonAnID,
                    TenMonAn = r.MonAn.TenMon,
                    SoLuongBan = r.SoLuongBan,
                    DonGiaBan = r.DonGiaBan,
                    ThanhTien = Convert.ToInt32(r.SoLuongBan * r.DonGiaBan)
                }).ToList();

                hoaDonChiTiet = new BindingList<DanhSachHoaDon_ChiTiet>(ct);
            }

            dataGridView.DataSource = hoaDonChiTiet;
            BatTatChucNang();
        }
        
        private void btnXacNhanBan_Click(object sender, EventArgs e)
        {
            if (cboMonAn.SelectedIndex == -1) return;
            if (numSoLuongBan.Value <= 0) return;

            int maMon = Convert.ToInt32(cboMonAn.SelectedValue.ToString());
            var itemExisted = hoaDonChiTiet.FirstOrDefault(x => x.MonAnID == maMon);

            if (itemExisted != null)
            {
                // Món đã tồn tại -> Cập nhật số lượng
                itemExisted.SoLuongBan = (short)numSoLuongBan.Value;
                itemExisted.DonGiaBan = (int)numDonGiaBan.Value;
                itemExisted.ThanhTien = itemExisted.SoLuongBan * itemExisted.DonGiaBan;
                dataGridView.Refresh();
            }
            else
            {
                // Món mới -> Thêm dòng mới
                hoaDonChiTiet.Add(new DanhSachHoaDon_ChiTiet
                {
                    ID = 0,
                    HoaDonID = id,
                    MonAnID = maMon,
                    TenMonAn = cboMonAn.Text,
                    SoLuongBan = (short)numSoLuongBan.Value,
                    DonGiaBan = (int)numDonGiaBan.Value,
                    ThanhTien = (int)(numSoLuongBan.Value * numDonGiaBan.Value)
                });
            }
            BatTatChucNang();
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (dataGridView.CurrentRow == null) return;
            var item = (DanhSachHoaDon_ChiTiet)dataGridView.CurrentRow.DataBoundItem;
            hoaDonChiTiet.Remove(item);
            BatTatChucNang();
        }

        private void btnLuuHoaDon_Click(object sender, EventArgs e)
        {
            if (cboNhanVien.SelectedIndex == -1 || cboKhachHang.SelectedIndex == -1)
            {
                MessageBox.Show("Vui lòng chọn đầy đủ Nhân viên và Khách hàng!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (id != 0) // NẾU LÀ SỬA
            {
                var hd = context.HoaDon.Find(id);
                hd.NhanVienID = Convert.ToInt32(cboNhanVien.SelectedValue.ToString());
                hd.KhachHangID = Convert.ToInt32(cboKhachHang.SelectedValue.ToString());
                hd.GhiChuHoaDon = txtGhiChuHoaDon.Text;
                context.SaveChanges();

                var old = context.HoaDon_ChiTiet.Where(r => r.HoaDonID == id).ToList();
                context.HoaDon_ChiTiet.RemoveRange(old);

                foreach (var item in hoaDonChiTiet.ToList())
                {
                    HoaDon_ChiTiet ct = new HoaDon_ChiTiet();
                    ct.HoaDonID = id;
                    ct.MonAnID = item.MonAnID;
                    ct.SoLuongBan = item.SoLuongBan;
                    ct.DonGiaBan = item.DonGiaBan;
                    context.HoaDon_ChiTiet.Add(ct);
                }
                context.SaveChanges();
            }
            else // NẾU LÀ THÊM MỚI
            {
                HoaDon hd = new HoaDon();
                hd.NhanVienID = Convert.ToInt32(cboNhanVien.SelectedValue.ToString());
                hd.KhachHangID = Convert.ToInt32(cboKhachHang.SelectedValue.ToString());
                hd.NgayLap = DateTime.Now;
                hd.GhiChuHoaDon = txtGhiChuHoaDon.Text;
                context.HoaDon.Add(hd);
                context.SaveChanges();

                foreach (var item in hoaDonChiTiet.ToList())
                {
                    HoaDon_ChiTiet ct = new HoaDon_ChiTiet();
                    ct.HoaDonID = hd.ID;
                    ct.MonAnID = item.MonAnID;
                    ct.SoLuongBan = item.SoLuongBan;
                    ct.DonGiaBan = item.DonGiaBan;
                    context.HoaDon_ChiTiet.Add(ct);
                }
                context.SaveChanges();
            }

            MessageBox.Show("Lưu hóa đơn thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }

        private void btnInHoaDon_Click(object sender, EventArgs e)
        {

        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn thoát?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                this.Close();
            }
        }
    }
}