using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace QuanLyQuanAn.Data
{
    internal class MonAn
    {
        public int ID { get; set; }
        public int LoaiMonAnID { get; set; }
        public string TenMon { get; set; }
        public int DonGia { get; set; }
        public string? MoTa { get; set; }
        public string? HinhAnh { get; set; }
        public int SoLuong { get; set; }

        public virtual LoaiMonAn LoaiMonAn { get; set; } = null!;

        [NotMapped]
        public class DanhSachMonAn
        {
            public int ID { get; set; }
            public int LoaiMonAnID { get; set; }
            public string TenLoai { get; set; }
            public string TenMonAn { get; set; }
            public int? DonGia { get; set; }
            public int? SoLuong { get; set; }
            public string? MoTa { get; set; }
            public string? HinhAnh { get; set; }
        }
    }
}