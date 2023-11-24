using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp3
{
    public partial class TimKiemGiaoVien : Form
    {
        ProcessDatabase db = new ProcessDatabase();
        public TimKiemGiaoVien()
        {
            InitializeComponent();
        }

        private void cbbTrinhDo_Click(object sender, EventArgs e)
        {
            cbbTrinhDo.DataSource = db.DocBang("Select * from TrinhDo");
            cbbTrinhDo.ValueMember = "MaTrinhDo";
            cbbTrinhDo.DisplayMember = "MaTrinhDo";
        }

        private void cbbChuyenMon_Click(object sender, EventArgs e)
        {
            cbbChuyenMon.DataSource = db.DocBang("Select * from ChuyenMon");
            cbbChuyenMon.ValueMember = "MaChuyenMon";
            cbbChuyenMon.DisplayMember = "MaChuyenMon";
        }

        private void cbbMonDay_Click(object sender, EventArgs e)
        {
            cbbMonDay.DataSource = db.DocBang("Select * from MonHoc");
            cbbMonDay.ValueMember = "MaMon";
            cbbMonDay.DisplayMember = "MaMon";
        }
        public bool ktraDL()
        {
            bool k = true;
            if (txtTenGV.Text.Trim().Equals("") || cbbTrinhDo.Text.Trim().Equals("") || cbbChuyenMon.Text.Trim().Equals("")
                || cbbMonDay.Text.Trim().Equals(""))
            {
                MessageBox.Show("Hãy nhập đủ dữ liệu");
                k = false;
            }
            return k;
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            if (ktraDL())
            {
                string sql = "SELECT GV.MaGV, TenGV, GioiTinh, DiaChi, DienThoai, MaMon, XepLoai FROM GiaoVien GV JOIN GiaoVienMonHoc GVMH ON GVMH.MaGV = GV.MaGV WHERE TenGV =N'" +
                txtTenGV.Text + "' AND MaTrinhDo ='" + cbbTrinhDo.Text + "' AND MaChuyenMon ='" + cbbChuyenMon.Text + "' " + "AND MaMon ='" + cbbMonDay.Text + "'";
                dataGridView.DataSource = db.DocBang(sql);
                dataGridView.Columns[0].HeaderText = "Mã GV";
                dataGridView.Columns[1].HeaderText = "Tên GV";
                dataGridView.Columns[2].HeaderText = "Giới tính";
                dataGridView.Columns[3].HeaderText = "Địa chỉ";
                dataGridView.Columns[4].HeaderText = "Điện thoại";
                dataGridView.Columns[5].HeaderText = "Mã môn";
                dataGridView.Columns[6].HeaderText = "Xếp loại";
            }    
        }

        private void btnLamMoi_Click(object sender, EventArgs e)
        {
            txtTenGV.Text = "";
            cbbChuyenMon.Text = "";
            cbbTrinhDo.Text = "";
            cbbMonDay.Text = "";
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn thoát không?", "Thông báo",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                this.Close();
        }
    }
}
