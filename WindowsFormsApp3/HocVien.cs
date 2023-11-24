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
    public partial class HocVien : Form
    {
        ProcessDatabase db = new ProcessDatabase();
        public HocVien()
        {
            InitializeComponent();
        }

        private void HocVien_Load(object sender, EventArgs e)
        {
            dataGridView.DataSource = db.DocBang("select * from HocVien");
            dataGridView.Columns[0].HeaderText = "Mã học viên";
            dataGridView.Columns[1].HeaderText = "Tên học viên";
            dataGridView.Columns[2].HeaderText = "Mã lớp";
            dataGridView.Columns[3].HeaderText = "Ngày sinh";
            dataGridView.Columns[4].HeaderText = "Giới tính";
            dataGridView.Columns[5].HeaderText = "Địa chỉ";
            dataGridView.Columns[6].HeaderText = "Mã nghề";
            dataGridView.Columns[7].HeaderText = "Điện thoại";
            dataGridView.Columns[8].HeaderText = "Ngày nộp học phí";
            dataGridView.Columns[9].HeaderText = "Điểm";
            dataGridView.Columns[10].HeaderText = "Mã xếp loại";
        }

        private void cbbMaLop_Click(object sender, EventArgs e)
        {
            cbbMaLop.DataSource = db.DocBang("Select * from LopHoc");
            cbbMaLop.ValueMember = "MaLop";
            cbbMaLop.DisplayMember = "MaLop";
        }

        private void cbbMaNghe_Click(object sender, EventArgs e)
        {
            cbbMaNghe.DataSource = db.DocBang("Select * from NgheNghiep");
            cbbMaNghe.ValueMember = "MaNghe";
            cbbMaNghe.DisplayMember = "MaNghe";
        }

        private void cbbMaXepLoai_Click(object sender, EventArgs e)
        {
            cbbMaXepLoai.DataSource = db.DocBang("Select * from XepLoai");
            cbbMaXepLoai.ValueMember = "MaXepLoai";
            cbbMaXepLoai.DisplayMember = "MaXepLoai";
        }
        public bool ktraDL()
        {
            bool k = true;
            if (txtMaHocVien.Text.Trim().Equals("") || txtTenHocVien.Text.Trim().Equals("") || cbbMaLop.Text.Trim().Equals("") || dtpNgaySinh.Text.Trim().Equals("")
                || cbbGioiTinh.Text.Trim().Equals("") || txtDiaChi.Text.Trim().Equals("") || cbbMaNghe.Text.Trim().Equals("") || txtDienThoai.Text.Trim().Equals("")
                || dtpNgayNopHocPhi.Text.Trim().Equals("") || txtDiem.Text.Trim().Equals("") || cbbMaXepLoai.Text.Trim().Equals(""))
            {
                MessageBox.Show("Hãy nhập đủ dữ liệu");
                k = false;
            }
            return k;
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            if (ktraDL())
            {
                string sql;
                sql = "SELECT MaHocVien FROM HocVien WHERE MaHocVien='" + txtMaHocVien.Text + "'";
                SqlDataAdapter MyData = new SqlDataAdapter(sql, db.Con);
                DataTable table = new DataTable();
                MyData.Fill(table);
                if (table.Rows.Count > 0)
                {
                    MessageBox.Show("Mã học viên này đã tồn tại");
                    return;
                }
                //Thực hiện chèn thêm mới
                sql = "INSERT INTO HocVien VALUES ('" + txtMaHocVien.Text + "',N'" + txtTenHocVien.Text + "','" +
                cbbMaLop.Text + "','" + dtpNgaySinh.Text + "',N'" + cbbGioiTinh.Text + "',N'" + txtDiaChi.Text +
                "','" + cbbMaNghe.Text + "','" + txtDienThoai.Text + "','" + dtpNgayNopHocPhi.Text + "'," + float.Parse(txtDiem.Text) +
                ",'" + cbbMaXepLoai.Text + "')";
                db.CapNhat(sql);
                dataGridView.DataSource = db.DocBang("select * from HocVien");
            }
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            if (ktraDL())
            {
                string sql;
                sql = "UPDATE HocVien SET MaHocVien='" + txtMaHocVien.Text + "',TenHocVien=N'" + txtTenHocVien.Text +
                    "',MaLop='" + cbbMaLop.Text +
                    "',NgaySinh='" + dtpNgaySinh.Text + "',GioiTinh=N'" + cbbGioiTinh.Text +
                    "',DiaChi=N'" + txtDienThoai.Text + "',NgayNopHocPhi='" + dtpNgayNopHocPhi.Text +
                    "',Diem=" + float.Parse(txtDiem.Text) + ",MaXepLoai='" + cbbMaXepLoai.Text + "'";
                db.CapNhat(sql);
                dataGridView.DataSource = db.DocBang("select * from HocVien");
            }
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            string sql;
            if (MessageBox.Show("Bạn có muốn xóa không?", "Thông báo",
            MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                sql = "DELETE HocVien WHERE MaHocVien='" + txtMaHocVien.Text + "'";
                db.CapNhat(sql);
                dataGridView.DataSource = db.DocBang("select * from HocVien");
            }
        }

        private void btnLamMoi_Click(object sender, EventArgs e)
        {
            txtMaHocVien.Text = "";
            txtTenHocVien.Text = "";
            cbbMaLop.Text = "";
            dtpNgaySinh.Text = "";
            cbbGioiTinh.Text = "";
            txtDiaChi.Text = "";
            cbbMaNghe.Text = "";
            txtDienThoai.Text = "";
            dtpNgayNopHocPhi.Text = "";
            txtDiem.Text = "";
            cbbMaXepLoai.Text = "";
            dataGridView.DataSource = db.DocBang("select * from HocVien");
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn thoát không?", "Thông báo",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                this.Close();
        }

        private void dataGridView_Click(object sender, EventArgs e)
        {
            txtMaHocVien.Text = dataGridView.CurrentRow.Cells[0].Value.ToString();
            txtTenHocVien.Text = dataGridView.CurrentRow.Cells[1].Value.ToString();
            cbbMaLop.Text = dataGridView.CurrentRow.Cells[2].Value.ToString();
            dtpNgaySinh.Text = dataGridView.CurrentRow.Cells[3].Value.ToString();
            cbbGioiTinh.Text = dataGridView.CurrentRow.Cells[4].Value.ToString();
            txtDiaChi.Text = dataGridView.CurrentRow.Cells[5].Value.ToString();
            cbbMaNghe.Text = dataGridView.CurrentRow.Cells[6].Value.ToString();
            txtDienThoai.Text = dataGridView.CurrentRow.Cells[7].Value.ToString();
            dtpNgayNopHocPhi.Text = dataGridView.CurrentRow.Cells[8].Value.ToString();
            txtDiem.Text = dataGridView.CurrentRow.Cells[9].Value.ToString();
            cbbMaXepLoai.Text = dataGridView.CurrentRow.Cells[10].Value.ToString();
        }
    }
}
