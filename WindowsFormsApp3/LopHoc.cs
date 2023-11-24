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
    public partial class LopHoc : Form
    {
        ProcessDatabase db = new ProcessDatabase();
        public LopHoc()
        {
            InitializeComponent();
        }

        private void LopHoc_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = db.DocBang("select * from LopHoc");
            dataGridView1.Columns[0].HeaderText = "Mã lớp";
            dataGridView1.Columns[1].HeaderText = "Tên lớp";
            dataGridView1.Columns[2].HeaderText = "Mã môn";
            dataGridView1.Columns[3].HeaderText = "Mã giáo viên";
            dataGridView1.Columns[4].HeaderText = "Ngày BD";
            dataGridView1.Columns[5].HeaderText = "Ngày KT";
            dataGridView1.Columns[6].HeaderText = "Mã ca";
            dataGridView1.Columns[7].HeaderText = "Mã thời gian";
            dataGridView1.Columns[8].HeaderText = "Mã phòng";
            dataGridView1.Columns[9].HeaderText = "Sĩ số";
        }

        private void cbbMaMon_Click(object sender, EventArgs e)
        {
            cbbMaMon.DataSource = db.DocBang("Select * from MonHoc");
            cbbMaMon.ValueMember = "MaMon";
            cbbMaMon.DisplayMember = "MaMon";
        }

        private void cbbMaMon_SelectedIndexChanged(object sender, EventArgs e)
        {
            cbbMaGV.DataSource = db.DocBang("Select GiaoVien.MaGV, TenGV, MaMon from GiaoVien join GiaoVienMonHoc " +
                "on GiaoVien.MaGV = GiaoVienMonHoc.MaGV where MaMon='" + cbbMaMon.Text + "'");
            cbbMaGV.ValueMember = "MaGV";
            cbbMaGV.DisplayMember = "MaGV";
        }

        private void cbbMaCa_Click(object sender, EventArgs e)
        {
            cbbMaCa.DataSource = db.DocBang("Select * from CaHoc");
            cbbMaCa.ValueMember = "MaCa";
            cbbMaCa.DisplayMember = "MaCa";
        }

        private void cbbMaTG_Click(object sender, EventArgs e)
        {
            cbbMaTG.DataSource = db.DocBang("Select * from ThoiGian");
            cbbMaTG.ValueMember = "MaTG";
            cbbMaTG.DisplayMember = "MaTG";
        }

        private void cbbMaPhong_Click(object sender, EventArgs e)
        {
            cbbMaPhong.DataSource = db.DocBang("Select * from PhongHoc");
            cbbMaPhong.ValueMember = "MaPhong";
            cbbMaPhong.DisplayMember = "MaPhong";
        }

        public bool ktraDL()    //Khi thêm và sửa bắt buộc phải nhập đủ thông tin
        {
            bool k = true;
            if (txtMaLop.Text.Trim().Equals("") || txtTenLop.Text.Trim().Equals("") || cbbMaGV.Text.Trim().Equals("") || dtpNgayBD.Text.Trim().Equals("")
                || dtpNgayKT.Text.Trim().Equals("") || cbbMaCa.Text.Trim().Equals("") || cbbMaTG.Text.Trim().Equals("") || cbbMaPhong.Text.Trim().Equals(""))
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
                sql = "SELECT MaLop FROM LopHoc WHERE MaLop=N'" + txtMaLop.Text + "'";
                SqlDataAdapter MyData = new SqlDataAdapter(sql, db.Con);
                DataTable table = new DataTable();
                MyData.Fill(table);
                if (table.Rows.Count > 0)
                {
                    MessageBox.Show("Mã lớp này đã tồn tại");
                    return;
                }
                //Thực hiện chèn thêm mới
                sql = "INSERT INTO LopHoc VALUES ('" + txtMaLop.Text + "',N'" + txtTenLop.Text + "','" +
                cbbMaMon.Text + "','" + cbbMaGV.Text + "','" + dtpNgayBD.Text + "','" + dtpNgayKT.Text +
                "','" + cbbMaCa.Text + "','" + cbbMaTG.Text + "','" + cbbMaPhong.Text +
                "'," + int.Parse(txtSiSo.Text) + ")";
                db.CapNhat(sql);
                dataGridView1.DataSource = db.DocBang("select * from LopHoc");
            }
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            if (ktraDL())
            {
                string sql;
                sql = "UPDATE LopHoc SET MaLop='" + txtMaLop.Text + "',TenLop=N'" + txtTenLop.Text +
                    "',MaMon='" + cbbMaMon.Text + "',MaGV='" + cbbMaGV.Text + "',NgayBD='" + dtpNgayBD.Text +
                    "',NgayKT='" + dtpNgayKT.Text + "',MaCa='" + cbbMaCa.Text + "',MaThoiGian='" + cbbMaTG.Text +
                    "',MaPhong='" + cbbMaPhong.Text + "',SiSo=" + int.Parse(txtSiSo.Text) +
                    " WHERE MaLop='" + txtMaLop.Text + "'";
                db.CapNhat(sql);
                dataGridView1.DataSource = db.DocBang("select * from LopHoc");
            }
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            string sql;
            if (MessageBox.Show("Bạn có muốn xóa không?", "Thông báo",
            MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                sql = "DELETE LopHoc WHERE MaLop='" + txtMaLop.Text + "'";
                db.CapNhat(sql);
                dataGridView1.DataSource = db.DocBang("select * from LopHoc");
            }
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            TimKiemLopHoc tklh = new TimKiemLopHoc();
            tklh.Show();
        }

        private void btnLamMoi_Click(object sender, EventArgs e)
        {
            txtMaLop.Text = "";
            txtTenLop.Text = "";
            cbbMaMon.Text = "";
            cbbMaGV.Items.Clear();
            cbbMaGV.Text = "";
            dtpNgayBD.Value = DateTime.Now;
            dtpNgayKT.Value = DateTime.Now;
            cbbMaCa.Text = "";
            cbbMaTG.Text = "";
            cbbMaPhong.Text = "";
            txtSiSo.Text = "0";
            dataGridView1.DataSource = db.DocBang("select * from LopHoc");
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn thoát không?", "Thông báo",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                this.Close();
        }

        private void dataGridView1_Click(object sender, EventArgs e)
        {
            txtMaLop.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            txtTenLop.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            cbbMaMon.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            cbbMaGV.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            dtpNgayBD.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            dtpNgayKT.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            cbbMaCa.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            cbbMaTG.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            cbbMaPhong.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
            txtSiSo.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
        }
    }
}
