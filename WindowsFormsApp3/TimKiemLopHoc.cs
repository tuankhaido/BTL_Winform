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
    public partial class TimKiemLopHoc : Form
    {
        ProcessDatabase db = new ProcessDatabase();
        public TimKiemLopHoc()
        {
            InitializeComponent();
        }

        private void cbbMonHoc_Click(object sender, EventArgs e)
        {
            cbbMonHoc.DataSource = db.DocBang("Select * from MonHoc");
            cbbMonHoc.ValueMember = "MaMon";
            cbbMonHoc.DisplayMember = "MaMon";
        }

        private void cbbGV_Click(object sender, EventArgs e)
        {
            cbbGV.DataSource = db.DocBang("Select * from GiaoVien");
            cbbGV.ValueMember = "MaGV";
            cbbGV.DisplayMember = "MaGV";
        }
        public bool ktraDL()
        {
            bool k = true;
            if (cbbMonHoc.Text.Trim().Equals("") || cbbGV.Text.Trim().Equals("") || txtSiSo.Text.Trim().Equals(""))
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
                string monHoc = cbbMonHoc.Text; // Tiêu chí môn học
                string giaoVien = cbbGV.Text;  // Tiêu chí giáo viên
                int siSo = int.Parse(txtSiSo.Text);  // Tiêu chí sĩ số

                // Xây dựng truy vấn SQL để tìm kiếm dựa trên các tiêu chí đã cung cấp
                string sql = "SELECT * FROM LopHoc WHERE 1 = 1";

                if (!string.IsNullOrEmpty(monHoc))
                {
                    sql += $" AND MaMon = N'{monHoc}'";
                }

                if (!string.IsNullOrEmpty(giaoVien))
                {
                    sql += $" AND MaGV = '{giaoVien}'";
                }

                if (siSo > 0)
                {
                    sql += $" AND SiSo <= {siSo}";
                }

                // Thực hiện tìm kiếm và cập nhật DataGridView với kết quả tìm kiếm
                dataGridView1.DataSource = db.DocBang(sql);
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
        }

        private void btnLamMoi_Click(object sender, EventArgs e)
        {
            cbbMonHoc.Text = "";
            cbbGV.Text = "";
            txtSiSo.Text = "";
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn thoát không?", "Thông báo",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                this.Close();
        }
    }
}
