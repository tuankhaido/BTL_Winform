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
    public partial class GiaoVien : Form
    {
        ProcessDatabase db = new ProcessDatabase();
        string imagePath;
        public GiaoVien()
        {
            InitializeComponent();
        }

        private void GiaoVien_Load(object sender, EventArgs e)
        {
            dataGridView.DataSource = db.DocBang("select * from GiaoVien");
            dataGridView.Columns[0].HeaderText = "Mã GV";
            dataGridView.Columns[1].HeaderText = "Tên GV";
            dataGridView.Columns[2].HeaderText = "Ngày sinh";
            dataGridView.Columns[3].HeaderText = "Giới tính";
            dataGridView.Columns[4].HeaderText = "Địa chỉ";
            dataGridView.Columns[5].HeaderText = "Điện thoại";
            dataGridView.Columns[6].HeaderText = "Mã chức danh";
            dataGridView.Columns[7].HeaderText = "Mã chuyên môn";
            dataGridView.Columns[8].HeaderText = "Mã cơ quan";
            dataGridView.Columns[9].HeaderText = "Mã trình độ";
            dataGridView.Columns[10].HeaderText = "Ảnh";
        }

        private void cbbChucDanh_Click(object sender, EventArgs e)
        {
            cbbChucDanh.DataSource = db.DocBang("Select * from ChucDanh");
            cbbChucDanh.ValueMember = "MaChucDanh";
            cbbChucDanh.DisplayMember = "MaChucDanh";
        }

        private void cbbChuyenMon_Click(object sender, EventArgs e)
        {
            cbbChuyenMon.DataSource = db.DocBang("Select * from ChuyenMon");
            cbbChuyenMon.ValueMember = "MaChuyenMon";
            cbbChuyenMon.DisplayMember = "MaChuyenMon";
        }

        private void cbbCQ_Click(object sender, EventArgs e)
        {
            cbbCQ.DataSource = db.DocBang("Select * from CoQuan");
            cbbCQ.ValueMember = "MaCQ";
            cbbCQ.DisplayMember = "MaCQ";
        }

        private void cbbTrinhDo_Click(object sender, EventArgs e)
        {
            cbbTrinhDo.DataSource = db.DocBang("Select * from TrinhDo");
            cbbTrinhDo.ValueMember = "MaTrinhDo";
            cbbTrinhDo.DisplayMember = "MaTrinhDo";
        }

        public bool ktraDL()
        {
            bool k = true;
            if (txtMaGV.Text.Trim().Equals("") || txtTenGV.Text.Trim().Equals("") || dtpNgaySinh.Text.Trim().Equals("") || cbbGioiTinh.Text.Trim().Equals("")
                || txtDiaChi.Text.Trim().Equals("") || txtDienThoai.Text.Trim().Equals("") || cbbChucDanh.Text.Trim().Equals("") || cbbChuyenMon.Text.Trim().Equals("")
                || cbbCQ.Text.Trim().Equals("") || cbbTrinhDo.Text.Trim().Equals(""))
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
                string sql, txtAnh;
                sql = "SELECT MaGV FROM GiaoVien WHERE MaGV=N'" + txtMaGV.Text + "'";
                SqlDataAdapter MyData = new SqlDataAdapter(sql, db.Con);
                DataTable table = new DataTable();
                MyData.Fill(table);
                if (table.Rows.Count > 0)
                {
                    MessageBox.Show("Mã giáo viên này đã tồn tại");
                    return;
                }
                //Thêm đường dẫn ảnh
                if (pictureBox.Image != null)
                {
                    txtAnh = imagePath;
                }
                else
                    txtAnh = "ko";
                dataGridView.DataSource = db.DocBang("select * from GiaoVien");
            }
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            if (ktraDL())
            {
                string sql, txtAnh;
                //Thêm đường dẫn ảnh
                if (pictureBox.Image != null)
                {
                    txtAnh = imagePath;
                }
                else
                    txtAnh = "ko";
                sql = "UPDATE GiaoVien SET MaGV='" + txtMaGV.Text + "',TenGV=N'" + txtTenGV.Text +
                    "',NgaySinh='" + dtpNgaySinh.Text +
                    "',GioiTinh=N'" + cbbGioiTinh.Text + "',DiaChi=N'" + txtDiaChi.Text +
                    "',DienThoai='" + txtDienThoai.Text + "',MaChucDanh='" + cbbChucDanh.Text +
                    "',MaChuyenMon='" + cbbChuyenMon.Text + "',MaCQ='" + cbbCQ.Text +
                    "',MaTrinhDo='" + cbbTrinhDo.Text + "',Anh='" + txtAnh + "' WHERE MaGV='" + txtMaGV.Text + "'";
                db.CapNhat(sql);
                dataGridView.DataSource = db.DocBang("select * from GiaoVien");
            }
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            string sql;
            if (MessageBox.Show("Bạn có muốn xóa không?", "Thông báo",
            MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                sql = "DELETE GiaoVien WHERE MaGV='" + txtMaGV.Text + "'";
                db.CapNhat(sql);
                dataGridView.DataSource = db.DocBang("select * from GiaoVien");
                pictureBox.Image = null;
            }
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            TimKiemGiaoVien tkgv = new TimKiemGiaoVien();
            tkgv.Show();
        }

        private void btnLamMoi_Click(object sender, EventArgs e)
        {
            txtMaGV.Text = "";
            txtTenGV.Text = "";
            dtpNgaySinh.Value = DateTime.Now;
            cbbGioiTinh.Text = "";
            txtDiaChi.Text = "";
            txtDienThoai.Text = "";
            cbbChucDanh.Text = "";
            cbbChuyenMon.Text = "";
            cbbCQ.Text = "";
            cbbTrinhDo.Text = "";
            pictureBox.Image = null;
            dataGridView.DataSource = db.DocBang("select * from GiaoVien");
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn thoát không?", "Thông báo",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                this.Close();
        }

        private void btnThemAnh_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Image Files (.png; *.jpg; *.jpeg; *.gif; *.bmp)|.png; *.jpg; *.jpeg; *.gif; *.bmp";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Lấy đường dẫn của file ảnh đã chọn
                imagePath = openFileDialog.FileName;

                // Tạo một đối tượng Image từ file ảnh
                Image image = Image.FromFile(imagePath);

                // Đặt ảnh vào PictureBox
                pictureBox.Image = image;
            }
        }

        private void dataGridView_Click(object sender, EventArgs e)
        {
            txtMaGV.Text = dataGridView.CurrentRow.Cells[0].Value.ToString();
            txtTenGV.Text = dataGridView.CurrentRow.Cells[1].Value.ToString();
            dtpNgaySinh.Text = dataGridView.CurrentRow.Cells[2].Value.ToString();
            cbbGioiTinh.Text = dataGridView.CurrentRow.Cells[3].Value.ToString();
            txtDiaChi.Text = dataGridView.CurrentRow.Cells[4].Value.ToString();
            txtDienThoai.Text = dataGridView.CurrentRow.Cells[5].Value.ToString();
            cbbChucDanh.Text = dataGridView.CurrentRow.Cells[6].Value.ToString();
            cbbChuyenMon.Text = dataGridView.CurrentRow.Cells[7].Value.ToString();
            cbbCQ.Text = dataGridView.CurrentRow.Cells[8].Value.ToString();
            cbbTrinhDo.Text = dataGridView.CurrentRow.Cells[9].Value.ToString();
            if (dataGridView.CurrentRow.Cells[10].Value.ToString() != "ko")
            {
                imagePath = dataGridView.CurrentRow.Cells[10].Value.ToString();
                Image image = Image.FromFile(imagePath);
                pictureBox.Image = image;
            }
            else
                pictureBox.Image = null;
        }
    }
}
