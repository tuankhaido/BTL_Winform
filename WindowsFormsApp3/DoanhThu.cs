using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
//toi la khai
namespace WindowsFormsApp3
{
    public partial class DoanhThu : Form
    {
        ProcessDatabase db = new ProcessDatabase();
        DataTable dt = new DataTable();
        public DoanhThu()
        {
            InitializeComponent();
        }

        private void cbbMaLop_Click(object sender, EventArgs e)
        {
            cbbMaLop.DataSource = db.DocBang("Select * from LopHoc");
            cbbMaLop.ValueMember = "MaLop";
            cbbMaLop.DisplayMember = "MaLop";
        }

        private void btnXem_Click(object sender, EventArgs e)
        {
            string sql;
            sql = "select LopHoc.MaLop, TenLop, sum(HocPhi) as DoanhThu from HocVien join LopHoc " +
                "on HocVien.MaLop = LopHoc.MaLop join MonHoc " +
                "on LopHoc.MaMon = MonHoc.MaMon " +
                "where LopHoc.MaLop='" + cbbMaLop.Text + "' " + "group by LopHoc.MaLop, LopHoc.TenLop";
            dt = db.DocBang(sql);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].HeaderText = "Mã lớp";
            dataGridView1.Columns[1].HeaderText = "Tên lớp";
            dataGridView1.Columns[2].HeaderText = "Doanh thu";
        }

        private void btnXuatFileExcel_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0) //TH có dữ liệu để ghi
            {
                //Khai báo và khởi tạo các đối tượng
                Excel.Application exApp = new Excel.Application();
                Excel.Workbook exBook =
               exApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet exSheet = (Excel.Worksheet)exBook.Worksheets[1];
                //header
                Excel.Range header = (Excel.Range)exSheet.Cells[1, 2];
                exSheet.get_Range("B1:F1").Merge(true);
                header.Font.Size = 13;
                header.Font.Bold = true;
                header.Font.Color = Color.Red;
                header.Value = "DOANH THU THEO MÃ LỚP " + cbbMaLop.Text + "";
                //Định dạng tiêu đề bảng

                exSheet.get_Range("A3:C3").Font.Bold = true;
                exSheet.get_Range("A3:C3").HorizontalAlignment =
               Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                exSheet.get_Range("A3").Value = "STT";
                exSheet.get_Range("B3").Value = "Mã lớp";
                exSheet.get_Range("C3").Value = "Tên lớp";
                exSheet.get_Range("D3").Value = "Doanh thu";

                //In dữ liệu
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    exSheet.get_Range("A" + (i + 4).ToString() + ":D" + (i +
                   4).ToString()).Font.Bold = false;
                    exSheet.Cells[1, 3].EntireColumn.AutoFit();
                    exSheet.get_Range("A" + (i + 4).ToString()).Value = (i +
                   1).ToString();
                    exSheet.get_Range("B" + (i + 4).ToString()).Value =
                   dt.Rows[i]["MaLop"].ToString();
                    exSheet.get_Range("C" + (i + 4).ToString()).Value =
                   dt.Rows[i]["TenLop"].ToString();
                    exSheet.get_Range("D" + (i + 4).ToString()).Value =
                   dt.Rows[i]["DoanhThu"].ToString();
                }
                for (int col = 1; col <= dt.Columns.Count; col++)
                {
                    exSheet.Columns[col].AutoFit();
                }
                exSheet.Name = "DoanhThu";
                exBook.Activate(); //Kích hoạt file Excel
                                   //Thiết lập các thuộc tính của SaveFileDialog
                dlgSave1.Filter = "Excel Document(*.xlsx)|*.xlsx |Word Document(*.doc) | *.doc | All files(*.*) | *.* ";
                dlgSave1.FilterIndex = 1;
                dlgSave1.AddExtension = true;
                dlgSave1.DefaultExt = ".xlsx";
                if (dlgSave1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    exBook.SaveAs(dlgSave1.FileName.ToString());//Lưu file Excel
                exApp.Quit();//Thoát khỏi ứng dụng
            }
            else
                MessageBox.Show("Không có danh sách doanh thu để in");
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn thoát không?", "Thông báo",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                this.Close();
        }
    }
}
