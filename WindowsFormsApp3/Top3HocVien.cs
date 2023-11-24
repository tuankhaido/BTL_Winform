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

namespace WindowsFormsApp3
{
    public partial class Top3HocVien : Form
    {
        ProcessDatabase db = new ProcessDatabase();
        DataTable dt = new DataTable();
        public Top3HocVien()
        {
            InitializeComponent();
        }

        private void btnXem_Click(object sender, EventArgs e)
        {
            string sql = "SELECT * FROM (SELECT LopHoc.MaLop, HocVien.MaHocVien, HocVien.TenHocVien, HocVien.Diem, " +
                "ROW_NUMBER() OVER (PARTITION BY LopHoc.MaLop ORDER BY HocVien.Diem DESC) AS RowNum " +
                "FROM LopHoc INNER JOIN HocVien ON LopHoc.MaLop = HocVien.MaLop) AS T WHERE T.RowNum <= 3";
            dt = db.DocBang(sql);
            dataGridView.DataSource = dt;
            dataGridView.Columns[0].HeaderText = "Mã lớp";
            dataGridView.Columns[1].HeaderText = "Mã học viên";
            dataGridView.Columns[2].HeaderText = "Tên học viên";
            dataGridView.Columns[3].HeaderText = "Điểm";
            dataGridView.Columns[4].HeaderText = "Top";
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
                header.Value = "DANH SÁCH 3 HỌC VIÊN ĐẠT ĐIỂM CAO NHẤT CỦA TỪNG LỚP";
                //Định dạng tiêu đề bảng

                exSheet.get_Range("A3:E3").Font.Bold = true;
                exSheet.get_Range("A3:E3").HorizontalAlignment =
               Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                exSheet.get_Range("A3").Value = "Top";
                exSheet.get_Range("B3").Value = "Mã học viên";
                exSheet.get_Range("C3").Value = "Tên học viên";
                exSheet.get_Range("D3").Value = "Mã lớp";
                exSheet.get_Range("E3").Value = "Điểm";


                //In dữ liệu
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    exSheet.get_Range("A" + (i + 4).ToString() + ":D" + (i +
                   4).ToString()).Font.Bold = false;
                    exSheet.Cells[1, 3].EntireColumn.AutoFit();
                    exSheet.get_Range("A" + (i + 4).ToString()).Value = (i +
                   1).ToString();
                    exSheet.get_Range("B" + (i + 4).ToString()).Value =
                   dt.Rows[i]["MaHocVien"].ToString();
                    exSheet.get_Range("C" + (i + 4).ToString()).Value =
                   dt.Rows[i]["TenHocVien"].ToString();
                    exSheet.get_Range("D" + (i + 4).ToString()).Value =
                   dt.Rows[i]["MaLop"].ToString();
                    exSheet.get_Range("E" + (i + 4).ToString()).Value =
                   dt.Rows[i]["Diem"].ToString();
                }
                for (int col = 1; col <= dt.Columns.Count; col++)
                {
                    exSheet.Columns[col].AutoFit();
                }
                exSheet.Name = "Top3";
                exBook.Activate(); //Kích hoạt file Excel
                                   //Thiết lập các thuộc tính của SaveFileDialog
                dlgSave3.Filter = "Excel Document(*.xlsx)|*.xlsx |Word Document(*.doc) | *.doc | All files(*.*) | *.* ";
                dlgSave3.FilterIndex = 1;
                dlgSave3.AddExtension = true;
                dlgSave3.DefaultExt = ".xlsx";
                if (dlgSave3.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    exBook.SaveAs(dlgSave3.FileName.ToString());//Lưu file Excel
                exApp.Quit();//Thoát khỏi ứng dụng
            }
            else
                MessageBox.Show("Không có danh sách điểm để in");

        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn thoát không?", "Thông báo",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                this.Close();
        }
    }
}
