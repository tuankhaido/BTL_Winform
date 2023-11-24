using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp3
{
    internal class ProcessDatabase
    {
        SqlConnection con = new SqlConnection(@"Data Source=DESKTOP-69DU6AV;Initial Catalog=QuanLyHocVien;Integrated Security=True");

        public SqlConnection Con { get => con; set => con = value; }

        public void KetNoi()
        {
            con = new SqlConnection(@"Data Source=DESKTOP-69DU6AV;Initial Catalog=QuanLyHocVien;Integrated Security=True");
            //if (con.State != ConnectionState.Open)
            con.Open();
        }

        public void DongKetNoi()
        {
            if (con.State != ConnectionState.Closed)
                con.Close();
            con.Dispose();
        }

        public DataTable DocBang(string sql)
        {
            KetNoi();
            DataTable tb = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            da.Fill(tb);
            DongKetNoi();
            return tb;
        }

        public void CapNhat(string sql)
        {
            SqlCommand cm = new SqlCommand();
            KetNoi();
            cm.CommandText = sql;
            cm.Connection = con;
            cm.ExecuteNonQuery();
            DongKetNoi();
            cm.Dispose();
        }

        public void Trigger()
        {
            KetNoi();
            SqlCommand cmd = new SqlCommand(); //Đối tượng để thực hiện lệnh
            cmd.CommandText = "create or alter trigger setSiSo on HocVien for insert, update as " +
                "begin " +
                "declare @malop nvarchar(50) " +
                "select @malop = MaLop from inserted " +
                "update LopHoc set SiSo = (select count(MaHocVien) from HocVien join LopHoc " +
                "on HocVien.MaLop = LopHoc.MaLop " +
                "where HocVien.MaLop = @malop) " +
                "where MaLop = @malop " +
                "end";
            cmd.Connection = con;
            try
            {
                cmd.ExecuteNonQuery(); //Thực hiện câu lệnh
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            DongKetNoi();
        }
    }
}
