using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp3
{
    public partial class Form1 : Form
    {
        ProcessDatabase db = new ProcessDatabase();
        public Form1()
        {
            InitializeComponent();
            db.Trigger();
        }

        private void danhSáchGiáoViênToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GiaoVien gv = new GiaoVien();
            gv.Show();
        }

        private void danhSáchLớpHọcToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LopHoc lh = new LopHoc();
            lh.Show();
        }

        private void báoCáoDoanhThuTheoLớpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DoanhThu dt = new DoanhThu();
            dt.Show();
        }

        private void báoCáoDanhSáchLớpTheoPhòngToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LopTheoPhong ltp = new LopTheoPhong();
            ltp.Show();
        }

        private void danhSáchHọcViênToolStripMenuItem_Click(object sender, EventArgs e)
        {
            HocVien hv = new HocVien();
            hv.Show();
        }

        private void báoCáoĐiểmTheoLớpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DiemHocVien dhv = new DiemHocVien();
            dhv.Show();
        }

        private void báoCáoCácHọcViênĐiểmCaoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Top3HocVien t3hv = new Top3HocVien();
            t3hv.Show();
        }
    }
}
