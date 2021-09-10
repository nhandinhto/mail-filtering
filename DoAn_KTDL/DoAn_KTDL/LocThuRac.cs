using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace DoAn_KTDL
{
    public partial class LocThuRac : Form
    {
        public LocThuRac()
        {
            InitializeComponent();
        }
        XuLy xl = new XuLy();
        private void LocThuRac_Load(object sender, EventArgs e)
        {

        }


        private void btnTrain_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (xl.train())
                {
                    label1.Text = "Train Hoàn Tất";
                    lbAccuracy.Text = XuLy.Accuracy().ToString() + "%";
                }
                else
                {
                    label1.Text = "Tran không thành công";
                }
            }
            catch
            {
            }
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "All Text File (*.txt)|*.txt|All Document (*.pdf)|*.pdf";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //Gán đường dẫn lên textbox
                textBox1.Text = openFileDialog1.FileName;
                //Đọc dữ liệu từ file txt dùng 
                StreamReader sr = new StreamReader(openFileDialog1.FileName);
                //Đọc dữ liệu text và gán vào richTextBox
                richTextBox1.Text = sr.ReadToEnd();
                // Đóng luồng
                sr.Close();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = textBox1.Text;
            string test = richTextBox1.Text;
            int kq = XuLy.duDoan(test);
            if (kq == 1)
                lbDD.Text = "Thư Không Phải Là Thư Rác";
            else
                lbDD.Text = "Thư Là Thư Rác";
        }
    }
}
