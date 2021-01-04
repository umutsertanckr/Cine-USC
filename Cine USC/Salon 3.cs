using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using FlexCel.XlsAdapter;
using FlexCel.Core;

namespace Cine_USC
{
    public partial class Salon_3 : Form
    {
        
        DataTable dt_s1, dt_sk1;
        BindingSource bs_s1, bs_sk1;
        OleDbDataAdapter adapt;
        OleDbConnection bagla;
        string calistir;
        string d = "Dolu", b="Boş";
        OleDbCommand kom = new OleDbCommand();
        public Salon_3()
        {
            InitializeComponent();
        }
        void goster()
        {
            dt_s1 = new DataTable();
            adapt = new OleDbDataAdapter(calistir, bagla);
            adapt.Fill(dt_s1);                       
            bs_s1 = new BindingSource();
            bs_s1.DataSource = dt_s1;
            dataGridView1.DataSource = bs_s1;
            textBox3.DataBindings.Clear();
            textBox3.DataBindings.Add("text", bs_s1, "Referans_Num");
        }
       void koltuk()
        {
            dt_sk1 = new DataTable();
            adapt = new OleDbDataAdapter(calistir, bagla);
            adapt.Fill(dt_sk1);
            bs_sk1 = new BindingSource();
            bs_sk1.DataSource = dt_sk1;
            dataGridView2.DataSource = bs_sk1;
        }
        private void Salon_3_Load(object sender, EventArgs e)
        {
            try
            {
                bagla = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=CineUSC.mdb");
                if (bagla.State == ConnectionState.Closed)
                {
                    bagla.Open();
                }
            }
            catch
            {
                MessageBox.Show("Veritabanı ile bağlantı sağlanamadı!");
            }
            calistir = "Select *from Salon_3";
            goster();
            calistir = "Select *from s3_koltuk";
            koltuk();
            durum();
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[0]["Koltuk_No"].ToString() == "-1-" && dt_sk1.Rows[0]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-1-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox1.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-1-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }

        private void pictureBox1_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[0]["Koltuk_No"].ToString() == "-1-" && dt_sk1.Rows[0]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-1-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox1.ImageLocation = "koltuk.png";
                textBox2.Text += "-1-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[1]["Koltuk_No"].ToString() == "-2-" && dt_sk1.Rows[1]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-2-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox2.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-2-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox2_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[1]["Koltuk_No"].ToString() == "-2-" && dt_sk1.Rows[1]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-2-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox2.ImageLocation = "koltuk.png";
                textBox2.Text += "-2-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[2]["Koltuk_No"].ToString() == "-3-" && dt_sk1.Rows[2]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-3-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox3.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-3-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox3_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[2]["Koltuk_No"].ToString() == "-3-" && dt_sk1.Rows[2]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-3-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox3.ImageLocation = "koltuk.png";
                textBox2.Text += "-3-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[3]["Koltuk_No"].ToString() == "-4-" && dt_sk1.Rows[3]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-4-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox4.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-4-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox4_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[3]["Koltuk_No"].ToString() == "-4-" && dt_sk1.Rows[3]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-4-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox4.ImageLocation = "koltuk.png";
                textBox2.Text += "-4-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[4]["Koltuk_No"].ToString() == "-5-" && dt_sk1.Rows[4]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-5-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox5.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-5-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox5_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[4]["Koltuk_No"].ToString() == "-5-" && dt_sk1.Rows[4]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-5-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox5.ImageLocation = "koltuk.png";
                textBox2.Text += "-5-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox6_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[5]["Koltuk_No"].ToString() == "-6-" && dt_sk1.Rows[5]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-6-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox6.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-6-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox6_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[5]["Koltuk_No"].ToString() == "-6-" && dt_sk1.Rows[5]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-6-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox6.ImageLocation = "koltuk.png";
                textBox2.Text += "-6-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox7_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[6]["Koltuk_No"].ToString() == "-7-" && dt_sk1.Rows[6]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-7-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox7.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-7-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox7_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[6]["Koltuk_No"].ToString() == "-7-" && dt_sk1.Rows[6]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-7-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox7.ImageLocation = "koltuk.png";
                textBox2.Text += "-7-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox8_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[7]["Koltuk_No"].ToString() == "-8-" && dt_sk1.Rows[7]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-8-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox8.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-8-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox8_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[7]["Koltuk_No"].ToString() == "-8-" && dt_sk1.Rows[7]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-8-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox8.ImageLocation = "koltuk.png";
                textBox2.Text += "-8-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox9_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[8]["Koltuk_No"].ToString() == "-9-" && dt_sk1.Rows[8]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-9-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox9.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-9-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox9_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[8]["Koltuk_No"].ToString() == "-9-" && dt_sk1.Rows[8]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-9-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox9.ImageLocation = "koltuk.png";
                textBox2.Text += "-9-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox10_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[9]["Koltuk_No"].ToString() == "-10-" && dt_sk1.Rows[9]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-10-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox10.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-10-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox10_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[9]["Koltuk_No"].ToString() == "-10-" && dt_sk1.Rows[9]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-10-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox10.ImageLocation = "koltuk.png";
                textBox2.Text += "-10-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox11_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[10]["Koltuk_No"].ToString() == "-11-" && dt_sk1.Rows[10]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-11-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox11.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-11-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox11_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[10]["Koltuk_No"].ToString() == "-11-" && dt_sk1.Rows[10]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-11-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox11.ImageLocation = "koltuk.png";
                textBox2.Text += "-11-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox12_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[11]["Koltuk_No"].ToString() == "-12-" && dt_sk1.Rows[11]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-12-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox12.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-12-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox12_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[11]["Koltuk_No"].ToString() == "-12-" && dt_sk1.Rows[11]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-12-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox12.ImageLocation = "koltuk.png";
                textBox2.Text += "-12-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox13_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[12]["Koltuk_No"].ToString() == "-13-" && dt_sk1.Rows[12]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-13-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox13.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-13-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox13_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[12]["Koltuk_No"].ToString() == "-13-" && dt_sk1.Rows[12]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-13-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox13.ImageLocation = "koltuk.png";
                textBox2.Text += "-13-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox14_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[13]["Koltuk_No"].ToString() == "-14-" && dt_sk1.Rows[13]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-14-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox14.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-14-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox14_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[13]["Koltuk_No"].ToString() == "-14-" && dt_sk1.Rows[13]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-14-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox14.ImageLocation = "koltuk.png";
                textBox2.Text += "-14-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox15_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[14]["Koltuk_No"].ToString() == "-15-" && dt_sk1.Rows[14]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-15-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox15.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-15-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox15_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[14]["Koltuk_No"].ToString() == "-15-" && dt_sk1.Rows[14]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-15-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox15.ImageLocation = "koltuk.png";
                textBox2.Text += "-15-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox16_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[15]["Koltuk_No"].ToString() == "-16-" && dt_sk1.Rows[15]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-16-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox16.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-16-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox16_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[15]["Koltuk_No"].ToString() == "-16-" && dt_sk1.Rows[15]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-16-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox16.ImageLocation = "koltuk.png";
                textBox2.Text += "-16-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox17_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[16]["Koltuk_No"].ToString() == "-17-" && dt_sk1.Rows[16]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-17-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox17.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-17-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox17_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[16]["Koltuk_No"].ToString() == "-17-" && dt_sk1.Rows[16]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-17-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox17.ImageLocation = "koltuk.png";
                textBox2.Text += "-17-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox18_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[17]["Koltuk_No"].ToString() == "-18-" && dt_sk1.Rows[17]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-18-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox18.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-18-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox18_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[17]["Koltuk_No"].ToString() == "-18-" && dt_sk1.Rows[17]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-18-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox18.ImageLocation = "koltuk.png";
                textBox2.Text += "-18-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox19_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[18]["Koltuk_No"].ToString() == "-19-" && dt_sk1.Rows[18]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-19-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox19.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-19-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox19_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[18]["Koltuk_No"].ToString() == "-19-" && dt_sk1.Rows[18]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-19-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox19.ImageLocation = "koltuk.png";
                textBox2.Text += "-19-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox20_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[19]["Koltuk_No"].ToString() == "-20-" && dt_sk1.Rows[19]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-20-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox20.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-20-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox20_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[19]["Koltuk_No"].ToString() == "-20-" && dt_sk1.Rows[19]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-20-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox20.ImageLocation = "koltuk.png";
                textBox2.Text += "-20-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox21_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[20]["Koltuk_No"].ToString() == "-21-" && dt_sk1.Rows[20]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-21-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox21.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-21-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox21_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[20]["Koltuk_No"].ToString() == "-21-" && dt_sk1.Rows[20]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-21-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox21.ImageLocation = "koltuk.png";
                textBox2.Text += "-21-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox22_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[21]["Koltuk_No"].ToString() == "-22-" && dt_sk1.Rows[21]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-22-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox22.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-22-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox22_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[21]["Koltuk_No"].ToString() == "-22-" && dt_sk1.Rows[21]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-22-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox22.ImageLocation = "koltuk.png";
                textBox2.Text += "-22-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox23_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[22]["Koltuk_No"].ToString() == "-23-" && dt_sk1.Rows[22]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-23-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox23.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-23-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox23_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[22]["Koltuk_No"].ToString() == "-23-" && dt_sk1.Rows[22]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-23-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox23.ImageLocation = "koltuk.png";
                textBox2.Text += "-23-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox24_Click(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[23]["Koltuk_No"].ToString() == "-24-" && dt_sk1.Rows[23]["Koltuk_Durumu"].ToString() == b)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + d + "'Where Koltuk_No= '-24-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox24.ImageLocation = "koltukdolu.png";
                textBox2.Text += "-24-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }
        private void pictureBox24_DoubleClick(object sender, EventArgs e)
        {
            if (dt_sk1.Rows[23]["Koltuk_No"].ToString() == "-24-" && dt_sk1.Rows[23]["Koltuk_Durumu"].ToString() == d)
            {
                kom.Connection = bagla;
                kom.CommandText = "Update s3_koltuk Set Koltuk_Durumu = '" + b + "'Where Koltuk_No= '-24-'";
                kom.ExecuteNonQuery();
                kom.Dispose();
                pictureBox24.ImageLocation = "koltuk.png";
                textBox2.Text += "-24-";
                calistir = "Select *from s3_koltuk";
                koltuk();
            }
            else { }
        }

        

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                kom.Connection = bagla;
                kom.CommandText = "insert into Salon_3 (Adı_Soyadı,Koltuk_Numarası) values (" +
              " '" + textBox1.Text + "','" + textBox2.Text + "')";
                kom.ExecuteNonQuery();  // komutu çalıştır.
                kom.Dispose();   // hafızayı boşalt

                MessageBox.Show("Bilet Satıldı!", "Cine USC");
            }

            catch (Exception hata) 
            {
                MessageBox.Show(hata.Message);
            }
            calistir = "Select *from Salon_3";
            goster();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            XlsFile excel = new XlsFile(true);
            excel.NewFile();

            excel.SetCellValue(1, 2, "Salon 3");
            int ek = 3;
            for (int i = 1; i <= dataGridView1.ColumnCount; i++)
            {
                excel.SetCellValue(3, i, dataGridView1.Columns[i - 1].Name);
            }

            for (int i = 1; i <= dataGridView1.RowCount - 1; i++)
            {

                for (int k = 1; k <= dataGridView1.ColumnCount; k++)
                {
                    excel.SetCellValue(i + ek, k, dataGridView1[k - 1, i - 1].Value.ToString());

                }
            }
            saveFileDialog1.Filter = ".xlsx|.xlsx";
            saveFileDialog1.ShowDialog();
            string yol2 = saveFileDialog1.FileName;

            try
            {
                excel.Save("" + yol2 + "");
            }
            catch
            {
                MessageBox.Show("Aktarma Başarısız!"); return;
            }
            MessageBox.Show("Tablo Aktarıldı!");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult cevap;
                cevap = MessageBox.Show("Kayıtı silmek istediğinizden emin misiniz?", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (cevap == DialogResult.Yes)
                {
                    kom.Connection = bagla;
                    kom.CommandText = "delete from Salon_3 where Referans_Num=" + dataGridView1.CurrentRow.Cells["Referans_Num"].Value +"";
                    kom.ExecuteNonQuery();
                    kom.Dispose();
                    calistir = "select * from Salon_3";
                    MessageBox.Show("Biletiniz İptal Edildi!");
                    goster();
                }
            }
            catch (Exception hata)
            {
                MessageBox.Show(hata.Message);
            }
            calistir = "Select *from Salon_3";
            goster();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox2.Clear();
        }
        void durum()
        {
            if (dt_sk1.Rows[0]["Koltuk_No"].ToString() == "-1-" && dt_sk1.Rows[0]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox1.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox1.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[1]["Koltuk_No"].ToString() == "-2-" && dt_sk1.Rows[1]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox2.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox2.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[2]["Koltuk_No"].ToString() == "-3-" && dt_sk1.Rows[2]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox3.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox3.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[3]["Koltuk_No"].ToString() == "-4-" && dt_sk1.Rows[3]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox4.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox4.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[4]["Koltuk_No"].ToString() == "-5-" && dt_sk1.Rows[4]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox5.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox5.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[5]["Koltuk_No"].ToString() == "-6-" && dt_sk1.Rows[5]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox6.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox6.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[6]["Koltuk_No"].ToString() == "-7-" && dt_sk1.Rows[6]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox7.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox7.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[7]["Koltuk_No"].ToString() == "-8-" && dt_sk1.Rows[7]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox8.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox8.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[8]["Koltuk_No"].ToString() == "-9-" && dt_sk1.Rows[8]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox9.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox9.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[9]["Koltuk_No"].ToString() == "-10-" && dt_sk1.Rows[9]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox10.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox10.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[10]["Koltuk_No"].ToString() == "-11-" && dt_sk1.Rows[10]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox11.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox11.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[11]["Koltuk_No"].ToString() == "-12-" && dt_sk1.Rows[11]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox12.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox12.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[12]["Koltuk_No"].ToString() == "-13-" && dt_sk1.Rows[12]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox13.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox13.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[13]["Koltuk_No"].ToString() == "-14-" && dt_sk1.Rows[13]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox14.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox14.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[14]["Koltuk_No"].ToString() == "-15-" && dt_sk1.Rows[14]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox15.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox15.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[15]["Koltuk_No"].ToString() == "-16-" && dt_sk1.Rows[15]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox16.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox16.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[16]["Koltuk_No"].ToString() == "-17-" && dt_sk1.Rows[16]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox17.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox17.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[17]["Koltuk_No"].ToString() == "-18-" && dt_sk1.Rows[17]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox18.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox18.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[18]["Koltuk_No"].ToString() == "-19-" && dt_sk1.Rows[18]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox19.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox19.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[19]["Koltuk_No"].ToString() == "-20-" && dt_sk1.Rows[19]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox20.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox20.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[20]["Koltuk_No"].ToString() == "-21-" && dt_sk1.Rows[20]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox21.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox21.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[21]["Koltuk_No"].ToString() == "-22-" && dt_sk1.Rows[0]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox22.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox22.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[22]["Koltuk_No"].ToString() == "-23-" && dt_sk1.Rows[22]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox23.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox23.ImageLocation = "koltuk.png"; }
            if (dt_sk1.Rows[23]["Koltuk_No"].ToString() == "-24-" && dt_sk1.Rows[23]["Koltuk_Durumu"].ToString() == d)
            {
                pictureBox24.ImageLocation = "koltukdolu.png";
            }
            else { pictureBox24.ImageLocation = "koltuk.png"; }

        }
    }
}
