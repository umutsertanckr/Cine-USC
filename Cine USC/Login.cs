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

namespace Cine_USC
{
    public partial class Login : Form
    {
        OleDbConnection bagla;
        OleDbDataReader oku;
        string calistir;
        public Login()
        {
            InitializeComponent();
        }

        private void Login_Load(object sender, EventArgs e)
        {
       
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bagla = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=CineUSC.mdb");
            bagla.Open();
            calistir = "SELECT * FROM login";
            OleDbCommand kom = new OleDbCommand(calistir,bagla);
            kom.Connection = bagla;
            kom.CommandText = "SELECT * FROM login where Kullanıcı_Adı='" + textBox1.Text + "' AND Sifre='" + textBox2.Text + "'";
            oku = kom.ExecuteReader();
            if (oku.Read())
            {
                MessageBox.Show("Tebrikler! Başarılı bir şekilde giriş yaptınız.");
                Form1 ana = new Form1();
                ana.Show();
                this.Visible = false;
            }
            else
            {
                MessageBox.Show("Kullanıcı adı veya şifre hatalı kontrol ediniz.");
            }
            bagla.Close();
        }
    }
}
