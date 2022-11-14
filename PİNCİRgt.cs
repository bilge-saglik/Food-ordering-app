using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;
using System.Windows.Forms;
using System.IO;
using System.Configuration;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace Pincir
{

    public partial class Form1 : Form

    {
        public Form1()
        {
            InitializeComponent();

        }
        static string conString = "Server= ServerAD ; Database=DBad; User Id = IDad ; password=sifre; ";
        SqlConnection baglanti = new SqlConnection(conString);

        private void Form1_Load(object sender, EventArgs e)

          
        {

        SqlConnection con = new SqlConnection(conString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = con;
            cmd.CommandText = "SELECT * FROM urunler";
            con.Open();

            SqlDataReader dr2 = cmd.ExecuteReader();

            //ArrayList Isimler = new ArrayList();
            int i = 0;
            while (dr2.Read())
            {
                // checkedListBox1.Items.Add(dr2["urun_ad"]);
                checkedListBox1.DisplayMember = "Text";
                checkedListBox1.ValueMember = "Value";
                checkedListBox1.Items.Insert(i, new CheckedListMy{Text = dr2["urun_ad"].ToString()+" "+ dr2["urun_fyt"].ToString() + " TL", Value = Convert.ToInt32(dr2["urun_id"]) });
                i++;
            }

            dr2.Close();
            con.Close();           

            comboBox1.Items.Clear();

            baglanti.Open();
            SqlCommand sorgu = new SqlCommand(" SELECT * FROM Sehirler ", baglanti);
            sorgu.ExecuteNonQuery();

            DataTable dt = new DataTable();
            SqlDataAdapter adp = new SqlDataAdapter(sorgu);
            adp.Fill(dt);

            foreach (DataRow dr in dt.Rows)
            {
                ComboboxItem item = new ComboboxItem();
                item.Text = dr["SehirAdi"].ToString();
                item.Value = Convert.ToInt32(dr["SehirId"]);

                comboBox1.Items.Add(item);
                comboBox1.SelectedIndex = 0;
            }


        }
        public class ComboboxItem
        {
            public string Text { get; set; }
            public object Value { get; set; }

            public override string ToString()
            {
                return Text;
            }
        }
       
        public class CheckedListMy
        {
            public string Text { get; set; }
            public int Value { get; set; }
        }

        [Obsolete]

        public class mylistbox
        {
            public string Text { get; set; }
            public int Value { get; set; }
        }
        private void button1_Click(object sender, EventArgs e) 
        {
            tabControl1.SelectedIndex = 2;
            SqlConnection baglanti = new SqlConnection(conString);
         
            if (baglanti.State == ConnectionState.Closed)
                baglanti.Open();
           
           


            foreach (var item in checkedListBox1.CheckedItems)
            {
                string kayit = "insert into sepet(sepet_urun, sepet_fyt,ref_urunid) values (@sepeturun, @sepetfyt,@ref_urunid)";
                SqlCommand komut = new SqlCommand(kayit, baglanti);
                CheckedListMy row = (CheckedListMy)item;
                MessageBox.Show(row.Text + ": " + row.Value);

                string cmdtext = "select * from urunler where urun_id=" + row.Value + "";
                SqlConnection con = new SqlConnection(conString);
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = cmdtext;
                con.Open();
                double fiyat = 0;
                SqlDataReader dr2 = cmd.ExecuteReader();

                //ArrayList Isimler = new ArrayList();
                int i = 0;
                while (dr2.Read())
                {
                    fiyat = Convert.ToDouble(dr2["urun_fyt"]);
                }
                con.Close();

                komut.Parameters.AddWithValue("@ref_urunid", row.Value);
                komut.Parameters.AddWithValue("@sepeturun", row.Text);


                komut.Parameters.AddWithValue("@sepetfyt", fiyat);
                if (baglanti.State == ConnectionState.Closed)
                    baglanti.Open();
                komut.ExecuteNonQuery();
                baglanti.Close();
               
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 1;

        }
        int musteriid = 0;

        [Obsolete]
        private void button3_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 3;

            if (textBox2.Text.Length > 12)
            {

                MessageBox.Show("12 krakterden fazla telefon no girilemez !");
                return;

            }
            else
            {
                if (baglanti.State == ConnectionState.Open)
                    baglanti.Close();
                baglanti.Open();
                string kayit = "insert into musteri(mus_ad_soyad,mus_tel,mus_adres)  output INSERTED.musteri_id values (@musteriad,@mustel,@musadres)";
                SqlCommand komut = new SqlCommand(kayit, baglanti);

                komut.Parameters.AddWithValue("@musteriad", textBox1.Text);


                komut.Parameters.AddWithValue("@mustel", textBox2.Text);

                komut.Parameters.AddWithValue("@musadres", textBox3.Text);
                musteriid = (int)komut.ExecuteScalar();
                baglanti.Close();
                MessageBox.Show("Kaydınız gerçekleşti!");
                string aktar = textBox1.Text;
                string aktar1 = textBox2.Text;
                string aktar2 = textBox3.Text;
                listBox2.Items.Add(aktar);
                listBox2.Items.Add(aktar1);
                listBox2.Items.Add(aktar2);
                for (int n = listBox2.Items.Count - 1; n >= 0; --n)
                {
                    string removelistitem = "all";
                    if (listBox2.Items[n].ToString().Contains(removelistitem))
                    {
                        listBox2.Items.RemoveAt(n);
                    }
                }
            }

            if (baglanti.State == ConnectionState.Closed)
                baglanti.Open();
            SqlConnection con = new SqlConnection(conString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = con;
            cmd.CommandText = "SELECT * FROM sepet";
            con.Open();
            SqlDataReader dr3 = cmd.ExecuteReader();

            int i = 0;
            int total = 0;
            while (dr3.Read())
            {
                listBox1.DisplayMember = "Text";
                listBox1.ValueMember = "Value";
                listBox1.Items.Insert(i,  new {Text= dr3["sepet_urun"].ToString(), Value = Convert.ToInt32(dr3["sepet_id"]) }) ;
                
                total = total + Convert.ToInt32(dr3["sepet_fyt"]);
                
            }
            label2.Text = total.ToString();
            dr3.Close();
            con.Close();

            baglanti.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 2;
           
        }
        int urunid = 0;
        private void button5_Click(object sender, EventArgs e)
        {

            tabControl1.SelectedIndex = 4;

            baglanti.Open();
            SqlDataAdapter da = new SqlDataAdapter("SELECT * FROM siparis", "server = serverAD; database = DBad; UID = IDad; password = sifre");
            DataSet ds = new DataSet();
            da.Fill(ds, "siparis");
            //dataGridView1.Columns.Remove("ref_urun_id");
            dataGridView1.DataSource = ds.Tables["siparis"].DefaultView;

            string icerik = string.Empty;
            for (int index = 0; index < listBox1.Items.Count; index++)
            {
                icerik += listBox1.Items[index].ToString() + ",";
            }
            string kayit = "insert into siparis( ref_mus_id, siparis_icerik, toplam_fiyat ) values (@refmusid,@siparisicerik,@toplamfiyat)";
            SqlCommand komut = new SqlCommand(kayit, baglanti);

            komut.Parameters.AddWithValue("@siparisicerik", icerik);

            komut.Parameters.AddWithValue("@refmusid", musteriid);
            komut.Parameters.AddWithValue("@toplamfiyat", label2.Text );
            komut.ExecuteNonQuery();
            baglanti.Close();
            MessageBox.Show("Kayıt Tamamlandı");
          

        }

        private void button6_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Siparişiniz oluşturuldu!");
            baglanti.Open();
            string silmeSorgusu = "DELETE from sepet";
            SqlCommand silKomutu = new SqlCommand(silmeSorgusu, baglanti);
            silKomutu.Parameters.AddWithValue("@sepetid", listBox1.Text);
            silKomutu.ExecuteNonQuery();
            baglanti.Close();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 4;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            MessageBox.Show("online ödeme sayfasına aktarılıyor...");
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox2.Items.Clear();
            comboBox2.Text = "";
            comboBox3.Items.Clear();
            comboBox3.Text = "";
            int selectedValue = Convert.ToInt32((comboBox1.SelectedItem as ComboboxItem).Value.ToString());
            SqlCommand sorgu = new SqlCommand(" SELECT * FROM Ilceler where SehirId=" + selectedValue + " ", baglanti);
            sorgu.ExecuteNonQuery();
            DataTable dt = new DataTable();
            SqlDataAdapter adp = new SqlDataAdapter(sorgu);
            adp.Fill(dt);

            foreach (DataRow dr in dt.Rows)
            {
                ComboboxItem item = new ComboboxItem
                {
                    Text = dr["IlceAdi"].ToString(),
                    Value = Convert.ToInt32(dr["ilceId"])
                };


                comboBox2.Items.Add(item);

            }
            textBox3.Text = string.Empty;
            textBox3.Text = comboBox1.Text + "," + comboBox2.Text + "," + comboBox3.Text;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {


            comboBox3.Items.Clear();

            //comboBox3.SelectedText = "";
            comboBox3.Text = "";
            if (baglanti.State == ConnectionState.Open)
                baglanti.Close();
            baglanti.Open();
            int v = Convert.ToInt32((comboBox2.SelectedItem as ComboboxItem).Value.ToString());
            int selectedValue = v;
            SqlCommand sorgu = new SqlCommand(" SELECT * FROM SemtMah where ilceId =" +
                " " + selectedValue + " "
                , baglanti);
            sorgu.ExecuteNonQuery();
            DataTable dt = new DataTable();
            SqlDataAdapter adp = new SqlDataAdapter(sorgu);
            adp.Fill(dt); SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sorgu);
            adp.Fill(dt);

            foreach (DataRow dr in dt.Rows)
            {

                ComboboxItem item = new ComboboxItem
                {
                    Text = dr["MahalleAdi"].ToString(),
                    Value = Convert.ToInt32(dr["SemtMahId"])
                };
                comboBox3.Items.Add(item);

            }
            textBox3.Text = string.Empty;
            textBox3.Text = comboBox1.Text + "," + comboBox2.Text + "," + comboBox3.Text;

        }


        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox3.Text = string.Empty;
            textBox3.Text = comboBox1.Text + "," + comboBox2.Text + "," + comboBox3.Text;
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 1;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 0;
        }


        private void button11_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 5;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 3;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
            object Missing = Type.Missing;
            Workbook workbook = excel.Workbooks.Add(Missing);
            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];
            int StartCol = 1;
            int StartRow = 1;
            for (int j = 0; j < dataGridView1.Columns.Count; j++)
            {
                Range myRange = (Range)sheet1.Cells[StartRow, StartCol + j];
                myRange.Value2 = dataGridView1.Columns[j].HeaderText;
            }
            StartRow++;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    Range myRange = (Range)sheet1.Cells[StartRow + i, StartCol + j];
                    myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                    myRange.Select();
                }
            }

        }

        
    }
}


