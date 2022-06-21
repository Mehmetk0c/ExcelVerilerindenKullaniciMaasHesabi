using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace exel_veri_çekme
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        DataTableCollection dtable;
        OleDbConnection data;
        public static double[] mesai = { 216.0, 198.0, 225.0, 198.0, 207.0, 184.5, 207.0, 225.0, 211.5, 184.5, 216.0, 211.5, 216.0, 198.0 };

        private void button1_Click(object sender, EventArgs e)
        {


            using (OpenFileDialog openFile = new OpenFileDialog() { Title = "Excel Dosyaları" })
            {
                if (openFile.ShowDialog() == DialogResult.OK)
                {

                    textBox1.Text = openFile.FileName;
                    data = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox1.Text + ";Extended Properties='Excel 12.0 Xml; HDR=YES';");
                }

            }
            Veriler();
        }

        void Veriler()
        {
            data.Open();
            OleDbDataAdapter vt = new OleDbDataAdapter("Select * from [Sayfa1$]", data);
            DataTable dtable = new DataTable();
            vt.Fill(dtable);

            foreach (DataRow row in dtable.Rows)
            {

                int personelid = Convert.ToInt32(row["IDPersonel"]);
                int cihazid = Convert.ToInt32(row["IDCihaz"]);
                string ldate = row["Saat"].ToString();
                int mesai = Convert.ToInt32(row["Sütun4"]);
                string ad = row["Ad"].ToString();
                Veri a = new Veri(personelid, ldate, cihazid, mesai, ad);
                Veri.Ekle(a);
            }


            data.Close();

            var query = Veri.Veriler().GroupBy(x => new { x.Ad, x.Gun, x.Ay, x.Yıl })
                .Select(y => new { Personel = y.Key.Ad, Gun = y.Key.Gun, Ay = y.Key.Ay, Yıl = y.Key.Yıl, Mesai = ((y.Max(a => a.Saat) - y.Min(a => a.Saat))) }).OrderBy(x => x.Personel)
                ;

            var query2 = query
                .GroupBy(x => new { x.Personel, x.Ay, x.Yıl })
                .Select(y => new veri1()
                {
                    Personel = y.Key.Personel,
                    Ay = y.Key.Ay,
                    Yıl = y.Key.Yıl,
                    ToplamMesai = Math.Round(new TimeSpan(y.Sum(a => a.Mesai.Ticks)).TotalHours, 3),
                     AyMesai = (y.Key.Yıl == 2018 ? mesai[11 + y.Key.Ay] : mesai[y.Key.Ay - 1]),
                    Fark= Math.Round(Math.Round(new TimeSpan(y.Sum(a => a.Mesai.Ticks)).TotalHours, 3) - (y.Key.Yıl == 2018 ? mesai[11 + y.Key.Ay] : mesai[y.Key.Ay - 1]), 3),
                    Maas = Math.Round(
                        Math.Round(Math.Round(new TimeSpan(y.Sum(a => a.Mesai.Ticks)).TotalHours, 3) - (y.Key.Yıl == 2018 ? mesai[11 + y.Key.Ay] : mesai[y.Key.Ay - 1]), 3) > 0 ?Math.Round(Math.Round(new TimeSpan(y.Sum(a => a.Mesai.Ticks)).TotalHours, 3) - (y.Key.Yıl == 2018 ? mesai[11 + y.Key.Ay] : mesai[y.Key.Ay - 1]), 3)*20+3000.00 :3000.00+ Math.Round(Math.Round(new TimeSpan(y.Sum(a => a.Mesai.Ticks)).TotalHours, 3) - (y.Key.Yıl == 2018 ? mesai[11 + y.Key.Ay] : mesai[y.Key.Ay - 1]), 3)*10, 2)

                })
                .AsEnumerable()
                ;

            var source = new BindingSource(query2.ToList(), null);

            dataGridView2.DataSource = new MySortableBindingList<veri1>(query2.ToArray());
            var results = dataGridView2.DataSource;


        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int index = dataGridView2.SelectedCells[0].RowIndex;
            textBox2.Text = dataGridView2.Rows[index].Cells[0].Value.ToString();
            textBox3.Text = dataGridView2.Rows[index].Cells[3].Value.ToString();


        }

        private void button3_Click(object sender, EventArgs e)
        {
            Veriler();
        }
        public string parcala(string parca)
            {
            string deger = "";
            string[] dizi = parca.Split(' ');
            for (int i = 0; i < dizi.Length; i++)
            {
                deger += dizi[i];
            }
            if (deger != "")
                return deger;
            else
                return parca;

            }

        private void button2_Click(object sender, EventArgs e)
        {
            DateTime  tarih= dateTimePicker1.Value.Date;

            DateTime tarih2 = dateTimePicker2.Value.Date;
            if (tarih <= tarih2)
            {
                if (textBox6.Text == "" || textBox6.Text == String.Empty)
                {
                    var query = Veri.Veriler().GroupBy(x => new { x.Ad, x.Gun, x.Ay, x.Yıl })
                    .Select(y => new { Personel = y.Key.Ad, Gun = y.Key.Gun, Ay = y.Key.Ay, Yıl = y.Key.Yıl, Mesai = ((y.Max(a => a.Saat) - y.Min(a => a.Saat))) }).OrderBy(x => x.Personel)
                    ;

                    var query2 = query
                        .GroupBy(x => new { x.Personel, x.Ay, x.Yıl, x.Gun })
                        .Select(y => new veri2()
                        {
                            Personel = y.Key.Personel,
                            Ay = y.Key.Ay,
                            Yıl = y.Key.Yıl,
                            GunlukMesai = Math.Round(new TimeSpan(y.Sum(a => a.Mesai.Ticks)).TotalHours, 3),
                            Tarih = new DateTime(y.Key.Yıl, y.Key.Ay, y.Key.Gun)


                        })
                        .Where(a => (a.Tarih.Date >= tarih && a.Tarih.Date <= tarih2)).OrderBy(x => x.Personel)
                        ;

                    dataGridView2.DataSource = new MySortableBindingList<veri2>(query2.ToArray());

                }
                else
                {
                    var query = Veri.Veriler().GroupBy(x => new { x.Ad, x.Gun, x.Ay, x.Yıl })
                    .Select(y => new { Personel = y.Key.Ad, Gun = y.Key.Gun, Ay = y.Key.Ay, Yıl = y.Key.Yıl, Mesai = ((y.Max(a => a.Saat) - y.Min(a => a.Saat))) }).OrderBy(x => x.Personel)
                    ;

                    var query2 = query
                        .GroupBy(x => new { x.Personel, x.Ay, x.Yıl, x.Gun })
                        .Select(y => new veri2()
                        {
                            Personel = y.Key.Personel,
                            Ay = y.Key.Ay,
                            Yıl = y.Key.Yıl,
                            GunlukMesai = Math.Round(new TimeSpan(y.Sum(a => a.Mesai.Ticks)).TotalHours, 3),
                            Tarih = new DateTime(y.Key.Yıl, y.Key.Ay, y.Key.Gun)


                        })
                        .Where(a => (a.Tarih.Date >= tarih && a.Tarih.Date <= tarih2) && parcala(a.Personel).ToLower() == parcala(textBox6.Text).ToLower()).OrderBy(x => x.Personel)
                        ;

                    dataGridView2.DataSource = new MySortableBindingList<veri2>(query2.ToArray());
                }
            }else if (tarih > tarih2) 
            {
                MessageBox.Show("Geçerli Tarih Giriniz!");
            }
        }
            

       
    }

  
}
