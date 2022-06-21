using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace exel_veri_çekme
{
    class Veri
    {
        private int personelID;
        private DateTime date;
        private int cihazID;
        private int mesai;
        private DateTime tarih;
        private TimeSpan saat;
        private int ay;
        private int yıl;
        private int gun;
        private string ad;


        private static List<Veri> veriler = new List<Veri>();

        public Veri(int PersonelID, string tarih, int CihazID, int mesai,string ad)
        {
            this.personelID = PersonelID;
            this.date = Convert.ToDateTime(tarih);
            this.cihazID = CihazID;
            this.mesai = mesai;
            this.tarih = date.Date;
            this.saat = date.TimeOfDay;
            this.ay = date.Month;
            this.yıl = date.Year;
            this.gun = date.Day;
            this.ad = ad;
        }
        public int PersonelID { get { return personelID; } set { personelID = value; } }
        public DateTime Date { get { return date; } set { date = value; tarih = date.Date; saat = date.TimeOfDay; } }
        public int CihazID { get { return cihazID; } set { cihazID = value; } }
        public int Mesai { get { return mesai; } set { mesai = value; } }
        public DateTime Tarih { get { return tarih; } }
        public TimeSpan Saat { get { return saat; } }
        public int Ay { get { return ay; } }
        public int Yıl { get { return yıl; } }
        public int Gun { get { return gun; } }
        public string Ad { get { return ad; } set { ad = value; } }

        public static List<Veri> Veriler()
        {
            return veriler;
        }
        public static void Ekle(Veri a)
        {
            veriler.Add(a);
        }

    }
    class veri1
    {
        public string Personel { get; set; }
        public int Ay { get; set; }

        public int Yıl { get; set; }

        public double ToplamMesai { get; set; }

        public double AyMesai { get; set; }

       public double Fark { get; set; }
        public double Maas { get; set; }
    }
    class veri2
    {
        public string Personel { get; set; }
        public int Ay { get; set; }

        public int Yıl { get; set; }

        public double GunlukMesai { get; set; }

        public DateTime Tarih { get; set; }    

    }
}

