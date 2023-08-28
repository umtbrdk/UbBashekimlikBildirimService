using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace UbBashekimlikBildirimService
{
    public partial class UbBashekimlikBildirimService : ServiceBase
    {
        public UbBashekimlikBildirimService()
        {
            InitializeComponent();
        }
        string dbAdres, dbSifre, dbKullAdi;
        string _gonderen = ""; string _pass = ""; string _port = ""; string _ssl = ""; string _host = ""; string _name = "";string _hastaneAdi;
        int _timerSure = 10;
        Timer timerGunluk;
        DateTime kontrolGun;
        bool gunlukKontrol = false;
        bool haftalikKontrol = false;


        protected override void OnStart(string[] args)
        {
            LogKlasorCreate();
            MailListOku();

            BaglantiAyarlari();

           
            MailAyarlariOkuSQL();

            kontrolGun = DateTime.Now;

            timerGunluk = new Timer();
            timerGunluk.Interval = _timerSure * 1000; // 10 saniye
            timerGunluk.Elapsed += new ElapsedEventHandler(TmrGunlukKontrol);
            timerGunluk.Enabled = true;

            LogInsert("ServisStart", "Servis çalışmaya başladı. Süre : " + _timerSure + " saniye.. MAIL_WL_SMTP_TIMER key ile değiştirebilirsiniz. Ver.23/05 13:25");
        }
        private void TmrGunlukKontrol(object sender, EventArgs e)
        {
            if (kontrolGun.ToShortDateString() != DateTime.Now.ToShortDateString())
            {
                gunlukKontrol = false;
                haftalikKontrol = false;
                kontrolGun = DateTime.Now;
            }

            DateTime bugunGunluk = DateTime.Now;

            if (bugunGunluk.ToShortTimeString() == "09:00" && gunlukKontrol == false)
            {
                gunlukKontrol = true;
                GunlukRaporSQL();
            }

            if (bugunGunluk.DayOfWeek == DayOfWeek.Saturday && bugunGunluk.ToShortTimeString() == "10:00" && haftalikKontrol == false )
            {
                haftalikKontrol = true;
                HaftalikRaporSQL();
            }
        }

     
        protected override void OnStop()
        {
            LogInsert("ServisStop", "Servis Durduruldu");
        }

        void GunlukRaporSQL()
        {
            gunlukKontrol = true;

            string bugun = DateTime.Now.AddDays(-1).ToShortDateString();
            string htmlTable = "<h3>Sayın yetkili; </h3>";
            htmlTable += "<h3>" + _hastaneAdi + " " + bugun + " tarihi 00:00 ile 23:59 saatleri arasındaki veriler aşağıdaki gibidir.</h3>";
            htmlTable += "<table border='1'><tr>";

            string cmdtxt = @"select ROWNUM SIRA_NO, UPPER(ISLEM) ISLEM, ADET, 
            CASE WHEN to_char(CIRO,'FM999G999G999D00L', 'NLS_NUMERIC_CHARACTERS = '',.''') = ',00TL' THEN ''
            ELSE to_char(CIRO,'FM999G999G999D00L', 'NLS_NUMERIC_CHARACTERS = '',.''') END ""CIRO TL"" FROM(
            select sira_no, islem, count(adet) adet, sum(ciro) ciro from hastane.vw_bashekimlik_istatistik
            where tarih BETWEEN TRUNC(SYSDATE-1) AND TRUNC(SYSDATE-1) +INTERVAL '23' HOUR + INTERVAL '59' MINUTE
            group by sira_no, islem
            order by sira_no)";

            try
            {
                using (OracleConnection conn = new OracleConnection(Database.connstr))
                using (OracleCommand cmd = new OracleCommand(cmdtxt, conn))
                {
                    conn.Open();

                    using (OracleDataReader reader = cmd.ExecuteReader())
                    {
                        System.Data.DataTable dataTable = new System.Data.DataTable();
                        dataTable.Load(reader);

                        foreach (DataColumn column in dataTable.Columns)
                        {
                            htmlTable += "<th>" + column.ColumnName + "</th>";
                        }
                        htmlTable += "</tr>";
                        foreach (DataRow row in dataTable.Rows)
                        {
                            htmlTable += "<tr>";
                            foreach (DataColumn column in dataTable.Columns)
                            {
                                htmlTable += "<td>" + row[column] + "</td>";
                            }
                            htmlTable += "</tr>";
                        }
                        htmlTable += "</table>";
                        htmlTable += "<p> </p>";
                        htmlTable += "<p><em>Not: Bu özet rapor HBYS Birimi tarafından otomatik hazırlanıp tarafınıza gönderilmektedir. Problem olduğunu düşünüyorsanız hbys@tinaztepe.com üzerinden irtibata geçebilirsiniz.</em></p>";

                    }
                    MailGonder(htmlTable,bugun, bugun, "G");
                    cmd.Dispose();
                    conn.Close();
                    LogInsert("GunlukRaporSonuc","Günlük Rapor başarı ile gönderildi.");
                }
            }
            catch (Exception ex)
            {
                LogInsert("GunlukRaporSQL Hata : ", ex.Message);
            }
        }
        void HaftalikRaporSQL()
        {
            haftalikKontrol = true;

            string tar1 = DateTime.Now.AddDays(-7).ToShortDateString();
            string tar2 = DateTime.Now.AddDays(-1).ToShortDateString();

            string htmlTable = "<h3>Sayın yetkili; </h3>";
            htmlTable += "<h3>"+ _hastaneAdi + " " + tar1 + " 00:00 ile " + tar2 + " 23:59 saatleri arasındaki veriler aşağıdaki gibidir.</h3>";
            htmlTable += "<table border='1'><tr>";

            string cmdtxt = @"select ROWNUM SIRA_NO, UPPER(ISLEM) ISLEM, ADET, 
            CASE WHEN to_char(CIRO,'FM999G999G999D00L', 'NLS_NUMERIC_CHARACTERS = '',.''') = ',00TL' THEN ''
            ELSE to_char(CIRO,'FM999G999G999D00L', 'NLS_NUMERIC_CHARACTERS = '',.''') END ""CIRO TL"" FROM(
            select sira_no, islem, count(adet) adet, sum(ciro) ciro from hastane.vw_bashekimlik_istatistik
            where tarih BETWEEN TRUNC(SYSDATE) -7 AND TRUNC(SYSDATE) - INTERVAL '1' DAY + INTERVAL '23:59:59' HOUR TO SECOND
            group by sira_no, islem
            order by sira_no)";

            try
            {
                using (OracleConnection conn = new OracleConnection(Database.connstr))
                using (OracleCommand cmd = new OracleCommand(cmdtxt, conn))
                {
                    conn.Open();

                    using (OracleDataReader reader = cmd.ExecuteReader())
                    {
                        System.Data.DataTable dataTable = new System.Data.DataTable();
                        dataTable.Load(reader);

                        foreach (DataColumn column in dataTable.Columns)
                        {
                            htmlTable += "<th>" + column.ColumnName + "</th>";
                        }
                        htmlTable += "</tr>";
                        foreach (DataRow row in dataTable.Rows)
                        {
                            htmlTable += "<tr>";
                            foreach (DataColumn column in dataTable.Columns)
                            {
                                htmlTable += "<td>" + row[column] + "</td>";
                            }
                            htmlTable += "</tr>";
                        }
                        htmlTable += "</table>";
                        htmlTable += "<p> </p>";
                        htmlTable += "<p><em>Not: Bu özet rapor HBYS Birimi tarafından otomatik hazırlanıp tarafınıza gönderilmektedir. Problem olduğunu düşünüyorsanız hbys@tinaztepe.com üzerinden irtibata geçebilirsiniz.</em></p>";
                        htmlTable = htmlTable.Replace("BUGUN YATISI YAPILAN", "BU HAFTA YATISI YAPILAN");
                    }
                    MailGonder(htmlTable,tar1, tar2, "H");
                    cmd.Dispose();
                    conn.Close();
                    LogInsert("GunlukRaporSonuc", "Haftalık Rapor başarı ile gönderildi.");
                }
            }
            catch (Exception ex)
            {
                LogInsert("HaftalikRaporSQL Hata : ", ex.Message);
            }
        }

        private void MailGonder(string html, string tarih1, string tarih2, string type)
        {
            try
            {
                string subj = "";
                if (type == "G")
                {
                    subj = tarih1 + " Tarihli " + _hastaneAdi + " Günlük Özet";
                }
                if (type == "H")
                {
                    subj = tarih1 + " ile " + tarih2 + " Arası " + _hastaneAdi + " Haftalık Özet";
                }
                MailMessage mail = new MailMessage();
                SmtpClient smtpClient = new SmtpClient(_host, Convert.ToInt32(_port));
                smtpClient.Credentials = new NetworkCredential(_gonderen, _pass);
                smtpClient.EnableSsl = false;
                mail.From = new MailAddress(_gonderen, "Yönetim Bilgilendirme Servisi");

                foreach (string ePostaAdres in gelenVeri)
                {
                    if (ePostaAdres.Contains("@"))
                    {
                        mail.To.Add(ePostaAdres);
                    }
                }

                //mail.To.Add("zehra.ozirmakli@tinaztepe.com");
                //mail.To.Add("ecetonguc.kockesen@tinaztepe.com");
                //mail.To.Add("serap.uluirmak@tinaztepe.com");

                mail.Subject = subj;
                mail.Body = html;
                mail.IsBodyHtml = true;
                smtpClient.Send(mail);
            }
            catch (Exception ex)
            {
                LogInsert("MailGonder Hata : ", ex.Message);
            }
            
        }
        void LogKlasorCreate()
        {
            string logKlasor = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOG");
            if (!Directory.Exists(logKlasor))
            {
                Directory.CreateDirectory(logKlasor);
                LogInsert("Klasör Oluşturma", "LOG klasörü oluşturuldu.");
            }
        }
        void LogInsert(string baslik, string msj)
        {
            //LogKlasorCreate();
            string time = DateTime.Now.ToShortDateString();
            string log = baslik + ";" + msj + ";" + DateTime.Now;

            string directory = AppDomain.CurrentDomain.BaseDirectory + @"\LOG";
            string filePath = Path.Combine(directory, "UbBashekimlikBildirimService_" + time + ".log");

            if (!File.Exists(filePath))
            {
                using (StreamWriter writer = new StreamWriter(filePath))
                {
                    writer.WriteLine(log);
                }
            }
            else // Dosya varsa altına tarih eklenir
            {
                using (StreamWriter writer = File.AppendText(filePath))
                {
                    writer.WriteLine(log);
                }
            }
        }


        /// <summary>
        /// ayarlar
        /// </summary>
        public void BaglantiAyarlari() // bağlantı ayarları
        {
            foreach (string mail in gelenVeri)
            {
                if (mail.Contains("@"))
                {
                    LogInsert("Otomatik Mail Gönderilecek Adres", mail);
                }
            }
            string directory = AppDomain.CurrentDomain.BaseDirectory;
            string filePath = Path.Combine(directory, "UbBashekimlikBildirimService.ini");

            // Dosya yoksa oluşturulur ve tarih yazılır
            if (!File.Exists(filePath))
            {
                LogInsert("BağlantıAyarları", "UbBashekimlikBildirimService.ini bulunamadı. Yüklü olduğu klasöre örnek dosya oluşturuluyor.. Lütfen düzenleyin.");
                using (StreamWriter writer = new StreamWriter(filePath))
                {
                    writer.Write("DB=" + "172.0.0.1:1521/orcl" + ";");
                    writer.Write("KullAdi=" + "enabiz" + ";");
                    writer.Write("KullSifre=" + "enabiz");
                }
                this.Stop();
            }
            else
            {
                string veri = "";
                try
                {
                    veri = File.ReadAllText(filePath);
                    //using (StreamReader sr = new StreamReader("UBMailService.ini"))
                    //{
                    //    string line;
                    //    while ((line = sr.ReadLine()) != null)
                    //    {
                    //        veri = line;
                    //    }
                    //    sr.Close();
                    //}
                }
                catch (Exception ex)
                {
                    LogInsert("Veri : ", ex.Message);

                }

                LogInsert("Veri : ", veri);
                dbAdres = veri.Substring(3, veri.IndexOf(";") - 3);
                dbSifre = veri.Substring(veri.IndexOf("KullSifre") + 10, veri.Length - (veri.IndexOf("KullSifre") + 10));
                dbKullAdi = veri.Substring(veri.IndexOf("KullAdi=") + 8, (veri.LastIndexOf(";")) - (veri.IndexOf("KullAdi=") + 8));
                Database.ConnStr(dbAdres, dbSifre, dbKullAdi);
                LogInsert("BaglantiAyarlari", "Bağlantı ayarları okundu." + Database.connstr);
            }
        }

        void MailAyarlariOkuSQL()
        {
            string cmdtxtGonderen = @"select deger from hastane.hastanekey where key = 'MAIL_WL_SMTP_GONDEREN'";
            string cmdtxtPass = @"select deger from hastane.hastanekey where key = 'MAIL_WL_SMTP_PASS'";
            string cmdtxtPort = @"select deger from hastane.hastanekey where key = 'MAIL_WL_SMTP_PORT'";
            string cmdtxtSSL = @"select deger from hastane.hastanekey where key = 'MAIL_WL_SMTP_SSL'";
            string cmdtxtHost = @"select deger from hastane.hastanekey where key = 'MAIL_WL_SMTP_HOST'";
            string cmdtxtName = @"select deger from hastane.hastanekey where key = 'MAIL_WL_SMTP_NAME'";
            string cmdtxtTimer = @"select deger from hastane.hastanekey where key = 'MAIL_WL_SMTP_TIMER'";
            string cmdtxtHastane = @"SELECT initcap(HASTANE_ADI) hastane_adi FROM HASTANE.HASTANE WHERE HASTANE_NO = 1";

            try
            {
                using (OracleConnection conn = new OracleConnection(Database.connstr))
                using (OracleCommand cmd = new OracleCommand(cmdtxtGonderen, conn))
                {
                    conn.Open();

                    // reader is IDisposable and should be closed
                    using (OracleDataReader dr = cmd.ExecuteReader())
                    {
                        //List<String> items = new List<String>();
                        if (dr.HasRows)
                        {
                            while (dr.Read())
                            {
                                _gonderen = (dr.GetString(0));
                            }
                        }
                    }
                }

                using (OracleConnection conn = new OracleConnection(Database.connstr))
                using (OracleCommand cmd = new OracleCommand(cmdtxtPass, conn))
                {
                    conn.Open();

                    // reader is IDisposable and should be closed
                    using (OracleDataReader dr = cmd.ExecuteReader())
                    {
                        //List<String> items = new List<String>();
                        if (dr.HasRows)
                        {
                            while (dr.Read())
                            {
                                _pass = (dr.GetString(0));
                            }
                        }
                    }

                }

                using (OracleConnection conn = new OracleConnection(Database.connstr))
                using (OracleCommand cmd = new OracleCommand(cmdtxtPort, conn))
                {
                    conn.Open();

                    // reader is IDisposable and should be closed
                    using (OracleDataReader dr = cmd.ExecuteReader())
                    {
                        //List<String> items = new List<String>();
                        if (dr.HasRows)
                        {
                            while (dr.Read())
                            {
                                _port = (dr.GetString(0));
                            }
                        }
                    }
                }

                using (OracleConnection conn = new OracleConnection(Database.connstr))
                using (OracleCommand cmd = new OracleCommand(cmdtxtSSL, conn))
                {
                    conn.Open();

                    // reader is IDisposable and should be closed
                    using (OracleDataReader dr = cmd.ExecuteReader())
                    {
                        //List<String> items = new List<String>();
                        if (dr.HasRows)
                        {
                            while (dr.Read())
                            {
                                _ssl = (dr.GetString(0));
                            }
                        }
                    }
                }

                using (OracleConnection conn = new OracleConnection(Database.connstr))
                using (OracleCommand cmd = new OracleCommand(cmdtxtHost, conn))
                {
                    conn.Open();

                    // reader is IDisposable and should be closed
                    using (OracleDataReader dr = cmd.ExecuteReader())
                    {
                        //List<String> items = new List<String>();
                        if (dr.HasRows)
                        {
                            while (dr.Read())
                            {
                                _host = (dr.GetString(0));
                            }
                        }
                    }
                }

                using (OracleConnection conn = new OracleConnection(Database.connstr))
                using (OracleCommand cmd = new OracleCommand(cmdtxtName, conn))
                {
                    conn.Open();

                    // reader is IDisposable and should be closed
                    using (OracleDataReader dr = cmd.ExecuteReader())
                    {
                        //List<String> items = new List<String>();
                        if (dr.HasRows)
                        {
                            while (dr.Read())
                            {
                                _name = (dr.GetString(0));
                            }
                        }
                    }
                }

                using (OracleConnection conn = new OracleConnection(Database.connstr))
                using (OracleCommand cmd = new OracleCommand(cmdtxtTimer, conn))
                {
                    conn.Open();

                    // reader is IDisposable and should be closed
                    using (OracleDataReader dr = cmd.ExecuteReader())
                    {
                        //List<String> items = new List<String>();
                        if (dr.HasRows)
                        {
                            while (dr.Read())
                            {
                                _timerSure = Convert.ToInt32((dr.GetString(0)));
                            }
                        }
                    }
                }

                using (OracleConnection conn = new OracleConnection(Database.connstr))
                using (OracleCommand cmd = new OracleCommand(cmdtxtHastane, conn))
                {
                    conn.Open();

                    // reader is IDisposable and should be closed
                    using (OracleDataReader dr = cmd.ExecuteReader())
                    {
                        //List<String> items = new List<String>();
                        if (dr.HasRows)
                        {
                            while (dr.Read())
                            {
                                _hastaneAdi = (dr.GetString(0));
                            }
                        }
                    }
                }
                LogInsert("Key okundu : ", _host + _gonderen + _pass + _port + _ssl + _name + _timerSure + _hastaneAdi);
            }
            catch (Exception ex)
            {
                LogInsert("Key Okuma Hatası : ", ex.Message);
                this.Stop();
            }
        }
        string[] gelenVeri;

        void MailListOku()
        {
            try
            {
                string directory = AppDomain.CurrentDomain.BaseDirectory;
                string filePath = Path.Combine(directory, "UbBildirimMailList.ini");

                // Dosya yoksa oluşturulur ve tarih yazılır
                if (!File.Exists(filePath))
                {
                    LogInsert("UbBildirimMailList", "UbBildirimMailList.ini bulunamadı. Yüklü olduğu klasöre örnek dosya oluşturuluyor.. Lütfen düzenleyin.");
                    using (StreamWriter writer = new StreamWriter(filePath))
                    {
                        writer.Write("umtbrdk@yahoo.com");
                    }
                    this.Stop();
                }
                else
                {
                    gelenVeri = File.ReadAllLines(filePath);
                    string mailList = "";

                    foreach (string mail in gelenVeri)
                    {
                        if (mail.Contains("@"))
                        {
                            mailList += mail + ";";
                        }
                    }
                    LogInsert("UbBildirimMailList", "Mail listesi okundu : " + mailList);
                }
            }
            catch (Exception ex)
            {
                LogInsert("UbBildirimMailList", "Mail listesi Hata : " + ex.Message);
            }

        }
    }
}
