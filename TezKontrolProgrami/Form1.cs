using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TezKontrolProgrami
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            label1.Visible = false;
            label2.Visible = false;
            checkBox1.Visible = false;
        }
        OpenFileDialog file = new OpenFileDialog();

        string Dosyayolu = "";
       
        
        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "Word Dosyası |*.docx";
            file.ShowDialog();
            Dosyayolu = file.FileName;
            Dosyayolu = Dosyayolu.Replace("\\", "/");
            textBox2.Text = Dosyayolu;
        } // dosyaları c# aktarma


        List<string> bilgi = new List<string>();
        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
            dataGridView1.Rows.Clear();
            richTextBox1.Text = "";
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            object miss = System.Reflection.Missing.Value;
            object path = textBox2.Text;
            object readOnly = true;
            Microsoft.Office.Interop.Word.Document docs = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
            string totaltext = "";
            dataGridView1.ColumnCount = 1;
            dataGridView1.Columns[0].Name = "Bilgiler";

            

            int j = 0;
            for (int i = 0; i < docs.Paragraphs.Count; i++)
            {
                
                if (docs.Paragraphs[i + 1].Range.Font.Size == 13 &&docs.Paragraphs[i + 1].Range.Font.Bold==-1 && docs.Paragraphs[i + 1].Range.Text.ToString() != "\r" && docs.Paragraphs[i + 1].Range.Text.ToString() != "\n")
                {
                    if (j == 0)// ilk gelen verinin kapak yazısı olmasını sağlama
                    {
                        dataGridView1.Rows.Add();
                        dataGridView1.Rows.Add();
                        dataGridView1.Rows[j].Cells[0].Value = "Kapak";
                        bilgi.Add(totaltext);
                        totaltext = "";
                        j++;
                        dataGridView1.Rows[j].Cells[0].Value = docs.Paragraphs[i + 1].Range.Text.ToString();
                        j++;
                    }
                    else // 2. gelen başlığın adını çekmek
                    {
                        bilgi.Add(totaltext);
                        totaltext = "";
                        dataGridView1.Rows.Add();
                        dataGridView1.Rows[j].Cells[0].Value = docs.Paragraphs[i + 1].Range.Text.ToString(); // başlığı datagridview aktar
                        j++;
                    }
                    
                }
                else // eğer gelen veri başlık değilse alt başlık olarak almayı sağlamak için kurulan yapı
                {
                    totaltext+= docs.Paragraphs[i + 1].Range.Text.ToString(); // alt başlığı çekme
                }
            }
            bilgi.Add(totaltext);
            MessageBox.Show("Dosya Aktarımı Başarıyla gerçekleşmiştir", "Bilgilendirme", MessageBoxButtons.OK, MessageBoxIcon.Information);
        } // dosya yolu seçme

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                richTextBox1.Text = "";
                richTextBox1.Text = bilgi[dataGridView1.CurrentRow.Index];
            }
            catch
            {
                MessageBox.Show("Boş Satır Seçtiniz Lütfen Dolu Satırlardan birini Seçin", "Bilgilendirme", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } 
        } // alt metinlere inme

        private void button3_Click(object sender, EventArgs e)
        {

            label1.Visible = true;
            label2.Visible = true;
            checkBox1.Visible = true;


            string kaynakca = AltBaslıgıBulma("KAYNAKLAR");
            int kaynaksayisi = Kaynaksayısı(kaynakca);
            label1.Text = kaynaksayisi.ToString() + "Kaynak Sayısı  ";

            //çift tırnak kontrolü
            int cifttirnak = 0;
            cifttirnak = ciftTirnak();
            label2.Text = "Çift tırnak içinde 50 den fazla kelime olan söz sayısı" + cifttirnak;


            //önsöz ilk paragrafında teşekkür kelimesi var mı ?
            string onsoz = AltBaslıgıBulma("ÖNSÖZ");
            bool Cevap = onsozTesekkur(onsoz);
            checkBox1.Enabled = false;
            checkBox1.Checked = Cevap;
           
        } // raporlama

        private bool onsozTesekkur(string onsoz)
        {
            bool cevap;
            string tut="";
            for(int i = 0; i < onsoz.Length; i++)
            {
                tut += onsoz[i];
                try
                {
                    if (onsoz[i] =='.' && onsoz[i+1]=='\r')
                    {
                        break;
                    }
                }
                catch
                {

                }
            }
            int sayi =tut.IndexOf("teşekkür", 0, tut.Length);

            if (sayi == -1)
                cevap = false;
            else
                cevap = true;
            return cevap;
        } // Onsoz ilk paragraf kontrol

        public int ciftTirnak()
        {
            int tirnakuygunmu = 0;
            foreach (string eleman in bilgi)
            {
                string Eleman;
                Eleman = eleman;
                int sayac = 0;
                bool giris = false;
                for(int i = 0; i < Eleman.Length; i++)
                {
                    

                    if (Eleman[i] == '“'||giris==true)
                    {
                        giris = true;
                        Console.WriteLine(Eleman[i]);
                        sayac++;
                    }
                    if (Eleman[i] == '”'||giris==false)
                    {
                        giris = false;
                        if (sayac > 50)
                        {
                            tirnakuygunmu++;
                        }
                        sayac = 0;
                    }
                }

            }
            return tirnakuygunmu;
        } //çift tırnak kontrol



        public string AltBaslıgıBulma(string baslık)
        {
            baslık += "\r";
            int index=0;
            for (int i = 0; i < dataGridView1.RowCount-1; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString() == baslık)
                {
                    index = i+1;
                }
            }
            int j = 0;
            string Eleman="";
            foreach (string eleman in bilgi)
            {
                j++;
                if (index == j)
                    Eleman = eleman;
                
            }
            return Eleman;




        } // alt başlıkları bulma 



        public int Kaynaksayısı(string kaynakca)
        {
            int sayac = 0;
            
            for(int i = 0; i < kaynakca.Length; i++)
            {
                if (kaynakca[i] == '-')
                {
                    if(char.IsDigit(kaynakca[i - 1]) == true)
                    {
                        if(kaynakca[ i + 1] == ')')
                        {
                            sayac++;
                        }
                    }
                    
                    
                }
            }
            return sayac;
        } // kaynak sayısını bulma
    }
}




