using cdn_api;
using Hydra;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static ZZnaCSV.Towary;

namespace ZZnaCSV
{
    public partial class AktualizacjaCen : Form
    {
        
        string connectionString;
        List<ListaTowar> lista;
        

        public AktualizacjaCen(List<ListaTowar> listaTowars)
        {
            
            
            try
            {
                if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 16) == "ERPXL_PoloMarket")
                    connectionString = @"Data Source = SRV1; Initial Catalog = ERPXL_PoloMarket; User ID = admin; Password=Dupa2003!";
                else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 22) == "ERPXL_Tuflowers_Mobile")
                    connectionString = @"Data Source = TF-SQL; Initial Catalog = ERPXL_Tuflowers_Mobile; User ID = onesupport; Password=Ancr3d0$$db";
                else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 24) == "ERPXL_Tuflowers_20220104")
                    connectionString = @"Data Source = TF-SQL; Initial Catalog = ERPXL_Tuflowers_20220104; User ID = onesupport; Password=Ancr3d0$$db";
                else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 24) == "ERPXL_Tuflowers_20220908")
                    connectionString = @"Data Source = TF-SQL; Initial Catalog = ERPXL_Tuflowers_20220908; User ID = onesupport; Password=Ancr3d0$$db";
                else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 24) == "ERPXL_Tuflowers_20220331")
                    connectionString = @"Data Source = TF-SQL; Initial Catalog = ERPXL_Tuflowers_20220331; User ID = onesupport; Password=Ancr3d0$$db";
                else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 15) == "ERPXL_Tuflowers")
                    connectionString = @"Data Source = TF-SQL; Initial Catalog = ERPXL_Tuflowers; User ID = onesupport; Password=Ancr3d0$$db";
                //connectionString = @"Data Source = TF-SQL; Initial Catalog = ERPXL_Tuflowers_Mobile;persist security info=True; Integrated Security=SSPI;";
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message, "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {

                InitializeComponent();
                numericUpDown1.Value = 5;
                numericUpDown2.Value = 7;
                numericUpDown3.Value = Convert.ToDecimal(0.3);
                numericUpDown4.Value = 55;
                numericUpDown5.Value = 12;
                string[] stringRodzaje = new[] { "VANS", "HURTOWA", "MASTER" };
                Rodzaje.Items.AddRange(stringRodzaje);
                lista = listaTowars;
                dodajTowary(listaTowars);
            }
            catch (Exception e)
            {
                MessageBox.Show("Błąd w AktualizacjaCen.cs: "+e.Message+"/n"+e.StackTrace, "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (tabela1.RowCount == 1)
                {
                    MessageBox.Show("Brak towarów do wyświetlenia", "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Dispose(true);
                  
                }

        }

       



        private void AktualizacjaCen_Load(object sender, EventArgs e)
        {

        }

        private void buttonBufor_Click(object sender, EventArgs e)
        {
            try
            {
                for (int counter = 0; counter < (tabela1.Rows.Count) - 1; counter++)
                {
                
                    tabela1.Rows[counter].Cells["BuforPro"].Value = Convert.ToDouble(numericUpDown1.Value);
                    tabela1.Rows[counter].Cells["Bufor"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value) * Convert.ToDouble(numericUpDown1.Value) / 100 + Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value), 2);
                    tabela1.Rows[counter].Cells["Odpad"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["Bufor"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["OdpadPro"].Value) / 100 + Convert.ToDouble(tabela1.Rows[counter].Cells["Bufor"].Value), 2);
                    tabela1.Rows[counter].Cells["CenaTraColumn"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["Odpad"].Value) + Convert.ToDouble(tabela1.Rows[counter].Cells["Transport"].Value);
                    tabela1.Rows[counter].Cells["WartoscTraColumn"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaTraColumn"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["IloscColumn"].Value);
                    if (Convert.ToString(tabela1.Rows[counter].Cells["NarzutPro"].Value) != "")
                    {
                        tabela1.Rows[counter].Cells["CenaNetto"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaTraColumn"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["NarzutPro"].Value) / 100 + Convert.ToDouble(tabela1.Rows[counter].Cells["CenaTraColumn"].Value), 2);
                    }
                    else if(Convert.ToString(tabela1.Rows[counter].Cells["NarzutZl"].Value) != "")
                    {
                        tabela1.Rows[counter].Cells["CenaNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaTraColumn"].Value) + Convert.ToDouble(tabela1.Rows[counter].Cells["NarzutZl"].Value);
                    }
                    else
                        tabela1.Rows[counter].Cells["CenaNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaTraColumn"].Value);

                    tabela1.Rows[counter].Cells["WartoscCenaColumn"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["IloscColumn"].Value), 2);
                    tabela1.Rows[counter].Cells["Brutto"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["VatPro"].Value) / 100 + Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value), 2);
                   
                    if (Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value) != 0)
                    {
                        tabela1.Rows[counter].Cells["CenaNetto2"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value)/(1+ Convert.ToDouble(tabela1.Rows[counter].Cells["VatPro"].Value) / 100),2);
                        tabela1.Rows[counter].Cells["MarzaPro"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) * 100, 2);
                        tabela1.Rows[counter].Cells["ZyskNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value);
                        tabela1.Rows[counter].Cells["Przecena"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["ZyskNetto"].Value);

                    }
                    else
                    {
                        tabela1.Rows[counter].Cells["ZyskNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value);
                        tabela1.Rows[counter].Cells["Przecena"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["ZyskNetto"].Value);
                        tabela1.Rows[counter].Cells["MarzaPro"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) * 100, 2);
                        tabela1.Rows[counter].Cells["CenaNetto2"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value);

                    }
                    tabela1.Rows[counter].Cells["CenaRabat"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["RabatPro"].Value) / 100, 2);
                    tabela1.Rows[counter].Cells["MarzaRabat"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[counter].Cells["CenaRabat"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[counter].Cells["CenaRabat"].Value) * 100, 2);
               
                }
            }
            catch (Exception en)
            {
                MessageBox.Show("Błędne obliczenia dla zmiany wartości Buforu: (" + en.Message + ")\nSprawdź, czy użyto przecinka jako separatora dziesiętnego");
            }

        }

        private void buttonOdpad_Click(object sender, EventArgs e)
        {
            try { 
            for (int counter = 0; counter < (tabela1.Rows.Count) - 1; counter++)
            {
                tabela1.Rows[counter].Cells["OdpadPro"].Value = Convert.ToDouble(numericUpDown2.Value);
                tabela1.Rows[counter].Cells["Odpad"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["Bufor"].Value) * Convert.ToDouble(numericUpDown2.Value) / 100 + Convert.ToDouble(tabela1.Rows[counter].Cells["Bufor"].Value), 2);
                tabela1.Rows[counter].Cells["CenaTraColumn"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["Odpad"].Value) + Convert.ToDouble(tabela1.Rows[counter].Cells["Transport"].Value);
                tabela1.Rows[counter].Cells["WartoscTraColumn"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaTraColumn"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["IloscColumn"].Value), 2);
                if (Convert.ToString(tabela1.Rows[counter].Cells["NarzutPro"].Value) != "")
                {
                    tabela1.Rows[counter].Cells["CenaNetto"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaTraColumn"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["NarzutPro"].Value) / 100 + Convert.ToDouble(tabela1.Rows[counter].Cells["CenaTraColumn"].Value), 2);
                }
                else if (Convert.ToString(tabela1.Rows[counter].Cells["NarzutZl"].Value) != "")
                {
                    tabela1.Rows[counter].Cells["CenaNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaTraColumn"].Value) + Convert.ToDouble(tabela1.Rows[counter].Cells["NarzutZl"].Value);
                }
                else
                    tabela1.Rows[counter].Cells["CenaNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaTraColumn"].Value);
                tabela1.Rows[counter].Cells["WartoscCenaColumn"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["IloscColumn"].Value), 2);
                tabela1.Rows[counter].Cells["Brutto"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["VatPro"].Value) / 100 + Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value), 2);
                //tabela1.Rows[counter].Cells["ZyskNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value);
                //tabela1.Rows[counter].Cells["Przecena"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["ZyskNetto"].Value);
                //tabela1.Rows[counter].Cells["MarzaPro"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) * 100, 2);
                if (Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value) != 0)
                {
                       // tabela1.Rows[counter].Cells["MarzaPro"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value) * 100, 2);
                        tabela1.Rows[counter].Cells["CenaNetto2"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value) / (1 + Convert.ToDouble(tabela1.Rows[counter].Cells["VatPro"].Value) / 100), 2);
                        //tabela1.Rows[counter].Cells["ZyskNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value);
                        tabela1.Rows[counter].Cells["MarzaPro"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) * 100, 2);
                        tabela1.Rows[counter].Cells["ZyskNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value);
                        tabela1.Rows[counter].Cells["Przecena"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["ZyskNetto"].Value);
                    }
                else
                {
                        tabela1.Rows[counter].Cells["ZyskNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value);
                        tabela1.Rows[counter].Cells["Przecena"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["ZyskNetto"].Value);
                        tabela1.Rows[counter].Cells["MarzaPro"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) * 100, 2);
                        tabela1.Rows[counter].Cells["CenaNetto2"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value);
                }
                tabela1.Rows[counter].Cells["CenaRabat"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["RabatPro"].Value) / 100, 2);
                tabela1.Rows[counter].Cells["MarzaRabat"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[counter].Cells["CenaRabat"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[counter].Cells["CenaRabat"].Value) * 100, 2);

            }
            }
            catch (Exception en)
            {
                MessageBox.Show("Błędne obliczenia dla zmiany wartości Odpadu: (" + en.Message + ")\nSprawdź, czy użyto przecinka jako separatora dziesiętnego");
            }
        }

        private void buttonTrans_Click(object sender, EventArgs e)
        {
            try { 
            for (int counter = 0; counter < (tabela1.Rows.Count) - 1; counter++)
            {
                tabela1.Rows[counter].Cells["Transport"].Value = Convert.ToDouble(numericUpDown3.Value);
                tabela1.Rows[counter].Cells["CenaTraColumn"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["Odpad"].Value) + Convert.ToDouble(numericUpDown3.Value);
                tabela1.Rows[counter].Cells["WartoscTraColumn"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaTraColumn"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["IloscColumn"].Value);
                if (Convert.ToString(tabela1.Rows[counter].Cells["NarzutPro"].Value) != "")
                {
                    tabela1.Rows[counter].Cells["CenaNetto"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaTraColumn"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["NarzutPro"].Value) / 100 + Convert.ToDouble(tabela1.Rows[counter].Cells["CenaTraColumn"].Value), 2);
                }
                else if (Convert.ToString(tabela1.Rows[counter].Cells["NarzutZl"].Value) != "")
                {
                    tabela1.Rows[counter].Cells["CenaNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaTraColumn"].Value) + Convert.ToDouble(tabela1.Rows[counter].Cells["NarzutZl"].Value);
                }
                else
                    tabela1.Rows[counter].Cells["CenaNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaTraColumn"].Value);
                tabela1.Rows[counter].Cells["WartoscCenaColumn"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["IloscColumn"].Value), 2);
                tabela1.Rows[counter].Cells["Brutto"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["VatPro"].Value) / 100 + Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value), 2);
                //tabela1.Rows[counter].Cells["ZyskNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value);
                //tabela1.Rows[counter].Cells["Przecena"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["ZyskNetto"].Value);
                //tabela1.Rows[counter].Cells["MarzaPro"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) * 100, 2);
                if (Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value) != 0)
                {
                        //tabela1.Rows[counter].Cells["MarzaPro"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value) * 100, 2);
                        tabela1.Rows[counter].Cells["CenaNetto2"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value) / (1 + Convert.ToDouble(tabela1.Rows[counter].Cells["VatPro"].Value) / 100), 2);
                        //tabela1.Rows[counter].Cells["ZyskNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value);
                        tabela1.Rows[counter].Cells["MarzaPro"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) * 100, 2);
                        tabela1.Rows[counter].Cells["ZyskNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value);
                        tabela1.Rows[counter].Cells["Przecena"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["ZyskNetto"].Value);
                    }
                else
                {
                        tabela1.Rows[counter].Cells["ZyskNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value);
                        tabela1.Rows[counter].Cells["Przecena"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["ZyskNetto"].Value);
                        tabela1.Rows[counter].Cells["MarzaPro"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) * 100, 2);
                        tabela1.Rows[counter].Cells["CenaNetto2"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value);
                }
                tabela1.Rows[counter].Cells["CenaRabat"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["RabatPro"].Value) / 100, 2);
                tabela1.Rows[counter].Cells["MarzaRabat"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[counter].Cells["CenaRabat"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[counter].Cells["CenaRabat"].Value) * 100, 2);

            }
            }
            catch (Exception en)
            {
                MessageBox.Show("Błędne obliczenia dla zmiany wartości Transportu: (" + en.Message + ")\nSprawdź, czy użyto przecinka jako separatora dziesiętnego");
            }
        }

        private void buttonNarzut_Click(object sender, EventArgs e)
        {
            try { 
            for (int counter = 0; counter < (tabela1.Rows.Count) - 1; counter++)
            {
                if (Convert.ToString(tabela1.Rows[counter].Cells["NarzutPro"].Value) != "")
                {
                    tabela1.Rows[counter].Cells["NarzutPro"].Value = Convert.ToDouble(numericUpDown4.Value);
                    tabela1.Rows[counter].Cells["CenaNetto"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaTraColumn"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["NarzutPro"].Value) / 100 + Convert.ToDouble(tabela1.Rows[counter].Cells["CenaTraColumn"].Value), 2);
                }
                tabela1.Rows[counter].Cells["WartoscCenaColumn"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["IloscColumn"].Value), 2);
                tabela1.Rows[counter].Cells["Brutto"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["VatPro"].Value) / 100 + Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value), 2);
               //tabela1.Rows[counter].Cells["ZyskNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value);
                //tabela1.Rows[counter].Cells["Przecena"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["ZyskNetto"].Value);
                //tabela1.Rows[counter].Cells["MarzaPro"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) * 100, 2);
                if (Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value) != 0)
                {
                        //tabela1.Rows[counter].Cells["MarzaPro"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value) * 100, 2);
                        tabela1.Rows[counter].Cells["CenaNetto2"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value) / (1 + Convert.ToDouble(tabela1.Rows[counter].Cells["VatPro"].Value) / 100), 2);
                        //tabela1.Rows[counter].Cells["ZyskNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value);
                        tabela1.Rows[counter].Cells["MarzaPro"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) * 100, 2);
                        tabela1.Rows[counter].Cells["ZyskNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value);
                        tabela1.Rows[counter].Cells["Przecena"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["ZyskNetto"].Value);
                    }
                else
                {
                        tabela1.Rows[counter].Cells["ZyskNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value);
                        tabela1.Rows[counter].Cells["Przecena"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["ZyskNetto"].Value);
                        tabela1.Rows[counter].Cells["MarzaPro"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) * 100, 2);
                        tabela1.Rows[counter].Cells["CenaNetto2"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value);
                }
                tabela1.Rows[counter].Cells["CenaRabat"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["RabatPro"].Value) / 100, 2);
            }
            }
            catch (Exception en)
            {
                MessageBox.Show("Błędne obliczenia dla zmiany wartości Narzutu %: (" + en.Message + ")\nSprawdź, czy użyto przecinka jako separatora dziesiętnego");
            }
        }

        private void buttonRabat_Click(object sender, EventArgs e)
        {
            try { 
            for (int counter = 0; counter < (tabela1.Rows.Count) - 1; counter++)
            {
                tabela1.Rows[counter].Cells["RabatPro"].Value = Convert.ToDouble(numericUpDown5.Value);
                tabela1.Rows[counter].Cells["CenaRabat"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["RabatPro"].Value) / 100, 2);
                tabela1.Rows[counter].Cells["MarzaRabat"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[counter].Cells["CenaRabat"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[counter].Cells["CenaRabat"].Value) * 100, 2);
            }
            }
            catch (Exception en)
            {
                MessageBox.Show("Błędne obliczenia dla zmiany wartości Rabatu: (" + en.Message + ")\nSprawdź, czy użyto przecinka jako separatora dziesiętnego");
            }
        }

        

        void dodajTowary(List<ListaTowar> listaTowars)
        {
            string NazwaPelna = "";
            
                ArrayList row;
                foreach (var item in listaTowars)
                {
                try
                    {
                    string AEN = "SELECT [Twr_Ean] FROM CDN.TwrKarty WHERE Twr_GIDTyp=" + item.GIDTyp + " and Twr_GIDFirma=" + item.GIDFirma + " and  Twr_GIDNumer=" + item.GIDNumer + " and Twr_GIDLp=" + item.GIDLp;
                    string Kod = "SELECT [Twr_Kod] FROM CDN.TwrKarty WHERE Twr_GIDTyp=" + item.GIDTyp + " and Twr_GIDFirma=" + item.GIDFirma + " and  Twr_GIDNumer=" + item.GIDNumer + " and Twr_GIDLp=" + item.GIDLp;
                    string Nazwa = "SELECT [Twr_Nazwa] FROM CDN.TwrKarty WHERE Twr_GIDTyp=" + item.GIDTyp + " and Twr_GIDFirma=" + item.GIDFirma + " and  Twr_GIDNumer=" + item.GIDNumer + " and Twr_GIDLp=" + item.GIDLp;
                    string Ilosc = "SELECT SUM([TwZ_Ilosc]) FROM CDN.TwrZasoby WHERE TwZ_TwrTyp=" + item.GIDTyp + " and TwZ_TwrFirma=" + item.GIDFirma + " and  TwZ_TwrNumer=" + item.GIDNumer + " and TwZ_TwrLp=" + item.GIDLp;
                    //string cenaZakup = "SELECT SUM(TwZ_KsiegowaNetto)/SUM([TwZ_Ilosc]) FROM CDN.TwrZasoby WHERE TwZ_TwrTyp=" + item.GIDTyp + " and TwZ_TwrFirma=" + item.GIDFirma + " and  TwZ_TwrNumer=" + item.GIDNumer + " and TwZ_TwrLp=" + item.GIDLp;
                    string cenaZakup = "SELECT TOP(1) (TwZ_KsiegowaNetto)/([TwZ_Ilosc]) FROM CDN.TwrZasoby WHERE TwZ_TwrTyp=" + item.GIDTyp + " and TwZ_TwrFirma=" + item.GIDFirma + " and  TwZ_TwrNumer=" + item.GIDNumer + " and TwZ_TwrLp=" + item.GIDLp + " ORDER BY TwZ_DataP DESC";
                    string VAT = "SELECT [Twr_StawkaPodSpr] FROM [CDN].[TwrKarty] "
                                + "WHERE Twr_GIDTyp = " + item.GIDTyp + " and Twr_GIDFirma = " + item.GIDFirma + " and Twr_GIDNumer = " + item.GIDNumer + " and Twr_GIDLp = " + item.GIDLp;
                    string Rabat = "if((SELECT [Twr_BezRabatu] FROM CDN.TwrKarty WHERE Twr_GIDTyp=" + item.GIDTyp + " and Twr_GIDFirma=" + item.GIDFirma + " and  Twr_GIDNumer=" + item.GIDNumer + " and Twr_GIDLp=" + item.GIDLp+")= 1) SELECT 'true' else SELECT 'false' ";
                    string Flagi = "SELECT TAP_Flagi FROM cdn.TwrAplikacje WHERE TAP_ObiTyp=" + item.GIDTyp + " and  TAP_ObiNumer = " + item.GIDNumer;
                    int[] i_flagi = czyPromocja(ConnectToSQL(connectionString, Flagi));
                    NazwaPelna = ConnectToSQL(connectionString, Nazwa);
                    double dIlosc = Convert.ToDouble(ConnectToSQL(connectionString, Ilosc));
                    double dCenaZakup = Convert.ToDouble(ConnectToSQL(connectionString, cenaZakup));
                    double dVat= Convert.ToDouble(ConnectToSQL(connectionString, VAT));
                    double dBufor = dCenaZakup * 5.00 / 100 + dCenaZakup;
                    double dOdpad = 7.00 / 100 * dBufor + dBufor;
                    double dZakupTrans = 0.3 + dOdpad;
                    double dWartoscCena = dIlosc * dZakupTrans;
                    string NarzutPro = "";
                    string NarzutZl = "";
                    znajdzNarzut(ref NarzutPro, ref NarzutZl, item.GIDTyp, item.GIDNumer, item.GIDFirma, item.GIDLp);
                    double cenaNetto = 0.0;

                    if (NarzutPro != "")
                    {
                        cenaNetto = dZakupTrans * Convert.ToDouble(NarzutPro) / 100 + dZakupTrans;
                    }
                    else
                    {
                        if (NarzutZl == "")
                            cenaNetto = dZakupTrans;
                        else
                            cenaNetto = dZakupTrans + Convert.ToDouble(NarzutZl);
                    }
                    double dWartoscNetto = cenaNetto * dIlosc;
                    double dBrutto = dVat / 100 * cenaNetto + cenaNetto;
                    double dCenaReczna = 0.0;
                    double dZysk = 0.0;//cenaNetto - dZakupTrans;
                    double dMarza = 0.0;//(cenaNetto - dZakupTrans) / cenaNetto * 100;
                    double dCenaNetto = 0.0;
                    if (dCenaReczna != 0)
                    {
                        dCenaNetto = dCenaReczna;
                    }
                    else
                    {
                        dZysk = cenaNetto - dCenaZakup;
                        dMarza = (cenaNetto - dCenaZakup) / cenaNetto * 100;
                        dCenaNetto = cenaNetto;
                    }
                    double dCenaRabat = dCenaNetto - dCenaNetto * 12 / 100;
                    double dMarzaRabat = (dCenaRabat - dCenaZakup) / dCenaRabat * 100;
                    row = new ArrayList();
                    row.Add(ConnectToSQL(connectionString, Kod));
                    row.Add(ConnectToSQL(connectionString, Nazwa));
                    row.Add(dIlosc);
                    row.Add(Math.Round(dCenaZakup, 2));
                    row.Add(5.00);
                    row.Add(Math.Round(dBufor, 2));
                    row.Add(Math.Round(dOdpad, 2));
                    row.Add(Math.Round(dZakupTrans, 2));
                    row.Add(NarzutPro);
                    row.Add(NarzutZl);
                    row.Add(Math.Round(cenaNetto, 2));
                    row.Add(dVat);
                    row.Add(Math.Round(dBrutto, 2));
                    row.Add(0.00);
                    row.Add(Math.Round(dZysk, 2));
                    row.Add(Math.Round(dZysk, 2));
                    row.Add(Math.Round(dMarza, 2));
                    row.Add(Math.Round(dCenaNetto, 2));
                    row.Add(Convert.ToBoolean(ConnectToSQL(connectionString, Rabat)));
                    row.Add(12);
                    row.Add(Math.Round(dCenaRabat, 2));
                    row.Add(Math.Round(dMarzaRabat, 2));
                    row.Add(Convert.ToBoolean(i_flagi[0]));
                    row.Add(ConnectToSQL(connectionString, AEN));
                    row.Add(7.00);
                    row.Add(0.3);
                    row.Add(Math.Round(dWartoscCena, 2));
                    row.Add(Math.Round(dWartoscNetto, 2));
                    row.Add(i_flagi[1]);
                    tabela1.Rows.Add(row.ToArray());
                    tabela1.Rows[tabela1.RowCount - 1].DefaultCellStyle.BackColor = System.Drawing.Color.White;

                }
                catch (Exception e)
                {

                    MessageBox.Show("Błędne dane dla towaru: " + NazwaPelna + " (" + e.Message + ")");



                }
            }

            
           
            
        }

        public void znajdzNarzut( ref string NarzutPro, ref string NarzutZl,  int GIDTyp, int GIDNumer, int GIDFirma, short GIDLp)
        {
            int lp = 0;
            do
            {
                string zlQuery = "SELECT [Atr_Wartosc] FROM [CDN].[Atrybuty] WHERE Atr_AtkId = (SELECT[AtK_ID] FROM[CDN].[AtrybutyKlasy] WHERE AtK_Nazwa = 'Domyślny Narzut Zł') and [Atr_ObiTyp]=" + GIDTyp + " and [Atr_ObiFirma]=" + GIDFirma + " and [Atr_ObiNumer]=" + GIDNumer + " and [Atr_ObiLp]= " +lp ;
                string proQuery = "SELECT [Atr_Wartosc] FROM [CDN].[Atrybuty] WHERE Atr_AtkId = (SELECT[AtK_ID] FROM[CDN].[AtrybutyKlasy] WHERE AtK_Nazwa = 'Domyślny Narzut %') and [Atr_ObiTyp]=" + GIDTyp + " and [Atr_ObiFirma]=" + GIDFirma + " and [Atr_ObiNumer]=" + GIDNumer + " and [Atr_ObiLp]= "+lp;
                NarzutPro = ConnectToSQL(connectionString, proQuery);
                NarzutZl = ConnectToSQL(connectionString, zlQuery);

                string typQuery = "SELECT [TwG_GrOTyp] FROM [CDN].[TwrGrupy] WHERE TwG_GIDTyp = "+GIDTyp+" and TwG_GIDFirma = "+GIDFirma+" and TwG_GIDNumer = "+GIDNumer+ " and TwG_GIDLp="+ GIDLp;
                string firmaQuery = "SELECT [TwG_GrOFirma] FROM [CDN].[TwrGrupy] WHERE TwG_GIDTyp = " + GIDTyp + " and TwG_GIDFirma = " + GIDFirma + " and TwG_GIDNumer = " + GIDNumer + " and TwG_GIDLp=" + GIDLp;
                string numerQuery = "SELECT [TwG_GrONumer] FROM [CDN].[TwrGrupy] WHERE TwG_GIDTyp = " + GIDTyp + " and TwG_GIDFirma = " + GIDFirma + " and TwG_GIDNumer = " + GIDNumer + " and TwG_GIDLp=" + GIDLp;
                string lpQuery = "SELECT [TwG_GrOLp] FROM [CDN].[TwrGrupy] WHERE TwG_GIDTyp = " + GIDTyp + " and TwG_GIDFirma = " + GIDFirma + " and TwG_GIDNumer = " + GIDNumer + " and TwG_GIDLp=" + GIDLp;

                
                if (ConnectToSQL(connectionString, typQuery) == "")
                    break;
                GIDTyp = Convert.ToInt32(ConnectToSQL(connectionString, typQuery));
                GIDNumer = Convert.ToInt32(ConnectToSQL(connectionString, numerQuery));
                GIDFirma = Convert.ToInt32(ConnectToSQL(connectionString, firmaQuery));
                GIDLp = Convert.ToInt16(ConnectToSQL(connectionString, lpQuery));
                //MessageBox.Show(proQuery + " Wynik: " + NarzutPro);
                lp = 1;

            } while (NarzutZl == "" && NarzutPro == "");
        }

        public int[] czyPromocja(string flagi)
        {
            ushort f =Convert.ToUInt16(flagi);
            //Byte b = Convert.ToByte(f);
            int czyProm = (f & (1 << 2));
            if(czyProm==4)
            {
                f -= 4;
            }
            //MessageBox.Show("flagi -4: " + f);
            return new int[] { czyProm,Convert.ToInt32(f)};
        }

        public string ConnectToSQL(string connection_string, string query)
        {
            string str = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, "Query");
                    if (dataSet.Tables.Count == 0)
                    {
                        return "";
                    }
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                    {
                        str = row[0].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                str = str + "Bład SQL:" + ex.ToString() + "\n";
            }
            return str;
        }

        private void buttonZapisz_Click(object sender, EventArgs e)
        {
            // logowanie();
            // zmianaCeny();
            // cdn_api.cdn_api.XLLogout(1);
            int lp = 0;
            string cell = "";
            int cennikID = 0;
            if (Convert.ToString(Rodzaje.SelectedItem) == "" )
            {
                MessageBox.Show("Nie wybrano rodzaju ceny", "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if(Convert.ToString(Rodzaje.SelectedItem) == "VANS")
            {
                lp = 1;
                cell = "Brutto";
                cennikID = 1;
                zapisz(lp, cell, cennikID);
            }
            else if(Convert.ToString(Rodzaje.SelectedItem) == "HURTOWA")
            {
                lp = 2;
                cell = "Brutto";
                cennikID = 2;
                zapisz(lp, cell, cennikID);
            }
            else if(Convert.ToString(Rodzaje.SelectedItem)=="MASTER")
            {
                lp = 3;
                cell = "Brutto";
                cennikID = 17;
                zapisz(lp, cell, cennikID);

            }

            //ActiveForm.Close();
            //Dispose(true);
            //Application.Exit();
           
        }
        
        void zapisz(int lp, string cell, int cennikID)
        {
            string date1 = "1800-12-28";
            string date2 = DateTime.Now.ToString("yyyy-MM-dd");
            string query5 = "SELECT DATEDIFF(day, '" + date1 + "','" + date2 + "'); ";
            int data = Convert.ToInt32(ConnectToSQL(connectionString, query5));

            string tempCell = cell;
            for (int counter = 0; counter < (tabela1.Rows.Count) - 1; counter++)
            {
                cell = tempCell;
                if (/*lp == 1 &&*/ Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value) != 0)
                    cell = "CenaReczna";
                // string czyIstnieje = "SELECT [TwC_Wartosc] FROM [CDN].[TwrCeny] WHERE TwC_TwrTyp=" + lista[counter].GIDTyp + " and TwC_TwrFirma=" + lista[counter].GIDFirma + " and  TwC_TwrNumer=" + lista[counter].GIDNumer + " and TwC_TwrLp="+lp +" and TwC_TcnId="+cennikID;
                string czyIstnieje = "SELECT [TwC_Wartosc] FROM [CDN].[TwrCeny] join CDN.TwrKarty on TwC_TwrTyp=Twr_GIDTyp and TwC_TwrNumer = Twr_GIDNumer and TwC_TwrFirma=Twr_GIDFirma WHERE Twr_Kod = '" + Convert.ToString(tabela1.Rows[counter].Cells["kod"].Value)+" ' and TwC_TwrLp=" + lp + " and TwC_TcnId=" + cennikID;

                if (Convert.ToString(ConnectToSQL(connectionString, czyIstnieje)) == "")
                {
                    //insertInto
                     string insert = "INSERT INTO [CDN].[TwrCeny] ([TwC_TwrTyp],[TwC_TwrFirma],[TwC_TwrNumer],[TwC_TwrLp],[TwC_Waluta],[TwC_NrKursu],[TwC_Marza],[TwC_Zaok] ,[TwC_Offset] ,[TwC_RodzajPodstawy] ,[TwC_Wartosc] ,[TwC_Aktualizacja] ,[TwC_Priorytet],[TwC_CzasModyfikacji],[TwC_KosztyDodatkowe],[TwC_DokTyp],[TwC_DokFirma],[TwC_DokNumer],[TwC_DokLp],[TwC_TcnId],[TwC_DataOd],[TwC_KntTyp] ,[TwC_KntFirma] ,[TwC_KntNumer]  ,[TwC_KntLp]  ,[TwC_Aktywny])" +
                    "VALUES (" + lista[counter].GIDTyp + " , " + lista[counter].GIDFirma + " , " + lista[counter].GIDNumer + " , "+lp+ " , 'PLN' , 1 , 0, 0.0001, 0, 0, " + Convert.ToString(tabela1.Rows[counter].Cells[cell].Value).Replace(",",".") + ", 0, 0,"+data+", 0, 0, 0, 0, 0, "+cennikID+","+data+",0, 0, 0, 0, 0)";
                    InsertInto(insert);
                   // zmianaCeny(lp, cell, cennikID);
                }
                else
                {
                    //Update
                    //string update = "UPDATE [CDN].[TwrCeny] SET [TwC_Wartosc] = " + Convert.ToString(tabela1.Rows[counter].Cells[cell].Value).Replace(",", ".") + " WHERE TwC_TwrTyp=" + lista[counter].GIDTyp + " and TwC_TwrFirma=" + lista[counter].GIDFirma + " and  TwC_TwrNumer=" + lista[counter].GIDNumer + " and TwC_TwrLp="+lp+" and TwC_TcnId=" + cennikID;
                    string update = "UPDATE t SET [TwC_Wartosc] = " + Convert.ToString(tabela1.Rows[counter].Cells[cell].Value).Replace(",", ".") + ", TwC_CzasModyfikacji=datediff(second,'19900101',getdate()) FROM CDN.TwrCeny as t join CDN.TwrKarty on TwC_TwrTyp = Twr_GIDTyp and TwC_TwrNumer = Twr_GIDNumer and TwC_TwrFirma = Twr_GIDFirma WHERE Twr_Kod = '" + Convert.ToString(tabela1.Rows[counter].Cells["kod"].Value) + "' and TwC_TwrLp=" + lp + " and TwC_TcnId=" + cennikID;
                    Update(update, counter, cell);
                }

                try
                {
                    int Rabat = 0;
                    if (Convert.ToBoolean(tabela1.Rows[counter].Cells["CheckRabat"].Value) == true)
                        Rabat = 1;
                    string updateRabatSQL = "UPDATE cdn.TwrKarty set Twr_BezRabatu=" + Rabat + ", Twr_LastModL = datediff(second,'19900101',getdate()) WHERE Twr_Kod = '" + Convert.ToString(tabela1.Rows[counter].Cells["kod"].Value) + "'";
                    UpdateRabat(updateRabatSQL);
                }
                catch(Exception e)
                {
                    MessageBox.Show("Błąd zmiany Twr_BezRabatu: "+e.Message ,"Komunikat",
                  MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                try
                {
                    int promocja = 0;
                    if (Convert.ToBoolean(tabela1.Rows[counter].Cells["CheckPromocja"].Value) == true)
                        promocja = 4;
                    int flagi = Convert.ToInt32(tabela1.Rows[counter].Cells["Flagi"].Value) + promocja;
                    string updateFlagiSQL = "UPDATE cdn.TwrAplikacje set TAP_Flagi=" + flagi + ", TAP_CzasModyfikacji=datediff(second,'19900101',getdate()) WHERE TAP_ObiTyp=" + lista[counter].GIDTyp + " and  TAP_ObiNumer = " + lista[counter].GIDNumer;
                    UpdateRabat(updateFlagiSQL);
                }
                catch (Exception e)
                {
                    MessageBox.Show("Błąd zmiany TAP_Flagi: " + e.Message, "Komunikat",
                  MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            MessageBox.Show("Zapisano nowe ceny", "Komunikat",
                  MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /*void logowanie()
        {
            string XLSQLServer = "192.168.1.202";
            string XLSQLDatabase = "Tuflowers";//"ERPXL_PoloMarket"
            string XLUser = "ADMIN"; //
            string XLPass = "NE!w0rk12!";
            XLLoginInfo_20202 nowyUzytkownik = new XLLoginInfo_20202();
            nowyUzytkownik.Wersja = 20202;
            nowyUzytkownik.Winieta = 1;
            nowyUzytkownik.TrybWsadowy = 1;
            nowyUzytkownik.UtworzWlasnaSesje = 1;
            nowyUzytkownik.SerwerKlucza = XLSQLServer;
            nowyUzytkownik.Baza = XLSQLDatabase;
            nowyUzytkownik.OpeIdent = XLUser;
            nowyUzytkownik.OpeHaslo = XLPass;
            //nowyUzytkownik.PortHTTPSerweraKlucza = 1;
            //nowyUzytkownik.RodzajSerweraKlucza = 0;
            // nowyUzytkownik.SerwerKlucza = "SRV1::5000139268";
           // nowyUzytkownik.ProgramID = "ONES004 - Zmiana Cen";

            int sesja = 1;
            int blad = cdn_api.cdn_api.XLLogin(nowyUzytkownik, ref sesja);
            if (blad != 0)
            {
                MessageBox.Show("Błąd logowania: " + blad, "Komunikat",
                   MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            
        }

        void zmianaCeny(int lp, string cell, int cennikID)
        {
            logowanie();
            for (int counter = 0; counter < (tabela1.Rows.Count) - 1; counter++)
            {
                XLCenaInfo_20202 cena = new XLCenaInfo_20202();
                cena.Wersja = 20202;
                cena.GIDFirma = lista[counter].GIDFirma;
                cena.GIDLp = lp;
                cena.GIDNumer = lista[counter].GIDNumer;
                cena.GIDTyp = lista[counter].GIDTyp;
                //cena.Numer = 1;
                //cena.Marza = Convert.ToString(tabela1.Rows[counter].Cells["MarzaPro"].Value);
                cena.Wartosc = Convert.ToString(tabela1.Rows[counter].Cells[cell].Value);
                cena.TcnId = cennikID;
                
                int blad=cdn_api.cdn_api.XLDodajCene(1,cena);
                if (blad != 0)
                {
                    MessageBox.Show("Błąd zmiany ceny: " + blad+ " dla "+ Convert.ToString(tabela1.Rows[counter].Cells["Nazwa"].Value), "Komunikat",
                       MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            cdn_api.cdn_api.XLLogout(1);

        }*/

        public string Update(string sql, int counter, string cell)
        {
            try
            {
                using (var connection = new SqlConnection(connectionString))
                {
                    using (var command = new SqlCommand(sql, connection))
                    {
                        command.Parameters.Add("@wartosc", SqlDbType.Decimal).Value = Convert.ToDouble(tabela1.Rows[counter].Cells[cell].Value);
                        
                        // repeat for all variables....
                        connection.Open();
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return sql;
        }

        public string UpdateRabat(string sql)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    connection.Open();
                    adapter.UpdateCommand = new SqlCommand(sql, connection);
                    adapter.UpdateCommand.ExecuteNonQuery();
                    //MessageBox.Show("Row updated !! ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return sql;
        }
        public string InsertInto(string sql)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    connection.Open();
                    adapter.InsertCommand = new SqlCommand(sql, connection);
                    adapter.InsertCommand.ExecuteNonQuery();
                    //MessageBox.Show("Row inserted !! ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString() +"\n"+ ex.InnerException, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return sql;
        }

        private void tabela1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try { 
            tabela1[e.ColumnIndex, e.RowIndex].Style
               .SelectionBackColor = Color.Blue;
                tabela1[e.ColumnIndex, e.RowIndex].Value= Convert.ToString(tabela1[e.ColumnIndex, e.RowIndex].Value).Replace(".", ",");
            tabela1.Rows[e.RowIndex].Cells["Bufor"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["Zakup"].Value) * Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["BuforPro"].Value) / 100 + Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["Zakup"].Value), 2);
            tabela1.Rows[e.RowIndex].Cells["Odpad"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["Bufor"].Value) * Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["OdpadPro"].Value) / 100 + Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["Bufor"].Value), 2);
            tabela1.Rows[e.RowIndex].Cells["CenaTraColumn"].Value = Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["Odpad"].Value) + Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["Transport"].Value);
            tabela1.Rows[e.RowIndex].Cells["WartoscTraColumn"].Value = Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["CenaTraColumn"].Value) * Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["IloscColumn"].Value);
            if (Convert.ToString(tabela1.Rows[e.RowIndex].Cells["NarzutPro"].Value) != "")
            {
                tabela1.Rows[e.RowIndex].Cells["CenaNetto"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["CenaTraColumn"].Value) * Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["NarzutPro"].Value) / 100 + Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["CenaTraColumn"].Value), 2);
            }
            else if (Convert.ToString(tabela1.Rows[e.RowIndex].Cells["NarzutZl"].Value) != "")
            {
                tabela1.Rows[e.RowIndex].Cells["CenaNetto"].Value = Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["CenaTraColumn"].Value) + Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["NarzutZl"].Value);
            }
            else
                tabela1.Rows[e.RowIndex].Cells["CenaNetto"].Value = Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["CenaTraColumn"].Value);
            tabela1.Rows[e.RowIndex].Cells["WartoscCenaColumn"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["CenaNetto"].Value) * Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["IloscColumn"].Value), 2);
            tabela1.Rows[e.RowIndex].Cells["Brutto"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["CenaNetto"].Value) * Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["VatPro"].Value) / 100 + Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["CenaNetto"].Value), 2);
           
            if (Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["CenaReczna"].Value) != 0)
            {
                    tabela1.Rows[e.RowIndex].Cells["CenaNetto2"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["CenaReczna"].Value) / (1 + Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["VatPro"].Value) / 100), 2);
                    tabela1.Rows[e.RowIndex].Cells["MarzaPro"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["CenaNetto2"].Value) - Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["CenaNetto2"].Value) * 100, 2);
                    tabela1.Rows[e.RowIndex].Cells["ZyskNetto"].Value = Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["CenaNetto2"].Value) - Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["Zakup"].Value);
                    tabela1.Rows[e.RowIndex].Cells["Przecena"].Value = Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["ZyskNetto"].Value);
                }
            else
            {
                    tabela1.Rows[e.RowIndex].Cells["ZyskNetto"].Value = Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["CenaNetto"].Value) - Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["Zakup"].Value);
                    tabela1.Rows[e.RowIndex].Cells["Przecena"].Value = Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["ZyskNetto"].Value);
                    tabela1.Rows[e.RowIndex].Cells["MarzaPro"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["CenaNetto"].Value) - Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["CenaNetto"].Value) * 100, 2);
                    tabela1.Rows[e.RowIndex].Cells["CenaNetto2"].Value = Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["CenaNetto"].Value);
            }
            tabela1.Rows[e.RowIndex].Cells["CenaRabat"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["CenaNetto2"].Value) - Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["CenaNetto2"].Value) * Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["RabatPro"].Value) / 100, 2);
            tabela1.Rows[e.RowIndex].Cells["MarzaRabat"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["CenaRabat"].Value) - Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[e.RowIndex].Cells["CenaRabat"].Value) * 100, 2);

            }
            catch (Exception en)
            {
                MessageBox.Show( en.Message + "\nSprawdź, czy użyto przecinka jako separatora dziesiętnego");
            }
        }

        private void buttonNarzutZl_Click(object sender, EventArgs e)
        {
            try { 
            for (int counter = 0; counter < (tabela1.Rows.Count) - 1; counter++)
            {
                if ( Convert.ToString(tabela1.Rows[counter].Cells["NarzutPro"].Value) == "" || Convert.ToDouble(tabela1.Rows[counter].Cells["NarzutPro"].Value) == 0)
                {
                    tabela1.Rows[counter].Cells["NarzutZl"].Value = Convert.ToDouble(numericUpDown6.Value);
                    tabela1.Rows[counter].Cells["CenaNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaTraColumn"].Value) + Convert.ToDouble(tabela1.Rows[counter].Cells["NarzutZl"].Value);
                    tabela1.Rows[counter].Cells["WartoscCenaColumn"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["IloscColumn"].Value), 2);
                    tabela1.Rows[counter].Cells["Brutto"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["VatPro"].Value) / 100 + Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value), 2);
                    
                    if (Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value) != 0)
                    {
                            //tabela1.Rows[counter].Cells["MarzaPro"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value) * 100, 2);
                            tabela1.Rows[counter].Cells["CenaNetto2"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value);
                            //tabela1.Rows[counter].Cells["ZyskNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaReczna"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value);
                            tabela1.Rows[counter].Cells["MarzaPro"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) * 100, 2);
                            tabela1.Rows[counter].Cells["ZyskNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value);
                            tabela1.Rows[counter].Cells["Przecena"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["ZyskNetto"].Value);
                        }
                    else
                    {
                            tabela1.Rows[counter].Cells["ZyskNetto"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value);
                            tabela1.Rows[counter].Cells["Przecena"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["ZyskNetto"].Value);
                            tabela1.Rows[counter].Cells["MarzaPro"].Value = Math.Round((Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["Zakup"].Value)) / Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value) * 100, 2);
                            tabela1.Rows[counter].Cells["CenaNetto2"].Value = Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto"].Value);
                    }
                    tabela1.Rows[counter].Cells["CenaRabat"].Value = Math.Round(Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) - Convert.ToDouble(tabela1.Rows[counter].Cells["CenaNetto2"].Value) * Convert.ToDouble(tabela1.Rows[counter].Cells["RabatPro"].Value) / 100, 2);
                }
            }
            }
            catch (Exception en)
            {
                MessageBox.Show("Błędne obliczenia dla zmiany wartości Narzutu Zł: (" + en.Message + ")\nSprawdź, czy użyto przecinka jako separatora dziesiętnego");
            }
        }
    }
}
