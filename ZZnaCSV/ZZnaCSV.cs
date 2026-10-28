using Hydra;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;
using ZZnaCSV;
using System.Collections.Generic;
using static ZZnaCSV.Towary;
using System.Text;
using System;

[assembly: CallbackAssemblyDescription(Name: "ONE-004 Ceny",
                                           Description: "Aktualizacja cen",
                                           Author: "OneSupport",
                                           Version: "1.0",
                                           SystemVersion: "2021.0.2",
                                           Created: "21-08-2021")]
namespace ZZnaCVS
{
    [SubscribeProcedure(Procedures._TwrListaGrup, "Aktualizacja cen")]
    public class AktualizujCene : Callback
    {
        private ClaWindow RapButton;
        private ClaWindow zakladkaTabWewnetrzne;
        private ClaWindow filtrBtn;
        private List<ListaTowar> listaTowars = new List<ListaTowar>();
        int listID;
        public override void Init()
        {
            
            AddSubscription(true, 0, Events.JustAfterWindowOpening, new TakeEventDelegate(OnOpenWindow));
            AddSubscription(false, 0, Events.ResizeWindow, new TakeEventDelegate(Resize));
        }



        private bool OnOpenWindow(Procedures ProcId, int ControlId, Events Event)
        {

            ClaWindow Parent = GetWindow();
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            Rectangle WndBnd = GetWindow().Bounds;
            zakladkaTabWewnetrzne = GetWindow().AllChildren["?TabKody"];
            RapButton = zakladkaTabWewnetrzne.Children.Add(ControlTypes.button);
            RapButton.Visible = true;
            //RapButton.TextRaw = "Ceny";
            RapButton.ToolTipRaw = "Aktualizuj ceny";
            RapButton.IconRaw = "C:\\OneSupport\\money.ico";
            //RapButton.IconRaw = "C:\\Program Files (x86)\\Comarch ERP XL 2021.0\\money.ico";
        

            filtrBtn = GetWindow().AllChildren["?ButtonZamienniki"];                            //pobiera współrzędne przycisku zapisz
            RapButton.Bounds = new Rectangle(filtrBtn.Bounds.X+18, filtrBtn.Bounds.Y, 20, 15);
            listID = GetWindow().AllChildren["?List"].Id;      //tu sprawdzic
            
            RapButton.OnAfterMouseDown += new TakeEventDelegate(BtnOnMouseDown);

            //pobrane z pola w zakładce Raporty
            return true;
        }

        public bool Resize(Procedures ProcId, int ControlId, Events Event)
        {
            filtrBtn = GetWindow().AllChildren["?ButtonZamienniki"];                            //pobiera współrzędne przycisku zapisz
            RapButton.Bounds = new Rectangle(filtrBtn.Bounds.X + 18, filtrBtn.Bounds.Y, 20, 15);

            return true;
        }

        public bool BtnOnMouseDown(Procedures ProcId, int ControlId, Events Event)
        {

            try
            {

                ADODB._Recordset recordset = Runtime.WindowController.GetQueueMarked((int)ProcId, listID, GetCallbackThread());
                // if(recordset!=null)
                // MessageBox.Show(Convert.ToString(listID), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Towary(recordset);
                /* foreach(var item in listaTowars)
                 {
                     MessageBox.Show("GIDTyp "+Convert.ToString(item.GIDTyp), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                     MessageBox.Show("GIDFirma " + Convert.ToString(item.GIDFirma), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                     MessageBox.Show("GIDNumer " + Convert.ToString(item.GIDNumer), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                     MessageBox.Show("GIDLp " + Convert.ToString(item.GIDLp), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                 }*/
                //Application.Run(new AktualizacjaCen(listaTowars));
                AktualizacjaCen aktualizator = new AktualizacjaCen(listaTowars);
                aktualizator.ShowDialog();
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Błąd aktualizacji danych: " + ex.Message +"\n"+ex.StackTrace, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            listaTowars.Clear();

            return true;
        }

        bool Towary(ADODB._Recordset recordset)
        {
            
            try
            {
                StringBuilder dokumenty = new StringBuilder();
                if (recordset != null)
                {
                    int count = recordset.RecordCount;

                    if (count != 0)
                    {
                        recordset.MoveFirst();

                        while (recordset.EOF == false)
                        {
                            ADODB.Fields fields = recordset.Fields;

                            listaTowars.Add(new ListaTowar()
                            {
                                GIDTyp = Convert.ToInt16(fields[0].Value),
                                GIDFirma = Convert.ToInt32(fields[1].Value),
                                GIDNumer = Convert.ToInt32(fields[2].Value),
                                GIDLp = Convert.ToInt16(fields[3].Value),
                            });
                            recordset.MoveNext();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Nie zaznaczono towarów. Operacja anulowana.", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return true;
                    }
                }
                else
                {
                    MessageBox.Show("Nie zaznaczono towarów. Operacja anulowana.", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Błąd pobierania zaznaczonych towarów --> " + ex.Message, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return true;
            }
        }
        #region
        public override void Cleanup()
        {
            
        }
    }
}

#endregion