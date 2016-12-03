using Excel;
using Novacode;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.Threading;

namespace Massen_Mailer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ProgrammPfad = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            Zufallsgenerator = new Random();
            BereitsVersendet = false;
            SMTPServerIP = "...";//Hier IP des SMTP Servers angeben
        }
        private Mail[] AlleMails { get; set; }
        private string SMTPServerIP { get; set; }
        internal string ProgrammPfad { get; set; }
        private List<long> AlleZeiten { get; set; }
        private Random Zufallsgenerator { get; set; }
        private bool TestEmailAdresseFehler { get; set; }
        private bool BereitsVersendet { get; set; }
        internal List<string>[] AlleFehler { get; set; }
        private void TestButton_Click(object sender, EventArgs e)
        {
            try
            {
                new Thread(delegate ()
                {
                    OberflächeDeaktivieren(false);
                    var TestEmailAnzahl = 3;
                    var Fehler = string.Empty;
                    if (AlleMails.Length < 3)
                    {
                        TestEmailAnzahl = AlleMails.Length;
                    }
                    if (!AlleMailsÜberprüfen())
                    {
                        MessageBox.Show("Es sind bei der Überprüfung der Daten Fehler vorgekommen. Diese wurden in folgende Datei geschrieben:\n\n" + ProgrammPfad + @"\EmailAdressenFehler.txt" + "\n\nBitte beheben Sie diese Fehler zuerst. Es wurden keine Emails versandt.");
                    }
                    else
                    {
                        Ladebalken.Invoke(new Action<int>(s => { Ladebalken.Maximum = s; }), TestEmailAnzahl);
                        {
                            Parallel.For(0, TestEmailAnzahl, new ParallelOptions { MaxDegreeOfParallelism = TestEmailAnzahl }, i =>
                            {
                                AlleMails[i].TextBearbeiten(this);
                            });
                            for (int i = 0; i < TestEmailAnzahl; i++)
                            {
                                Fehler = AlleMails[i].EmailSenden(SMTPServerIP, TestEmailAdresseTextBox.Text, true);
                                while (!Fehler.Equals(string.Empty) && (i != TestEmailAnzahl))
                                {
                                    switch (Fehler)
                                    {
                                        case "1":
                                            WarteAnimation(i + 1, TestEmailAnzahl);
                                            break;
                                        default:
                                            MessageBox.Show("Fehler: " + Fehler);
                                            i = TestEmailAnzahl;
                                            break;
                                    }
                                    Fehler = AlleMails[i].EmailSenden(SMTPServerIP, TestEmailAdresseTextBox.Text, true);
                                }
                                if (i != TestEmailAnzahl)
                                {
                                    StatusLabel.Invoke(new Action<string>(s => { StatusLabel.Text = s; }), string.Format("Status: {0}/{1} versendet", i + 1, TestEmailAnzahl));
                                    Ladebalken.Invoke(new Action<int>(s => { Ladebalken.Value = s; }), i + 1);
                                }
                            }
                        }
                    }
                    OberflächeDeaktivieren(true);
                    GC.Collect();
                }).Start();
            }
            catch (Exception ex)
            {
                Program.MeldeFehler(ex.Message + "\n" + ex.StackTrace);
                Environment.Exit(1);
            }
        }
        private bool WarteAnimation(int Pos, int MaxPos)
        {
            for (int i = 0; i < 30; i++)
            {
                StatusLabel.Invoke(new Action<string>(s => { StatusLabel.Text = s; }), string.Format("Status: {0}/{1} Keine Verbindung. Versuche erneut" + (new StringBuilder(i + 1).Insert(0, ".", i + 1).ToString()), Pos, MaxPos));
                Thread.Sleep(1000);
            }
            GC.Collect();
            return true;
        }
        private void StartenButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (!BereitsVersendet || (BereitsVersendet && (MessageBox.Show("Sie haben die Emails schon einmal versendet. Sind Sie sich sicher, dass Sie das noch einmal tun möchten?", "Bestätigung", MessageBoxButtons.YesNo) == DialogResult.Yes)))
                {
                    new Thread(delegate ()
                    {
                        OberflächeDeaktivieren(false);
                        AlleZeiten = new List<long>();
                        if (!AlleMailsÜberprüfen())
                        {
                            MessageBox.Show("Es sind bei der Überprüfung der Daten Fehler vorgekommen. Diese wurden in die Datei:\n\n" + ProgrammPfad + @"\EmailAdressenFehler.txt" + "\n\ngeschrieben. Bitte beheben Sie diese Fehler zuerst. Es wurden keine Emails versandt.");
                        }
                        else
                        {
                            var GefundeneFehler = new List<string>();
                            var Parallelität = 10;
                            if (AlleMails.Length < 10)
                            {
                                Parallelität = AlleMails.Length;
                            };
                            Ladebalken.Invoke(new Action<int>(s => { Ladebalken.Minimum = s; }), 0);
                            Ladebalken.Invoke(new Action<int>(s => { Ladebalken.Maximum = s; }), AlleMails.Length);
                            int ProcessedEmailCount = 0;
                            for (int i = 0; i < AlleMails.Length; i++)
                            {
                                var CurWatch = new Stopwatch();
                                CurWatch.Start();
                                AlleMails[i].TextBearbeiten(this);
                                CurWatch.Stop();
                                StatusLabel.Invoke(new Action<string>(s => { StatusLabel.Text = s; }), string.Format("Status: {0}/{1} vorbereitet {2}", ProcessedEmailCount, AlleMails.Length, ETABerechnen(CurWatch.ElapsedMilliseconds / Parallelität, AlleMails.Length - 1 - ProcessedEmailCount)));
                                Ladebalken.Invoke(new Action<int>(s => { Ladebalken.Value = s; }), ProcessedEmailCount);
                                Interlocked.Increment(ref ProcessedEmailCount);
                                if (i % 50 == 0)
                                {
                                    GC.Collect();
                                }
                            }
                            AlleZeiten = new List<long>();
                            var Watch = new Stopwatch();
                            var Rückgabewert = string.Empty;
                            for (int i = 0; i < AlleMails.Length; i++)
                            {
                                Watch.Restart();
                                Rückgabewert = AlleMails[i].EmailSenden(SMTPServerIP, string.Empty, false);
                                while (Rückgabewert == "1")
                                {
                                    Rückgabewert = AlleMails[i].EmailSenden(SMTPServerIP, string.Empty, false);
                                    WarteAnimation(i + 1, AlleMails.Length);
                                    Watch.Reset();
                                }
                                if (!Rückgabewert.Equals(string.Empty))
                                {
                                    GefundeneFehler.Add("Zeile: " + (i + 2).ToString() + ", Spalte: Empfänger, Wert: " + AlleMails[i].Empfänger + ", Ursache: " + Rückgabewert);
                                }
                                Watch.Stop();
                                StatusLabel.Invoke(new Action<string>(s => { StatusLabel.Text = s; }), string.Format("Status: {0}/{1} versendet {2}", i + 1, AlleMails.Length, ETABerechnen(Watch.ElapsedMilliseconds, AlleMails.Length - 1 - i)));
                                Ladebalken.Invoke(new Action<int>(s => { Ladebalken.Value = s; }), i + 1);
                                if (i % 30 == 0)
                                {
                                    GC.Collect();
                                }
                            }
                            if (File.Exists(ProgrammPfad + @"\EmailAdressenFehler.txt"))
                            {
                                File.Delete(ProgrammPfad + @"\EmailAdressenFehler.txt");
                            }
                            if (GefundeneFehler.Count > 0)
                            {
                                File.WriteAllLines(ProgrammPfad + @"\EmailAdressenFehler.txt", GefundeneFehler);
                                MessageBox.Show("Es sind beim Versand der Emails Fehler vorgekommen. Diese wurden in folgende Datei geschrieben:\n\n" + ProgrammPfad + @"\EmailAdressenFehler.txt" + "\n\nAlle anderen Emails wurden erfolgreich verschickt.");
                            }
                            else
                            {
                                MessageBox.Show("Vorgang ohne Fehler abgeschlossen.");
                            }
                            BereitsVersendet = true;
                        }
                        OberflächeDeaktivieren(true);
                        GC.Collect();
                    }).Start();
                }
            }
            catch (Exception ex)
            {
                Program.MeldeFehler(ex.Message + "\n" + ex.StackTrace);
                Environment.Exit(1);
            }
        }
        private void OberflächeDeaktivieren(bool Aktivieren)
        {
            if (Aktivieren)
            {
                CSVPfadSuchenButton.Invoke(new Action<bool>(s => { CSVPfadSuchenButton.Enabled = s; }), true);
                TestEmailAdresseTextBox.Invoke(new Action<bool>(s => { TestEmailAdresseTextBox.Enabled = s; }), true);
                if (AdressePrüfen(TestEmailAdresseTextBox.Text))
                {
                    TestButton.Invoke(new Action<bool>(s => { TestButton.Enabled = s; }), true);
                }
                StartenButton.Invoke(new Action<bool>(s => { StartenButton.Enabled = s; }), true);
                StatusLabel.Invoke(new Action<string>(s => { StatusLabel.Text = s; }), "Status: 0/0 versendet ETA: 0m 0s");
                Ladebalken.Invoke(new Action<int>(s => { Ladebalken.Minimum = s; }), 0);
                if (!ReferenceEquals(AlleMails, null))
                {
                    Ladebalken.Invoke(new Action<int>(s => { Ladebalken.Maximum = s; }), AlleMails.Length);
                }
                else
                {
                    Ladebalken.Invoke(new Action<int>(s => { Ladebalken.Maximum = s; }), 1);
                }
                Ladebalken.Invoke(new Action<int>(s => { Ladebalken.Value = s; }), 0);
            }
            else
            {
                CSVPfadSuchenButton.Invoke(new Action<bool>(s => { CSVPfadSuchenButton.Enabled = s; }), false);
                TestEmailAdresseTextBox.Invoke(new Action<bool>(s => { TestEmailAdresseTextBox.Enabled = s; }), false);
                TestButton.Invoke(new Action<bool>(s => { TestButton.Enabled = s; }), false);
                StartenButton.Invoke(new Action<bool>(s => { StartenButton.Enabled = s; }), false);
            }
        }
        private void TestEmailAdresseTextBox_TextChanged(object sender, EventArgs e)
        {
            if ((CSVPfadTextBox.TextLength > 0) && AdressePrüfen(TestEmailAdresseTextBox.Text))
            {
                TestButton.Enabled = true;
            }
            else
            {
                TestButton.Enabled = false;
            }
        }
        internal bool AdressePrüfen(string Adresse)
        {
            if (!Adresse.Contains("@") || !Adresse.Contains("."))
            {
                return false;
            }
            else if (Adresse.Length < 6)
            {
                return false;
            }
            else if ((Adresse.Split('@').Length - 1) > 1)
            {
                return false;
            }
            else if (Adresse.Contains('"') || Adresse.Contains('(') || Adresse.Contains(')') || Adresse.Contains(',') || Adresse.Contains(':') || Adresse.Contains(';') || Adresse.Contains('<') || Adresse.Contains('>') || Adresse.Contains('[') || Adresse.Contains(']') || Adresse.Contains('\\') || Adresse.Contains("..") || Adresse.Contains(" "))
            {
                return false;
            }
            else if (!Encoding.ASCII.GetString(Encoding.ASCII.GetBytes(Adresse)).Equals(Adresse, StringComparison.InvariantCultureIgnoreCase))
            {
                return false;
            }
            return true;
        }
        private void CSVPfadSuchenButton_Click(object sender, EventArgs e)
        {
            try
            {
                var openCSVDialog = new OpenFileDialog();
                openCSVDialog.Title = "Wählen Sie die CSV Datei aus";
                openCSVDialog.Filter = "CSV Datei|*.xlsx";
                openCSVDialog.InitialDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                openCSVDialog.Multiselect = false;
                var DialogErgebnis = openCSVDialog.ShowDialog(this);
                if (DialogErgebnis == DialogResult.OK)
                {
                    CSVPfadTextBox.Text = openCSVDialog.FileName;
                    if (AlleMailsEinlesen(openCSVDialog.FileName))
                    {
                        if (AdressePrüfen(TestEmailAdresseTextBox.Text))
                        {
                            TestButton.Enabled = true;
                        }
                        StartenButton.Enabled = true;
                    }
                    else
                    {
                        MessageBox.Show("Ungültige CSV Datei");
                    }
                }
            }
            catch (Exception ex)
            {
                Program.MeldeFehler(ex.Message + "\n" + ex.StackTrace);
                Environment.Exit(1);
            }
        }
        private bool AlleMailsEinlesen(string Pfad)
        {
            try
            {
                var Exelsheets = Workbook.Worksheets(Pfad).ToArray();
                var ExelRows = Exelsheets[0].Rows;
                var SpaltenNamen = ExelRows[0].Cells;
                AlleMails = new Mail[ExelRows.Length - 1];
                if (ExelRows.Length > 1)
                {
                    Parallel.For(1, ExelRows.Length, new ParallelOptions { MaxDegreeOfParallelism = ExelRows.Length - 1 }, i =>
                    {
                        AlleMails[i - 1] = new Mail();
                        for (int i2 = 0; i2 < ExelRows[i].Cells.Length; i2++)
                        {
                            if (!ReferenceEquals(ExelRows[i].Cells[i2], null) && !ExelRows[i].Cells[i2].Text.Equals(string.Empty))
                            {
                                var AktuelleSpaltenDaten = ExelRows[i].Cells[i2];
                                AlleMails[i - 1].InfoHinzufügen(SpaltenNamen[i2].Text, AktuelleSpaltenDaten.Text, ProgrammPfad);
                            }
                            else if (i2 > 6)
                            {
                                AlleMails[i - 1].AustauschBegriffe.Add(new AustauschBegriff(SpaltenNamen[i2].Text, " "));
                            }
                        }
                    });
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                Program.MeldeFehler(ex.Message + "\n" + ex.StackTrace);
                Environment.Exit(1);
                return false;
            }
        }
        private bool AlleMailsÜberprüfen()
        {
            try
            {
                AlleFehler = new List<string>[AlleMails.Length];
                if (File.Exists(ProgrammPfad + @"\EmailAdressenFehler.txt"))
                {
                    File.Delete(ProgrammPfad + @"\EmailAdressenFehler.txt");
                }
                var GefundeneFehler = new List<string>();
                for (int i = 0; i < AlleMails.Length; i++)
                {
                    if (AlleMails[i].AustauschBegriffe.Count > 0)
                    {
                        for (int i2 = 0; i2 < AlleMails[i].AustauschBegriffe.Count; i2++)
                        {
                            for (int i3 = 0; i3 < AlleMails[i].AustauschBegriffe.Count; i3++)
                            {
                                if ((i2 != i3) && AlleMails[i].AustauschBegriffe[i2].AltesWort.StartsWith(AlleMails[i].AustauschBegriffe[i3].AltesWort))
                                {
                                    GefundeneFehler.Add("Variablename " + AlleMails[i].AustauschBegriffe[i2].AltesWort + " und Variablename " + AlleMails[i].AustauschBegriffe[i3].AltesWort + " klingen zu ähnlich");
                                    goto GoOn;
                                }
                            }
                        }
                    }
                }
                GoOn:
                Parallel.For(0, AlleMails.Length, new ParallelOptions { MaxDegreeOfParallelism = AlleMails.Length }, i =>
                {
                    AlleFehler[i] = new List<string>();
                    if (AlleMails[i].MinimumErfüllt)
                    {
                        AlleMails[i].AdressenPrüfen(this, i);
                    }
                    else
                    {
                        AlleFehler[i].Add("Zeile: " + (i + 2).ToString() + ", Minimum nicht erfüllt");
                    }
                });
                for (int i = 0; i < AlleFehler.Length; i++)
                {
                    if (AlleFehler[i].Count > 0)
                    {
                        GefundeneFehler.AddRange(AlleFehler[i]);
                    }
                }
                if (GefundeneFehler.Count > 0)
                {
                    File.WriteAllLines(ProgrammPfad + @"\EmailAdressenFehler.txt", GefundeneFehler);
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                Program.MeldeFehler(ex.Message + "\n" + ex.StackTrace);
                Environment.Exit(1);
                return true;
            }
        }
        private string ETABerechnen(long LetzteZeitmessung, int ÜbrigeAnzahl)
        {
            if (ÜbrigeAnzahl <= 0)
            {
                ÜbrigeAnzahl = 1;
            }
            AlleZeiten.Add(LetzteZeitmessung);
            var Durchschnitt = AlleZeiten.Average();
            var ETASekunden = (Durchschnitt * ÜbrigeAnzahl) / 1000;
            return string.Format("ETA: {0}", TimeSpan.FromSeconds(ETASekunden).ToString("g"));
        }
        internal string RandomText(int Länge)
        {
            return new string(Enumerable.Repeat("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", Länge).Select(s => s[Zufallsgenerator.Next(s.Length)]).ToArray());
        }
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Environment.Exit(0);
        }
        private void Form1_Shown(object sender, EventArgs e)
        {
            new Thread(delegate ()
            {
                OberflächeDeaktivieren(false);
                StatusLabel.Invoke(new Action<string>(s => { StatusLabel.Text = s; }), "Status: Überprüfe auf Updates");
                Program.Update();
                OberflächeDeaktivieren(true);
            }).Start();
        }
    }

    internal sealed class Mail
    {
        internal Mail()
        {
            AustauschBegriffe = new List<AustauschBegriff>();
        }
        private string Sender { get; set; }
        internal string Empfänger { get; set; }
        private string[] CC { get; set; }
        private string[] BCC { get; set; }
        private string BetreffTextDateiPfad { get; set; }
        private string BetreffText { get; set; }
        private string InhaltTextDateiPfad { get; set; }
        private string InhaltText { get; set; }
        private string[] AnhangDateiPfade { get; set; }
        private string[] BearbeiteteAnhangDateiPfade { get; set; }
        internal List<AustauschBegriff> AustauschBegriffe { get; set; }
        internal bool MinimumErfüllt { get; set; }
        internal void InfoHinzufügen(string Name, string Wert, string ProgrammPfad)
        {
            switch (Name.ToUpperInvariant())
            {
                case "ABSENDER":
                    Sender = Wert;
                    break;
                case "EMPFÄNGER":
                    Empfänger = Wert;
                    break;
                case "CC":
                    CC = Wert.Split(',');
                    break;
                case "BCC":
                    BCC = Wert.Split(',');
                    break;
                case "BETREFF":
                    if (File.Exists(ProgrammPfad + @"\" + Wert))
                    {
                        BetreffTextDateiPfad = Wert;
                    }
                    break;
                case "MAILTEXT":
                    if (File.Exists(ProgrammPfad + @"\" + Wert))
                    {
                        InhaltTextDateiPfad = Wert;
                    }
                    break;
                case "ANHANG":
                    AnhangDateiPfade = Wert.Split(',');
                    break;
                default:
                    AustauschBegriffe.Add(new AustauschBegriff(Name, Wert));
                    break;
            }
            if (!ReferenceEquals(Sender, null) && !ReferenceEquals(Empfänger, null) && !ReferenceEquals(BetreffTextDateiPfad, null) && !ReferenceEquals(InhaltTextDateiPfad, null))
            {
                MinimumErfüllt = true;
            }
            else
            {
                MinimumErfüllt = false;
            }
        }
        internal bool TextBearbeiten(Form1 HauptForm)
        {
            try
            {
                BetreffText = File.ReadAllText(HauptForm.ProgrammPfad + @"\" + BetreffTextDateiPfad, Encoding.Default).Replace(System.Environment.NewLine, string.Empty);
                InhaltText = File.ReadAllText(HauptForm.ProgrammPfad + @"\" + InhaltTextDateiPfad, Encoding.Default);
                string TempPfad;
                for (int i = 0; i < AustauschBegriffe.Count; i++)
                {
                    BetreffText = BetreffText.Replace(AustauschBegriffe[i].AltesWort, AustauschBegriffe[i].NeuesWort);
                    InhaltText = InhaltText.Replace(AustauschBegriffe[i].AltesWort, AustauschBegriffe[i].NeuesWort);
                }
                BearbeiteteAnhangDateiPfade = null;
                if (!ReferenceEquals(AnhangDateiPfade, null))
                {
                    BearbeiteteAnhangDateiPfade = new string[AnhangDateiPfade.Length];
                    AnhangDateiPfade.CopyTo(BearbeiteteAnhangDateiPfade, 0);
                    for (int i = 0; i < AnhangDateiPfade.Length; i++)
                    {
                        if (File.Exists(HauptForm.ProgrammPfad + @"\" + AnhangDateiPfade[i]))
                        {
                            if (AnhangDateiPfade[i].EndsWith(".docx"))
                            {
                                TempPfad = Path.GetTempPath() + @"Massen Mailer\" + HauptForm.RandomText(10) + @"\" + AnhangDateiPfade[i];
                                if (!Directory.Exists(Path.GetDirectoryName(TempPfad)))
                                {
                                    Directory.CreateDirectory(Path.GetDirectoryName(TempPfad));
                                }
                                else if (File.Exists(TempPfad))
                                {
                                    {
                                        File.Delete(TempPfad);
                                    }
                                }
                                File.Copy(HauptForm.ProgrammPfad + @"\" + AnhangDateiPfade[i], TempPfad);
                                var Dokument = DocX.Load(TempPfad);
                                for (int i2 = 0; i2 < AustauschBegriffe.Count; i2++)
                                {
                                    Dokument.ReplaceText(AustauschBegriffe[i2].AltesWort, AustauschBegriffe[i2].NeuesWort);
                                }
                                Dokument.Save();
                                Dokument.Dispose();
                                var oWord = new Word.Application();
                                oWord.Visible = false;
                                object oMissing = Missing.Value;
                                object isVisible = true;
                                object readOnly = false;
                                object oInput = TempPfad;
                                object oOutput = TempPfad.Replace(".docx", ".pdf");
                                object oFormat = WdSaveFormat.wdFormatPDF;
                                object oNoChanges = WdSaveOptions.wdDoNotSaveChanges;
                                if (File.Exists((string)oOutput))
                                {
                                    File.Delete((string)oOutput);
                                }
                                var oDoc = oWord.Documents.Open(ref oInput, ref oMissing, ref readOnly, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref isVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                                oDoc.Activate();
                                oDoc.SaveAs(ref oOutput, ref oFormat, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                                oWord.Quit(ref oNoChanges, ref oMissing, ref oMissing);
                                BearbeiteteAnhangDateiPfade[i] = (string)oOutput;
                            }
                            else if (!AnhangDateiPfade[i].Equals(string.Empty))
                            {
                                BearbeiteteAnhangDateiPfade[i] = HauptForm.ProgrammPfad + @"\" + AnhangDateiPfade[i];
                            }
                        }
                        else
                        {
                            BearbeiteteAnhangDateiPfade[i] = string.Empty;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Program.MeldeFehler(ex.Message + "\n" + ex.StackTrace);
                Environment.Exit(1);
            }
            return true;
        }
        internal bool AdressenPrüfen(Form1 HauptForm, int AktuellerMailIndex)
        {
            bool Rückgabewert = true;
            if (!HauptForm.AdressePrüfen(Sender))
            {
                HauptForm.AlleFehler[AktuellerMailIndex].Add("Zeile: " + (AktuellerMailIndex + 2).ToString() + ", Spalte: Sender, Wert: " + Sender);
                Rückgabewert = false;
            }
            if (!HauptForm.AdressePrüfen(Empfänger))
            {
                HauptForm.AlleFehler[AktuellerMailIndex].Add("Zeile: " + (AktuellerMailIndex + 2).ToString() + ", Spalte: Empfänger, Wert: " + Empfänger);
                Rückgabewert = false;
            }
            if (!ReferenceEquals(CC, null))
            {
                for (int i = 0; i < CC.Length; i++)
                {
                    if (!HauptForm.AdressePrüfen(CC[i]))
                    {
                        HauptForm.AlleFehler[AktuellerMailIndex].Add("Zeile: " + (AktuellerMailIndex + 2).ToString() + ", Spalte: CC, Wert: " + CC[i]);
                        Rückgabewert = false;
                    }
                }
            }
            if (!ReferenceEquals(BCC, null))
            {
                for (int i = 0; i < BCC.Length; i++)
                {
                    if (!HauptForm.AdressePrüfen(BCC[i]))
                    {
                        HauptForm.AlleFehler[AktuellerMailIndex].Add("Zeile: " + (AktuellerMailIndex + 2).ToString() + ", Spalte: BCC, Wert: " + BCC[i]);
                        Rückgabewert = false;
                    }
                }
            }
            if (!ReferenceEquals(AnhangDateiPfade, null))
            {
                for (int i = 0; i < AnhangDateiPfade.Length; i++)
                {
                    if (!File.Exists(HauptForm.ProgrammPfad + @"\" + AnhangDateiPfade[i]))
                    {
                        HauptForm.AlleFehler[AktuellerMailIndex].Add("Zeile: " + (AktuellerMailIndex + 2).ToString() + ", Spalte: Anhang, Wert: " + AnhangDateiPfade[i]);
                        Rückgabewert = false;
                    }
                }
            }
            return Rückgabewert;
        }
        internal string EmailSenden(string SMTPServerIP, string EmpfängerAdresse, bool Test)
        {
            try
            {
                System.Net.Mail.MailMessage Email;
                if (Test)
                {
                    Email = new System.Net.Mail.MailMessage(Sender, EmpfängerAdresse, BetreffText, InhaltText);
                }
                else
                {
                    Email = new System.Net.Mail.MailMessage(Sender, Empfänger, BetreffText, InhaltText);
                }
                Email.IsBodyHtml = false;
                Email.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
                Email.BodyEncoding = Encoding.Default;
                Email.HeadersEncoding = Encoding.Default;
                if (!Test)
                {
                    if (!ReferenceEquals(CC, null))
                    {
                        for (int i = 0; i < CC.Length; i++)
                        {
                            Email.CC.Add(new MailAddress(CC[i]));
                        }
                    }
                    if (!ReferenceEquals(BCC, null))
                    {
                        for (int i = 0; i < BCC.Length; i++)
                        {
                            Email.Bcc.Add(new MailAddress(BCC[i]));
                        }
                    }
                }
                if (!ReferenceEquals(BearbeiteteAnhangDateiPfade, null))
                {
                    for (int i = 0; i < BearbeiteteAnhangDateiPfade.Length; i++)
                    {
                        if (!BearbeiteteAnhangDateiPfade[i].Equals(string.Empty))
                        {
                            var FileInfo = new FileInfo(BearbeiteteAnhangDateiPfade[i]);
                            var Anhang = new Attachment(FileInfo.FullName, MediaTypeNames.Application.Octet);
                            var disposition = Anhang.ContentDisposition;
                            disposition.CreationDate = FileInfo.CreationTime;
                            disposition.ModificationDate = FileInfo.LastWriteTime;
                            disposition.ReadDate = FileInfo.LastAccessTime;
                            disposition.FileName = FileInfo.Name;
                            disposition.Size = FileInfo.Length;
                            disposition.DispositionType = DispositionTypeNames.Attachment;
                            Email.Attachments.Add(Anhang);
                        }
                    }
                }
                try
                {
                    var SMTPClient = new SmtpClient(SMTPServerIP, 25);
                    SMTPClient.EnableSsl = false;
                    SMTPClient.UseDefaultCredentials = true;
                    SMTPClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                    SMTPClient.DeliveryFormat = SmtpDeliveryFormat.International;
                    SMTPClient.Send(Email);
                    SMTPClient.Dispose();
                    return string.Empty;
                }
                catch (Exception ex)
                {
                    if (!ex.Message.Contains("5."))
                    {
                        return "1";
                    }
                    else
                    {
                        return ex.Message;
                    }
                }
            }
            catch (Exception ex)
            {
                Program.MeldeFehler(ex.Message + "\n" + ex.StackTrace);
                Environment.Exit(1);
                return "1";
            }
        }
    }

    internal sealed class AustauschBegriff
    {
        internal AustauschBegriff(string AltesWort, string NeuesWort)
        {
            this.AltesWort = AltesWort;
            this.NeuesWort = NeuesWort;
        }
        internal string AltesWort { get; set; }
        internal string NeuesWort { get; set; }
    }
}
