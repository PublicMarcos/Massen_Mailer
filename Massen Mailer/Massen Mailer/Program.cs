using Massen_Mailer.Properties;
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace Massen_Mailer
{
    static class Program
    {
        /// <summary>
        /// Der Haupteinstiegspunkt für die Anwendung.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            ÜberprüfeAufUpdate = true;
            if (!ReferenceEquals(args, null) && (args.Length == 1) && args[0].Equals("KeineUpdatesmehr", StringComparison.InvariantCultureIgnoreCase))
            {
                ÜberprüfeAufUpdate = false;
            }
            Application.Run(new Form1());
        }
        private static bool ÜberprüfeAufUpdate { get; set; }
        internal static bool MeldeFehler(string Fehlermeldung)
        {
            if (Fehlermeldung.Contains(@"\n"))
            {
                Fehlermeldung = Fehlermeldung.Replace(@"\n", "++++");
            }
            var Fehlermeldung2 = new string[1] { DateTime.Now.ToString("dd.MM.yy HH:mm:ss", CultureInfo.InvariantCulture) + "++++" + Assembly.GetExecutingAssembly().Location + "++++" + Environment.MachineName + "++++" + Environment.UserName + "++++" + Fehlermeldung };
            try
            {
                File.AppendAllLines(@"\\...\Interne_Programme$\Massen Mailer\Fehlermeldungen.txt", Fehlermeldung2, System.Text.Encoding.UTF8);//Hier den Netzwerkpfad für Updates eintragen
            }
            catch
            {
                MessageBoxA(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle, "Fehlermeldung konnte nicht gespeichert werden.", string.Empty, (uint)0x00000000L | (uint)0x00000010L | (uint)0x00000000L + (uint)0x00000000L + (uint)0x00040000L);
            }
            MessageBoxA(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle, "Ort:" + Assembly.GetExecutingAssembly().Location + "\nMeldung:\n\n" + Fehlermeldung, string.Empty, (uint)0x00000000L | (uint)0x00000010L | (uint)0x00000000L + (uint)0x00000000L + (uint)0x00040000L);
            return true;
        }
        internal static bool Update()
        {
            try
            {
                if (!ÜberprüfeAufUpdate)
                {
                    Thread.Sleep(3000);
                }
                if (File.Exists(Assembly.GetExecutingAssembly().Location + ".old"))
                {
                    File.Delete(Assembly.GetExecutingAssembly().Location + ".old");
                }
                if (File.Exists(Assembly.GetExecutingAssembly().Location.Replace(".exe", ".pdb") + ".old"))
                {
                    File.Delete(Assembly.GetExecutingAssembly().Location.Replace(".exe", ".pdb") + ".old");
                }
                if (!File.Exists(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\DocX.dll"))
                {
                    File.WriteAllBytes(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\DocX.dll", Resources._DocX);
                }
                if (!File.Exists(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Excel.dll"))
                {
                    File.WriteAllBytes(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Excel.dll", Resources._Excel);
                }
                if (ÜberprüfeAufUpdate)
                {
                    if (NetzwerkpfadÜberprüfen() && !FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).ProductVersion.Equals("0.0.0.0") && !File.ReadAllText(@"\\...\Interne_Programme$\Massen Mailer\Update.txt").Equals(FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).ProductVersion, StringComparison.InvariantCultureIgnoreCase))//Hier den Netzwerkpfad für Updates eintragen
                    {
                        File.Move(Assembly.GetExecutingAssembly().Location, Assembly.GetExecutingAssembly().Location + ".old");
                        File.Move(Assembly.GetExecutingAssembly().Location.Replace(".exe", ".pdb"), Assembly.GetExecutingAssembly().Location.Replace(".exe", ".pdb") + ".old");
                        File.Copy(@"\\...\Interne_Programme$\Massen Mailer\Massen Mailer.exe", Assembly.GetExecutingAssembly().Location);//Hier den Netzwerkpfad für Updates eintragen
                        File.Copy(@"\\...\Interne_Programme$\Massen Mailer\Massen Mailer.pdb", Assembly.GetExecutingAssembly().Location.Replace(".exe", ".pdb"));//Hier den Netzwerkpfad für Updates eintragen
                        var NeuerProzessInfos = new System.Diagnostics.ProcessStartInfo();
                        NeuerProzessInfos.FileName = Assembly.GetExecutingAssembly().Location;
                        NeuerProzessInfos.Arguments = " KeineUpdatesmehr";
                        NeuerProzessInfos.WorkingDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                        Process.Start(NeuerProzessInfos);
                        Environment.Exit(0);
                    }
                }
            }
            catch (Exception ex)
            {
                MeldeFehler(ex.Message + "\n" + ex.StackTrace);
                Environment.Exit(1);
            }
            return true;
        }
        private static bool NetzwerkpfadÜberprüfen()
        {
            try
            {
                var ZeitUhr = Stopwatch.StartNew();
                var OrdnerExistiert = false;
                var CheckThread = new Thread(delegate ()
                {
                    if (Directory.Exists(@"\\...\Interne_Programme$\Massen Mailer"))//Hier den Netzwerkpfad für Updates eintragen
                    {
                        OrdnerExistiert = true;
                    }
                });
                CheckThread.Start();
                while ((ZeitUhr.Elapsed.Seconds < 60) && (CheckThread.ThreadState == System.Threading.ThreadState.Running))
                {
                    Thread.Sleep(50);
                }
                if ((CheckThread.ThreadState == System.Threading.ThreadState.Running) || !OrdnerExistiert)
                {
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
                return false;
            }
        }
        [DllImport("User32.dll")]
        internal static extern int MessageBoxA(IntPtr hWnd, String text, String caption, uint type);
    }
}
