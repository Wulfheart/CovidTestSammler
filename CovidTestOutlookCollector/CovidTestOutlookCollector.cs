using iText.Kernel.Pdf;
using iText.Kernel.Utils;
using MsgReader.Outlook;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CovidTestOutlookCollector
{
    public partial class CovidTestOutlookCollector : Form
    {
        public CovidTestOutlookCollector()
        {
            InitializeComponent();
        }

        private void Auswertung_Click(object sender, EventArgs e)
        {
            try
            {

                // Open Folder
                var fbd = new FolderBrowserDialog();
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    
                    Auswertung.Enabled = false;
                    LogOutput.Clear();
                    List<string> files = Directory.GetFiles(fbd.SelectedPath).ToList();
                    // Are there .msg files?
                    Log("Suche nach .msg-Dateien.");
                    foreach (var f in files.ToList())
                    {
                        if (!f.EndsWith(".msg"))
                        {
                            files.Remove(f);
                        }
                    }
                    Log($"{files.Count} .msg-Dateien gefunden.");
                    if (files.Count <= 0)
                    {
                        throw new Exception("Es scheint, als ob im gewählten Ordner keine .msg-Dateien vorhanden sind.");
                    }
                    Log("Öffne bisherigen Stand.");
                    Config conf = new Config(fbd.SelectedPath);
                    Log("Bisheriger Stand geöffnet.");
                    Log("Filtere als bereits erledigt gekennzeichnete Dateien heraus.");
                    foreach (var f in files.ToList())
                    {
                        if (conf.Data.files.Contains(f))
                        {
                            Log($"Auslassen von {Path.GetFileName(f)}, weil diese Datei bereits als erledigt markiert wurde.");
                            files.Remove(f);
                        }
                    }
                    Log($"Filtern abgeschlossen. Es sind noch {files.Count} .msg-Dateien übrig.");
                    if (files.Count <= 0)
                    {
                        throw new Exception("Es scheint, als ob keine neuen .msg-Dateien vorhanden sind.");
                    }


                    // Import Messages
                    // !!!! This could go horribly wrong if there are too many new Message Files
                    Log("Importiere .msg-Dateien.");
                    List<Msg> Messages = new List<Msg>();
                    foreach (var f in files)
                    {
                        if (f.Contains("Password", StringComparison.OrdinalIgnoreCase))
                        {
                            Log($"Importiere {Path.GetFileName(f)} als Email, die ein Passwort enthält.");
                            Messages.Add(new PasswordMsg(f));
                        }
                        else
                        {
                            Log($"Importiere {Path.GetFileName(f)} als Email, die eine PDF enthält.");
                            Messages.Add(new PdfMsg(f));
                        }
                    }
                    Log($"Import abgeschlossen. Es wurden insgesamt {Messages.Count} Dateien importiert.");

                    Log("Filtere IDs heraus, die als bereits erledigt markiert wurden.");
                    // Remove all messages with ids that are already present
                    foreach (var msg in Messages.ToList())
                    {
                        if (conf.Data.ids.Contains(msg.ID))
                        {
                            Log($"Auslassen von {Path.GetFileName(msg.Path)}, weil diese ID ({msg.ID}) bereits als erledigt markiert wurde.");
                            conf.AddFile(msg.Path);
                            Messages.Remove(msg);
                        }
                    }
                    conf.Save();
                    Log($"Filtern abgeschlossen. Es sind noch {Messages.Count} Messages übrig.");
                    Log("Gruppiere Nachrichten nach Barcode.");
                    // Get all where at least a Password and PDF File is present
                    var groups = Messages.GroupBy(m => m.ID).Select(u => u.ToList()).Where(g =>
                    {
                        bool hasPdf = false;
                        bool hasPassword = false;
                        foreach (var item in g)
                        {
                            if (item.GetType().Equals(typeof(PasswordMsg)))
                            {
                                hasPassword = true;
                            }
                            else if (item.GetType().Equals(typeof(PdfMsg)))
                            {
                                hasPdf = true;
                            }
                        }
                        return hasPdf && hasPassword;
                    }).ToList();
                    Log($"Gruppieren abgeschlossen. Es sind insgesamt {groups.Count} einzigartige Barcodes.");
                    if (groups.Count <= 0)
                    {
                        throw new Exception("Es ist keine vollständige Gruppe mit Barcodes vorhanden. Vorgang abgeschlossen, ohne irgendwelche Dateien zu erstellen.");
                    }
                    Log("Starte Erstellung der einzelnen PDFs.");
                    List<Entity> entities = new List<Entity>();
                    string timeForPath = DateTime.Now.ToString("yyyy_MM_dd_hh_mm");
                    string folder = Path.Combine(fbd.SelectedPath, "PDFs", timeForPath);
                    List<string> EntityErrors = new List<string>();
                    for (int i = 0; i < groups.Count; i++)
                    {
                        try
                        {
                            Log($"Entity {i + 1} von {groups.Count}");
                            var group = groups[i];
                            Log($"Entity {i + 1} von {groups.Count}: Suche Nachricht mit Passwort.");
                            PasswordMsg passwordMsg = group.Where(g => g.GetType().Equals(typeof(PasswordMsg))).Cast<PasswordMsg>().First();
                            Log($"Entity {i + 1} von {groups.Count}: Nachricht mit Passwort gefunden.");
                            Log($"Entity {i + 1} von {groups.Count}: Suche Nachricht mit PDF.");
                            PdfMsg pdfMsg = group.Where(g => g.GetType().Equals(typeof(PdfMsg))).Cast<PdfMsg>().First();
                            Log($"Entity {i + 1} von {groups.Count}: Nachricht mit Passwort gefunden.");
                            Entity entity = new Entity(passwordMsg, pdfMsg);
                            Log($"Entity {i + 1} von {groups.Count}: Einlesen der Daten der PDF.");
                            entity.ParseData();
                            Log($"Entity {i + 1} von {groups.Count}: Es handelt sich um {entity.Forename} {entity.Surname}.");
                            Log($"Entity {i + 1} von {groups.Count}: Abspeichern der PDF im Ordner {folder}.");
                            entity.Save(folder);
                            Log($"Entity {i + 1} von {groups.Count}: Abspeichern der Daten in Exceldatei im Ordner {folder}.");
                            entity.WriteToExcel(folder, $"_Zusammenfassung_{timeForPath}.xlsx");
                            // Mark all in this group as done
                            conf.AddID(group[0].ID);
                            foreach (var g in group)
                            {
                                conf.AddFile(g.Path);

                            }
                            conf.Save();
                            Log($"Entity {i + 1} von {groups.Count} wurde als erledigt markiert.");
                            entities.Add(entity);
                        }
                        catch (Exception ex)
                        {
                            string error = ($"Entity {i + 1} von {groups.Count}: Fehler {ex.Message}.");
                            Log(error);
                            EntityErrors.Add(error);
                        }

                    }

                    Log($"Zusammenführen der PDFs im Ordner {folder}.");
                    Log($"Hole Liste aller Dateien.");
                    List<string> newFiles = new List<string>();
                    try
                    {
                        newFiles = new List<string>(Directory.GetFiles(folder));
                        Log($"{newFiles.Count} Dateien gefunden. Wähle PDFs aus.");

                    }
                    catch (Exception)
                    {
                        throw new Exception("Es scheint ein Problem beim Öffnen der PDF Dateien aufgetreten zu sein. Eventuelle Fehlermeldungen entnehmen Sie bitte dem Log.");
                    }
                    List<string> pdfs = newFiles.Where(s => s.EndsWith(".pdf")).ToList();
                    Log($"Es wurden {pdfs.Count} PDFs unter {newFiles.Count} Dateien gefunden.");
                    string PdfPath = Path.Combine(folder, $"_Zusammenfassung_Druck_{timeForPath}.pdf");
                    PdfDocument printPdf = new PdfDocument(new PdfWriter(PdfPath));
                    for (int j = 0; j < pdfs.Count; j++)
                    {
                        Log($"PDF {j + 1} von {pdfs.Count} anfügen.");
                        string path = pdfs[j];
                        PdfDocument doc = new PdfDocument(new PdfReader(path));
                        PdfMerger merger = new PdfMerger(printPdf);
                        merger.Merge(doc, 1, doc.GetNumberOfPages());
                        doc.Close();
                        Log($"PDF {j + 1} von {pdfs.Count} angefügt.");
                    }
                    printPdf.Close();
                    Log($"Zusammenfassung wurde erstellt unter {PdfPath}.");
                    for (int i = 0; i < EntityErrors.Count; i++)
                    {
                        Log($"Entity Error {i + 1} von {EntityErrors.Count}: {EntityErrors[i]}");
                    }
                    Log("Vorgang abgeschlossen.");
                    MessageBox.Show($"Vorgang abgeschlossen. Die erstellten Dateien können in dem Ordner {folder} abgerufen werden.");
                    if (EntityErrors.Count > 0)
                    {
                        MessageBox.Show($"Es wurden {EntityErrors.Count} Fehler bei der Erstellung der Auswertung (Entity-Error) festgestellt. Genauere Informationen entnehmen Sie bitte dem Log.");
                    }

                }


            }
            catch (Exception ex)
            {
                Log(ex.Message);
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Auswertung.Enabled = true;
            }
        }

        private void Log(string message)
        {
            LogOutput.AppendText(message + "\r\n");
            LogOutput.ScrollToCaret();
        }

        
    }
}
