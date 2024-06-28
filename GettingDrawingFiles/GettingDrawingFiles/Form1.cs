using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Inventor;

namespace GettingDrawingFiles
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = true;
            DialogResult dialogres = folderDlg.ShowDialog();
            DirectoryInfo dir = new DirectoryInfo(folderDlg.SelectedPath);
            FileInfo[] files = dir.GetFiles("*.idw");
            if (dialogres == DialogResult.OK)
            {
                textBox1.Text = folderDlg.SelectedPath;
                foreach(var file in files)
                {
                    textBox2.Text  = textBox2.Text + file.ToString() + "\r\n";
                }
            }
            Inventor.Application invApp = (Inventor.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application");
            try
            {
                foreach (var file in files)
                {
                    string fullfilepath = folderDlg.SelectedPath + "\\" + file;
                    Logger(fullfilepath);
                    Inventor.Document Doc = invApp.Documents.Open(fullfilepath, true);
                    DrawingDocument drwdoc = (DrawingDocument)Doc;
                    PropertySet drawprop = drwdoc.PropertySets["Inventor User Defined Properties"];

                    string PropVal = "";
                    string PropName = "";
                    Property propname = null;
                    string[] fieldnames = {"Revision Number", "1-1-Status", "2-1-Status", "3-1-Status", "4-1-Status", "5-1-Status", "6-1-Status", "7-1-Status", "8-1-Status", "Änderungstext1", "Änderungstext2",
             "Änderungstext3", "Änderungstext4", "Änderungstext5", "Änderungstext6", "Änderungstext7", "Änderungstext8", "1-3-Date", "2-3-Date", "3-3-Date", "4-3-Date", "5-3-Date", "6-3-Date", "7-3-Date", "8-3-Date",
             "1-4-Name", "2-4-Name", "3-4-Name", "4-4-Name", "5-4-Name", "6-4-Name", "7-4-Name", "8-4-Name"};
                    foreach (var names in fieldnames)
                    {
                        try
                        {
                            propname = drawprop[names];
                            Sheet osheet = drwdoc.Sheets[1];
                            DrawingView drawview = osheet.DrawingViews[1];
                            Doc = (Document)drawview.ReferencedDocumentDescriptor.ReferencedDocument;
                            PropertyAdd(Doc, propname, names, propname.Value);
                        }
                        catch
                        {
                            Sheet osheet = drwdoc.Sheets[1];
                            DrawingView drawview = osheet.DrawingViews[1];
                            Doc = (Document)drawview.ReferencedDocumentDescriptor.ReferencedDocument;
                            PropertyAdd(Doc, propname, names, "");
                        }
                    }
                    Doc.Save();
                    Doc.Close();
                    drwdoc.Close();
                }
                invApp.Quit();
            }
            catch(Exception ex)
            {
                Logger(ex.ToString());
            }
        }
        public static void PropertyAdd(Document Doc, Property prop,string propname, object propvalue)
        {
            bool propExists = true;
            PropertySet customPropSet = Doc.PropertySets["Inventor User Defined Properties"];
            try
            {
                prop = customPropSet[propname];
            }
            catch (Exception ex)
            {
                propExists = false;
            }
            if (!propExists)
            {
                prop = customPropSet.Add(propvalue, propname, null);
            }
            else
            {
                prop.Value = propvalue;
            }
        }
        public static void FindDyr(string path)
        {
            DirectoryInfo dir = new DirectoryInfo(@"C:\mani\logs\");
            if (!dir.Exists)
            {
                dir.Create();
            }
        }
        public static void Logger(string lines)
        {
            string path = "C:/mani/logs/";
            FindDyr(path);
            string fileName = DateTime.Now.Day.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Year.ToString() + "_Logs.txt";
            try
            {
                System.IO.StreamWriter file = new System.IO.StreamWriter(path + fileName, true);
                file.WriteLine(DateTime.Now.ToString() + ": " + lines);
                file.Close();
            }
            catch (Exception) { }
        }
    }
}
