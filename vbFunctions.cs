using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Web;
using System.Data;
using System.Net.Mail;
using System.Text.RegularExpressions;
namespace WebScrape
{
    class vbFunctions
    {
        Microsoft.Office.Interop.Outlook.MailItem selectedOutlookMessages()
        {
            Microsoft.Office.Interop.Outlook.Application myOlApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.Explorer objView = myOlApp.ActiveExplorer();
            Microsoft.Office.Interop.Outlook.MailItem olMail = (MailItem)objView.Selection;
            return olMail;
        }


        public List<string> Split(string str, string del)
        {
            List<string> ary = new List<string>(str.Split(new string[] { del }, StringSplitOptions.None));
            return ary;
        }


        

        public void ShowMessageBox(string text, string head = "")
        {

            MessageBoxButtons buttons = MessageBoxButtons.OK;
            MessageBox.Show(text, head, buttons);
        }

        public string FolderPicker()
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            string str = "";
            folderDlg.ShowNewFolderButton = true;
            DialogResult result = folderDlg.ShowDialog();

            if (result == DialogResult.OK)
            {
                Environment.SpecialFolder root = folderDlg.RootFolder;
                str = folderDlg.SelectedPath;
                string strSub = str.Substring(str.Length - 11, 11);
                if (strSub == "\\New folder")
                {
                    result = folderDlg.ShowDialog();
                    if (result == DialogResult.OK)
                    {
                        str = folderDlg.SelectedPath;
                    }
                }
            }
            return str;
        }






    }
}
