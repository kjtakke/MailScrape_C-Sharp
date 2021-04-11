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



//MessageBoxButtons buttons = MessageBoxButtons.YesNo;
//MessageBox.Show(str, "Folder Choice:", buttons);


namespace WebScrape

{

    class Main
    {
        const int ArrayDim = 18;
        const String FileLocation = "Documents";
        private string[,] Selected_mail_items = new string[,] { { "" }, { "" } };
        private string ext;
        private string exportString;
        private string filePathPicked;

        public void Attachments()
        {
            Microsoft.Office.Interop.Outlook.Application myOlApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.Explorer objView = myOlApp.ActiveExplorer();
            filePathPicked = FolderPicker();

            foreach (Microsoft.Office.Interop.Outlook.MailItem olMail in objView.Selection)
            {
                string FilePathConverter;
                try
                {

                    foreach (Microsoft.Office.Interop.Outlook.Attachment olAttachment in olMail.Attachments)
                    {
                        if (olAttachment.FileName != "")
                        {
                            //FilePathConverter = File_Exists(filePathPicked + "\\" + olAttachment.FileName.ToString());
                            FilePathConverter = filePathPicked + "\\" + olAttachment.FileName.ToString();
                            olAttachment.SaveAsFile(FilePathConverter);
                        }
                    }
                }
                finally { }
            }
        }


        string File_Exists(string fielPath)
        {

            string strFileExists; bool fileExists; string temp_FileName; string temp_FileName_Placeholder;
            string temp_FileExt; string temp_path; int i;

            strFileExists = Directory.GetFiles(fielPath, "*.*").ToString(); //Need to fix  Need to change to a bool
            if (strFileExists != "")
            {
                List<string> temp_FileArray = new List<string>(strFileExists.Split(new string[] { "." }, StringSplitOptions.None));
                temp_FileExt = temp_FileArray.Last().ToString();
                temp_FileName = temp_FileArray.First().ToString();
                temp_FileArray.Clear();
                temp_FileArray = new List<string>(fielPath.Split(new string[] { "\\" }, StringSplitOptions.None));
                temp_path = "";

                for (i = 0; i < temp_FileArray.Count - 1; i++)
                {
                    temp_path = temp_path + temp_FileArray[i].ToString() + "\\";
                }

                fileExists = true;
                temp_FileName_Placeholder = temp_FileName;
                i = 1;


                do
                {
                    temp_FileName_Placeholder = temp_FileName + "(" + i + ")";

                    if (Directory.GetFiles(fielPath, "*.*").ToString() + temp_FileExt != "")
                    {
                        i = i + 1;
                    }
                    else
                    {
                        fielPath = temp_path + temp_FileName_Placeholder + temp_FileExt;
                        fileExists = false;
                    }
                } while (fileExists = true);
            }
            else
            {
                fielPath = fielPath;

            }
            return fielPath;
        }














        string FolderPicker()
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

        void Mail_Scrape()
        {
            Mail_Scrape(); CleanText();
        }

        string jsonArray(string str, string del)
        {
            string tmpString;
            List<string> tempArray = new List<string>(str.Split(new string[] { del }, StringSplitOptions.None));
            tmpString = "[" + "\r" + "\u0020" + "\u0020" + "\u0020" + "\u0020";
            for (int i = 0; i < tempArray.Count; i++)
            {
                tempArray[i] = tempArray[i].Trim();
                if (i == tempArray.Count)
                {
                    tmpString = tmpString + "{\"email\":\"" + tempArray[i] + "\"}," + "\r" + "\u0020" + "\u0020" + "\u0020";
                }
                else
                {
                    tmpString = tmpString + "{\"email\":\"" + tempArray[i] + "\"}," + "\r" + "\u0020" + "\u0020" + "\u0020" + "\u0020";
                }
                tmpString = tmpString + "]";
            }
            return tmpString;
        }

        string FileName()
        {
            string FileDate; string UserName; 

            FileDate = DateTime.Now.ToString("yymmdd");
            UserName = Environment.UserName;
            List<string> tempArray = new List<string>(UserName.Split(new string[] { "." }, StringSplitOptions.None));
            UserName = "";

            for (int i = 0; i < tempArray.Count; i++){

                if (i == tempArray.Count)
                {
                    UserName = UserName + tempArray[i];
                }
                else
                {
                    UserName = UserName + tempArray[i] + "_";
                }
            }
            return FileDate + " - " + UserName + " - " + "Mail_Scrape" + ext;
        }

 
        public void CleanText()
        {
            int i; int j; string myString = "";

            for (i = 1; i < Selected_mail_items.Length; i++)
            {
                for (j = 0; j < ArrayDim; j++)
                {
                    Selected_mail_items[i, j] = Selected_mail_items[i, j].Replace("\"", "'");
                }
            }

            void get_Selected_mail_items()
            {
                Microsoft.Office.Interop.Outlook.Application myOlApp = new Microsoft.Office.Interop.Outlook.Application();
                Microsoft.Office.Interop.Outlook.Explorer objView = myOlApp.ActiveExplorer();
                int k = 1;

                foreach (Microsoft.Office.Interop.Outlook.MailItem olMail in objView.Selection) { i += 1; }

                Selected_mail_items = ResizeArray(ref Selected_mail_items, 0, i - 2, 18, 0);
                Selected_mail_items[0, 0] = "To";
                Selected_mail_items[0, 1] = "CC";
                Selected_mail_items[0, 2] = "Reply_Recipient_Names";
                Selected_mail_items[0, 3] = "Sender_Email_Address";
                Selected_mail_items[0, 4] = "Sender_Name";
                Selected_mail_items[0, 5] = "Sent_On_Behalf_Of_Name";
                Selected_mail_items[0, 6] = "Sender_Email_Type";
                Selected_mail_items[0, 7] = "Sent";
                Selected_mail_items[0, 8] = "Size"; ;
                Selected_mail_items[0, 9] = "Unread";
                Selected_mail_items[0, 10] = "Creation_Time";
                Selected_mail_items[0, 11] = "Last_Modification_Time";
                Selected_mail_items[0, 12] = "Sent_On";
                Selected_mail_items[0, 13] = "Received_Time";
                Selected_mail_items[0, 14] = "Importance";
                Selected_mail_items[0, 15] = "Received_By_Name";
                Selected_mail_items[0, 16] = "Received_On_Behalf_Of_Name";
                Selected_mail_items[0, 17] = "Subject";
                Selected_mail_items[0, 18] = "Body";
                k = 1;

                foreach (Microsoft.Office.Interop.Outlook.MailItem olMail in objView.Selection)
                {
                    try
                    {
                        if (olMail.To == null) { Selected_mail_items[k, 0] = ""; } else { Selected_mail_items[k, 0] = olMail.To.ToString(); }
                        if (olMail.CC == null) { Selected_mail_items[k, 1] = ""; } else { Selected_mail_items[k, 1] = olMail.CC.ToString(); }
                        if (olMail.ReplyRecipientNames == null) { Selected_mail_items[k, 2] = ""; } else { Selected_mail_items[k, 2] = olMail.ReplyRecipientNames.ToString(); }
                        if (olMail.SenderEmailAddress == null) { Selected_mail_items[k, 3] = ""; } else { Selected_mail_items[k, 3] = olMail.SenderEmailAddress.ToString(); }
                        if (olMail.SenderName == null) { Selected_mail_items[k, 4] = ""; } else { Selected_mail_items[k, 4] = olMail.SenderName.ToString(); }
                        if (olMail.SentOnBehalfOfName == null) { Selected_mail_items[k, 5] = ""; } else { Selected_mail_items[k, 5] = olMail.SentOnBehalfOfName.ToString(); }
                        if (olMail.SenderEmailType == null) { Selected_mail_items[k, 6] = ""; } else { Selected_mail_items[k, 6] = olMail.SenderEmailType.ToString(); }
                        if (olMail.Sent == null) { Selected_mail_items[k, 7] = ""; } else { Selected_mail_items[k, 7] = olMail.Sent.ToString(); }
                        if (olMail.Size == null) { Selected_mail_items[k, 8] = ""; } else { Selected_mail_items[k, 8] = olMail.Size.ToString(); }
                        if (olMail.UnRead == null) { Selected_mail_items[k, 9] = ""; } else { Selected_mail_items[k, 9] = olMail.UnRead.ToString(); }
                        if (olMail.CreationTime == null) { Selected_mail_items[k, 10] = ""; } else { Selected_mail_items[k, 10] = olMail.CreationTime.ToString(); }
                        if (olMail.LastModificationTime == null) { Selected_mail_items[k, 11] = ""; } else { Selected_mail_items[k, 11] = olMail.LastModificationTime.ToString(); }
                        if (olMail.SentOn == null) { Selected_mail_items[k, 12] = ""; } else { Selected_mail_items[k, 12] = olMail.SentOn.ToString(); }
                        if (olMail.ReceivedTime == null) { Selected_mail_items[k, 13] = ""; } else { Selected_mail_items[k, 13] = olMail.ReceivedTime.ToString(); }
                        if (olMail.Importance == null) { Selected_mail_items[k, 14] = ""; } else { Selected_mail_items[k, 14] = olMail.Importance.ToString(); }
                        if (olMail.ReceivedByName == null) { Selected_mail_items[k, 15] = ""; } else { Selected_mail_items[k, 15] = olMail.ReceivedByName.ToString(); }
                        if (olMail.ReceivedOnBehalfOfName == null) { Selected_mail_items[k, 16] = ""; } else { Selected_mail_items[k, 16] = olMail.ReceivedOnBehalfOfName.ToString(); }
                        if (olMail.Subject == null) { Selected_mail_items[k, 17] = ""; } else { Selected_mail_items[k, 17] = olMail.Subject.ToString(); }
                        if (olMail.Body == null) { Selected_mail_items[k, 18] = ""; } else { Selected_mail_items[k, 18] = olMail.Body.ToString(); }

                        k += 1;
                    }
                    finally { }
                }
            }

            

            void ShowMessageBox(string text, string head = "")
            {

                MessageBoxButtons buttons = MessageBoxButtons.OK;
                MessageBox.Show(text, head, buttons);
            }

            string[,] ResizeArray<T>(ref T[,] array, int padLeft, int padRight, int padTop, int padBottom)
            {
                int ow = array.GetLength(0);
                int oh = array.GetLength(1);
                int nw = ow + padLeft + padRight;
                int nh = oh + padTop + padBottom;

                int x0 = padLeft;
                int y0 = padTop;
                int x1 = x0 + ow - 1;
                int y1 = y0 + oh - 1;
                int u0 = -x0;
                int v0 = -y0;

                if (x0 < 0) x0 = 0;
                if (y0 < 0) y0 = 0;
                if (x1 >= nw) x1 = nw - 1;
                if (y1 >= nh) y1 = nh - 1;

                T[,] nArr = new T[nw, nh];
                for (int y = y0; y <= y1; y++)
                {
                    for (int x = x0; x <= x1; x++)
                    {
                        nArr[x, y] = array[u0 + x, v0 + y];
                    }
                }
                array = nArr;

                return (string[,])(object)nArr;
            }
        }




        

    }
}
