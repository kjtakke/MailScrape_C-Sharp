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
using System.Threading.Tasks;


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
        private int UboundSelectedMailItems;

        public void save_EmailsWithAttments()
        {
            save_Emails();

            Microsoft.Office.Interop.Outlook.Application myOlApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.Explorer objView = myOlApp.ActiveExplorer();

            foreach (Microsoft.Office.Interop.Outlook.MailItem olMail in objView.Selection)
            {
                string FilePathConverter;
                try
                {

                    foreach (Microsoft.Office.Interop.Outlook.Attachment olAttachment in olMail.Attachments)
                    {
                        if (olAttachment.FileName != "")
                        {

                            string subj = olMail.Subject.ToString();
                            if (subj.Length > 20) { subj = subj.Substring(0, 20); }

                            string FileName = olMail.SentOn.ToString("yyddmm-hhmmss") + "-" +
                                  olMail.SenderEmailAddress.ToString() + "-" +
                                  subj;
                            FileName = FileName.Replace("\\", " ");
                            FileName = FileName.Replace("/", " ");
                            FileName = FileName.Replace(".", " ");
                            FileName = FileName.Replace("|", " ");
                            FileName = FileName.Replace("*", " ");
                            FileName = FileName.Replace("*", " ");
                            FileName = FileName.Replace("?", " ");
                            FileName = FileName.Replace(":", " ");
                            FileName = FileName.Replace("<", " ");
                            FileName = FileName.Replace(">", " ");


                            Directory.CreateDirectory(filePathPicked + "\\" + FileName + "\\");
                            FilePathConverter = File_Exists(filePathPicked + "\\" + FileName + "\\" + olAttachment.FileName.ToString());
                            olAttachment.SaveAsFile(FilePathConverter);
                        }
                    }
                }
                finally { }
            }
        }


        public void save_Emails()
        {
            Microsoft.Office.Interop.Outlook.Application myOlApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.Explorer objView = myOlApp.ActiveExplorer();
            filePathPicked = FolderPicker();
            foreach (Microsoft.Office.Interop.Outlook.MailItem olMail in objView.Selection) 
            { 

                

                string subj = olMail.Subject.ToString();
                string FileName = olMail.SentOn.ToString("yyddmm-hhmmss") + "-" +
                                  olMail.SenderEmailAddress.ToString() + "-" +
                                  subj;
                FileName = FileName.Replace("\\", " ");
                FileName = FileName.Replace("/", " ");
                FileName = FileName.Replace(".", " ");
                FileName = FileName.Replace("|", " ");
                FileName = FileName.Replace("*", " ");
                FileName = FileName.Replace("*", " ");
                FileName = FileName.Replace("?", " ");
                FileName = FileName.Replace(":", " ");
                FileName = FileName.Replace("<", " ");
                FileName = FileName.Replace(">", " ");


                string savepath = filePathPicked + "\\" + FileName + ".txt";
                olMail.SaveAs(savepath, 0);

            };
        }

        public void JSON()
        {
            JSONPlane();

            Microsoft.Office.Interop.Outlook.Application myOlApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.Explorer objView = myOlApp.ActiveExplorer();

            foreach (Microsoft.Office.Interop.Outlook.MailItem olMail in objView.Selection)
            {
                string FilePathConverter;
                try
                {

                    foreach (Microsoft.Office.Interop.Outlook.Attachment olAttachment in olMail.Attachments)
                    {
                        if (olAttachment.FileName != "")
                        {

                            string subj = olMail.Subject.ToString();
                            if (subj.Length > 20) { subj = subj.Substring(0, 20); }

                            string FileName = olMail.SentOn.ToString("yyddmm-hhmmss") + "-" +
                                  olMail.SenderEmailAddress.ToString() + "-" +
                                  subj;
                            FileName = FileName.Replace("\\", " ");
                            FileName = FileName.Replace("/", " ");
                            FileName = FileName.Replace(".", " ");
                            FileName = FileName.Replace("|", " ");
                            FileName = FileName.Replace("*", " ");
                            FileName = FileName.Replace("*", " ");
                            FileName = FileName.Replace("?", " ");
                            FileName = FileName.Replace(":", " ");
                            FileName = FileName.Replace("<", " ");
                            FileName = FileName.Replace(">", " ");


                            Directory.CreateDirectory(filePathPicked + "\\" + FileName + "\\");
                            FilePathConverter = File_Exists(filePathPicked + "\\" + FileName + "\\" + olAttachment.FileName.ToString());
                            olAttachment.SaveAsFile(FilePathConverter);
                        }
                    }
                }
                finally { }
            }

        }



        public void JSONPlane()
        {
            Mail_Scrape();
            filePathPicked = FolderPicker();

            for(int i = 1; i < UboundSelectedMailItems + 1; i++)
            {
                exportString = "";
                string toStr = jsonArray(Selected_mail_items[i, 0], ";");
                string ccStr = jsonArray(Selected_mail_items[i, 1], ";");

                exportString = exportString + "{" + "\r" + "\u0020" +
                                              "\"people\" : {" + "\r" + "\u0020" + "\u0020" +
                                              "\"to\" : " + toStr + "," + "\r" + "\u0020" + "\u0020" +
                                              "\"cc\" : " + ccStr + "\r" + "\u0020" +
                                              "}," + "\r" + "\u0020";
                exportString = exportString +
                                "\"names\" : {" + "\r" + "\u0020" + "\u0020" +
                                    "\"ReplyRecipientNames\" : \"" + Selected_mail_items[i, 2] + "\"," + "\r" + "\u0020" + "\u0020" +
                                    "\"SenderName\" : \"" + Selected_mail_items[i, 4] + "\"," + "\r" + "\u0020" + "\u0020" +
                                    "\"SentOnBehalfOfName\" : \"" + Selected_mail_items[i, 5] + "\"," + "\r" + "\u0020" + "\u0020" +
                                    "\"ReceivedOnBehalfOfName\" : \"" + Selected_mail_items[i, 16] + "\"," + "\r" + "\u0020" + "\u0020" +
                                    "\"ReceivedByName\" : \"" + Selected_mail_items[i, 15] + "\"" + "\r" + "\u0020" +
                                "}," + "\r" + "\u0020";
                exportString = exportString +
                                "\"time\" : {" + "\r" + "\u0020" + "\u0020" +
                                    "\"CreationTime\" : \"" + Selected_mail_items[i, 10] + "\"," + "\r" + "\u0020" + "\u0020" +
                                    "\"LastModificationTime\" : \"" + Selected_mail_items[i, 11] + "\"," + "\r" + "\u0020" + "\u0020" +
                                    "\"SentOn\" : \"" + Selected_mail_items[i, 12] + "\"," + "\r" + "\u0020" + "\u0020" +
                                    "\"ReceivedTime\" : \"" + Selected_mail_items[i, 13] + "\"" + "\r" + "\u0020" +
                                "}," + "\r" + "\u0020";
                exportString = exportString +
                                "\"metadata\" : {" + "\r" + "\u0020" + "\u0020" +
                                    "\"SenderEmailType\" : \"" + Selected_mail_items[i, 6] + "\"," + "\r" + "\u0020" + "\u0020" +
                                    "\"Size\" : " + Selected_mail_items[i, 8] + "," + "\r" + "\u0020" + "\u0020" +
                                    "\"UnRead\" : " + Selected_mail_items[i, 9] + "," + "\r" + "\u0020" + "\u0020" +
                                    "\"Sent\" : " + Selected_mail_items[i, 7] + "," + "\r" + "\u0020" + "\u0020" +
                                    "\"Importance\" : " + Selected_mail_items[i, 14] + "\r" + "\u0020" +
                                "}," + "\r" + "\u0020";
                exportString = exportString +
                                "\"text\" : {" + "\r" + "\u0020" + "\u0020" +
                                        "\"Subject\" : \"" + Selected_mail_items[i, 17].Replace("\"", "'") + "\"," + "\r" + "\u0020" + "\u0020" +
                                        "\"Body\" : \"" + Selected_mail_items[i, 18].Replace("\"", "'") + "\"" + "\r" + "\u0020" +
                                    "}" + "\r" +
                            "}";

                string subj = Selected_mail_items[i, 17];
                if (subj.Length > 20){subj = subj.Substring(0, 20);}

                string FileName = Selected_mail_items[i, 12] + "-" + 
                                  Selected_mail_items[i, 4] + "-" +
                                  subj; 

                FileName = FileName.Replace("\\", " ");
                FileName = FileName.Replace("/", " ");
                FileName = FileName.Replace(".", " ");
                FileName = FileName.Replace("|", " ");
                FileName = FileName.Replace("*", " ");
                FileName = FileName.Replace("*", " ");
                FileName = FileName.Replace("?", " ");
                FileName = FileName.Replace(":", " ");
                FileName = FileName.Replace("<", " ");
                FileName = FileName.Replace(">", " ");

                using (StreamWriter writer = new StreamWriter(filePathPicked + "\\" + FileName + ".json"))
                {
                    writer.WriteLine(exportString);
                }
            }
        }

        public void CSV()
        {
            DataTable table = new DataTable("");
            DataColumn tableColumn = table.Columns.Add("To",typeof(string));
            table.Columns.Add("CC", typeof(string));
            table.Columns.Add("Reply_Recipient_Names", typeof(string));
            table.Columns.Add("Sender_Email_Address", typeof(string));
            table.Columns.Add("Sender_Name", typeof(string));
            table.Columns.Add("Sent_On_Behalf_Of_Name", typeof(string));
            table.Columns.Add("Sender_Email_Type", typeof(string));
            table.Columns.Add("Sent", typeof(string));
            table.Columns.Add("Size", typeof(string));
            table.Columns.Add("Unread", typeof(string));
            table.Columns.Add("Creation_Time", typeof(string));
            table.Columns.Add("Last_Modification_Time", typeof(string));
            table.Columns.Add("Sent_On", typeof(string));
            table.Columns.Add("Received_Time", typeof(string));
            table.Columns.Add("Importance", typeof(string));
            table.Columns.Add("Received_By_Name", typeof(string));
            table.Columns.Add("Received_On_Behalf_Of_Name", typeof(string));
            table.Columns.Add("Subject", typeof(string));
            table.Columns.Add("Body", typeof(string));

            Mail_Scrape();

            for (int i = 0; i < UboundSelectedMailItems; i++)
            {
                DataRow newRow = table.NewRow();
                for(int j = 0; j < 17; j++)
                {
                    newRow[j] = Selected_mail_items[i, j];

                }
                table.Rows.Add(newRow);
            }
            ext = ".csv";
            string FilePath = FolderPicker();
            FilePath = FilePath + "\\" + FileName();
            string csvTbl = ConvertToCSV(table);
            System.IO.File.WriteAllText(FilePath, ConvertToCSV(table));
        }

        string ConvertToCSV(DataTable dt)
        {
            StringBuilder sb = new StringBuilder("", 50);
            foreach (DataRow row in dt.Rows)
            {
                sb.AppendLine(String.Join(",", (
                    from i in row.ItemArray
                    select i.ToString()
                    .Replace("\"", "\"\"")
                    .Replace(",", "\\,")
                    .Replace(Environment.NewLine, "\\" + Environment.NewLine)
                    .Replace("\\", "\\\\")).ToArray()));
            }
            return sb.ToString();
        }

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
                            FilePathConverter = File_Exists(filePathPicked + "\\" + olAttachment.FileName.ToString());
                            //FilePathConverter = filePathPicked + "\\" + olAttachment.FileName.ToString();
                            olAttachment.SaveAsFile(FilePathConverter);
                        }
                    }
                }
                finally { }
            }
        }

        string File_Exists(string filePath)
        {
            string strFileExists; string temp_FileName; string temp_FileName_Placeholder;
            string temp_FileExt; string temp_path; int i = 1; int j;
            strFileExists = filePath;
            List<string> temp_FileArray = new List<string>(strFileExists.Split(new string[] { "." }, StringSplitOptions.None));
            temp_FileExt = temp_FileArray.Last().ToString();
            temp_FileName = temp_FileArray.First().ToString();
            temp_FileArray.Clear();
            temp_FileArray = new List<string>(filePath.Split(new string[] { "\\" }, StringSplitOptions.None));
            temp_path = "";
            
            for (j = 0; j < temp_FileArray.Count - 1; j++)
            {
                temp_path = temp_path + temp_FileArray[j].ToString() + "\\";
            }
            
            string[] pathArray = Directory.GetFiles(temp_path);
            temp_FileName_Placeholder = filePath;

            beginAgain:

            for (j = 0; j < pathArray.Length; j++)
            {
                if (temp_FileName_Placeholder  == pathArray[j])
                {
                    temp_FileName_Placeholder = temp_FileName + "(" + i + ")" + "." + temp_FileExt;
                    i += 1;
                    goto beginAgain;                    
                } 
            }
            return temp_FileName_Placeholder;
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
            get_Selected_mail_items();
            CleanText();
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
            int i; int j;
            
            for (i = 0; i < UboundSelectedMailItems; i++)
            {
                for (j = 0; j < ArrayDim; j++)
                {
                    Selected_mail_items[i, j] = Selected_mail_items[i, j].Replace("\"", "'");
                }
            }
        }

        void get_Selected_mail_items()
        {
            Microsoft.Office.Interop.Outlook.Application myOlApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.Explorer objView = myOlApp.ActiveExplorer();
            int k = 1;
            int i = 1; //NOT SURE IF THIS SHOULD BE ZERO
            foreach (Microsoft.Office.Interop.Outlook.MailItem olMail in objView.Selection) { try { i += 1; } finally { } }

            UboundSelectedMailItems = i - 1;
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

                    if (olMail.SentOn == null) { Selected_mail_items[k, 12] = ""; } else { Selected_mail_items[k, 12] = olMail.SentOn.ToString("yyddmm-hhmmss"); }

                    if (olMail.ReceivedTime == null) { Selected_mail_items[k, 13] = ""; } else { Selected_mail_items[k, 13] = olMail.ReceivedTime.ToString("yyddmm-hhmmss"); }
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

