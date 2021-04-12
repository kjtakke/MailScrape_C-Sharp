using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WebScrape
{
    public partial class Ribbon1
    {

        Main MyScraoe = new Main();
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void CSV_Click(object sender, RibbonControlEventArgs e)
        {
            Main mn = new Main();
            mn.CSV();
        }

        private void JSON_Click(object sender, RibbonControlEventArgs e)
        {
            Main mn = new Main();
            mn.JSONPlane();
        }

        private void Text_Click(object sender, RibbonControlEventArgs e)
        {
            Main mn = new Main();
            mn.save_Emails();
        }

        private void Attachments_Click(object sender, RibbonControlEventArgs e)
        {
            Main mn = new Main();
            mn.Attachments();
        }

        private void jsonAtt_Click(object sender, RibbonControlEventArgs e)
        {
            Main mn = new Main();
            mn.JSON();
        }

        private void TextAtt_Click(object sender, RibbonControlEventArgs e)
        {
            Main mn = new Main();
            mn.save_EmailsWithAttments();
        }
    }
}
