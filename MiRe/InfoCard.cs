using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MiRe
{
    public partial class InfoCard : UserControl
    {
        public InfoCard(string[] text)
        {
            InitializeComponent();

            this.Width = 600;
            this.Height = 1000;
            this.richTextBox1.Width = 550;
            this.richTextBox1.Height = 950;
            this.richTextBox1.Multiline = true;

            foreach (string s in text)
            {
                richTextBox1.AppendText(s);
            }
            //richTextBox1.Text = string.Join("", text);


        }
    }
}
