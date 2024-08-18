using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelFriend
{
    public partial class frmInputText : Form
    {
        protected string title { get; set; }
        protected string HelpText { get; set; } = null;
        protected int MAX_LENGTH { get; set; }
        public string Answer { get; set; }
        public frmInputText(string title = "ورود داده", string HelpText = null, int max_length = 200)
        {
            this.title = title;
            this.HelpText = HelpText;
            this.MAX_LENGTH = max_length;
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Answer = textBox1.Text;
            Close();
        }

        private void frmInputText_Load(object sender, EventArgs e)
        {
            
            this.Text = title;
            this.Answer = "";
            textBox1.MaxLength = MAX_LENGTH;

            if (HelpText != null)
            {
                lblHelp.Text = HelpText;
            }
        }
    }
}
