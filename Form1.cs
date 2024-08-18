using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelFriend
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string cellSource = "";
            string cellDestination = "";
            if (txtPath.Text == "")
            {
                MessageBox.Show("لطفا مسیر فایل اکسل را انتخاب نمایید");
                return;
            }


            var frm = new frmInputText("ستون مبدا میلادی را وارد نمایید", HelpText: "مثلا A", max_length: 1);
            if (frm.ShowDialog() == DialogResult.OK)
            {
                cellSource = frm.Answer;

                var frm2 = new frmInputText("ستون مقصد شمسی را وارد نمایید", HelpText: "مثلا B", max_length: 1);
                if (frm2.ShowDialog() == DialogResult.OK)
                {
                    cellDestination = frm2.Answer;
                    try
                    {
                        convertDateMiladiToShamsi(cellSource, cellDestination);

                        MessageBox.Show("Done", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {

                    }

                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string cellSource = "";
            string cellDestination = "";
            if (txtPath.Text == "")
            {
                MessageBox.Show("لطفا مسیر فایل اکسل را انتخاب نمایید");
                return;
            }

            var frm = new frmInputText("ستون مبدا شمسی را وارد نمایید", HelpText: "مثلا A", max_length:1);
            if(frm.ShowDialog() == DialogResult.OK)
            {
                cellSource = frm.Answer;

                var frm2 = new frmInputText("ستون مقصد میلادی را وارد نمایید", HelpText: "مثلا B", max_length: 1);
                if (frm2.ShowDialog() == DialogResult.OK)
                {
                    cellDestination = frm2.Answer;
                    try
                    {
                        convertDateShamsiToMiladi(cellSource, cellDestination);

                        MessageBox.Show("Done", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {

                    }
                
                }
            }
        }

        private void btnPath_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtPath.Text = openFileDialog1.FileName;

                loadSheets();
            }
        }


        void loadSheets()
        {
            var wb = new XLWorkbook(txtPath.Text);
            cmbSheets.Items.Clear();
            foreach (var sh in wb.Worksheets)
            {
                cmbSheets.Items.Add(sh.Name);
            }

            if (cmbSheets.Items.Count > 0)
            {
                cmbSheets.SelectedIndex = 0;
            }

            wb.Dispose();
        }
        
        void convertDateShamsiToAge(string sourceColumn = "B", string destColumn = "C")
        {
            if (cmbSheets.Items.Count == 0) return;

            using (var wb = new XLWorkbook(txtPath.Text))
            {
                var sheet = wb.Worksheet(cmbSheets.Text);

                int index = 1;
                while (true)
                {
                    var cSource = sheet.Cell(index, sourceColumn).Value;
                    if (index >= 20 && cSource.ToString() == "")
                    {
                        break;
                    }

                    CultureInfo persianCulture = new CultureInfo("fa-IR");
                    try
                    {
                        DateTime persianDateTime = DateTime.ParseExact(cSource.ToString(), "yyyy/MM/dd", persianCulture);

                        sheet.Cell(index, destColumn).Value = Math.Floor((DateTime.Now - persianDateTime).TotalDays / 365);
                    }
                    catch (Exception ex)
                    {
                        //ignore
                    }


                    index++;
                }

                wb.Save();// commit
            }
        }

        void convertDateShamsiToMiladi(string sourceColumn = "B", string destColumn = "C")
        {
            if(cmbSheets.Items.Count == 0) return;

            using(var wb = new XLWorkbook(txtPath.Text))
            {
                var sheet = wb.Worksheet(cmbSheets.Text);

                int index = 1;
                while (true)
                {
                    var cSource = sheet.Cell(index, sourceColumn).Value;
                    if(index >= 20 && cSource.ToString() == "")
                    {
                        break;
                    }

                    CultureInfo persianCulture = new CultureInfo("fa-IR");
                    try
                    {
                        DateTime persianDateTime = DateTime.ParseExact(cSource.ToString(), "yyyy/MM/dd", persianCulture);

                        sheet.Cell(index, destColumn).Value = persianDateTime.ToString("yyyy/MM/dd");
                    }
                    catch (Exception ex)
                    {
                        //ignore
                    }
                    

                    index ++;
                }

                wb.Save();// commit
            }
        }

        void convertDateMiladiToShamsi(string sourceColumn = "B", string destColumn = "C")
        {
            if (cmbSheets.Items.Count == 0) return;

            using (var wb = new XLWorkbook(txtPath.Text))
            {
                var sheet = wb.Worksheet(cmbSheets.Text);

                int index = 1;
                while (true)
                {
                    var cSource = sheet.Cell(index, sourceColumn).Value;
                    if (index >= 20 && cSource.ToString() == "")
                    {
                        break;
                    }

                    try
                    {
                        DateTime dt = DateTime.Parse(cSource.ToString());
                        PersianCalendar pr = new PersianCalendar();
                        string fa_format = ""
                            + pr.GetYear(dt) + "/" +
                            +pr.GetMonth(dt) + "/" +
                            +pr.GetDayOfMonth(dt)
                            ;



                        sheet.Cell(index, destColumn).Value = fa_format;
                    }
                    catch (Exception ex)
                    {
                        //ignore
                    }


                    index++;
                }

                wb.Save();// commit
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //loadSheets();

           //convertDateShamsiToMiladi();

        }

        private void btnAgeFromShamsi_Click(object sender, EventArgs e)
        {
            string cellSource = "";
            string cellDestination = "";
            if (txtPath.Text == "")
            {
                MessageBox.Show("لطفا مسیر فایل اکسل را انتخاب نمایید");
                return;
            }

            var frm = new frmInputText("ستون مبدا شمسی را وارد نمایید", HelpText: "مثلا A", max_length: 1);
            if (frm.ShowDialog() == DialogResult.OK)
            {
                cellSource = frm.Answer;

                var frm2 = new frmInputText("ستون مقصد سن را وارد نمایید", HelpText: "مثلا B", max_length: 1);
                if (frm2.ShowDialog() == DialogResult.OK)
                {
                    cellDestination = frm2.Answer;
                    try
                    {
                        convertDateShamsiToAge(cellSource, cellDestination);

                        MessageBox.Show("Done", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {

                    }

                }
            }
        }
    }
}
