using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using System;
using System.CodeDom;
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
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        string showCellDLG(string msg)
        {
            var frm = new frmInputText(msg, HelpText: "مثلا A", max_length: 1);
            if (frm.ShowDialog() == DialogResult.OK)
            {
                return frm.Answer;
            }

            throw new Exception("Not Selected !");
        }

        private void button3_Click(object sender, EventArgs e)
        {
    
            if (txtPath.Text == "")
            {
                MessageBox.Show("لطفا مسیر فایل اکسل را انتخاب نمایید","Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            try
            {
                var cellSourceA = showCellDLG("مبدا شروع شمسی ؟");
                var cellSourceB = showCellDLG("مبدا پایان شمسی ؟");
                var cellDestinationYear = showCellDLG("مقصد سال ؟");
                var cellDestinationMonth = showCellDLG("مقصد ماه ؟");
                var cellDestinationDay = showCellDLG("مقصد روز ؟");

               


                convertDateShamsiDiffToYMD(cellSourceA, cellSourceB, 
                    cellDestinationYear, 
                    cellDestinationMonth, 
                    cellDestinationDay);

                MessageBox.Show("Done", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);



            }
            catch (Exception ex) { }
        





        }

        private void button2_Click(object sender, EventArgs e)
        {
            string cellSource = "";
            string cellDestination = "";
            if (txtPath.Text == "")
            {
                MessageBox.Show("لطفا مسیر فایل اکسل را انتخاب نمایید", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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

        void convertDateShamsiDiffToYMD(string sourceColumnA = "A", string sourceColumnB = "B", string destColumnY = "C", string destColumnM = "D", string destColumnD = "E")
        {
            if (cmbSheets.Items.Count == 0) return;

            using (var wb = new XLWorkbook(txtPath.Text))
            {
                var sheet = wb.Worksheet(cmbSheets.Text);

                int index = 1;
                while (true)
                {
                    var cSource = sheet.Cell(index, sourceColumnA).Value;
                    if (index >= 20 && cSource.ToString() == "")
                    {
                        break;
                    }
                    try
                    {
                        
                        var dtA = Utility.getShamsiToMiladi(sheet.Cell(index, sourceColumnA).Value.GetText());
                        var dtB = Utility.getShamsiToMiladi(sheet.Cell(index, sourceColumnB).Value.GetText());

                        if(dtA != null && dtB != null)
                        {
                            var dtDiffNull = (dtA - dtB);
                            if(dtDiffNull != null)
                            {
                                TimeSpan dtDiff = dtDiffNull.Value;
                                long days = Convert.ToInt64(dtDiffNull?.TotalDays);
                                long year = (long)days / 365;
                                long month = (long)((days - 365 * year) / 30);
                                long day = days - 365 * year - 30 * month;

                                sheet.Cell(index, destColumnY).Value = year;
                                sheet.Cell(index, destColumnM).Value = month;
                                sheet.Cell(index, destColumnD).Value = day;
                                
                            }

                        }
                    }
                    catch (Exception ex) { /* date parse exception */ }
                    index++;
                }

                wb.Save();// commit
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //txtPath.Text = @"C:\Users\Altyn Inche 01\Desktop\b.xlsx";
            //loadSheets();

            //convertDateShamsiDiffToYMD("C", "D", "E", "F", "G");

        }

        private void btnAgeFromShamsi_Click(object sender, EventArgs e)
        {
            string cellSource = "";
            string cellDestination = "";
            if (txtPath.Text == "")
            {
                MessageBox.Show("لطفا مسیر فایل اکسل را انتخاب نمایید", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void btrnMiladiToShamsi_Click(object sender, EventArgs e)
        {
            string cellSource = "";
            string cellDestination = "";
            if (txtPath.Text == "")
            {
                MessageBox.Show("لطفا مسیر فایل اکسل را انتخاب نمایید", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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
    }
}
