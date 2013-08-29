using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Collections;
using System.Drawing.Printing;
using System.IO;

namespace ReceptionScreen
{
    public partial class ReceptionScreen : Form
    {
        private bool _waitingListShow = true;
        private bool _nextTicketShow = true;
        private string _nextTicketText = "A000-0";
        public static int[] TotalTicketSold = new int[4];
        private int _nop;
        private string _printedTicketText = "A000-0";
        private readonly ArrayList _arrayList = new ArrayList();
        private bool _parsable = true;
        public static int[] TotalCutomers = new int[4];
        private Font _printFont;
        private static DateTime _time = DateTime.Now;
        private const string Format = "ddd MMM d HH:mm yyyy";
        private static int _typeA, _typeB, _typeC;
        private static int _timeRange;
        private static int _oldTimeRange = -1;
        private static int saveCount = 0;

        public ReceptionScreen()
        {
            if (Microsoft.VisualBasic.Interaction.InputBox("Password: ", "Protection") != "19931993")
            {
                MessageBox.Show(@"Wrong password! Sorry hacker :P");
                Application.Exit();
                Application.ExitThread();
            }
            else
            {
                AppDomain.CurrentDomain.ProcessExit += OnProcessExit;
                InitializeComponent();
            }
        }

        private void ReceptionScreen_Load(object sender, EventArgs e)
        {
            //UI element anchoring
            Rectangle screen = Screen.PrimaryScreen.Bounds;
            DateTime time = DateTime.Now;
            const string format = "ddd MMM d HH:mm yyyy";
            _waitingListShow = true;
            WaitingList.Hide();
            _nextTicketShow = true;
            lbl_DateTime.Text = time.ToString(format);
            lbl_CompanyName.Left = screen.Width/2 - lbl_CompanyName.Width/2 + screen.Width/20;
            lbl_CompanyName.Top = screen.Height/13 - lbl_CompanyName.Height/13;
            pic_CompanyPicture.Left = screen.Width/20 - pic_CompanyPicture.Width/20;
            pic_CompanyPicture.Top = screen.Height/20 - pic_CompanyPicture.Height/20;
            pic_face.Left = screen.Width - pic_face.Width - screen.Width/60;
            pic_face.Top = screen.Height/20 - pic_face.Height/30;
            lbl_TopBreakLine.Left = 0;
            lbl_TopBreakLine.Top = screen.Height/20 + pic_CompanyPicture.Height;

            lbl_NextTicket.Left = screen.Width/2 - lbl_NextTicket.Width/2 + screen.Width/20;
            lbl_NextTicket.Top = lbl_TopBreakLine.Top + screen.Height/10;

            lbl_LabelNextTicket.Left = screen.Width/7 - lbl_LabelNextTicket.Width/7;
            lbl_LabelNextTicket.Top = lbl_NextTicket.Top + lbl_NextTicket.Height - lbl_LabelNextTicket.Height*2 -
                                      screen.Height/60;

            lbl_LabelPrintTicket.Left = lbl_LabelNextTicket.Left;
            lbl_LabelPrintTicket.Top = lbl_LabelNextTicket.Top + lbl_LabelNextTicket.Height + screen.Height/7;

            txb_NOF.Left = lbl_NextTicket.Left + screen.Width/20;
            txb_NOF.Top = lbl_LabelPrintTicket.Top;

            btn_CreateTicket.Left = txb_NOF.Left + txb_NOF.Width + screen.Width/100;
            btn_CreateTicket.Top = txb_NOF.Top;

            chk_printEnable.Left = btn_CreateTicket.Left + btn_CreateTicket.Width + screen.Width/100;
            chk_printEnable.Top = btn_CreateTicket.Top + screen.Height/300;

            chk_printEnable.Hide();

            btn_Analysis.Left = screen.Width/100;
            btn_Analysis.Top = screen.Height - btn_Analysis.Height*4 - screen.Height/100;

            btn_Save.Left = btn_Analysis.Left + btn_Analysis.Width + screen.Width/100;
            btn_Save.Top = btn_Analysis.Top;
            btn_Save.Hide();

            lbl_LabelPrintedNum.Top = lbl_LabelPrintTicket.Top + screen.Height/12;
            lbl_LabelPrintedNum.Left = lbl_LabelPrintTicket.Left;
            lbl_PrintedNumber.Top = lbl_LabelPrintedNum.Top;

            lbl_PrintedNumber.Left = txb_NOF.Left;
            lbl_BottomBreakLine.Top = lbl_PrintedNumber.Top + lbl_PrintedNumber.Height + screen.Height/30;
            lbl_ContactInfo.Top = btn_Analysis.Top;
            lbl_ContactInfo.Left = screen.Width/2 - lbl_ContactInfo.Width/2;
            WaitingList.Height = screen.Height*2/5;
            WaitingList.Left = screen.Width - WaitingList.Width - screen.Width/50;
            WaitingList.Top = lbl_TopBreakLine.Top + lbl_TopBreakLine.Height + screen.Height/100;
            lbl_DateTime.Left = screen.Width - lbl_DateTime.Width - screen.Width/100;
            lbl_DateTime.Top = btn_Analysis.Top;
            lbl_waitingTickets.Top = WaitingList.Top + WaitingList.Height + screen.Height/100;
            lbl_waitingTickets.Left = screen.Width - lbl_waitingTickets.Width - screen.Width/100;

            lbl_displayTotal.Top = lbl_waitingTickets.Top + lbl_waitingTickets.Height + screen.Height/100;
            lbl_displayTotal.Left = screen.Width - lbl_waitingTickets.Width - screen.Width/50;
            lbl_displayTotal.Hide();

            txb_NOF.Focus();
            btn_Analysis.Visible = false;
            ExcelDoc.ReadDoc();
        }

        private void btn_CreateTicket_Click(object sender, EventArgs e)
        {
            if (txb_NOF.Text == null)
            {
                txb_NOF.BackColor = Color.Red;
                return;
            }
            try
            {
                _nop = int.Parse(txb_NOF.Text);
            }
            catch
            {
                txb_NOF.BackColor = Color.Red;
                _parsable = false;
            }
            if (_parsable && _nop > 0 && _nop < 100)
            {
                txb_NOF.BackColor = Color.White;
                txb_NOF.Clear();
                TotalTicketSold[_timeRange]++;
                TotalCutomers[_timeRange] += _nop;

                char type = 'A';
                if (_nop > 0 && _nop <= 5)
                {
                    type = 'A';
                    _typeA++;
                }
                else if (_nop > 5 && _nop <= 10)
                {
                    type = 'B';
                    _typeB++;
                }
                else if (_nop > 10)
                {
                    type = 'C';
                    _typeC++;
                }

                int totalTicketSoldOfDay = TotalTicketSold[0] + TotalTicketSold[1] + TotalTicketSold[2] +
                                           TotalTicketSold[3];
                _printedTicketText = string.Format("{0}{1:D3}-{2}", type, totalTicketSoldOfDay, _nop);
                lbl_PrintedNumber.Text = _printedTicketText;

                _arrayList.Add(_printedTicketText);
                WaitingList.Items.Clear();
                WaitingList.Items.AddRange(_arrayList.ToArray());
                lbl_waitingTickets.Text = string.Format(
                    "| A : {0} | B : {1} | C : {2} |\r\n            -> Total : {3}", _typeA, _typeB, _typeC,
                    _arrayList.ToArray().Length);

                if (chk_printEnable.Checked)
                {
                    var pd = new PrintDocument();
                    pd.PrintPage += PrintTicket;
                    pd.Print();
                }
            }
            else
                txb_NOF.BackColor = Color.Red;
            _parsable = true;
            txb_NOF.Focus();
        }

        private void ReceptionScreen_MouseMove(object sender, MouseEventArgs e)
        {
            if (btn_Analysis.Bounds.Contains(e.Location) && !btn_Analysis.Visible)
            {
                btn_Analysis.Show();
            }
            else
            {
                btn_Analysis.Hide();
            }

            if (btn_Save.Bounds.Contains(e.Location) && !btn_Save.Visible)
            {
                btn_Save.Show();
            }
            else
            {
                btn_Save.Hide();
            }

            if (WaitingList.Bounds.Contains(e.Location) && !WaitingList.Visible && _waitingListShow)
            {
                WaitingList.Show();
            }
            else
            {
                WaitingList.Hide();
            }

            if (chk_printEnable.Bounds.Contains(e.Location) && !chk_printEnable.Visible)
            {
                chk_printEnable.Show();
            }
            else
            {
                chk_printEnable.Hide();
            }

            var totalT = TotalTicketSold[0] + TotalTicketSold[1] + TotalTicketSold[2] + TotalTicketSold[3];
            var totalC = TotalCutomers[0] + TotalCutomers[1] + TotalCutomers[2] + TotalCutomers[3];

            if (lbl_displayTotal.Bounds.Contains(e.Location) && !lbl_displayTotal.Visible)
            {
                lbl_displayTotal.Text = string.Format("Total tickets: {0}\r\nTotal customers: {1}", totalT, totalC);
                lbl_displayTotal.Show();
            }
            else
            {
                lbl_displayTotal.Hide();
            }
        }

        private void btn_Analysis_Click(object sender, EventArgs e)
        {
            if (_nextTicketShow)
            {
                _waitingListShow = false;

                int entireDayTicket = TotalTicketSold[0] + TotalTicketSold[1] + TotalTicketSold[2] + TotalTicketSold[3];
                int entireDayCus = TotalCutomers[0] + TotalCutomers[1] + TotalCutomers[2] + TotalCutomers[3];

                lbl_NextTicket.Text = string.Format(@"=========================================================" + @"
0am to 2am: " + @"
Total number of tickets sold: {0}   Total number customers: {1}
9am to 3pm: " + "\r\nTotal number of tickets sold: {2}   Total number customers: {3}\r\n5pm to 9pm: " +
                                                    "\r\nTotal number of tickets sold: {4}   Total number customers: {5}\r\n9pm to 12am: " +
                                                    "\r\nTotal number of tickets sold: {6}   Total number customers: {7}\r\nEntire day: " +
                                                    "\r\nTotal number of tickets sold: {8}   Total number customers: {9}\r\n=========================================================",
                                                    TotalTicketSold[0], TotalCutomers[0], TotalTicketSold[1],
                                                    TotalCutomers[1], TotalTicketSold[2], TotalCutomers[2],
                                                    TotalTicketSold[3], TotalCutomers[3], entireDayTicket, entireDayCus);
                lbl_NextTicket.Font = new Font(lbl_NextTicket.Font.FontFamily.Name, 12);
                _nextTicketShow = false;
            }
            else
            {
                lbl_NextTicket.Text = _nextTicketText;
                lbl_NextTicket.Font = new Font(lbl_NextTicket.Font.FontFamily.Name, 140);
                _nextTicketShow = true;
                _waitingListShow = true;
            }
        }

        private void WaitingList_DoubleClick(object sender, EventArgs e)
        {
            if (WaitingList.SelectedItem != null)
            {
                DialogResult dialogResult = MessageBox.Show(@"Next cutomer is: " + WaitingList.SelectedItem,
                                                            @"Are you sure?", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    //do something
                    _nextTicketText = WaitingList.SelectedItem.ToString();
                    _arrayList.Remove(WaitingList.SelectedItem.ToString());
                    WaitingList.Items.Clear();
                    WaitingList.Items.AddRange(_arrayList.ToArray());
                    lbl_NextTicket.Text = _nextTicketText;

                    if (_nextTicketText.StartsWith("A"))
                    {
                        _typeA--;
                    }
                    else if (_nextTicketText.StartsWith("B"))
                    {
                        _typeB--;
                    }
                    else if (_nextTicketText.StartsWith("C"))
                    {
                        _typeC--;
                    }

                    lbl_waitingTickets.Text = string.Format("| A : {0} | B : {1} | C : {2} |\r\nTotal : {3}", _typeA,
                                                            _typeB, _typeC, _arrayList.ToArray().Length);
                }
                else if (dialogResult == DialogResult.No)
                {
                    txb_NOF.Focus();
                }
            }
        }

        public static void OnProcessExit(object sender, EventArgs e)
        {
            ExcelDoc.WriteDoc();
        }

        // The PrintPage event is raised for each page to be printed. 
        private void PrintTicket(object sender, PrintPageEventArgs ev)
        {
            var stringList = new ArrayList
                {
                    "\r\n",
                    "                 www.dragonlegend.ca",
                    "                      905-940-1133",
                    "     25 Lanark Rd. Markham, ON, L3R 8E8",
                    "====================================",
                    "\r\n",
                    "                 Your Ticket Number Is: ",
                    string.Format("{0}", _printedTicketText),
                    "\r\n",
                    "\r\n",
                    "\r\n",
                    "\r\n",
                    "====================================",
                    string.Format("           Date: {0}", _time.ToString(Format))
                };

            var stringListArray = stringList.ToArray();

            var count = 0;
            var count2 = 0;
            float leftMargin = ev.MarginBounds.Left;
            float topMargin = ev.MarginBounds.Top;
            _printFont = new Font("Arial", 8);
            string line;

            // Calculate the number of lines per page.
            var linesPerPage = ev.MarginBounds.Height/
                               _printFont.GetHeight(ev.Graphics);

            // Print each line of the file. 
            while (count < linesPerPage && count2 < stringListArray.Length &&
                   ((line = stringListArray[count2].ToString()) != null))
            {
                float yPos = topMargin + (count*
                                          _printFont.GetHeight(ev.Graphics));
                var parts = line.Split('-');
                parts[0] = parts[0].Trim();
                if (parts[0].Length == 4 && (parts[0].Contains('A') || parts[0].Contains('B') || parts[0].Contains('C')))
                {
                    _printFont = new Font("Arial", 40);
                    if (parts[1].Length == 1)
                        line = " " + line;
                }
                ev.Graphics.DrawString(line, _printFont, Brushes.Black,
                                       leftMargin, yPos, new StringFormat());
                _printFont = new Font("Arial", 8);
                count++;
                count2++;
            }

            // If more lines exist, print another page. 
            ev.HasMorePages = false;
        }

        private void tm_Update_Tick(object sender, EventArgs e)
        {
            DateTime time = DateTime.Now;
            const string format = "ddd MMM d HH:mm yyyy";
            lbl_DateTime.Text = time.ToString(format);
            txb_NOF.Focus();
            int hour = time.Hour;

            if (hour >= 0 && hour < 9)
            {
                _timeRange = 0;
            }
            else if (hour >= 9 && hour < 17)
            {
                _timeRange = 1;
            }
            else if (hour >= 17 && hour < 21)
            {
                _timeRange = 2;
            }
            else if (hour >= 21 && hour < 24)
            {
                _timeRange = 3;
            }

            if (_oldTimeRange == 3 && _timeRange == 0)
            {
                Application.Restart();
            }
            _oldTimeRange = _timeRange;

            saveCount++;
            if (saveCount >= 600000)
            {
                ExcelDoc.WriteDoc();
                saveCount = 0;
            }
        }

        private void btn_Save_Click(object sender, EventArgs e)
        {
            ExcelDoc.WriteDoc();
        }
    }

    public class ExcelDoc
    {
        public static Microsoft.Office.Interop.Excel.Application App = null;
        public static Microsoft.Office.Interop.Excel.Workbook Workbook = null;
        public static Microsoft.Office.Interop.Excel.Worksheet Worksheet = null;
        public static Microsoft.Office.Interop.Excel.Range WorkSheetRange = null;
        public static bool SheetAlreadyExist = true;


        public static void ReadDoc()
        {
            var time = DateTime.Now;
            try
            {
                App = new Microsoft.Office.Interop.Excel.Application {Visible = true};
                Workbook = App.Workbooks.Open("DragonLegendSalesReport-AutoGen", Type.Missing, true, Type.Missing, "19931993");
                Worksheet = (Microsoft.Office.Interop.Excel.Worksheet) Workbook.Sheets[1];

                var temp = "blah";
                const string format = "ddd MMM d, yyyy";
                var i = 4;

                var strArray = new string[13];
                var hasTodayInfo = true;

                while (temp != time.ToString(format))
                {
                    var range = Worksheet.Range["B" + i, "M" + i];
                    var myvalues = (Array) range.Cells.Value;
                    strArray = myvalues.OfType<object>().Select(x => x.ToString()).ToArray();

                    if (strArray.Length == 0)
                    {
                        //Excel doesn't have today's info yet
                        hasTodayInfo = false;
                        break;
                    }

                    temp = strArray[0];

                    i++;
                }

                if (hasTodayInfo)
                {
                    ReceptionScreen.TotalTicketSold[0] += Convert.ToInt32(strArray[1]);
                    ReceptionScreen.TotalTicketSold[1] += Convert.ToInt32(strArray[3]);
                    ReceptionScreen.TotalTicketSold[2] += Convert.ToInt32(strArray[5]);
                    ReceptionScreen.TotalTicketSold[3] += Convert.ToInt32(strArray[7]);

                    ReceptionScreen.TotalCutomers[0] += Convert.ToInt32(strArray[2]);
                    ReceptionScreen.TotalCutomers[1] += Convert.ToInt32(strArray[4]);
                    ReceptionScreen.TotalCutomers[2] += Convert.ToInt32(strArray[6]);
                    ReceptionScreen.TotalCutomers[3] += Convert.ToInt32(strArray[8]);

                    if (Convert.ToInt32(strArray[9]) != ReceptionScreen.TotalTicketSold[0] +
                                                        ReceptionScreen.TotalTicketSold[1] +
                        ReceptionScreen.TotalTicketSold[2] +
                        ReceptionScreen.TotalTicketSold[3])
                    {
                        throw new InvalidDataException();
                    }

                    if (Convert.ToInt32(strArray[10]) != ReceptionScreen.TotalCutomers[0] +
                                                         ReceptionScreen.TotalCutomers[1] +
                        ReceptionScreen.TotalCutomers[2] +
                        ReceptionScreen.TotalCutomers[3])
                    {
                        throw new InvalidDataException();
                    }

                    App.Quit();
                }
                else
                {
                    App.Quit();
                    WriteDoc();
                }

            }
            catch (Exception)
            {
                App.Quit();

                Console.Write(@"File doesn't exist");
                SheetAlreadyExist = false;
                CreateDoc();
            }
        }

        public static void CreateDoc()
        {
            App = new Microsoft.Office.Interop.Excel.Application {Visible = true};
            Workbook = App.Workbooks.Add(1);
            Worksheet = (Microsoft.Office.Interop.Excel.Worksheet) Workbook.Sheets[1];

            //creates the main header
            CreateHeaders(2, 2, "Sales Report", "B2", "M2", 2, "YELLOW", true, 10, "n");
            //creates subheaders
            CreateHeaders(3, 2, "Date", "B3", "C3", 0, "GAINSBORO", true, 10, "");

            CreateHeaders(3, 4, "TIC:0|2", "D3", "D3", 0, "GAINSBORO", true, 10, "");
            CreateHeaders(3, 5, "CUS:0|2", "E3", "E3", 0, "GAINSBORO", true, 10, "");

            CreateHeaders(3, 6, "TIC:9|15", "F3", "F3", 0, "GAINSBORO", true, 10, "");
            CreateHeaders(3, 7, "CUS:9|15", "G3", "G3", 0, "GAINSBORO", true, 10, "");

            CreateHeaders(3, 8, "TIC:17|21", "H3", "H3", 0, "GAINSBORO", true, 10, "");
            CreateHeaders(3, 9, "CUS:17|21", "I3", "I3", 0, "GAINSBORO", true, 10, "");

            CreateHeaders(3, 10, "TIC:21|24", "J3", "J3", 0, "GAINSBORO", true, 10, "");
            CreateHeaders(3, 11, "CUS:21|24", "K3", "K3", 0, "GAINSBORO", true, 10, "");

            CreateHeaders(3, 12, "TIC TOTAL", "L3", "L3", 0, "GAINSBORO", true, 10, "");
            CreateHeaders(3, 13, "CUS TOTAL", "M3", "M3", 0, "GAINSBORO", true, 10, "");

            //add Data to cells
            const string format = "ddd MMM d, yyyy";
            var time = DateTime.Now;

            CreateHeaders(4, 2, "", "B4", "C4", 2, "WHITE", false, 10, "n");
            AddData(4, 2, time.ToString(format), "B4", "C4", "");

            CreateHeaders(4, 4, "", "D4", "D4", 2, "WHITE", false, 10, "n");
            AddData(4, 4, ReceptionScreen.TotalTicketSold[0].ToString(), "D4", "D4", "#,##0");

            CreateHeaders(4, 5, "", "E4", "E4", 2, "WHITE", false, 10, "n");
            AddData(4, 5, ReceptionScreen.TotalCutomers[0].ToString(), "E4", "E4", "#,##0");

            CreateHeaders(4, 6, "", "F4", "F4", 2, "WHITE", false, 10, "n");
            AddData(4, 6, ReceptionScreen.TotalTicketSold[1].ToString(), "F4", "F4", "#,##0");

            CreateHeaders(4, 7, "", "G4", "G4", 2, "WHITE", false, 10, "n");
            AddData(4, 7, ReceptionScreen.TotalCutomers[1].ToString(), "G4", "G4", "#,##0");

            CreateHeaders(4, 8, "", "H4", "H4", 2, "WHITE", false, 10, "n");
            AddData(4, 8, ReceptionScreen.TotalTicketSold[2].ToString(), "H4", "H4", "#,##0");

            CreateHeaders(4, 9, "", "I4", "I4", 2, "WHITE", false, 10, "n");
            AddData(4, 9, ReceptionScreen.TotalCutomers[2].ToString(), "I4", "I4", "#,##0");

            CreateHeaders(4, 10, "", "J4", "J4", 2, "WHITE", false, 10, "n");
            AddData(4, 10, ReceptionScreen.TotalTicketSold[3].ToString(), "J4", "J4", "#,##0");

            CreateHeaders(4, 11, "", "K4", "K4", 2, "WHITE", false, 10, "n");
            AddData(4, 11, ReceptionScreen.TotalCutomers[3].ToString(), "K4", "K4", "#,##0");

            CreateHeaders(4, 12, "", "L4", "L4", 2, "WHITE", false, 10, "n");
            var total = ReceptionScreen.TotalTicketSold[0] + ReceptionScreen.TotalTicketSold[1] +
                        ReceptionScreen.TotalTicketSold[2] + ReceptionScreen.TotalTicketSold[3];
            AddData(4, 12, total.ToString(), "L4", "L4", "#,##0");

            CreateHeaders(4, 13, "", "M4", "M4", 2, "WHITE", false, 10, "n");
            total = ReceptionScreen.TotalCutomers[0] + ReceptionScreen.TotalCutomers[1] +
                    ReceptionScreen.TotalCutomers[2] + ReceptionScreen.TotalCutomers[3];
            AddData(4, 13, total.ToString(), "M4", "M4", "#,##0");

            App.DisplayAlerts = false;
            Workbook.SaveAs("DragonLegendSalesReport-AutoGen",Type.Missing,"19931993",Type.Missing,true);
            App.Quit();
        }

        public static void WriteDoc()
        {
            var activeWriteRowNum = 4;
            App = new Microsoft.Office.Interop.Excel.Application {Visible = true};
            Workbook = App.Workbooks.Open("DragonLegendSalesReport-AutoGen", Type.Missing, false, Type.Missing, "19931993", Type.Missing,true);
            Worksheet = (Microsoft.Office.Interop.Excel.Worksheet) Workbook.Sheets[1];

            var temp = "blah";
            const string format = "ddd MMM d, yyyy";
            var time = DateTime.Now;

            var hasTodayInfo = true;

            while (temp != time.ToString(format))
            {
                var range = Worksheet.Range["B" + activeWriteRowNum, "M" + activeWriteRowNum];
                var myvalues = (Array) range.Cells.Value;
                var strArray = myvalues.OfType<object>().Select(x => x.ToString()).ToArray();

                if (strArray.Length == 0)
                {
                    hasTodayInfo = false;
                    break;
                }

                temp = strArray[0];
                activeWriteRowNum++;
            }

            if (hasTodayInfo)
                activeWriteRowNum--;
            //creates the main header
            CreateHeaders(2, 2, "Sales Report", "B2", "M2", 2, "YELLOW", true, 10, "n");
            //creates subheaders
            CreateHeaders(3, 2, "Date", "B3", "C3", 0, "GAINSBORO", true, 10, "");

            CreateHeaders(3, 4, "TIC:0|2", "D3", "D3", 0, "GAINSBORO", true, 10, "");
            CreateHeaders(3, 5, "CUS:0|2", "E3", "E3", 0, "GAINSBORO", true, 10, "");

            CreateHeaders(3, 6, "TIC:9|15", "F3", "F3", 0, "GAINSBORO", true, 10, "");
            CreateHeaders(3, 7, "CUS:9|15", "G3", "G3", 0, "GAINSBORO", true, 10, "");

            CreateHeaders(3, 8, "TIC:17|21", "H3", "H3", 0, "GAINSBORO", true, 10, "");
            CreateHeaders(3, 9, "CUS:17|21", "I3", "I3", 0, "GAINSBORO", true, 10, "");

            CreateHeaders(3, 10, "TIC:21|24", "J3", "J3", 0, "GAINSBORO", true, 10, "");
            CreateHeaders(3, 11, "CUS:21|24", "K3", "K3", 0, "GAINSBORO", true, 10, "");

            CreateHeaders(3, 12, "TIC TOTAL", "L3", "L3", 0, "GAINSBORO", true, 10, "");
            CreateHeaders(3, 13, "CUS TOTAL", "M3", "M3", 0, "GAINSBORO", true, 10, "");

            //add Data to cells

            CreateHeaders(activeWriteRowNum, 2, "", "B" + activeWriteRowNum, "C" + activeWriteRowNum, 2, "WHITE", false,
                          10, "n");
            AddData(activeWriteRowNum, 2, time.ToString(format), "B" + activeWriteRowNum, "C" + activeWriteRowNum, "");

            CreateHeaders(activeWriteRowNum, 4, "", "D" + activeWriteRowNum, "D" + activeWriteRowNum, 2, "WHITE", false,
                          10, "n");
            AddData(activeWriteRowNum, 4, ReceptionScreen.TotalTicketSold[0].ToString(), "D" + activeWriteRowNum,
                    "D" + activeWriteRowNum, "#,##0");

            CreateHeaders(activeWriteRowNum, 5, "", "E" + activeWriteRowNum, "E" + activeWriteRowNum, 2, "WHITE", false,
                          10, "n");
            AddData(activeWriteRowNum, 5, ReceptionScreen.TotalCutomers[0].ToString(), "E" + activeWriteRowNum,
                    "E" + activeWriteRowNum, "#,##0");

            CreateHeaders(activeWriteRowNum, 6, "", "F" + activeWriteRowNum, "F" + activeWriteRowNum, 2, "WHITE", false,
                          10, "n");
            AddData(activeWriteRowNum, 6, ReceptionScreen.TotalTicketSold[1].ToString(), "F" + activeWriteRowNum,
                    "F" + activeWriteRowNum, "#,##0");

            CreateHeaders(activeWriteRowNum, 7, "", "G" + activeWriteRowNum, "G" + activeWriteRowNum, 2, "WHITE", false,
                          10, "n");
            AddData(activeWriteRowNum, 7, ReceptionScreen.TotalCutomers[1].ToString(), "G" + activeWriteRowNum,
                    "G" + activeWriteRowNum, "#,##0");

            CreateHeaders(activeWriteRowNum, 8, "", "H" + activeWriteRowNum, "H" + activeWriteRowNum, 2, "WHITE", false,
                          10, "n");
            AddData(activeWriteRowNum, 8, ReceptionScreen.TotalTicketSold[2].ToString(), "H" + activeWriteRowNum,
                    "H" + activeWriteRowNum, "#,##0");

            CreateHeaders(activeWriteRowNum, 9, "", "I" + activeWriteRowNum, "I" + activeWriteRowNum, 2, "WHITE", false,
                          10, "n");
            AddData(activeWriteRowNum, 9, ReceptionScreen.TotalCutomers[2].ToString(), "I" + activeWriteRowNum,
                    "I" + activeWriteRowNum, "#,##0");

            CreateHeaders(activeWriteRowNum, 10, "", "J" + activeWriteRowNum, "J" + activeWriteRowNum, 2, "WHITE", false,
                          10, "n");
            AddData(activeWriteRowNum, 10, ReceptionScreen.TotalTicketSold[3].ToString(), "J" + activeWriteRowNum,
                    "J" + activeWriteRowNum, "#,##0");

            CreateHeaders(activeWriteRowNum, 11, "", "K" + activeWriteRowNum, "K" + activeWriteRowNum, 2, "WHITE", false,
                          10, "n");
            AddData(activeWriteRowNum, 11, ReceptionScreen.TotalCutomers[3].ToString(), "K" + activeWriteRowNum,
                    "K" + activeWriteRowNum, "#,##0");

            CreateHeaders(activeWriteRowNum, 12, "", "L" + activeWriteRowNum, "L" + activeWriteRowNum, 2, "WHITE", false,
                          10, "n");
            var total = ReceptionScreen.TotalTicketSold[0] + ReceptionScreen.TotalTicketSold[1] +
                        ReceptionScreen.TotalTicketSold[2] + ReceptionScreen.TotalTicketSold[3];
            AddData(activeWriteRowNum, 12, total.ToString(), "L" + activeWriteRowNum, "L" + activeWriteRowNum, "#,##0");

            CreateHeaders(activeWriteRowNum, 13, "", "M" + activeWriteRowNum, "M" + activeWriteRowNum, 2, "WHITE", false,
                          10, "n");
            total = ReceptionScreen.TotalCutomers[0] + ReceptionScreen.TotalCutomers[1] +
                    ReceptionScreen.TotalCutomers[2] + ReceptionScreen.TotalCutomers[3];
            AddData(activeWriteRowNum, 13, total.ToString(), "M" + activeWriteRowNum, "M" + activeWriteRowNum, "#,##0");

            App.DisplayAlerts = false;
            Workbook.SaveAs("DragonLegendSalesReport-AutoGen", Type.Missing, "19931993", Type.Missing, true);
            App.Quit();
        }

        public static void CreateHeaders(int row, int col, string htext, string cell1,
                                         string cell2, int mergeColumns, string b, bool font, int size, string
                                                                                                            fcolor)
        {
            Worksheet.Cells[row, col] = htext;
            WorkSheetRange = Worksheet.Range[cell1, cell2];
            WorkSheetRange.Merge(mergeColumns);
            switch (b)
            {
                case "YELLOW":
                    WorkSheetRange.Interior.Color = Color.Yellow.ToArgb();
                    break;
                case "GRAY":
                    WorkSheetRange.Interior.Color = Color.Gray.ToArgb();
                    break;
                case "GAINSBORO":
                    WorkSheetRange.Interior.Color =
                        Color.Gainsboro.ToArgb();
                    break;
                case "Turquoise":
                    WorkSheetRange.Interior.Color =
                        Color.Turquoise.ToArgb();
                    break;
                case "PeachPuff":
                    WorkSheetRange.Interior.Color =
                        Color.PeachPuff.ToArgb();
                    break;

            }

            WorkSheetRange.Borders.Color = Color.Black.ToArgb();
            WorkSheetRange.Font.Bold = font;
            WorkSheetRange.ColumnWidth = size;
            WorkSheetRange.Font.Color = fcolor.Equals("") ? Color.White.ToArgb() : Color.Black.ToArgb();
        }

        public static void AddData(int row, int col, string data,
                                   string cell1, string cell2, string format)
        {
            Worksheet.Cells[row, col] = data;
            WorkSheetRange = Worksheet.Range[cell1, cell2];
            WorkSheetRange.Borders.Color = Color.Black.ToArgb();
            WorkSheetRange.NumberFormat = format;
        }
    }
}
