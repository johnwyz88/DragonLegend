using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;
using System.Drawing.Printing;
using System.IO;
using System.Data.SqlClient;

namespace ReceptionScreen
{
    public partial class ReceptionScreen : Form
    {
        private bool waitingListShow = true;
        private bool nextTicketShow = true;
        private string nextTicketText = "A000-0";
        public static int[] totalTicketSold = new int[4];
        private int NOP = 0;
        private string printedTicketText = "A000-0";
        private ArrayList arrayList = new ArrayList();
        private bool parsable = true;
        public static int[] totalCutomers = new int[4];
        private int totalWaitTime = 0;
        private float avgWaitTime = 0;
        private ticket newTicket;
        private Font printFont;
        private static DateTime time = DateTime.Now;
        private static string format = "ddd MMM d HH:mm yyyy";
        private static int typeA, typeB, typeC;
        private static int timeRange;
        private static int oldTimeRange = -1;

        public ReceptionScreen()
        {
            InitializeComponent();
        }

        private void ReceptionScreen_Load(object sender, EventArgs e)
        {
            //UI element positions
            AppDomain.CurrentDomain.ProcessExit += new EventHandler(OnProcessExit); 
            Rectangle screen = Screen.PrimaryScreen.Bounds;
            DateTime time = DateTime.Now;
            string format = "ddd MMM d HH:mm yyyy";
            waitingListShow = true;
            WaitingList.Hide();
            nextTicketShow = true;
            lbl_DateTime.Text = time.ToString(format);
            lbl_CompanyName.Left = screen.Width / 2 - lbl_CompanyName.Width / 2 + screen.Width / 20;
            lbl_CompanyName.Top = screen.Height / 13 - lbl_CompanyName.Height / 13;
            pic_CompanyPicture.Left = screen.Width / 20 - pic_CompanyPicture.Width / 20;
            pic_CompanyPicture.Top = screen.Height / 20 - pic_CompanyPicture.Height / 20;
            lbl_TopBreakLine.Left = 0;
            lbl_TopBreakLine.Top = screen.Height / 20 + pic_CompanyPicture.Height;
            lbl_LabelNextTicket.Left = screen.Width / 7 - lbl_LabelNextTicket.Width / 7;
            lbl_LabelNextTicket.Top = lbl_NextTicket.Top + lbl_NextTicket.Height - lbl_LabelNextTicket.Height * 2;
            lbl_NextTicket.Left = screen.Width / 2 - lbl_NextTicket.Width / 2 + screen.Width / 20;
            lbl_NextTicket.Top = lbl_TopBreakLine.Top + lbl_TopBreakLine.Height / 2;
            lbl_LabelPrintTicket.Left = lbl_LabelNextTicket.Left;
            btn_CreateTicket.Left = lbl_NextTicket.Left + txb_NOF.Width + screen.Width / 17;   
            btn_Reset.Left = screen.Width / 100;
            btn_Reset.Top = screen.Height - btn_Reset.Height * 4 - screen.Height / 100;
            lbl_LabelPrintedNum.Top = lbl_LabelPrintTicket.Top + screen.Height / 12;
            lbl_LabelPrintedNum.Left = lbl_LabelPrintTicket.Left;
            lbl_PrintedNumber.Top = lbl_LabelPrintedNum.Top;
            txb_NOF.Left = lbl_NextTicket.Left + screen.Width / 20;
            lbl_PrintedNumber.Left = txb_NOF.Left;
            lbl_BottomBreakLine.Top = lbl_PrintedNumber.Top + lbl_PrintedNumber.Height + screen.Height / 30;
            lbl_ContactInfo.Top = btn_Reset.Top;
            lbl_ContactInfo.Left = screen.Width / 2 - lbl_ContactInfo.Width / 2;
            WaitingList.Left = screen.Width - WaitingList.Width - screen.Width / 50;
            WaitingList.Top = lbl_TopBreakLine.Top + lbl_TopBreakLine.Height + screen.Height / 100;
            lbl_DateTime.Left = screen.Width - lbl_DateTime.Width - screen.Width / 100;
            lbl_DateTime.Top = btn_Reset.Top;
            btn_Analysis.Top = btn_Reset.Top;
            btn_Analysis.Left = btn_Reset.Left + btn_Reset.Width + screen.Width / 100;
            lbl_waitingTickets.Top = WaitingList.Top + WaitingList.Height + screen.Height / 100;
            lbl_waitingTickets.Left = screen.Width - lbl_waitingTickets.Width - screen.Width / 100;
            txb_NOF.Focus();
            btn_Reset.Visible = false;
            btn_Analysis.Visible = false;
            ExcelDoc.readDoc();
        }

        private void btn_CreateTicket_Click(object sender, EventArgs e)
        {
            if (txb_NOF.Text == null)
            {
                txb_NOF.BackColor = Color.Red;
                return;
            }
            else
            {
                try
                {
                    NOP = int.Parse(txb_NOF.Text);
                }
                catch
                {
                    txb_NOF.BackColor = Color.Red;
                    parsable = false;
                }
            }
            if (parsable && NOP > 0 && NOP < 100)
            {
                txb_NOF.BackColor = Color.White;
                txb_NOF.Clear();
                totalTicketSold[timeRange]++;
                totalCutomers[timeRange] += NOP;

                char type = 'A';
                if (NOP > 0 && NOP <= 5)
                {
                    type = 'A';
                    typeA++;
                }
                else if (NOP > 5 && NOP <= 10)
                {
                    type = 'B';
                    typeB++;
                }
                else if (NOP > 10)
                {
                    type = 'C';
                    typeC++;
                }

                int totalTicketSoldOfDay = totalTicketSold[0] + totalTicketSold[1] + totalTicketSold[2] + totalTicketSold[3];
                printedTicketText = string.Format("{0}{1:D3}-{2}", type, totalTicketSoldOfDay, NOP);
                lbl_PrintedNumber.Text = printedTicketText;

                newTicket = new ticket();
                newTicket.text = printedTicketText;
                newTicket.startTime = DateTime.Now.TimeOfDay.TotalSeconds;

                arrayList.Add(printedTicketText);
                WaitingList.Items.Clear();
                WaitingList.Items.AddRange(arrayList.ToArray());
                lbl_waitingTickets.Text = string.Format("| A : {0} | B : {1} | C : {2} |\r\n            -> Total : {3}", typeA, typeB, typeC, arrayList.ToArray().Length.ToString());

                PrintDocument pd = new PrintDocument();
                pd.PrintPage += new PrintPageEventHandler
                   (this.PrintTicket);
                pd.Print();
            }
            else
                txb_NOF.BackColor = Color.Red;
            parsable = true;
            txb_NOF.Focus();
        }

        private void ReceptionScreen_MouseMove(object sender, MouseEventArgs e)
        {
            if (btn_Reset.Bounds.Contains(e.Location) && !btn_Reset.Visible)
            {
                btn_Reset.Show();
            }
            else
            {
                btn_Reset.Hide();
            }
            if (btn_Analysis.Bounds.Contains(e.Location) && !btn_Analysis.Visible)
            {
                btn_Analysis.Show();
            }
            else
            {
                btn_Analysis.Hide();
            }
            if (WaitingList.Bounds.Contains(e.Location) && !WaitingList.Visible && waitingListShow)
            {
                WaitingList.Show();
            }
            else
            {
                WaitingList.Hide();
            }
        }

        private void btn_Reset_Click(object sender, EventArgs e)
        {
            Application.Restart();
        }       

        private void btn_Analysis_Click(object sender, EventArgs e)
        {
            if (nextTicketShow)
            {
                waitingListShow = false;

                int entireDayTicket = totalTicketSold[0] + totalTicketSold[1] + totalTicketSold[2] + totalTicketSold[3];
                int entireDayCus = totalCutomers[0] + totalCutomers[1] + totalCutomers[2] + totalCutomers[3];

                lbl_NextTicket.Text = "=========================================================" +
                    "\r\n0am to 2am: " +
                    "\r\nTotal number of tickets sold: " + totalTicketSold[0].ToString() +
                    "   Total number customers: " + totalCutomers[0].ToString() +
                    "\r\n9am to 3pm: " +
                    "\r\nTotal number of tickets sold: " + totalTicketSold[1].ToString() +
                    "   Total number customers: " + totalCutomers[1].ToString() +
                    "\r\n5pm to 9pm: " +
                    "\r\nTotal number of tickets sold: " + totalTicketSold[2].ToString() +
                    "   Total number customers: " + totalCutomers[2].ToString() +
                    "\r\n9pm to 12am: " +
                    "\r\nTotal number of tickets sold: " + totalTicketSold[3].ToString() +
                    "   Total number customers: " + totalCutomers[3].ToString() +
                    "\r\nEntire day: " +
                    "\r\nTotal number of tickets sold: " + entireDayTicket.ToString() +
                    "   Total number customers: " + entireDayCus.ToString()
                    + "\r\n=========================================================";
                lbl_NextTicket.Font = new Font(lbl_NextTicket.Font.FontFamily.Name, 12);
                nextTicketShow = false;
            }
            else
            {
                lbl_NextTicket.Text = nextTicketText;
                lbl_NextTicket.Font = new Font(lbl_NextTicket.Font.FontFamily.Name, 140);
                nextTicketShow = true;
                waitingListShow = true;
            }
        }

        private void WaitingList_DoubleClick(object sender, EventArgs e)
        {
            if (WaitingList.SelectedItem != null)
            {
                DialogResult dialogResult = MessageBox.Show("Next cutomer is: " + WaitingList.SelectedItem.ToString(), "Are you sure?", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    //do something
                    nextTicketText = WaitingList.SelectedItem.ToString();
                    arrayList.Remove(WaitingList.SelectedItem.ToString());
                    WaitingList.Items.Clear();
                    WaitingList.Items.AddRange(arrayList.ToArray());
                    lbl_NextTicket.Text = nextTicketText;
                    newTicket.endTime = DateTime.Now.TimeOfDay.TotalSeconds;
                    totalWaitTime = System.Convert.ToInt32(newTicket.endTime - newTicket.startTime);

                    if (nextTicketText.StartsWith("A"))
                    {
                        typeA--;
                    }
                    else if (nextTicketText.StartsWith("B"))
                    {
                        typeB--;
                    }
                    else if (nextTicketText.StartsWith("C"))
                    {
                        typeC--;
                    }

                    lbl_waitingTickets.Text = string.Format("| A : {0} | B : {1} | C : {2} |\r\nTotal : {3}", typeA, typeB, typeC, arrayList.ToArray().Length.ToString());
                }
                else if (dialogResult == DialogResult.No)
                {
                    txb_NOF.Focus();
                }
            }
        }

        private static void OnProcessExit(object sender, EventArgs e)
        {
            ExcelDoc.app = new Microsoft.Office.Interop.Excel.Application();
            ExcelDoc.app.Visible = true;
            ExcelDoc.workbook = ExcelDoc.app.Workbooks.Open("DragonLegendSalesReport-AutoGen");
            ExcelDoc.worksheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelDoc.workbook.Sheets[1];
            ExcelDoc.writeDoc();
        }

        // The PrintPage event is raised for each page to be printed. 
        private void PrintTicket(object sender, PrintPageEventArgs ev)
        {
            ArrayList stringList = new ArrayList();

            stringList.Add("\r\n");
            stringList.Add("                 www.dragonlegend.ca");
            stringList.Add("                      905-940-1133");
            stringList.Add("     25 Lanark Rd. Markham, ON, L3R 8E8");
            stringList.Add("====================================");
            stringList.Add("\r\n");
            stringList.Add("                 Your Ticket Number Is: ");
            stringList.Add(string.Format("{0}", printedTicketText));
            stringList.Add("\r\n");
            stringList.Add("\r\n");
            stringList.Add("\r\n");
            stringList.Add("\r\n");
            stringList.Add("====================================");
            stringList.Add(string.Format("           Date: {0}", time.ToString(format)));

            var stringListArray = stringList.ToArray();

            float linesPerPage = 0;
            float yPos = 0;
            int count = 0;
            int count2 = 0;
            float leftMargin = ev.MarginBounds.Left;
            float topMargin = ev.MarginBounds.Top;
            printFont = new Font("Arial", 8);
            string line = null;

            // Calculate the number of lines per page.
            linesPerPage = ev.MarginBounds.Height /
               printFont.GetHeight(ev.Graphics);

            // Print each line of the file. 
            while (count < linesPerPage && count2 < stringListArray.Length &&
               ((line = stringListArray[count2].ToString()) != null))
            {
                yPos = topMargin + (count *
                   printFont.GetHeight(ev.Graphics));
                var parts = line.Split('-');
                parts[0] = parts[0].Trim();
                if (parts[0].Length == 4 && (parts[0].Contains('A') || parts[0].Contains('B') || parts[0].Contains('C')))
                {
                    printFont = new Font("Arial", 40);
                    if (parts[1].Length == 1)
                        line = " " + line;
                }
                ev.Graphics.DrawString(line, printFont, Brushes.Black,
                   leftMargin, yPos, new StringFormat());
                printFont = new Font("Arial", 8);
                count++;
                count2++;
            }

            // If more lines exist, print another page. 
            ev.HasMorePages = false;
        }

        private void tm_Update_Tick(object sender, EventArgs e)
        {
            DateTime time = DateTime.Now;
            string format = "ddd MMM d HH:mm yyyy";
            lbl_DateTime.Text = time.ToString(format);
            txb_NOF.Focus();
            int hour = time.Hour;
            
            if (hour >= 0 && hour < 2)
            {
                timeRange = 0;
            }
            else if (hour >= 9 && hour < 15)
            {
                timeRange = 1;
            }
            else if (hour >= 17 && hour < 21)
            {
                timeRange = 2;
            }
            else if (hour >= 21 && hour < 24)
            {
                timeRange = 3;
            }

            if (oldTimeRange == 3 && timeRange == 0)
            {
                //Store Data

                //Reset tickets
                totalTicketSold[0] = 0;
                totalTicketSold[1] = 0;
                totalTicketSold[2] = 0;
                totalTicketSold[3] = 0;
            }
            oldTimeRange = timeRange;
        }
    }

    public class ticket
    {
        public string text { get; set; }
        public double startTime { get; set; }
        public double endTime { get; set; }
    }

    public class ExcelDoc
    {
        public static Microsoft.Office.Interop.Excel.Application app = null;
        public static Microsoft.Office.Interop.Excel.Workbook workbook = null;
        public static Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
        public static Microsoft.Office.Interop.Excel.Range workSheet_range = null;
        public static bool sheetAlreadyExist = true;

        public static void readDoc()
        {
            try
            {
                app = new Microsoft.Office.Interop.Excel.Application();
                app.Visible = true;
                workbook = app.Workbooks.Open("DragonLegendSalesReport-AutoGen");
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

                string temp = "blah";
                string format = "ddd MMM d, yyyy";
                DateTime time = DateTime.Now;
                int i = 4;
                string[] strArray = new string[13];

                while (temp != time.ToString(format))
                {
                    Microsoft.Office.Interop.Excel.Range range = worksheet.get_Range("B" + i.ToString(), "M" + i.ToString());
                    System.Array myvalues = (System.Array)range.Cells.Value;
                    strArray = myvalues.OfType<object>().Select(x => x.ToString()).ToArray();

                    temp = strArray[0];
                    i++;
                }

                ReceptionScreen.totalTicketSold[0] += System.Convert.ToInt32(strArray[1]);
                ReceptionScreen.totalTicketSold[1] += System.Convert.ToInt32(strArray[3]);
                ReceptionScreen.totalTicketSold[2] += System.Convert.ToInt32(strArray[5]);
                ReceptionScreen.totalTicketSold[3] += System.Convert.ToInt32(strArray[7]);

                ReceptionScreen.totalCutomers[0] += System.Convert.ToInt32(strArray[2]);
                ReceptionScreen.totalCutomers[1] += System.Convert.ToInt32(strArray[4]);
                ReceptionScreen.totalCutomers[2] += System.Convert.ToInt32(strArray[6]);
                ReceptionScreen.totalCutomers[3] += System.Convert.ToInt32(strArray[8]);

                if (System.Convert.ToInt32(strArray[9]) != ReceptionScreen.totalTicketSold[0] +
                    ReceptionScreen.totalTicketSold[1] + ReceptionScreen.totalTicketSold[2] +
                    ReceptionScreen.totalTicketSold[3])
                {
                    throw new InvalidDataException();
                }

                if (System.Convert.ToInt32(strArray[10]) != ReceptionScreen.totalCutomers[0] +
                        ReceptionScreen.totalCutomers[1] + ReceptionScreen.totalCutomers[2] +
                        ReceptionScreen.totalCutomers[3])
                {
                    throw new InvalidDataException();
                }

                ExcelDoc.app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelDoc.app);
                ExcelDoc.app = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception e)
            {
                ExcelDoc.app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelDoc.app);
                ExcelDoc.app = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Console.Write("File doesn't exist");
                sheetAlreadyExist = false;
                createDoc();
                writeDoc();
            }
            finally
            {
            }           
        }

        public static void createDoc()
        {
            try
            {
                app = new Microsoft.Office.Interop.Excel.Application();
                app.Visible = true;
                workbook = app.Workbooks.Add(1);
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
            }
            catch (Exception e)
            {
                Console.Write("Error");
            }
            finally
            {
            }
        }

        public static void writeDoc()
        {

            //creates the main header
            createHeaders(2, 2, "Sales Report", "B2", "M2", 2, "YELLOW", true, 10, "n");
            //creates subheaders
            createHeaders(3, 2, "Date", "B3", "C3", 0, "GAINSBORO", true, 10, "");

            createHeaders(3, 4, "TIC:0|2", "D3", "D3", 0, "GAINSBORO", true, 10, "");
            createHeaders(3, 5, "CUS:0|2", "E3", "E3", 0, "GAINSBORO", true, 10, "");

            createHeaders(3, 6, "TIC:9|15", "F3", "F3", 0, "GAINSBORO", true, 10, "");
            createHeaders(3, 7, "CUS:9|15", "G3", "G3", 0, "GAINSBORO", true, 10, "");

            createHeaders(3, 8, "TIC:17|21", "H3", "H3", 0, "GAINSBORO", true, 10, "");
            createHeaders(3, 9, "CUS:17|21", "I3", "I3", 0, "GAINSBORO", true, 10, "");

            createHeaders(3, 10, "TIC:21|24", "J3", "J3", 0, "GAINSBORO", true, 10, "");
            createHeaders(3, 11, "CUS:21|24", "K3", "K3", 0, "GAINSBORO", true, 10, "");

            createHeaders(3, 12, "TIC TOTAL", "L3", "L3", 0, "GAINSBORO", true, 10, "");
            createHeaders(3, 13, "CUS TOTAL", "M3", "M3", 0, "GAINSBORO", true, 10, "");

            //add Data to cells
            string format = "ddd MMM d, yyyy";
            DateTime time = DateTime.Now;

            createHeaders(4, 2, "", "B4", "C4", 2, "WHITE", false, 10, "n");
            addData(4, 2, time.ToString(format), "B4", "C4", "");

            createHeaders(4, 4, "", "D4", "D4", 2, "WHITE", false, 10, "n");
            addData(4, 4, ReceptionScreen.totalTicketSold[0].ToString(), "D4", "D4", "#,##0");

            createHeaders(4, 5, "", "E4", "E4", 2, "WHITE", false, 10, "n");
            addData(4, 5, ReceptionScreen.totalCutomers[0].ToString(), "E4", "E4", "#,##0");

            createHeaders(4, 6, "", "F4", "F4", 2, "WHITE", false, 10, "n");
            addData(4, 6, ReceptionScreen.totalTicketSold[1].ToString(), "F4", "F4", "#,##0");

            createHeaders(4, 7, "", "G4", "G4", 2, "WHITE", false, 10, "n");
            addData(4, 7, ReceptionScreen.totalCutomers[1].ToString(), "G4", "G4", "#,##0");

            createHeaders(4, 8, "", "H4", "H4", 2, "WHITE", false, 10, "n");
            addData(4, 8, ReceptionScreen.totalTicketSold[2].ToString(), "H4", "H4", "#,##0");

            createHeaders(4, 9, "", "I4", "I4", 2, "WHITE", false, 10, "n");
            addData(4, 9, ReceptionScreen.totalCutomers[2].ToString(), "I4", "I4", "#,##0");

            createHeaders(4, 10, "", "J4", "J4", 2, "WHITE", false, 10, "n");
            addData(4, 10, ReceptionScreen.totalTicketSold[3].ToString(), "J4", "J4", "#,##0");

            createHeaders(4, 11, "", "K4", "K4", 2, "WHITE", false, 10, "n");
            addData(4, 11, ReceptionScreen.totalCutomers[3].ToString(), "K4", "K4", "#,##0");

            createHeaders(4, 12, "", "L4", "L4", 2, "WHITE", false, 10, "n");
            var total = ReceptionScreen.totalTicketSold[0] + ReceptionScreen.totalTicketSold[1] + ReceptionScreen.totalTicketSold[2] + ReceptionScreen.totalTicketSold[3];
            addData(4, 12, total.ToString(), "L4", "L4", "#,##0");

            createHeaders(4, 13, "", "M4", "M4", 2, "WHITE", false, 10, "n");
            total = ReceptionScreen.totalCutomers[0] + ReceptionScreen.totalCutomers[1] + ReceptionScreen.totalCutomers[2] + ReceptionScreen.totalCutomers[3];
            addData(4, 13, total.ToString(), "M4", "M4", "#,##0");

            //add end row
            createHeaders(5, 2, "", "B5", "M5", 2, "GRAY", true, 10, "");

            ExcelDoc.app.DisplayAlerts = false;
            ExcelDoc.workbook.SaveAs("DragonLegendSalesReport-AutoGen");
            ExcelDoc.app.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelDoc.app);
            ExcelDoc.app = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public static void createHeaders(int row, int col, string htext, string cell1,
        string cell2, int mergeColumns, string b, bool font, int size, string
        fcolor)
        {
            worksheet.Cells[row, col] = htext;
            workSheet_range = worksheet.get_Range(cell1, cell2);
            workSheet_range.Merge(mergeColumns);
            switch (b)
            {
                case "YELLOW":
                    workSheet_range.Interior.Color = System.Drawing.Color.Yellow.ToArgb();
                    break;
                case "GRAY":
                    workSheet_range.Interior.Color = System.Drawing.Color.Gray.ToArgb();
                    break;
                case "GAINSBORO":
                    workSheet_range.Interior.Color =
            System.Drawing.Color.Gainsboro.ToArgb();
                    break;
                case "Turquoise":
                    workSheet_range.Interior.Color =
            System.Drawing.Color.Turquoise.ToArgb();
                    break;
                case "PeachPuff":
                    workSheet_range.Interior.Color =
            System.Drawing.Color.PeachPuff.ToArgb();
                    break;
                default:
                    //  workSheet_range.Interior.Color = System.Drawing.Color..ToArgb();
                    break;
            }

            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
            workSheet_range.Font.Bold = font;
            workSheet_range.ColumnWidth = size;
            if (fcolor.Equals(""))
            {
                workSheet_range.Font.Color = System.Drawing.Color.White.ToArgb();
            }
            else
            {
                workSheet_range.Font.Color = System.Drawing.Color.Black.ToArgb();
            }
        }

        public static void addData(int row, int col, string data,
            string cell1, string cell2, string format)
        {
            worksheet.Cells[row, col] = data;
            workSheet_range = worksheet.get_Range(cell1, cell2);
            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
            workSheet_range.NumberFormat = format;
        }
    }
}
