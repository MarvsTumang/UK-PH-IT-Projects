using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ArrowGlobal
{
    public partial class MainForm : Form
    {
        public static CustomProgressBar pbarMain;
        public static CustomProgressBar pbarSub;
        public static ToolStripStatusLabel tsStatus;
        public static DataGridView dataGridView;

        public MainForm()
        {
            InitializeComponent();
            pbarMain = cpbMain;
            pbarSub = cpbSub;
            tsStatus = tsLabel;
            dataGridView = dgView;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            Thread thread = new Thread(new ThreadStart(Process));
            thread.Start();
        }

        private void Process()
        {
            Thread.Sleep(1000);
            DataTable table;
            List<PhoneNumber> pNumbers;
            List<CCLN> cclns;
            FormControl.MaxPercSubProc = 100.00 / 5;

            string textFile = @"C:\Users\mtumang\Documents\Projects\Arrow\TestData\Text File Template.txt";
            string teleAppend = @"C:\Users\mtumang\Documents\Projects\Arrow\TestData\Trace Tele-Appends - Template.xls";
            string cclnFile = @"C:\Users\mtumang\Documents\Projects\Arrow\TestData\CCL_Number_Mapping.xlsx";

            string[] customColumns = new string[]
            {
                "ArrowKey||txt|",
                "CustAcctNo||txt|",
                "DefaultDate|||01/01/1900",
                "DefaultAmount|||0.00",
                "Debtor1_addressLine1|Address line 1||",
                "Debtor1_addressLine2|Address line 2||",
                "Debtor1_addressLine3|City||",
                "Debtor1_addressLine4|County||",
                "Debtor1_postcode|Post Code||",
            };

            table = Table.FromTextFile(textFile, customColumns: customColumns);

            FormControl.ViewData(table);

            int[] columns = new int[] { 1, 2, 4, 5 };
            pNumbers = ExcelFile.ToList<PhoneNumber>(teleAppend, columns);

            //string arrowKeys = String.Join(", ", pNumbers.AsEnumerable().Select(c => c.ArrowKey).ToArray());
            
            columns = new int[] { 1, 2 };
            cclns = ExcelFile.ToList<CCLN>(cclnFile, columns);

            Table.InsertColumns(table, new string[] { "NumberMobile1|43||", "NumberWork1|44||", "NumberHome1|45||", "Entity|51|int|2913", "CCLN|52|int|" });

            double subPerc = 0.00;
            double stepPerc = FormControl.MaxPercSubProc / table.Rows.Count;

            FormControl.SetStatus("Formatting Data...");

            foreach (DataRow row in table.Rows)
            {
                subPerc = (table.Rows.IndexOf(row) + 1) / (double)table.Rows.Count * 100.00;
                FormControl.SetSubProgress(subPerc);
                FormControl.SetMainProgress(stepPerc);

                string arrowKey = row["arrowKey"].ToString().Trim().ToLower();
                PhoneNumber pNum = pNumbers.Find(p => p.ArrowKey.Trim().ToLower() == arrowKey);
                
                if (pNum != null)
                {
                    row["NumberMobile1"] = PadLeft(pNum.Mobile, 11);
                    row["NumberWork1"] = PadLeft(pNum.Work, 11);
                    row["NumberHome1"] = PadLeft(pNum.Home, 11);
                }

                string buyerName = row["BuyerName"].ToString().Trim().ToLower();
                CCLN ccln = cclns.Find(c => c.ArrowEntity.Trim().ToLower() == buyerName);

                if (ccln != null)
                {
                    row["CCLN"] = ccln.LicenseNumber;
                }

                row["HomePhoneNumber"] = PadLeft(row["HomePhoneNumber"], 11);
                row["WorkPhoneNumber"] = PadLeft(row["WorkPhoneNumber"], 11);
                row["MobilePhoneNumber"] = PadLeft(row["MobilePhoneNumber"], 11);

                if (!IsNullOrEmpty(row["Debtor1_addressLine5"]))
                {
                    row["County"] = String.Format("{0}, {1}", row["County"], row["Debtor1_addressLine5"]);
                }
            }

            FormControl.ViewData(table);
            
            string excelFileName = textFile.Remove(textFile.LastIndexOf('.'), 4) + " - To Load.xls";

            ExcelFile.SaveExcel(table, excelFileName);

            FormControl.SetStatus("Done!");

            //Thread.Sleep(3000);

            //Environment.Exit(0);
        }

        private string PadLeft(object obj, int totalWidth, char paddingChar = '0')
        {
            try
            {
                if (!String.IsNullOrEmpty(obj.ToString()))
                {
                    obj = obj.ToString().Trim().PadLeft(totalWidth, paddingChar);
                }
                else
                {
                    obj = String.Empty;
                }
            }
            catch 
            {
                obj = String.Empty;
            }

            return obj.ToString();
        }

        private bool IsNullOrEmpty(object obj)
        {
            if (obj != null && !String.IsNullOrEmpty(obj.ToString()))
            {
                return false;
            }
            else
            {
                obj = String.Empty;
                return true;
            }
        }
    }
}
