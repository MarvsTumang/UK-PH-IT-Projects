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

namespace ProjectTemplate
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
            //Rplace n to number of sub process
            //FormControl.MaxPercSubProc = 100.00 / n;

            

            FormControl.SetStatus("Done!");
            //Thread.Sleep(3000);
            //Environment.Exit(0);
        }
    }
}
