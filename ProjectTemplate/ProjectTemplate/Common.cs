using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ProjectTemplate
{
    #region Customize ProgressBar
    public enum ProgressBarDisplayText
    {
        Percentage,
        CustomText
    }

    public class CustomProgressBar : ProgressBar
    {
        //Property to set to decide whether to print a % or Text
        public ProgressBarDisplayText DisplayStyle { get; set; }

        //Property to hold the custom text
        public String CustomText { get; set; }

        public CustomProgressBar()
        {
            // Modify the ControlStyles flags
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Rectangle rect = ClientRectangle;
            Graphics g = e.Graphics;

            ProgressBarRenderer.DrawHorizontalBar(g, rect);
            rect.Inflate(-1, -1);
            if (Value > 0)
            {
                // As we doing this ourselves we need to draw the chunks on the progress bar
                Rectangle clip = new Rectangle(rect.X, rect.Y, (int)Math.Round(((float)Value / Maximum) * rect.Width), rect.Height);
                ProgressBarRenderer.DrawHorizontalChunks(g, clip);
            }

            // Set the Display text (Either a % amount or our custom text
            string text = DisplayStyle == ProgressBarDisplayText.Percentage ? Value.ToString() + '%' : CustomText;


            using (Font f = new Font(FontFamily.GenericSerif, 10))
            {

                SizeF len = g.MeasureString(text, f);
                // Calculate the location of the text (the middle of progress bar)
                // Point location = new Point(Convert.ToInt32((rect.Width / 2) - (len.Width / 2)), Convert.ToInt32((rect.Height / 2) - (len.Height / 2)));
                Point location = new Point(Convert.ToInt32((Width / 2) - len.Width / 2), Convert.ToInt32((Height / 2) - len.Height / 2));
                // The commented-out code will centre the text into the highlighted area only. This will centre the text regardless of the highlighted area.
                // Draw the custom text
                g.DrawString(text, f, Brushes.Black, location);
            }
        }
    }
    #endregion

    #region Excel related functions
    public static class ExcelFile
    {
        private static double mainPercStep;
        private static double subPerc;

        // Create List of Generic type from Excel file
        public static List<T> ToList<T>(string excelFilePath, int[] columns = null, int headerRow = 1, bool acceptDuplicate = false) where T : class, new()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-GB");
            Excel.Application xlApp = null;
            Excel.Workbook xlWb = null;
            subPerc = 0.00;
            string line;

            try
            {
                FormControl.SetStatus("Reading Excel File...");
                List<T> list = new List<T>();
                PropertyInfo[] props = null;
                PropertyInfo propInfo = null;
                xlApp = new Excel.Application();
                xlWb = xlApp.Workbooks.Open(excelFilePath);
                Excel.Worksheet xlWs = xlWb.Worksheets[1];
                Excel.Range xlRange = xlWs.UsedRange;
                int rowCount = xlRange.Rows.Count;
                int columnCount = (columns != null && columns.Length > 0) ? columns.Length : xlRange.Columns.Count;
                int recordCount = rowCount - headerRow;
                mainPercStep = FormControl.MaxPercSubProc / recordCount;

                T obj = new T();
                props = obj.GetType().GetProperties();

                if (columns == null || columns.Length == 0)
                {
                    columns = Enumerable.Range(1, columnCount).ToArray();
                }

                for (int row = 1; row <= recordCount; row++)
                {
                    subPerc = row / (double)recordCount * 100.00;
                    FormControl.SetSubProgress(subPerc);
                    FormControl.SetMainProgress(mainPercStep);

                    obj = new T();
                    line = string.Empty;
                    
                    for (int col = 0; col < columnCount; col++)
                    {
                        object value;
                        try
                        {
                            propInfo = props[col];
                            Type type = Nullable.GetUnderlyingType(propInfo.PropertyType) ?? propInfo.PropertyType;
                            value = Convert.ChangeType(xlWs.Cells[headerRow + row, columns[col]].Value, type);
                            propInfo.SetValue(obj, value, null);
                        }
                        catch //(Exception ex)
                        {
                            continue;
                        }

                        line += value != null ? value.ToString() : String.Empty;
                    }

                    if (!acceptDuplicate && list.Contains(obj) || String.IsNullOrEmpty(line.Trim()))
                    {
                        continue;
                    }

                    list.Add(obj);
                }

                return list;
            }
            catch
            {
                return null;
            }
            finally
            {
                xlWb.Close();
                xlApp.Quit();
            }
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public static void SaveExcel(DataTable dataTable, string fileName)
        {
            FormControl.SetStatus("Saving Excel File...");
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-GB");
            string dir = new FileInfo(fileName).DirectoryName;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWb = xlApp.Workbooks.Add(Type.Missing);
            Excel.Worksheet xlWs = xlWb.Worksheets[1];
            Excel.Range xlRange = null;
            Type[] types = new Type[dataTable.Columns.Count];
            int rowIndex = 0;
            int colIndex = 0;
            subPerc = 0.00;
            mainPercStep = FormControl.MaxPercSubProc / dataTable.Rows.Count;

            foreach (DataColumn column in dataTable.Columns)
            {
                xlRange = xlWs.Cells[1, column.Ordinal + 1] as Excel.Range;
                xlRange.Value = column.ColumnName;

                if (dataTable.Columns[colIndex].DataType == typeof(string))
                {
                    xlRange.EntireColumn.NumberFormat = "@";
                }
                else if (dataTable.Columns[colIndex].DataType == typeof(DateTime))
                {
                    xlRange.EntireColumn.NumberFormat = "dd/mm/yyyy hh:mm:ss";
                }

                colIndex++;
            }

            foreach (DataRow row in dataTable.Rows)
            {
                rowIndex = dataTable.Rows.IndexOf(row);
                subPerc = (rowIndex + 1) / (double) dataTable.Rows.Count * 100.00;

                FormControl.SetSubProgress(subPerc);
                FormControl.SetMainProgress(mainPercStep);

                xlRange = xlWs.Range[xlWs.Cells[rowIndex + 2, 1], xlWs.Cells[rowIndex + 2, dataTable.Columns.Count]];
                xlRange.Value = row.ItemArray;
            }

            xlWb.SaveAs(fileName, Excel.XlFileFormat.xlExcel8);
            xlWb.Close();
            xlApp.Quit();
        }
    }
    #endregion

    #region DataTable related functions
    public static class Table
    {
        private static double mainPercStep;

        // Transfer Text file data to DataTable
        public static DataTable FromTextFile(string textFilePath, char del = '|', string[] customColumns = null)
        {
            FormControl.SetStatus("Reading Text File...");
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-GB");
            DataTable dataTable = new DataTable();
            Type dataType;
            string[] lines = File.ReadAllLines(textFilePath);
            string[] headers = lines[0].Split(del);
            string headerName;
            object defaultValue;
            double subPerc = 0.0;
            
            mainPercStep = FormControl.MaxPercSubProc / (lines.Length - 1);

            foreach (string header in headers)
            {
                headerName = header;
                defaultValue = null;
                dataType = typeof(string);

                if (customColumns != null && customColumns.Length != 0)
                {
                    foreach (string column in customColumns)
                    {
                        string[] columns = column.Split(del);

                        if (header.Trim().ToLower().Equals(columns[0].Trim().ToLower()))
                        {
                            headerName = CreateColumn(dataTable, columns);
                            break;
                        }
                    }
                }

                if (!dataTable.Columns.Contains(headerName))
                {
                    DataColumn dataColumn = new DataColumn
                    {
                        DefaultValue = defaultValue,
                        DataType = dataType,
                        ColumnName = headerName
                    };

                    dataTable.Columns.Add(dataColumn);
                }
            }

            Type[] types = new Type[headers.Length];

            for (int row = 1; row < lines.Length; row++)
            {
                subPerc = row / (lines.Length - 1.00) * 100.00;
                FormControl.SetSubProgress(subPerc);
                FormControl.SetMainProgress(mainPercStep);

                string rowString = lines[row];

                if (rowString.Trim().Length == 0)
                {
                    continue;
                }

                DataRow dataRow = dataTable.Rows.Add();

                for (int col = 0; col < headers.Length; col++)
                {
                    if (row == 1)
                    {
                        types[col] = dataTable.Columns[col].DataType;
                    }

                    string[] columns = rowString.Split(del);

                    if (!String.IsNullOrEmpty(columns[col]))
                    {
                        if (types[col] == typeof(string))
                        {
                            dataRow[col] = columns[col];
                        }
                        else if (types[col] == typeof(int))
                        {
                            dataRow[col] = Int32.Parse(columns[col]);
                        }
                        else if (types[col] == typeof(DateTime))
                        {
                            dataRow[col] = DateTime.Parse(columns[col]);
                        }
                        else if (types[col] == typeof(double))
                        {
                            dataRow[col] = Double.Parse(columns[col]);
                        }
                        else if (types[col] == typeof(bool))
                        {
                            dataRow[col] = Boolean.Parse(columns[col]);
                        }
                        else
                        {
                            dataRow[col] = columns[col];
                        }
                    }
                }
            }

            return dataTable;
        }

        //Insert additional column/s to existing DataTable
        public static void InsertColumns(DataTable dataTable, string[] newColumns, char del = '|')
        {
            if (newColumns != null && newColumns.Length != 0)
            {
                foreach (string column in newColumns)
                {
                    string[] columns = column.Split(del);
                    CreateColumn(dataTable, columns);
                }
            }
        }

        private static string CreateColumn(DataTable dataTable, string[] columns)
        {
            DataColumn dataColumn;
            string headerName = columns[0];
            object defaultValue = null;
            Type dataType = typeof(string);
            int? ordinal = null;

            if (!String.IsNullOrEmpty(columns[1].Trim()))
            {
                try
                {
                    ordinal = Int32.Parse(columns[1]);
                }
                catch
                {
                    headerName = columns[1];
                }
            }

            if (!String.IsNullOrEmpty(columns[2].Trim()))
            {
                dataType = GetDataType(columns[2]);
            }

            if (!String.IsNullOrEmpty(columns[3].Trim()))
            {
                defaultValue = columns[3];
            }

            dataColumn = new DataColumn
            {
                DefaultValue = defaultValue,
                DataType = dataType,
                ColumnName = headerName
            };

            dataTable.Columns.Add(dataColumn);

            if (ordinal.HasValue)
            {
                dataTable.Columns[headerName].SetOrdinal((int)ordinal);
            }

            return headerName;
        }

        private static Type GetDataType(string stringType)
        {
            switch (stringType)
            {
                case "str":
                    stringType = "System.String";
                    break;
                case "int":
                    stringType = "System.Int32";
                    break;
                case "date":
                    stringType = "System.DateTime";
                    break;
                case "dbl":
                    stringType = "System.Double";
                    break;
                case "bool":
                    stringType = "System.Boolean";
                    break;
                default:
                    stringType = "System.String";
                    break;
            }

            return Type.GetType(stringType);
        }
    }
    #endregion

    #region Database Access
    public static class Db
    {
        private static string ConnString = ConfigurationManager.ConnectionStrings["ConnString"].ConnectionString;// { private get; set; }
        private static double mainPercStep;
        private static SqlConnection conn = new SqlConnection(ConnString);

        private static SqlDataReader Query(string query, out int count, Dictionary<string, object> parameters = null,  bool nonQuery = false)
        {
            conn.Open();

            using (SqlCommand cmd = new SqlCommand(query, conn))
            {
                if (query.Contains("ssrs_"))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                }

                if (parameters != null && parameters.Count > 0)
                {
                    foreach (var p in parameters)
                    {
                        cmd.Parameters.AddWithValue(p.Key, p.Value);
                    }
                }

                SqlDataReader rdr = cmd.ExecuteReader();
                count = rdr.Cast<object>().Count();
                rdr.Close();

                if (count > 0)
                {
                    return cmd.ExecuteReader(CommandBehavior.CloseConnection);
                }

                return null;
            }
        }

        public static List<T> ToList<T>(string query, Dictionary<string, object> parameters = null, bool acceptDuplicate = false) where T : class, new()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-GB");
            SqlDataReader rdr = null;
            List<T> list = new List<T>();
            int rowCount = 0;

            FormControl.SetStatus("Querying Database...");

            T obj = new T();
            PropertyInfo[] props = obj.GetType().GetProperties();
            PropertyInfo propInfo;

            rdr = Query(query, out rowCount, parameters);
            double stepSub = 100.00 / rowCount;
            mainPercStep = FormControl.MaxPercSubProc / rowCount;

            if (rdr.HasRows)
            {
                while (rdr.Read())
                {
                    FormControl.SetSubProgress(stepSub);
                    FormControl.SetMainProgress(mainPercStep);

                    obj = new T();

                    for (int i = 0; i < props.Length; i++)
                    {
                        propInfo = props[i];
                        propInfo.SetValue(obj, rdr.GetValue(i), null);
                    }

                    if (!acceptDuplicate && list.Contains(obj))
                    {
                        continue;
                    }

                    list.Add(obj);
                }
            }
            else
            {
                return null;
            }
            
            return list;
        }
    }
    #endregion

    #region Form Controls Access
    static class FormControl
    {
        public static double MaxPercSubProc { get; set; }
        private static double mainProgress;

        public static void SetStatus(string status)
        {
            MainForm.tsStatus.GetCurrentParent().Invoke(new System.Action(() =>
            {
                MainForm.tsStatus.Text = status;
            }));
        }

        public static void SetSubProgress(double percent)
        {
            //Thread.Sleep(50);

            MainForm.pbarSub.Invoke(new System.Action(() =>
            {
                MainForm.pbarSub.Value = (int) percent;
            }));

            if (percent == 100)
            {
                percent = 0;
            }
        }

        public static void SetMainProgress(double percent)
        {
            mainProgress += percent;
            //Thread.Sleep(50);

            MainForm.pbarMain.Invoke(new System.Action(() =>
            {
                MainForm.pbarMain.Value = (int)Math.Round(mainProgress, MidpointRounding.AwayFromZero);
            }));
        }

        public static void ViewData(DataTable table)
        {
            MainForm.dataGridView.Invoke(new Action(() =>
            {
                MainForm.dataGridView.DataSource = null;
                MainForm.dataGridView.DataSource = table;
            }));
        }
    }
    #endregion

    #region AES256 Encryption for local applications
    public static class AES256
    {
        private static byte[] AES_Encrypt(byte[] bytesToBeEncrypted, byte[] passwordBytes)
        {
            byte[] encryptedBytes = null;

            byte[] saltBytes = new byte[] { 0, 9, 2, 5, 8, 8, 9, 1, 6, 1, 0 };

            using (MemoryStream ms = new MemoryStream())
            {
                using (RijndaelManaged AES = new RijndaelManaged())
                {
                    AES.KeySize = 256;
                    AES.BlockSize = 128;

                    var key = new Rfc2898DeriveBytes(passwordBytes, saltBytes, 1000);
                    AES.Key = key.GetBytes(AES.KeySize / 8);
                    AES.IV = key.GetBytes(AES.BlockSize / 8);

                    AES.Mode = CipherMode.CBC;

                    using (var cs = new CryptoStream(ms, AES.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(bytesToBeEncrypted, 0, bytesToBeEncrypted.Length);
                        cs.Close();
                    }
                    encryptedBytes = ms.ToArray();
                }
            }

            return encryptedBytes;
        }

        public static string EncryptText(string input, string password)
        {
            // Get the bytes of the string
            byte[] bytesToBeEncrypted = Encoding.UTF8.GetBytes(input);
            byte[] passwordBytes = Encoding.UTF8.GetBytes(password);

            // Hash the password with SHA256
            passwordBytes = SHA256.Create().ComputeHash(passwordBytes);

            byte[] bytesEncrypted = AES_Encrypt(bytesToBeEncrypted, passwordBytes);

            string result = Convert.ToBase64String(bytesEncrypted);

            return result;
        }

        private static byte[] AES_Decrypt(byte[] bytesToBeDecrypted, byte[] passwordBytes)
        {
            byte[] decryptedBytes = null;

            byte[] saltBytes = new byte[] { 0, 9, 2, 5, 8, 8, 9, 1, 6, 1, 0 };

            using (MemoryStream ms = new MemoryStream())
            {
                using (RijndaelManaged AES = new RijndaelManaged())
                {
                    AES.KeySize = 256;
                    AES.BlockSize = 128;

                    var key = new Rfc2898DeriveBytes(passwordBytes, saltBytes, 1000);
                    AES.Key = key.GetBytes(AES.KeySize / 8);
                    AES.IV = key.GetBytes(AES.BlockSize / 8);

                    AES.Mode = CipherMode.CBC;

                    using (var cs = new CryptoStream(ms, AES.CreateDecryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(bytesToBeDecrypted, 0, bytesToBeDecrypted.Length);
                        cs.Close();
                    }
                    decryptedBytes = ms.ToArray();
                }
            }

            return decryptedBytes;
        }

        public static string DecryptText(string input, string password)
        {
            // Get the bytes of the string
            byte[] bytesToBeDecrypted = Convert.FromBase64String(input);
            byte[] passwordBytes = Encoding.UTF8.GetBytes(password);
            passwordBytes = SHA256.Create().ComputeHash(passwordBytes);

            byte[] bytesDecrypted = AES_Decrypt(bytesToBeDecrypted, passwordBytes);

            string result = Encoding.UTF8.GetString(bytesDecrypted);

            return result;
        }
    }
    #endregion
}
