using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;
using System.Windows.Forms.Design;

namespace ExcelTrayLibrary
{

    public enum MixType
    {
        Normal = 0,
        AsText = 1,
        AsTextExtended = 2
    }

    public enum UseHeader
    {
        YES,
        NO
    }

    public enum ExcelProvider
    {
        XLS,
        XLSX
    }

    /// <summary>
    /// A tray component which provides an easy way to create connection
    /// string for Excel files. Also includes a method for runtime use
    /// to get sheet names for the selected excel file.
    /// </summary>
    public class ExcelHelper : Component 
    {
        [Browsable(true), EditorBrowsable(EditorBrowsableState.Always),
         Description("Specifies .xls or .xlsx"), Category("Behavior")]
        public ExcelProvider Provider { get; set; }
        [Browsable(true), EditorBrowsable(EditorBrowsableState.Always), Category("Behavior"),
         Description("Indicates whether the first row in a sheet is data or column headers")]
        public UseHeader Headers { get; set; }
        [Browsable(true), EditorBrowsable(EditorBrowsableState.Always), Category("Behavior"), Description("IMEX Mode for Excel file")]
        public MixType Mode { get; set; } = MixType.AsText;
        private string _fileName;
        /// <summary>
        /// Set Excel file name and if no file name refresh properties e.g. FileExits and OleDbConnection string
        /// </summary>
        [Category("Behavior"), Description("Excel file to open"), Editor(typeof(ExcelFileNameEditor),typeof(System.Drawing.Design.UITypeEditor))]
        [RefreshProperties(RefreshProperties.All)]
        public string FileName
        {
            get => _fileName;
            set
            {
                _fileName = value;
                Provider = Path.GetExtension(value)?.ToLower() == ".xlsx" ? ExcelProvider.XLSX : ExcelProvider.XLS;
            }
        }
        [Browsable(true), EditorBrowsable(EditorBrowsableState.Always), Category("Data"), Description("U")]
        public string OleDbConnectionString => ConnectionString();
        private string _oleExelVersion = "8.0";
        private string _provider = "Microsoft.Jet.OLEDB.4.0";
        private void SetVersions(ExcelProvider pProvider)
        {
            if (pProvider == ExcelProvider.XLSX)
            {
                _oleExelVersion = "12.0";
                _provider = "Microsoft.ACE.OLEDB.12.0";
            }
            else
            {
                _oleExelVersion = "8.0";
                _provider = "Microsoft.Jet.OLEDB.4.0";
            }
        }
        /// <summary>
        /// True if file exists currently
        /// </summary>
        [Browsable(true), EditorBrowsable(EditorBrowsableState.Always), Category("Data"), Description("Check if excel file exists")]
        public bool FileExists => ExcelFileExists();

        private bool ExcelFileExists()
        {
            // is there a file in FileName property?
            if (string.IsNullOrWhiteSpace(FileName))
            {
                return false;
            }

            return File.Exists(FileName) ? true : false;

        }
        /// <summary>
        /// Get sheet names for the current excel file, sheet names with spaces will be escaped for us.
        /// </summary>
        /// <returns></returns>
        /// <remarks>
        /// OleDb sorts sheets names a-z so the sheet names
        /// will not be in ordinal position. Only way around this
        /// is to use Open XML, automation or a library such
        /// as GemBox, Spreadsheet light etc.
        /// 
        /// As coded there is no exception handling. We are working
        /// in the IDE so the exception will be caught a displayed.
        /// 
        /// See also
        /// https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.design.filenameeditor?redirectedfrom=MSDN&view=netframework-4.7.2
        /// </remarks>
        public List<string> SheetNames()
        {
            var sheetNames = new List<string>();

            using (var cn = new OleDbConnection(ConnectionString()))
            {

                cn.Open();

                DataTable dt = cn.GetSchema("Tables", new[] { null, null, null, "Table" });

                foreach (DataRow row in dt.Rows)
                {
                    sheetNames.Add(row.Field<string>("Table_Name"));
                }

            }

            return sheetNames;

        }
        /// <summary>
        /// Get sheet names for excel file
        /// </summary>
        /// <param name="pSheetName"></param>
        /// <returns></returns>
        public bool SheetExists(string pSheetName)
        {
            List<string> names = SheetNames();

            foreach (var sheet in names)
            {
                if (sheet.ToLower() == pSheetName.ToLower())
                {
                    return true;
                }
            }

            return false;

        }
        /// <summary>
        /// Contruct the Excel connection string from properties set in this component
        /// at design time.
        /// </summary>
        /// <returns></returns>
        public string ConnectionString()
        {
            if (ExcelFileExists())
            {
                SetVersions(Provider);
                return $"provider={_provider};data source={FileName};Extended Properties=\"Excel {_oleExelVersion};IMEX={Convert.ToInt32(Mode)};HDR={Headers.ToString()};\"";                
            }
            else
            {
                return "";
            }
        }
    }
    /// <summary>
    /// UITypeEditor used in ExcelProvider class which allows a developer
    /// to be presented with a file dialog to select an Excel file.
    /// </summary>
    /// <remarks>
    /// sender.Multiselect is by default false as we don't want more
    /// than one Excel file to be selected.
    /// </remarks>
    public class ExcelFileNameEditor : FileNameEditor
    {
        protected override void InitializeDialog(OpenFileDialog sender)
        {
            base.InitializeDialog(sender);
            sender.Filter = "Excel 2007|*.xlsx|Excel 2003|*.xls";
            sender.Title = "Select Excel File";
        }
    }
}