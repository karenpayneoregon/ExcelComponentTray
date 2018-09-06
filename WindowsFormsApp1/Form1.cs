using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            Shown += Form1_Shown;
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            listBox1.DataSource = excelHelper1.SheetNames();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            
            if (excelHelper1.FileExists && excelHelper1.SheetExists(listBox1.Text))
            {
                var dt = new DataTable();

                using (var cn = new OleDbConnection() { ConnectionString = excelHelper1.ConnectionString() })
                {
                    using (var cmd = new OleDbCommand() { Connection = cn })
                    {
                        cmd.CommandText = $"SELECT * FROM [{listBox1.Text}]";
                        cn.Open();

                        // ReSharper disable once AssignNullToNotNullAttribute
                        dt.Load(cmd.ExecuteReader());
                    }
                }

            }

        }
    }
}
