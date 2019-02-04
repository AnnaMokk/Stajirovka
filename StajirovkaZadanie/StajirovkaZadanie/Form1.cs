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
using Excel = Microsoft.Office.Interop.Excel;
using Access = Microsoft.Office.Interop.Access;

namespace StajirovkaZadanie
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox3.Text != "")
            {
                OleDbConnection connection;
                OleDbDataAdapter adapter;
                DataTable table;
                BindingSource bindingSource;
                InitializeComponent();
                OleDbConnectionStringBuilder stringBuilder = new OleDbConnectionStringBuilder();
                stringBuilder.DataSource = textBox3.Text;
                stringBuilder.Provider = "Microsoft.ACE.OLEDB.12.0";
                stringBuilder.Add("Extended Properties", "Excel 12.0 Xml;HDR=YES");
                connection = new OleDbConnection();
                connection.ConnectionString = stringBuilder.ConnectionString;
                adapter = new OleDbDataAdapter();
                OleDbCommand command = new OleDbCommand("SELECT * FROM [Лист1$]");
                command.Connection = connection;
                adapter.SelectCommand = command;
                table = new DataTable();
                adapter.Fill(table);
                bindingSource = new BindingSource();
                bindingSource.DataSource = table;
                textBox1.DataBindings.Add("Text", bindingSource, "Дата оплаты");
                textBox2.DataBindings.Add("Text", bindingSource, "Сумма");
                textBox4.DataBindings.Add("Text", bindingSource, "Информация");
            }
        }
    }
}
