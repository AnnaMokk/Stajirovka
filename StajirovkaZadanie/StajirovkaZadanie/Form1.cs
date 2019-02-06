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
        int j = 1;
        int i = 0;
        string innn;
        private void button3_Click(object sender, EventArgs e)
        {
            textBox10.Text = "  ";
            textBox8.Text = "  ";
            textBox7.Text = "  ";
            textBox2.Text = "  ";
            int k = 0;
            innn = "";
            string inf = "";
            if (textBox3.Text != "")
            {
                textBox1.Text = " ";
                Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
                Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(@textBox3.Text, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing); //открыть файл

                Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист
                var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку

                Excel.Application ObjWorkExcelNew = new Excel.Application();// создал книгу, можно с ней работать
                Excel.Workbook workBook = ObjWorkExcelNew.Workbooks.Add();
                Excel.Worksheet workSheet = workBook.ActiveSheet;
                string[,] list = new string[3, 10]; // массив значений с листа равен по размеру листу
                for (i = 0; i < 3; i++)
                { list[i, j] = ObjWorkSheet.Cells[j + 1, i + 1].Text.ToString(); }//считываем текст в строку//

                textBox1.Text += "     " + list[0, j];
                textBox8.Text += "     " + list[1, j];
                textBox7.Text += "     " + list[2, j];
                inf += textBox8.Text;
                int len = inf.Length;
                char[] z = inf.ToCharArray();
                for (int n = 0; n < len; n++)
                { if (k != 12)
                    {
                        if ((z[n] >= '0') && (z[n] <= '9'))
                        {
                            k++;
                            innn += z[n];
                        }
                        else
                        {
                            innn = "";
                            k = 0;
                        }
                    }

                }

                if (k == 12) textBox2.Text += innn;
                else textBox2.Text += " mmmmm " + k;
                innn = "";
                k = 0;



                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ObjWorkBook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ObjWorkExcel);

                ObjWorkBook = null;
                ObjWorkExcel = null;
            }
            j++;
         ///////////////////////////работа с ACCESS////////////////////////////////////////
         
            object ACName = "";
            object ACINN = "";
            object ACBOOL = "";
            object ACID = "";

           
                for (int i = 1; i <= 4; i++) /////Цикл проверки////////
                {

                    string ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Users\\Пользователь\\Documents\\GitHub\\Stajirovka\\база.mdb";
                    OleDbConnection con = new OleDbConnection(ConnectionString);
                    con.Open();

                    OleDbCommand outt = new OleDbCommand("SELECT [id],[name],[inn],[active] FROM users  Where [id]="+i, con);
                    OleDbDataReader rdr = outt.ExecuteReader();

                    while (rdr.Read())
                    {
                        ACName = rdr["name"];
                        ACINN = rdr["inn"];
                        ACBOOL = rdr["active"];
                    }

                    rdr.Close();
                    con.Close();

                    string GoalID = ACID.ToString();
                    string GoalN = ACName.ToString();
                    string GoalI = ACINN.ToString();
                    string GoalB = ACBOOL.ToString();
                    

                    if (GoalB == "Да")
                        {
                        textBox10.Text = "Да";
                            OleDbConnection con1 = new OleDbConnection(ConnectionString);
                            con1.Open();
                            string comandString = "INSERT INTO payment ([id_user],[sum],[date])VALUES('"+GoalID+"', '"+textBox7.Text.ToString()+"', '"+textBox1.Text.ToString()+"')";
                            OleDbCommand addID = new OleDbCommand(comandString, con1);
                            addID.ExecuteNonQuery();
                            con1.Close();
                        }
                        else
                        {
                        }
                        break;
                    }

                
            }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }


    }



//C:\Users\Пользователь\Documents\GitHub\Stajirovka\оплаты.xlsx   C:\Users\Пользователь\Documents\GitHub\Stajirovka\база.mdb//
