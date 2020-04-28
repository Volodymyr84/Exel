using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using Application = Microsoft.Office.Interop.Excel.Application;
using Excel = Microsoft.Office.Interop.Excel;


namespace test
{
    public partial class Form1 : Form
    {
        public Form1()
        {

            InitializeComponent();

            comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;

            System.Windows.Forms.ToolTip ToolTip1 = new System.Windows.Forms.ToolTip();
            ToolTip1.SetToolTip(this.button4, "Show file path in Explorer");
            ToolTip1.SetToolTip(this.button1, "Write data to a file");
            ToolTip1.SetToolTip(this.button2, "Clear all forms");
            ToolTip1.SetToolTip(this.button3, "Close the application");


        }
        
       public const string DEMOFILE = @"C:\Merz daily activities\Merz daily activities.xlsx";
        private void button1_Click(object sender, EventArgs e)
        {

            //const string DEMOFILE = @"C:\Users\Root\source\repos\test\test\Book1.xlsx";
            
            var application = new Microsoft.Office.Interop.Excel.Application();
            application.Visible = true; //показувати ексель;
            try // check if file existing
            {
                application.Workbooks.Open(DEMOFILE);
                var RangWorkbook = application.Workbooks.Item[1].Title;
            }
            catch
            {
                Microsoft.Office.Interop.Excel.Workbook workbook = application.Workbooks.Add();
                Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Worksheets[1];
                workbook.SaveAs(@"C:\Merz daily activities\Merz daily activities.xlsx");
            }
           
            application.DisplayAlerts = false;////Отключить отображение окон с сообщениями true- show


            int Active_sheet = 1;
            int cnt = application.Sheets.Count;
            string[] name = new string[cnt];
            for (int i = 1; i <= cnt; i++)
            {
                name[i - 1] = application.Sheets[i].Name;
            }

            string date = DateTime.UtcNow.ToString("d").Replace("/", ".");
            int count = name.Length;

            for (int i = 0; i < name.Length; i++)
            {
                if (name[i] != date )
                {
                    
                }
                else
                {
                    count--;
                    Active_sheet++;
                }
            }
            if (count == name.Length)
            {
                Worksheet addSheet = application.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                addSheet.Name = date;
                
            }


            //application.Worksheets.Select(N); // Select first WS

            if (name.Length == 1)//Якщо у документі присутня тільки одна сторінка то для Active_sheet встановлюємо її значення
            {
                Active_sheet = 1;
            }
            
            var sheet = application.Worksheets.get_Item(Active_sheet);// Вибераєм лист для запису даних
           
            ///
            ///<ToDO>
            ///Перепровірити на варіанті коли лист розміщений в середині книги
            ///

             Excel.Range range;
                       
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;



            range = sheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;
            int prom = rw;

            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                if (prom - 1 >= 1 || cl > 1)
                {
                    rCnt = rCnt + rw;
                }

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    sheet.Cells[rCnt, cCnt] = textBox1.Text;
                    cCnt++;
                    sheet.Cells[rCnt, cCnt] = comboBox1.Text;
                    cCnt++;
                    sheet.Cells[rCnt, cCnt] = textBox2.Text;
                }
                for (int i = 1; i <= 3; i++)
                    application.Columns[i].AutoFit();


            }
            ///
            /// ///
            ///// Видалення пустих сторінок
            ///
            
            
                                       // int CDS = cnt;// кількість сторінок в книзі
            int O = cnt;                         //перепресвоєння значення для cnt = application.Sheets.Count; після добавлення сторінки
            cnt = application.Sheets.Count;
            int[] delID = new int[cnt];// кількість сторінок в книзі

            name = new string[cnt];
            for (int i = 1; i <= cnt; i++)
            {
                name[i - 1] = application.Sheets[i].Name;
            }
            string[] delsheets = name;//назви сторінок в книзі
            bool sveech = false; // переключатель, якщо знайдена пуста сторінка, то змінюється його стан

            

                for (int i = 0; i < cnt; i++)//for (int i = 0; i < CDS; i++)
            {
                    int x1 = delsheets[i].Length - 1;
                    delsheets[i] = delsheets[i].Remove(x1);// відкидуєм останю цифру в назві листка, перезаписуєм масив новими значенями


                    if (delsheets[i] == "Аркуш" || delsheets[i] == "Sheet" || delsheets[i] == "Лист")
                    {
                        delID[i] = i + 1;
                        sveech = true;
                    }

                }
                                               
                if (sveech == true)
                {
                    for (int i = cnt - 1; i >= 0; i--)//for (int i = CDS - 1; i >= 0; i--)
                {
                        if (delID[i] != 0)
                        {
                            application.Worksheets[i + 1].Delete();
                        }
                    }
                }
                
            ///
            ///
            ///

                CLS();

            Workbook book = application.ActiveWorkbook;
            book.Save();
            //Threading
            Thread.Sleep(500);

            application.Quit();// Close Excel

            CloseProcess();
        }

        public void CloseProcess()
        {
            Process[] List;
            List = Process.GetProcessesByName("EXCEL");
            foreach (Process proc in List)
            {
                proc.Kill();
            }


        }

      
        public void CLS()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            comboBox1.Text = null;
        }

        private void button2_Click(object sender, EventArgs e) => CLS();

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var psi = new ProcessStartInfo();
            psi.FileName = @"C:\Merz daily activities\";
            Process.Start(psi);

        }
    
}
    }

