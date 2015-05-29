using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SQLite;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace GecikmeHesapla
{
    public partial class Form1 : Form
    {
        int[] dizi = new int[558];
        int[] dizi2 = new int[558];
        int[] dizi3 = new int[558];
        int[] dizi4 = new int[558];
        int[] dizi5 = new int[558];
        int[] dizi6 = new int[558];
        int[] dizi7 = new int[558];
        int[] dizi8 = new int[558];
        String file, time;
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Excel.Range range;
        public Form1()
        {
            InitializeComponent();
        }

        public void GecikmeHesapla()
        {
            
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            SQLiteConnection con = new SQLiteConnection("Data Source=db.db3;Version=3;Read Only=False;");
            SQLiteCommand q = new SQLiteCommand("select bodyno from bodyno", con);
            con.Open();
            SQLiteDataReader r = q.ExecuteReader();
            int i = 0;
            int body;
            while (r.Read())
            {
                dizi[i] = Convert.ToInt32(r["bodyno"]);
                dizi2[i] = i + 1;                
                i++;
            }
            r.Close();

            SQLiteCommand q2 = new SQLiteCommand("select * from bodyno", con);
            SQLiteDataReader r2 = q2.ExecuteReader();
            i = 0;
            while (r2.Read())
            {
                int b9 = Convert.ToInt32(r2["b9"]);
                body = Array.IndexOf(dizi, b9);
                dizi3[body] = i + 1;

                int u0 = Convert.ToInt32(r2["u0"]);
                body = Array.IndexOf(dizi, u0);
                dizi4[body] = i + 1;

                int u6 = Convert.ToInt32(r2["u6"]);
                body = Array.IndexOf(dizi, u6);
                dizi5[body] = i + 1;

                int t3 = Convert.ToInt32(r2["t3"]);
                body = Array.IndexOf(dizi, t3);
                dizi6[body] = i + 1;

                int p3 = Convert.ToInt32(r2["p3"]);
                body = Array.IndexOf(dizi, p3);
                dizi7[body] = i + 1;

                int p9 = Convert.ToInt32(r2["p9"]);
                body = Array.IndexOf(dizi, p9);
                dizi8[body] = i + 1;

                i++;
            }
            r2.Close();
            
            /*SQLiteCommand q2 = new SQLiteCommand("select b9 from bodyno", con);
            SQLiteDataReader r2 = q2.ExecuteReader();
            i = 0;
            while (r2.Read())
            {
                int b9 = Convert.ToInt32(r2["b9"]);
                int body = Array.IndexOf(dizi,b9);
                dizi[body, 2] = i + 1;
                i++;
            }
            r2.Close();*/
            
            con.Close();
            MessageBox.Show("Success!");
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void button2_Click(object sender, System.EventArgs e)
        {
            FileStream fs = new FileStream(@"C:\\Users\\Caner\\Desktop\\body sequence retention\\sonuc.xls", FileMode.Open, FileAccess.ReadWrite);
            HSSFWorkbook templateWorkbook = new HSSFWorkbook(fs);
            HSSFSheet sheet = (HSSFSheet)templateWorkbook.GetSheet("Sheet1");
            fs.Close();
            
            int row=0;
            SQLiteConnection con = new SQLiteConnection("Data Source=db.db3;Version=3;Read Only=False;");
            con.Open();
            try
            {
                SQLiteCommand q4 = new SQLiteCommand("select * from bodyno ", con);
                SQLiteDataReader r3 = q4.ExecuteReader();
                while(r3.Read())
                {
                    var rcell = sheet.CreateRow(row);
                    rcell.CreateCell(0).SetCellValue(Convert.ToInt32(r3["b0"]));
                    rcell.CreateCell(1).SetCellValue(Convert.ToInt32(r3["b9"]));
                    rcell.CreateCell(2).SetCellValue(Convert.ToInt32(r3["u0"]));
                    rcell.CreateCell(3).SetCellValue(Convert.ToInt32(r3["u6"]));
                    rcell.CreateCell(4).SetCellValue(Convert.ToInt32(r3["t3"]));
                    rcell.CreateCell(5).SetCellValue(Convert.ToInt32(r3["p3"]));
                    rcell.CreateCell(6).SetCellValue(Convert.ToInt32(r3["p9"]));
                    row++;
                }
                r3.Close();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                sheet.ForceFormulaRecalculation = true;
                fs = new FileStream(@"C:\\Users\\Caner\\Desktop\\body sequence retention\\sonuc.xls", FileMode.Open, FileAccess.ReadWrite);
                templateWorkbook.Write(fs);
                fs.Close();
                MessageBox.Show("Success!");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // Kontrolü göster
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                file = openFileDialog1.FileName;

                SQLiteConnection con3 = new SQLiteConnection("Data Source=db.db3;Version=3;Read Only=False;");
                con3.Open();
                SQLiteCommand com3 = new SQLiteCommand("select bodyno from bodyno", con3);
                SQLiteDataReader et = com3.ExecuteReader();
                if (et != null)
                {
                    SQLiteCommand q5 = new SQLiteCommand("delete from sqlite_sequence where name = 'bodyno'", con3);
                    q5.ExecuteNonQuery();
                    SQLiteCommand comd = new SQLiteCommand("delete from bodyno", con3);
                    comd.ExecuteNonQuery();
                }
                et.Close();

                FileStream fs = new FileStream(file, FileMode.Open, FileAccess.ReadWrite);
                XSSFWorkbook templateWorkbook = new XSSFWorkbook(fs);
                XSSFSheet sheet = (XSSFSheet)templateWorkbook.GetSheet("SIRA");
                sheet.ForceFormulaRecalculation = false;
                for (int rCnt = 1; rCnt < 559; rCnt++)
                {
                    SQLiteCommand comi = new SQLiteCommand("insert into bodyno(bodyno,b9,u0,u6,t3,p3) values(@bodyno,@b9,@u0,@u6,@t3,@p3)", con3);
                    comi.Parameters.AddWithValue("@bodyno", sheet.GetRow(rCnt).GetCell(1).NumericCellValue);
                    comi.Parameters.AddWithValue("@b9", sheet.GetRow(rCnt).GetCell(12).NumericCellValue);
                    comi.Parameters.AddWithValue("@u0", sheet.GetRow(rCnt).GetCell(15).NumericCellValue);
                    comi.Parameters.AddWithValue("@u6", sheet.GetRow(rCnt).GetCell(18).NumericCellValue);
                    comi.Parameters.AddWithValue("@t3", sheet.GetRow(rCnt).GetCell(21).NumericCellValue);
                    comi.Parameters.AddWithValue("@p3", sheet.GetRow(rCnt).GetCell(24).NumericCellValue);
                    comi.ExecuteNonQuery();
                }
                con3.Close();
                fs.Close();
                MessageBox.Show("Success!");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog2.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                String ed = openFileDialog2.FileName;
                

                SQLiteConnection con3 = new SQLiteConnection("Data Source=db.db3;Version=3;Read Only=False;");
                con3.Open();
                SQLiteCommand com3 = new SQLiteCommand("select saat from edminutes", con3);
                SQLiteDataReader et = com3.ExecuteReader();
                if (et != null)
                {
                    SQLiteCommand comd = new SQLiteCommand("delete from edminutes", con3);
                    comd.ExecuteNonQuery();
                }
                et.Close();
                try
                {
                    FileStream fs = new FileStream(ed, FileMode.Open, FileAccess.ReadWrite);
                        XSSFWorkbook templateWorkbook = new XSSFWorkbook(fs);
                        XSSFSheet sheet = (XSSFSheet)templateWorkbook.GetSheet("Sheet1");
                        sheet.ForceFormulaRecalculation = false;
                        for (int rCnt = 1; rCnt < 1337; rCnt++)
                        {
                            SQLiteCommand comi = new SQLiteCommand("insert into edminutes(saat,adet) values(@saat,@adet)", con3);
                            DateTime timeTemp = sheet.GetRow(rCnt).GetCell(0).DateCellValue;
                            String saat = timeTemp.ToString("HH:mm");
                            comi.Parameters.AddWithValue("@saat", saat);
                            comi.Parameters.AddWithValue("@adet", sheet.GetRow(rCnt).GetCell(1).NumericCellValue);
                            comi.ExecuteNonQuery();
                        }
                        con3.Close();
                        fs.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                MessageBox.Show("Success!");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SQLiteConnection con2 = new SQLiteConnection("Data Source=db.db3;Version=3;Read Only=False;");
            con2.Open();
            SQLiteCommand q5 = new SQLiteCommand("delete from sqlite_sequence where name = 'lastbodyno'", con2);
            q5.ExecuteNonQuery();
            SQLiteCommand q6 = new SQLiteCommand("delete from lastbodyno", con2);
            q6.ExecuteNonQuery();
            con2.Close();
            int bypassMin = -1000;//(int)numericUpDown1.Value;
            int bypassMax = -30;//(int)numericUpDown2.Value;
            int line1Min = -30;//(int)numericUpDown3.Value;
            int line1Max = -10;//(int)numericUpDown4.Value;
            int taktIn = 98;//Convert.ToInt32(textBox1.Text);
            int taktOut = 123;//Convert.ToInt32(textBox2.Text);
            int adet, lineSira = 0, lineSiraOut = 0, bypassAdet = 0, lin1Adet = 0, lin2Adet = 0, lin3Adet = 0, lin4Adet = 0, lin5Adet = 0, lin6Adet = 0, lin7Adet = 0;
            List<int> arrayBody = new List<int>();
            List<int> arrayB9 = new List<int>();
            DateTime date = new DateTime(11,11,11,9,0,0);
            SQLiteConnection con = new SQLiteConnection("Data Source=db.db3;Version=3;Read Only=False;");
            SQLiteCommand q = new SQLiteCommand("select bodyno,b9 from bodyno", con);
            con.Open();
            SQLiteDataReader r = q.ExecuteReader();
            int bodyno, b9;
            for (int sec = 1; sec <= 72000; sec++)
            {
                if (sec % taktIn == 0) {
                    r.Read();
                    if (!r.HasRows) {
                        MessageBox.Show("Araç bitti saniye: " + sec);
                        break;
                    }
                    bodyno = Convert.ToInt32(r["bodyno"]);
                    b9 = Convert.ToInt32(r["b9"]);
                    
                    if (bypassMin < b9 && b9 < bypassMax) {
                        bypass.Items.Add(bodyno);
                        if (arrayB9.Count != 0) {
                            for (int i = 0; i < arrayB9.Count; i++) {
                                arrayB9[i] = arrayB9[i] - 1;
                            }
                        }
                        log();
                        bypassAdet++;
                    }
                    else if (line1Min <= b9 && b9 <= line1Max)
                    {
                        line1.Items.Add(bodyno);
                            arrayBody.Add(bodyno);
                            arrayB9.Add(b9);
                            log();
                            lin1Adet++;
                    }
                    else {
                        if (lineSira == 0)
                        {
                            line2.Items.Add(bodyno);
                            arrayBody.Add(bodyno);
                            arrayB9.Add(b9);
                            lin2Adet++;
                            log();
                            lineSira++;
                        }
                        else if (lineSira == 1)
                        {
                            line3.Items.Add(bodyno);
                            arrayBody.Add(bodyno);
                            arrayB9.Add(b9);
                            lin3Adet++;
                            log();
                            lineSira++;
                        }
                        else if (lineSira == 2){
                            line4.Items.Add(bodyno);
                            arrayBody.Add(bodyno);
                            arrayB9.Add(b9);
                            lin4Adet++;
                            log();
                            lineSira++;
                        }
                        else if (lineSira == 3)
                        {
                            line5.Items.Add(bodyno);
                            arrayBody.Add(bodyno);
                            arrayB9.Add(b9);
                            lin5Adet++;
                            log();
                            lineSira++;
                        }
                        else if (lineSira == 4)
                        {
                            line6.Items.Add(bodyno);
                            arrayBody.Add(bodyno);
                            arrayB9.Add(b9);
                            lin6Adet++;
                            log();
                            lineSira++;
                        }
                        else if (lineSira == 5)
                        {
                            line7.Items.Add(bodyno);
                            arrayBody.Add(bodyno);
                            arrayB9.Add(b9);
                            lin7Adet++;
                            log();
                            lineSira = 0;
                        }
                    }
                }
                if (sec % taktOut == 0)
                {
                    date = date.AddSeconds(taktOut);
                    time = date.ToString("HH:mm");
                    SQLiteCommand q2 = new SQLiteCommand("select adet from edminutes where saat ='" + time + "' ", con);
                    SQLiteDataReader r2 = q2.ExecuteReader();
                    r2.Read();
                    adet = Convert.ToInt32(r2["adet"]);
                    r2.Close();
                    
                    if (bypass.Items.Count > 0)
                    {
                        bodyno = Convert.ToInt32(bypass.Items[0]);
                        SQLiteCommand q4 = new SQLiteCommand("select b9 from bodyno where bodyno ='" + bodyno + "' ", con);
                        SQLiteDataReader r3 = q4.ExecuteReader();
                        r3.Read();
                        b9 = Convert.ToInt32(r3["b9"]);
                        r3.Close();
                        bypass.Items.RemoveAt(0);
                        int oldb9 = b9;
                        b9 = adet + b9;
                        SQLiteCommand q3 = new SQLiteCommand("Insert Into lastbodyno(bodyno,b9,nb9,saat) values(@body,@ob9,@b9,@saat)", con);
                        q3.Parameters.AddWithValue("@body", bodyno);
                        q3.Parameters.AddWithValue("@ob9", oldb9);
                        q3.Parameters.AddWithValue("@b9", b9);
                        q3.Parameters.AddWithValue("@saat", time);
                        q3.ExecuteNonQuery();
                        log();
                    }
                    else if (line1.Items.Count > 0)
                    {
                        bodyno = Convert.ToInt32(line1.Items[0]);
                        b9 = arrayB9[arrayBody.IndexOf(bodyno)];
                        int oldb9 = b9;
                        b9 = adet + b9;
                        arrayB9[arrayBody.IndexOf(bodyno)] = b9;
                        line1.Items.RemoveAt(0);
                        SQLiteCommand q3 = new SQLiteCommand("Insert Into lastbodyno(bodyno,b9,nb9,saat) values(@body,@ob9,@b9,@saat)", con);
                        q3.Parameters.AddWithValue("@body", bodyno);
                        q3.Parameters.AddWithValue("@ob9", oldb9);
                        q3.Parameters.AddWithValue("@b9", b9);
                        q3.Parameters.AddWithValue("@saat", time);
                        q3.ExecuteNonQuery();
                        arrayBody.Remove(bodyno);
                        arrayB9.Remove(b9);
                        if (arrayB9.Count != 0) {
                            for (int i = 0; i < arrayB9.Count; i++) {
                                arrayB9[i] = arrayB9[i] - 1;
                            }
                        }
                        log();
                    }
                    else
                    {
                        if (line2.Items.Count > 0 && lineSiraOut == 0)
                        {
                            bodyno = Convert.ToInt32(line2.Items[0]);
                            line2.Items.RemoveAt(0);
                            b9 = arrayB9[arrayBody.IndexOf(bodyno)];
                            int oldb9 = b9;
                            SQLiteCommand q3 = new SQLiteCommand("Insert Into lastbodyno(bodyno,b9,nb9,saat) values(@body,@ob9,@b9,@saat)", con);
                            q3.Parameters.AddWithValue("@body", bodyno);
                            q3.Parameters.AddWithValue("@ob9", oldb9);
                            q3.Parameters.AddWithValue("@b9", b9);
                            q3.Parameters.AddWithValue("@saat", time);
                            q3.ExecuteNonQuery();
                            arrayBody.Remove(bodyno);
                            arrayB9.Remove(b9);
                            lineSiraOut++;
                            log();
                        }
                        else if (line3.Items.Count > 0 && lineSiraOut == 1)
                        {
                            bodyno = Convert.ToInt32(line3.Items[0]);
                            line3.Items.RemoveAt(0);
                            b9 = arrayB9[arrayBody.IndexOf(bodyno)];
                            int oldb9 = b9;
                            SQLiteCommand q3 = new SQLiteCommand("Insert Into lastbodyno(bodyno,b9,nb9,saat) values(@body,@ob9,@b9,@saat)", con);
                            q3.Parameters.AddWithValue("@body", bodyno);
                            q3.Parameters.AddWithValue("@ob9", oldb9);
                            q3.Parameters.AddWithValue("@b9", b9);
                            q3.Parameters.AddWithValue("@saat", time);
                            q3.ExecuteNonQuery();
                            arrayBody.Remove(bodyno);
                            arrayB9.Remove(b9);
                            lineSiraOut++;
                            log();
                        }
                        else if (line4.Items.Count > 0 && lineSiraOut == 2)
                        {
                            bodyno = Convert.ToInt32(line4.Items[0]);
                            line4.Items.RemoveAt(0);
                            b9 = arrayB9[arrayBody.IndexOf(bodyno)];
                            int oldb9 = b9;
                            SQLiteCommand q3 = new SQLiteCommand("Insert Into lastbodyno(bodyno,b9,nb9,saat) values(@body,@ob9,@b9,@saat)", con);
                            q3.Parameters.AddWithValue("@body", bodyno);
                            q3.Parameters.AddWithValue("@ob9", oldb9);
                            q3.Parameters.AddWithValue("@b9", b9);
                            q3.Parameters.AddWithValue("@saat", time);
                            q3.ExecuteNonQuery();
                            arrayBody.Remove(bodyno);
                            arrayB9.Remove(b9);
                            lineSiraOut++;
                            log();
                        }
                        else if (line5.Items.Count > 0 && lineSiraOut == 3)
                        {
                            bodyno = Convert.ToInt32(line5.Items[0]);
                            line5.Items.RemoveAt(0);
                            b9 = arrayB9[arrayBody.IndexOf(bodyno)];
                            int oldb9 = b9;
                            SQLiteCommand q3 = new SQLiteCommand("Insert Into lastbodyno(bodyno,b9,nb9,saat) values(@body,@ob9,@b9,@saat)", con);
                            q3.Parameters.AddWithValue("@body", bodyno);
                            q3.Parameters.AddWithValue("@ob9", oldb9);
                            q3.Parameters.AddWithValue("@b9", b9);
                            q3.Parameters.AddWithValue("@saat", time);
                            q3.ExecuteNonQuery();
                            arrayBody.Remove(bodyno);
                            arrayB9.Remove(b9);
                            lineSiraOut++;
                            log();
                        }
                        else if (line6.Items.Count > 0 && lineSiraOut == 4)
                        {
                            bodyno = Convert.ToInt32(line6.Items[0]);
                            line6.Items.RemoveAt(0);
                            b9 = arrayB9[arrayBody.IndexOf(bodyno)];
                            int oldb9 = b9;
                            SQLiteCommand q3 = new SQLiteCommand("Insert Into lastbodyno(bodyno,b9,nb9,saat) values(@body,@ob9,@b9,@saat)", con);
                            q3.Parameters.AddWithValue("@body", bodyno);
                            q3.Parameters.AddWithValue("@ob9", oldb9);
                            q3.Parameters.AddWithValue("@b9", b9);
                            q3.Parameters.AddWithValue("@saat", time);
                            q3.ExecuteNonQuery();
                            arrayBody.Remove(bodyno);
                            arrayB9.Remove(b9);
                            lineSiraOut++;
                            log();
                        }
                        else if (line7.Items.Count > 0 && lineSiraOut == 5)
                        {
                            bodyno = Convert.ToInt32(line7.Items[0]);
                            line7.Items.RemoveAt(0);
                            b9 = arrayB9[arrayBody.IndexOf(bodyno)];
                            int oldb9 = b9;
                            SQLiteCommand q3 = new SQLiteCommand("Insert Into lastbodyno(bodyno,b9,nb9,saat) values(@body,@ob9,@b9,@saat)", con);
                            q3.Parameters.AddWithValue("@body", bodyno);
                            q3.Parameters.AddWithValue("@ob9", oldb9);
                            q3.Parameters.AddWithValue("@b9", b9);
                            q3.Parameters.AddWithValue("@saat", time);
                            q3.ExecuteNonQuery();
                            arrayBody.Remove(bodyno);
                            arrayB9.Remove(b9);
                            lineSiraOut=0;
                            log();
                        }
                    }
                }
            }
            label3.Text = bypassAdet.ToString();
            label4.Text = lin1Adet.ToString();
            label5.Text = lin2Adet.ToString();
            label6.Text = lin3Adet.ToString();
            label7.Text = lin4Adet.ToString();
            label8.Text = lin5Adet.ToString();
            label9.Text = lin6Adet.ToString();
            label10.Text = lin7Adet.ToString();
            r.Close();
            con.Close();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            FileStream fs = new FileStream(@"C:\\Users\\Caner\\Desktop\\body sequence retention\\sonuc2.xlsx", FileMode.Open, FileAccess.ReadWrite);
            XSSFWorkbook templateWorkbook = new XSSFWorkbook(fs);
            XSSFSheet sheet = (XSSFSheet)templateWorkbook.GetSheet("Sheet1");
            fs.Close();

            List<int> b9s = new List<int>();
            int row = 0;
            SQLiteConnection con = new SQLiteConnection("Data Source=db.db3;Version=3;Read Only=False;");
            con.Open();
            try
            {
                SQLiteCommand q4 = new SQLiteCommand("select * from lastbodyno", con);
                SQLiteDataReader r3 = q4.ExecuteReader();
                while(r3.Read())
                {
                    var rcell = sheet.CreateRow(row);
                    rcell.CreateCell(0).SetCellValue(Convert.ToInt32(r3["bodyno"]));
                    rcell.CreateCell(1).SetCellValue(Convert.ToInt32(r3["nb9"]));
                    rcell.CreateCell(2).SetCellValue(r3["saat"].ToString());
                    row++;
                }
                r3.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                sheet.ForceFormulaRecalculation = true;
                fs = new FileStream(@"C:\\Users\\Caner\\Desktop\\body sequence retention\\sonuc2.xlsx", FileMode.Open, FileAccess.ReadWrite);
                templateWorkbook.Write(fs);
                fs.Close();
                MessageBox.Show("Success!");
            }
            con.Close();
        }
        public void log()
        {
            Console.WriteLine("time  : " + time);
            Console.WriteLine("bypass: " + bypass.Items.Count);
            Console.WriteLine("line1 : " + line1.Items.Count);
            Console.WriteLine("line2 : " + line2.Items.Count);
            Console.WriteLine("line3 : " + line3.Items.Count);
            Console.WriteLine("line4 : " + line4.Items.Count);
            Console.WriteLine("line5 : " + line5.Items.Count);
            Console.WriteLine("line6 : " + line6.Items.Count);
            Console.WriteLine("line7 : " + line7.Items.Count);
            Console.WriteLine("--------------------------------");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog2.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                String ed = openFileDialog2.FileName;


                SQLiteConnection con3 = new SQLiteConnection("Data Source=db.db3;Version=3;Read Only=False;");
                con3.Open();
                SQLiteCommand com3 = new SQLiteCommand("select bodyno from u0u6", con3);
                SQLiteDataReader et = com3.ExecuteReader();
                if (et != null)
                {
                    SQLiteCommand comd2 = new SQLiteCommand("delete from sirau6", con3);
                    comd2.ExecuteNonQuery();
                    SQLiteCommand comd = new SQLiteCommand("delete from u0u6", con3);
                    comd.ExecuteNonQuery();
                }
                et.Close();
                try
                {
                    FileStream fs = new FileStream(ed, FileMode.Open, FileAccess.ReadWrite);
                    XSSFWorkbook templateWorkbook = new XSSFWorkbook(fs);
                    XSSFSheet sheet = (XSSFSheet)templateWorkbook.GetSheet("BİRLEŞTİRİLMİŞ");
                    XSSFSheet sheet2 = (XSSFSheet)templateWorkbook.GetSheet("SIRA");
                    sheet.ForceFormulaRecalculation = false;
                    sheet2.ForceFormulaRecalculation = false;
                    for (int rCnt = 2; rCnt < 560; rCnt++)
                    {
                        SQLiteCommand comi2 = new SQLiteCommand("insert into sirau6(bodyno,u6) values(@bodyno,@u6)", con3);
                        comi2.Parameters.AddWithValue("@bodyno", sheet2.GetRow(rCnt-1).GetCell(1).NumericCellValue);
                        comi2.Parameters.AddWithValue("@u6", sheet2.GetRow(rCnt-1).GetCell(33).NumericCellValue);
                        comi2.ExecuteNonQuery();
                        SQLiteCommand comi = new SQLiteCommand("insert into u0u6(bodyno) values(@bodyno)", con3);
                        comi.Parameters.AddWithValue("@bodyno", sheet.GetRow(rCnt).GetCell(7).NumericCellValue);
                        comi.ExecuteNonQuery();
                    }
                    SQLiteCommand q = new SQLiteCommand("select bodyno from u0u6", con3);
                    SQLiteDataReader r = q.ExecuteReader();
                    while (r.Read())
                    {
                        SQLiteCommand comi = new SQLiteCommand("update u0u6 set u6 = (select u6 from sirau6 where bodyno = '" + Convert.ToInt32(r["bodyno"]) + "') where bodyno = '" + Convert.ToInt32(r["bodyno"]) + "'", con3);
                        comi.ExecuteNonQuery();
                    }
                    r.Close();

                    con3.Close();
                    fs.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                MessageBox.Show("Success!");
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            SQLiteConnection con = new SQLiteConnection("Data Source=db.db3;Version=3;Read Only=False;");
            SQLiteCommand q = new SQLiteCommand("select bodyno from u0u6", con);
            con.Open();
            SQLiteDataReader r = q.ExecuteReader();
            int i = 0;
            int body;
            while (r.Read())
            {
                dizi[i] = Convert.ToInt32(r["bodyno"]);
                i++;
            }
            r.Close();

            SQLiteCommand q2 = new SQLiteCommand("select u6 from u0u6", con);
            SQLiteDataReader r2 = q2.ExecuteReader();
            i = 0;
            while (r2.Read())
            {
                int u6 = Convert.ToInt32(r2["u6"]);
                body = Array.IndexOf(dizi, u6);
                if (body >= 0)
                {
                    dizi2[body] = body - i;
                    i++;
                }
            }
            r2.Close();
            SQLiteCommand com3 = new SQLiteCommand("select * from u0u6", con);
            SQLiteDataReader et = com3.ExecuteReader();
            if (et != null)
            {
                SQLiteCommand comd = new SQLiteCommand("delete from u0u6", con);
                comd.ExecuteNonQuery();
            }
            et.Close();
            for (int j = 0; j < dizi.Length; j++)
            {
                SQLiteCommand comi = new SQLiteCommand("insert into u0u6(bodyno,u6) values(@bodyno,@u6)", con);
                comi.Parameters.AddWithValue("@bodyno", dizi[j]);
                comi.Parameters.AddWithValue("@u6", dizi2[j]);
                comi.ExecuteNonQuery();
            }
            con.Close();
            MessageBox.Show("Success!");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            FileStream fs = new FileStream(@"C:\\Users\\Caner\\Desktop\\body sequence retention\\u0-u6.xlsx", FileMode.Open, FileAccess.ReadWrite);
            XSSFWorkbook templateWorkbook = new XSSFWorkbook(fs);
            XSSFSheet sheet = (XSSFSheet)templateWorkbook.GetSheet("Sheet1");
            fs.Close();

            SQLiteConnection con = new SQLiteConnection("Data Source=db.db3;Version=3;Read Only=False;");
            con.Open();
            try
            {
                SQLiteCommand q4 = new SQLiteCommand("select * from u0u6 ", con);
                SQLiteDataReader r3 = q4.ExecuteReader();
                int row = 0;
                while(r3.Read())
                {
                    var rcell = sheet.CreateRow(row);
                    rcell.CreateCell(0).SetCellValue(Convert.ToInt32(r3["bodyno"]));
                    rcell.CreateCell(1).SetCellValue(Convert.ToInt32(r3["u6"]));
                    row++;
                }
                r3.Close();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                sheet.ForceFormulaRecalculation = true;
                fs = new FileStream(@"C:\\Users\\Caner\\Desktop\\body sequence retention\\u0-u6.xlsx", FileMode.Open, FileAccess.ReadWrite);
                templateWorkbook.Write(fs);
                fs.Close();
                MessageBox.Show("Success!");
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog2.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                String ed = openFileDialog2.FileName;


                SQLiteConnection con3 = new SQLiteConnection("Data Source=db.db3;Version=3;Read Only=False;");
                con3.Open();
                SQLiteCommand com3 = new SQLiteCommand("select saat from prminutes", con3);
                SQLiteDataReader et = com3.ExecuteReader();
                if (et != null)
                {
                    SQLiteCommand comd = new SQLiteCommand("delete from prminutes", con3);
                    comd.ExecuteNonQuery();
                }
                et.Close();
                try
                {
                    FileStream fs = new FileStream(ed, FileMode.Open, FileAccess.ReadWrite);
                    XSSFWorkbook templateWorkbook = new XSSFWorkbook(fs);
                    XSSFSheet sheet = (XSSFSheet)templateWorkbook.GetSheet("Sayfa1");
                    sheet.ForceFormulaRecalculation = false;
                    for (int rCnt = 1; rCnt < 1442; rCnt++)
                    {
                        SQLiteCommand comi = new SQLiteCommand("insert into prminutes(saat,adet) values(@saat,@adet)", con3);
                        DateTime timeTemp = sheet.GetRow(rCnt).GetCell(0).DateCellValue;
                        String saat = timeTemp.ToString("HH:mm");
                        comi.Parameters.AddWithValue("@saat", saat);
                        comi.Parameters.AddWithValue("@adet", sheet.GetRow(rCnt).GetCell(1).NumericCellValue);
                        comi.ExecuteNonQuery();
                    }
                    con3.Close();
                    fs.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                MessageBox.Show("Success!");
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            SQLiteConnection con2 = new SQLiteConnection("Data Source=db.db3;Version=3;Read Only=False;");
            con2.Open();
            SQLiteCommand q5 = new SQLiteCommand("delete from sqlite_sequence where name = 'plastbodyno'", con2);
            q5.ExecuteNonQuery();
            SQLiteCommand q6 = new SQLiteCommand("delete from plastbodyno", con2);
            q6.ExecuteNonQuery();
            con2.Close();
            int line1Min = -1000;//(int)numericUpDown3.Value;
            int line1Max = -30;//(int)numericUpDown4.Value;
            int taktIn = 98;//Convert.ToInt32(textBox1.Text);
            int taktOut = 123;//Convert.ToInt32(textBox2.Text);
            int adet, lineSira = 0, lineSiraOut = 0, lin1Adet = 0, lin2Adet = 0, lin3Adet = 0;
            List<int> arrayBody = new List<int>();
            List<int> arrayU6 = new List<int>();
            DateTime date = new DateTime(11, 11, 11, 8, 0, 0);
            SQLiteConnection con = new SQLiteConnection("Data Source=db.db3;Version=3;Read Only=False;");
            SQLiteCommand q = new SQLiteCommand("select bodyno,u6 from u0u6", con);
            con.Open();
            SQLiteDataReader r = q.ExecuteReader();
            int bodyno, u6;
            for (int sec = 1; sec <= 82800; sec++)
            {
                if (sec % taktIn == 0)
                {
                    r.Read();
                    if (!r.HasRows)
                    {
                        MessageBox.Show("Araç bitti saniye: " + sec);
                        break;
                    }
                    bodyno = Convert.ToInt32(r["bodyno"]);
                    u6 = Convert.ToInt32(r["u6"]);

                    if (line1Min < u6 && u6 < line1Max)
                    {
                        line1.Items.Add(bodyno);
                        if (arrayU6.Count != 0)
                        {
                            for (int i = 0; i < arrayU6.Count; i++)
                            {
                                arrayU6[i] = arrayU6[i] - 1;
                            }
                        }
                        arrayBody.Add(bodyno);
                        arrayU6.Add(u6);
                        log();
                        lin1Adet++;
                    }
                    else
                    {
                        if (lineSira == 0)
                        {
                            line2.Items.Add(bodyno);
                            arrayBody.Add(bodyno);
                            arrayU6.Add(u6);
                            lin2Adet++;
                            log();
                            lineSira++;
                        }
                        else if (lineSira == 1)
                        {
                            line3.Items.Add(bodyno);
                            arrayBody.Add(bodyno);
                            arrayU6.Add(u6);
                            lin3Adet++;
                            log();
                            lineSira=0;
                        }
                    }
                }
                if (sec % taktOut == 0)
                {
                    date = date.AddSeconds(taktOut);
                    time = date.ToString("HH:mm");
                    SQLiteCommand q2 = new SQLiteCommand("select adet from prminutes where saat ='" + time + "' ", con);
                    SQLiteDataReader r2 = q2.ExecuteReader();
                    r2.Read();
                    adet = Convert.ToInt32(r2["adet"]);
                    r2.Close();
                    
                    if (line1.Items.Count > 0)
                    {
                        bodyno = Convert.ToInt32(line1.Items[0]);
                        u6 = arrayU6[arrayBody.IndexOf(bodyno)];
                        int oldb9 = u6;
                        u6 = adet + u6;
                        arrayU6[arrayBody.IndexOf(bodyno)] = u6;
                        line1.Items.RemoveAt(0);
                        SQLiteCommand q3 = new SQLiteCommand("Insert Into plastbodyno(bodyno,u6,saat) values(@body,@u6,@saat)", con);
                        q3.Parameters.AddWithValue("@body", bodyno);
                        q3.Parameters.AddWithValue("@u6", u6);
                        q3.Parameters.AddWithValue("@saat", time);
                        q3.ExecuteNonQuery();
                        arrayBody.Remove(bodyno);
                        arrayU6.Remove(u6);
                        log();
                    }
                    else
                    {
                        if (line2.Items.Count > 0 && lineSiraOut == 0)
                        {
                            bodyno = Convert.ToInt32(line2.Items[0]);
                            line2.Items.RemoveAt(0);
                            u6 = arrayU6[arrayBody.IndexOf(bodyno)];
                            int oldb9 = u6;
                            SQLiteCommand q3 = new SQLiteCommand("Insert Into plastbodyno(bodyno,u6,saat) values(@body,@u6,@saat)", con);
                            q3.Parameters.AddWithValue("@body", bodyno);
                            q3.Parameters.AddWithValue("@u6", u6);
                            q3.Parameters.AddWithValue("@saat", time);
                            q3.ExecuteNonQuery();
                            arrayBody.Remove(bodyno);
                            arrayU6.Remove(u6);
                            lineSiraOut++;
                            log();
                        }
                        else if (line3.Items.Count > 0 && lineSiraOut == 1)
                        {
                            bodyno = Convert.ToInt32(line3.Items[0]);
                            line3.Items.RemoveAt(0);
                            u6 = arrayU6[arrayBody.IndexOf(bodyno)];
                            int oldb9 = u6;
                            SQLiteCommand q3 = new SQLiteCommand("Insert Into plastbodyno(bodyno,u6,saat) values(@body,@u6,@saat)", con);
                            q3.Parameters.AddWithValue("@body", bodyno);
                            q3.Parameters.AddWithValue("@u6", u6);
                            q3.Parameters.AddWithValue("@saat", time);
                            q3.ExecuteNonQuery();
                            arrayBody.Remove(bodyno);
                            arrayU6.Remove(u6);
                            lineSiraOut=0;
                            log();
                        }
                    }
                }
            }
            label4.Text = lin1Adet.ToString();
            label5.Text = lin2Adet.ToString();
            label6.Text = lin3Adet.ToString();
            r.Close();
            con.Close();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            FileStream fs = new FileStream(@"C:\\Users\\Caner\\Desktop\\body sequence retention\\sonuc3.xlsx", FileMode.Open, FileAccess.ReadWrite);
            XSSFWorkbook templateWorkbook = new XSSFWorkbook(fs);
            XSSFSheet sheet = (XSSFSheet)templateWorkbook.GetSheet("Sheet1");
            fs.Close();

            List<int> u6s = new List<int>();
            int row = 0, nu6=0;
            SQLiteConnection con = new SQLiteConnection("Data Source=db.db3;Version=3;Read Only=False;");
            con.Open();
            try
            {
                SQLiteCommand q4 = new SQLiteCommand("select * from plastbodyno", con);
                SQLiteDataReader r3 = q4.ExecuteReader();
                while (r3.Read())
                {
                    var rcell = sheet.CreateRow(row);
                    rcell.CreateCell(0).SetCellValue(Convert.ToInt32(r3["bodyno"]));
                    rcell.CreateCell(1).SetCellValue(Convert.ToInt32(r3["u6"]));
                    rcell.CreateCell(2).SetCellValue(r3["saat"].ToString());
                    row++;
                }
                r3.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                sheet.ForceFormulaRecalculation = true;
                fs = new FileStream(@"C:\\Users\\Caner\\Desktop\\body sequence retention\\sonuc3.xlsx", FileMode.Open, FileAccess.ReadWrite);
                templateWorkbook.Write(fs);
                fs.Close();
                MessageBox.Show("Success!");
            }
            con.Close();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            // Kontrolü göster
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                file = openFileDialog1.FileName;
                SQLiteConnection con3 = new SQLiteConnection("Data Source=db.db3;Version=3;Read Only=False;");
                con3.Open();
                SQLiteCommand com3 = new SQLiteCommand("select bodyno from sirat3", con3);
                SQLiteDataReader et = com3.ExecuteReader();
                if (et != null)
                {
                    SQLiteCommand q5 = new SQLiteCommand("delete from sqlite_sequence where name = 'sirat3'", con3);
                    q5.ExecuteNonQuery();
                    SQLiteCommand comd = new SQLiteCommand("delete from sirat3", con3);
                    comd.ExecuteNonQuery();
                }
                et.Close();
                try
                {
                    FileStream fs = new FileStream(file, FileMode.Open, FileAccess.ReadWrite);
                    XSSFWorkbook templateWorkbook = new XSSFWorkbook(fs);
                    XSSFSheet sheet = (XSSFSheet)templateWorkbook.GetSheet("BİRLEŞTİRİLMİŞ");
                    XSSFSheet sheet2 = (XSSFSheet)templateWorkbook.GetSheet("SIRA");
                    sheet.ForceFormulaRecalculation = false;
                    sheet2.ForceFormulaRecalculation = false;
                    for (int rCnt = 2; rCnt < 560; rCnt++)
                    {
                        SQLiteCommand comi2 = new SQLiteCommand("insert into sirat3(bodyno,t3) values(@bodyno,@t3)", con3);
                        comi2.Parameters.AddWithValue("@bodyno", sheet2.GetRow(rCnt - 1).GetCell(1).NumericCellValue);
                        comi2.Parameters.AddWithValue("@t3", sheet2.GetRow(rCnt - 1).GetCell(36).NumericCellValue);
                        comi2.ExecuteNonQuery();
                        SQLiteCommand comi = new SQLiteCommand("insert into t3b0(bodyno) values(@bodyno)", con3);
                        comi.Parameters.AddWithValue("@bodyno", sheet.GetRow(rCnt).GetCell(9).NumericCellValue);
                        comi.ExecuteNonQuery();
                    }
                    SQLiteCommand q = new SQLiteCommand("select bodyno from t3b0", con3);
                    SQLiteDataReader r = q.ExecuteReader();
                    while (r.Read())
                    {
                        SQLiteCommand comi = new SQLiteCommand("update t3b0 set t3 = (select t3 from sirat3 where bodyno = '" + Convert.ToInt32(r["bodyno"]) + "') where bodyno = '" + Convert.ToInt32(r["bodyno"]) + "'", con3);
                        comi.ExecuteNonQuery();
                    }
                    r.Close();

                    con3.Close();
                    fs.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                MessageBox.Show("Success!");
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog2.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                String ed = openFileDialog2.FileName;


                SQLiteConnection con3 = new SQLiteConnection("Data Source=db.db3;Version=3;Read Only=False;");
                con3.Open();
                SQLiteCommand com3 = new SQLiteCommand("select saat from tcminutes", con3);
                SQLiteDataReader et = com3.ExecuteReader();
                if (et != null)
                {
                    SQLiteCommand comd = new SQLiteCommand("delete from tcminutes", con3);
                    comd.ExecuteNonQuery();
                }
                et.Close();
                try
                {
                    FileStream fs = new FileStream(ed, FileMode.Open, FileAccess.ReadWrite);
                    XSSFWorkbook templateWorkbook = new XSSFWorkbook(fs);
                    XSSFSheet sheet = (XSSFSheet)templateWorkbook.GetSheet("Araç Sayıları");
                    sheet.ForceFormulaRecalculation = false;
                    for (int rCnt = 1; rCnt < 12; rCnt++)
                    {
                        SQLiteCommand comi = new SQLiteCommand("insert into tcminutes(saat,adet) values(@saat,@adet)", con3);
                        DateTime timeTemp = sheet.GetRow(rCnt).GetCell(0).DateCellValue;
                        String saat = timeTemp.ToString("HH:mm");
                        comi.Parameters.AddWithValue("@saat", saat);
                        comi.Parameters.AddWithValue("@adet", sheet.GetRow(rCnt).GetCell(1).NumericCellValue);
                        comi.ExecuteNonQuery();
                    }
                    con3.Close();
                    fs.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                MessageBox.Show("Success!");
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            SQLiteConnection con2 = new SQLiteConnection("Data Source=db.db3;Version=3;Read Only=False;");
            con2.Open();
            SQLiteCommand q5 = new SQLiteCommand("delete from sqlite_sequence where name = 'lastopcoat'", con2);
            q5.ExecuteNonQuery();
            SQLiteCommand q6 = new SQLiteCommand("delete from lastopcoat", con2);
            q6.ExecuteNonQuery();
            con2.Close();
            int line1Min = -1000;//(int)numericUpDown3.Value;
            int line1Max = -30;//(int)numericUpDown4.Value;
            int taktIn = 98;//Convert.ToInt32(textBox1.Text);
            int taktOut = 123;//Convert.ToInt32(textBox2.Text);
            int adet=0, lineSira = 0, lineSiraOut = 0, lin1Adet = 0, lin2Adet = 0, lin3Adet = 0;
            List<int> arrayBody = new List<int>();
            List<int> arrayU6 = new List<int>();
            DateTime date = new DateTime(11, 11, 11, 8, 0, 0);
            SQLiteConnection con = new SQLiteConnection("Data Source=db.db3;Version=3;Read Only=False;");
            SQLiteCommand q = new SQLiteCommand("select * from t3b0", con);
            con.Open();
            SQLiteDataReader r = q.ExecuteReader();
            int bodyno, u6;
            for (int sec = 1; sec <= 54000; sec++)
            {
                if (sec % taktIn == 0)
                {
                    r.Read();
                    if (!r.HasRows)
                    {
                        MessageBox.Show("Araç bitti saniye: " + sec);
                        break;
                    }
                    bodyno = Convert.ToInt32(r["bodyno"]);
                    u6 = Convert.ToInt32(r["t3"]);

                    if (line1Min < u6 && u6 < line1Max)
                    {
                        line1.Items.Add(bodyno);
                        if (arrayU6.Count != 0)
                        {
                            for (int i = 0; i < arrayU6.Count; i++)
                            {
                                arrayU6[i] = arrayU6[i] - 1;
                            }
                        }
                        arrayBody.Add(bodyno);
                        arrayU6.Add(u6);
                        log();
                        lin1Adet++;
                    }
                    else
                    {
                            line2.Items.Add(bodyno);
                            arrayBody.Add(bodyno);
                            arrayU6.Add(u6);
                            lin2Adet++;
                            log();
                    }
                }
                if (sec % taktOut == 0)
                {
                    date = date.AddSeconds(taktOut);
                    time = date.ToString("HH:mm");
                    SQLiteCommand q2 = new SQLiteCommand("select adet from tcminutes where saat ='" + time + "' ", con);
                    SQLiteDataReader r2 = q2.ExecuteReader();
                    r2.Read();
                    if (r2.HasRows)
                        adet = Convert.ToInt32(r2["adet"]);
                    r2.Close();

                    if (line1.Items.Count > 0)
                    {
                        bodyno = Convert.ToInt32(line1.Items[0]);
                        u6 = arrayU6[arrayBody.IndexOf(bodyno)];
                        int oldb9 = u6;
                        u6 = adet + u6;
                        arrayU6[arrayBody.IndexOf(bodyno)] = u6;
                        line1.Items.RemoveAt(0);
                        SQLiteCommand q3 = new SQLiteCommand("Insert Into lastopcoat(bodyno,t3,saat) values(@body,@t3,@saat)", con);
                        q3.Parameters.AddWithValue("@body", bodyno);
                        q3.Parameters.AddWithValue("@t3", u6);
                        q3.Parameters.AddWithValue("@saat", time);
                        q3.ExecuteNonQuery();
                        arrayBody.Remove(bodyno);
                        arrayU6.Remove(u6);
                        log();
                    }
                    else if (line2.Items.Count > 0)
                    {
                        bodyno = Convert.ToInt32(line2.Items[0]);
                        line2.Items.RemoveAt(0);
                        u6 = arrayU6[arrayBody.IndexOf(bodyno)];
                        int oldb9 = u6;
                        SQLiteCommand q3 = new SQLiteCommand("Insert Into lastopcoat(bodyno,t3,saat) values(@body,@t3,@saat)", con);
                        q3.Parameters.AddWithValue("@body", bodyno);
                        q3.Parameters.AddWithValue("@t3", u6);
                        q3.Parameters.AddWithValue("@saat", time);
                        q3.ExecuteNonQuery();
                        arrayBody.Remove(bodyno);
                        arrayU6.Remove(u6);
                        log();
                    }
                }
            }
            label4.Text = lin1Adet.ToString();
            label5.Text = lin2Adet.ToString();
            r.Close();
            con.Close();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            FileStream fs = new FileStream(@"C:\\Users\\Caner\\Desktop\\body sequence retention\\sonuc4.xlsx", FileMode.Open, FileAccess.ReadWrite);
            XSSFWorkbook templateWorkbook = new XSSFWorkbook(fs);
            XSSFSheet sheet = (XSSFSheet)templateWorkbook.GetSheet("Sheet1");
            fs.Close();

            List<int> u6s = new List<int>();
            int row = 0, nu6 = 0;
            SQLiteConnection con = new SQLiteConnection("Data Source=db.db3;Version=3;Read Only=False;");
            con.Open();
            try
            {
                SQLiteCommand q4 = new SQLiteCommand("select * from lastopcoat", con);
                SQLiteDataReader r3 = q4.ExecuteReader();
                while (r3.Read())
                {
                    SQLiteCommand q5 = new SQLiteCommand("select t3 from t3p3 where bodyno = '" + r3["bodyno"] + "'", con);
                    SQLiteDataReader r4 = q5.ExecuteReader();
                    r4.Read();
                    if (r4.HasRows)
                        nu6 = Convert.ToInt32(r4["t3"]) + Convert.ToInt32(r3["t3"]);
                    else
                        nu6 = 0;
                    r4.Close();
                    
                    var rcell = sheet.CreateRow(row);
                    rcell.CreateCell(0).SetCellValue(Convert.ToInt32(r3["bodyno"]));
                    rcell.CreateCell(1).SetCellValue(nu6);
                    rcell.CreateCell(2).SetCellValue(r3["saat"].ToString());
                    row++;
                }
                r3.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                sheet.ForceFormulaRecalculation = true;
                fs = new FileStream(@"C:\\Users\\Caner\\Desktop\\body sequence retention\\sonuc4.xlsx", FileMode.Open, FileAccess.ReadWrite);
                templateWorkbook.Write(fs);
                fs.Close();
                MessageBox.Show("Success!");
            }
            con.Close();
        }

        private void button17_Click(object sender, EventArgs e)
        {
            SQLiteConnection con3 = new SQLiteConnection("Data Source=db.db3;Version=3;Read Only=False;");
            con3.Open();
            SQLiteCommand com3 = new SQLiteCommand("select bodyno from t3p3", con3);
            SQLiteDataReader et = com3.ExecuteReader();
            if (et != null)
            {
                SQLiteCommand q5 = new SQLiteCommand("delete from sqlite_sequence where name = 't3p3'", con3);
                q5.ExecuteNonQuery();
                SQLiteCommand comd = new SQLiteCommand("delete from t3p3", con3);
                comd.ExecuteNonQuery();
            }
            et.Close();
            try
            {
                FileStream fs = new FileStream(@"C:\\Users\\Caner\\Desktop\\body sequence retention\\1-5323-5880.xlsx", FileMode.Open, FileAccess.ReadWrite);
                XSSFWorkbook templateWorkbook = new XSSFWorkbook(fs);
                XSSFSheet sheet = (XSSFSheet)templateWorkbook.GetSheet("SIRA");
                sheet.ForceFormulaRecalculation = false;
                for (int rCnt = 1; rCnt < 559; rCnt++)
                {
                    SQLiteCommand comi = new SQLiteCommand("insert into t3p3(bodyno,t3) values(@bodyno,@t3)", con3);
                    comi.Parameters.AddWithValue("@bodyno", sheet.GetRow(rCnt).GetCell(1).NumericCellValue);
                    comi.Parameters.AddWithValue("@t3", sheet.GetRow(rCnt).GetCell(24).NumericCellValue);
                    comi.ExecuteNonQuery();
                }
                con3.Close();
                fs.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            MessageBox.Show("Success!");
        }
    }
}
