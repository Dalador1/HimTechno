using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;



namespace WindowsFormsApplication19
{



    public partial class Form1 : Form
    {

        SqlConnection sqlConnection;



        public Form1()
        {
            InitializeComponent();
        }

        string buf, buf1, buf2;
        string cell14, cell26, cell38, cell50, cell62, cell74, cell13, cell9; // Колонка А
        string cell15, cell27, cell39, cell51, cell63, cell75, cell86, cell8; // Колонка L
        string cell16, cell28, cell40, cell52, cell64, cell76, cell87, cell10; //Колонка AF 
        string cell17, cell29, cell41, cell53, cell65, cell77, cell88, cell12;//Колонка AS
        string cell18, cell30, cell42, cell54, cell66, cell78, cell89, cell11;//Колонка BE
        string cell19, cell31, cell43, cell55, cell67, cell79, cell90, cell7;//Колонка BU
        string cell20, cell32, cell44, cell56, cell68, cell80, cell91, cell3;//Колонка DH
        string cell21, cell33, cell45, cell57, cell69, cell81, cell92, cell2;//Колонка DS
        string cell22, cell34, cell46, cell58, cell70, cell82, cell93, cell1;//Колонка ED
        string cell23, cell35, cell47, cell59, cell71, cell83, cell94, cell6;//Колонка EO
        string cell24, cell36, cell48, cell60, cell72, cell84, cell95, cell5;//Колонка FC
        string cell25, cell37, cell49, cell61, cell73, cell85, cell96, cell4; //Колонка FN

        string Day, Mounth, Year,DayDostavki, MounthDostavki, YearDostavki; //Заполнение даты 
        string TTN;
        string driver, marka, number, company, marshrut;
        string Dropallowed, Dropdone, Dropallowname, Dropdonename, buhgalter;
        string names, places, priem, sda4a, pogryzka;
        string propis1, propis2;

        string OKPO; 

        public void WriteData()
        {
            Excel excel = new Excel(@"C:\TEXNO\test2.xls", 1);
            excel.WriteToColumnA(cell14, cell26, cell38, cell50, cell62, cell74, cell86, cell9);
            excel.WriteToColumnL(cell15, cell27, cell39, cell51, cell63, cell75, cell87, cell8);
            excel.WriteToColumnAF(cell16, cell28, cell40, cell52, cell64, cell76, cell87, cell10);
            excel.WriteToColumnAS(cell17, cell29, cell41, cell53, cell65, cell77, cell88, cell12);
            excel.WriteToColumnBE(cell18, cell30, cell42, cell54, cell66, cell78, cell89, cell11);
            excel.WriteToColumnBU(cell19, cell31, cell43, cell55, cell67, cell79, cell90, cell7);
            excel.WriteToColumnDH(cell20, cell32, cell44, cell56, cell68, cell80, cell91, cell3);
            excel.WriteToColumnDS(cell21, cell33, cell45, cell57, cell69, cell81, cell92, cell2);
            excel.WriteToColumnED(cell22, cell34, cell46, cell58, cell70, cell82, cell90, cell1);
            excel.WriteToColumnEO(cell23, cell35, cell47, cell59, cell71, cell83, cell91, cell6);
            excel.WriteToColumnFC(cell24, cell36, cell48, cell60, cell72, cell84, cell92, cell5);
            excel.WriteToColumnFN(cell25, cell37, cell49, cell61, cell73, cell85, cell93, cell4);
            excel.WriteToTTN(TTN);
            excel.WriteToCell(buf, buf1, buf2, cell19, cell31, cell43, cell55, cell67, cell79, cell90, cell7, Dropallowed, Dropdone, Dropallowname, Dropdonename, buhgalter, driver, company, marka, number, Day, Year, Mounth, names, places, priem, sda4a, DayDostavki, MounthDostavki, YearDostavki, marshrut,pogryzka, propis1, propis2);
            //excel.WriteDriver(driver);
           // excel.WriteToPoste(Dropallowed, Dropdone, Dropallowed, Dropdonename, buhgalter); 
            // cell19, cell31, cell43, cell55 , cell67, cell79, cell90 ,cell7
            excel.Close();

        }



        private async void Form1_Load(object sender, EventArgs e)
        {


            MessageBox.Show("Данная программа находится на стадии разработки. Перед расспечаткой ТН и формы ТОРГ-12 настоятельно рекомендуется проверить excel-файл в ручную");




            // TODO: данная строка кода позволяет загрузить данные в таблицу "databaseDataSet2.Customers". При необходимости она может быть перемещена или удалена.
            this.customersTableAdapter1.Fill(this.databaseDataSet2.Customers);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "databaseDataSet1.Customers". При необходимости она может быть перемещена или удалена.
            this.customersTableAdapter.Fill(this.databaseDataSet1.Customers);


            string connectionString = @"Data Source=(LocalDB)\v11.0;AttachDbFilename=C:\Users\1\Documents\Visual Studio 2012\Projects\WindowsFormsApplication19\WindowsFormsApplication19\Database.mdf;Integrated Security=True";
            sqlConnection = new SqlConnection(connectionString);
            await sqlConnection.OpenAsync();
            SqlDataReader sqlReader = null;
            SqlCommand command = new SqlCommand("SELECT *FROM [Customers]", sqlConnection);
            try
            {
                sqlReader = await command.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    if (checkBox1.Checked)

                    listBox1.Items.Add(Convert.ToString(sqlReader["id"]) + "   " + Convert.ToString(sqlReader["Name"]) + " " + Convert.ToString(sqlReader["Address"]) + " " + Convert.ToString(sqlReader["Number"]) + "   " + Convert.ToString(sqlReader["CA"]) + "   " + Convert.ToString(sqlReader["CRA"]) + " " + Convert.ToString(sqlReader["BankIC"]) + " " + Convert.ToString(sqlReader["Bank"]) + " " + Convert.ToString(sqlReader["INN"]) + " " + Convert.ToString(sqlReader["CPP"]) + " " + Convert.ToString(sqlReader["OKPO"]));

                    else

                    listBox1.Items.Add(Convert.ToString(sqlReader["Name"]) + " " + Convert.ToString(sqlReader["Address"]) + " " + Convert.ToString(sqlReader["Number"]) + "   " + Convert.ToString(sqlReader["CA"]) + "   " + Convert.ToString(sqlReader["CRA"]) + " " + Convert.ToString(sqlReader["BankIC"]) + " " + Convert.ToString(sqlReader["Bank"]) + " " + Convert.ToString(sqlReader["INN"]) + " " + Convert.ToString(sqlReader["CPP"]) + " " + Convert.ToString(sqlReader["OKPO"]));

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }


            //
            //
            //

            SqlDataReader sqlReader1 = null;
            SqlCommand cmd = new SqlCommand("SELECT *FROM [Delivers]", sqlConnection);
            try
            {
                sqlReader1 = await cmd.ExecuteReaderAsync();
                while (await sqlReader1.ReadAsync())
                {
                    if (checkBox1.Checked)

                    listBox2.Items.Add(Convert.ToString(sqlReader1["id"]) + "   " + Convert.ToString(sqlReader1["Name"]) + " " + Convert.ToString(sqlReader1["Address"]) + " " + Convert.ToString(sqlReader1["Number"]) + "   " + Convert.ToString(sqlReader1["CA"]) + "   " + Convert.ToString(sqlReader1["CRA"]) + " " + Convert.ToString(sqlReader1["BankIC"]) + " " + Convert.ToString(sqlReader1["Bank"]) + " " + Convert.ToString(sqlReader1["INN"]) + " " + Convert.ToString(sqlReader1["CPP"]) + " " + Convert.ToString(sqlReader1["OKPO"]));

                    else

                    listBox2.Items.Add(Convert.ToString(sqlReader1["Name"]) + " " + Convert.ToString(sqlReader1["Address"]) + " " + Convert.ToString(sqlReader1["Number"]) + "   " + Convert.ToString(sqlReader1["CA"]) + "   " + Convert.ToString(sqlReader1["CRA"]) + " " + Convert.ToString(sqlReader1["BankIC"]) + " " + Convert.ToString(sqlReader1["Bank"]) + " " + Convert.ToString(sqlReader1["INN"]) + " " + Convert.ToString(sqlReader1["CPP"]) + " " + Convert.ToString(sqlReader1["OKPO"]));


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            finally
            {
                if (sqlReader1 != null)
                    sqlReader1.Close();
            }

            //
            //
            //
            SqlDataReader sqlReader2 = null;
            SqlCommand cmd2 = new SqlCommand("SELECT *FROM [Buyers]", sqlConnection);
            try
            {
                sqlReader2 = await cmd2.ExecuteReaderAsync();
                while (await sqlReader2.ReadAsync())
                {
                    if (checkBox1.Checked)

                    listBox3.Items.Add(Convert.ToString(sqlReader2["id"]) + "   " + Convert.ToString(sqlReader2["Name"]) + " " + Convert.ToString(sqlReader2["Address"]) + " " + Convert.ToString(sqlReader2["Number"]) + "   " + Convert.ToString(sqlReader2["CA"]) + "   " + Convert.ToString(sqlReader2["CRA"]) + " " + Convert.ToString(sqlReader2["BankIC"]) + " " + Convert.ToString(sqlReader2["Bank"]) + " " + Convert.ToString(sqlReader2["INN"]) + " " + Convert.ToString(sqlReader2["CPP"]) + " " + Convert.ToString(sqlReader2["OKPO"]));

                    else

                    listBox3.Items.Add(Convert.ToString(sqlReader2["Name"]) + " " + Convert.ToString(sqlReader2["Address"]) + " " + Convert.ToString(sqlReader2["Number"]) + "   " + Convert.ToString(sqlReader2["CA"]) + "   " + Convert.ToString(sqlReader2["CRA"]) + " " + Convert.ToString(sqlReader2["BankIC"]) + " " + Convert.ToString(sqlReader2["Bank"]) + " " + Convert.ToString(sqlReader2["INN"]) + " " + Convert.ToString(sqlReader2["CPP"]) + " " + Convert.ToString(sqlReader2["OKPO"]));


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            finally
            {
                if (sqlReader2 != null)
                    sqlReader2.Close();
            }



        }
        









        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            buf = textBox1.Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            WriteData();

            MessageBox.Show("Данная программа находится в разработке.Необходимо ВРУЧНУЮ удалить цифры перед реквезитами грузоотправителей, грузополучателей и плательщиков, а так же проверить excel-файл. НЕОБХОДИМО В РУЧНУЮ ПЕРЕНЕСТИ КОД ОКПО ИЗ СТРОКИ РЕКВЕЗИТОВ В ОКНО <ОКПО> EXCEL-ФАЙЛА", "НАПОМИНАНИЕ"); 

            //SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            //saveFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            //saveFileDialog1.FilterIndex = 2;
            //saveFileDialog1.RestoreDirectory = true;

            //if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)

            //{
            //    Excel excel = new Excel(@"C:\TEXNO\test2.xls", 1);

            //   excel.wb.SaveAs(saveFileDialog1.FileName);
            
            //    excel.Close();

            //}



            //Excel excel = new Excel(@"C:\TEXNO\test2.xls", 1);
            //// excel.Save();
            //excel.SaveAs(@"C:\TEXNO\test2.xls");
            //excel.Close();
            //MessageBox.Show("Saved!");

            



        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

            buf1 = textBox2.Text;
        }



        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            buf2 = textBox3.Text;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
                sqlConnection.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
                sqlConnection.Close();
        }

        private async void button3_Click(object sender, EventArgs e)
        {

            if (label14.Visible)
                label14.Visible = false;



            if (!string.IsNullOrEmpty(textBox4.Text) && !string.IsNullOrWhiteSpace(textBox4.Text) &&
                !string.IsNullOrEmpty(textBox5.Text) && !string.IsNullOrWhiteSpace(textBox5.Text) &&
                !string.IsNullOrEmpty(textBox6.Text) && !string.IsNullOrWhiteSpace(textBox6.Text) &&
                !string.IsNullOrEmpty(textBox7.Text) && !string.IsNullOrWhiteSpace(textBox7.Text) &&
                !string.IsNullOrEmpty(textBox8.Text) && !string.IsNullOrWhiteSpace(textBox8.Text) &&
                !string.IsNullOrEmpty(textBox8.Text) && !string.IsNullOrWhiteSpace(textBox8.Text) &&
                !string.IsNullOrEmpty(textBox9.Text) && !string.IsNullOrWhiteSpace(textBox9.Text) &&
                !string.IsNullOrEmpty(textBox10.Text) && !string.IsNullOrWhiteSpace(textBox10.Text) &&
                !string.IsNullOrEmpty(textBox11.Text) && !string.IsNullOrWhiteSpace(textBox11.Text) &&
                !string.IsNullOrEmpty(textBox12.Text) && !string.IsNullOrWhiteSpace(textBox12.Text) &&
                !string.IsNullOrEmpty(textBox13.Text) && !string.IsNullOrWhiteSpace(textBox13.Text))
            {

                SqlCommand command = new SqlCommand("INSERT INTO [Customers] (Name,Address,Number,CA,CRA,BankIC,Bank,INN,CPP,OKPO)VALUES(@Name, @Address,@Number,@CA,@CRA,@BankIC,@Bank,@INN,@CPP,@OKPO)", sqlConnection);

                command.Parameters.AddWithValue("Name", textBox4.Text);
                command.Parameters.AddWithValue("Address", textBox5.Text);
                command.Parameters.AddWithValue("Number", textBox6.Text);
                command.Parameters.AddWithValue("CA", textBox7.Text);
                command.Parameters.AddWithValue("CRA", textBox8.Text);
                command.Parameters.AddWithValue("BankIC", textBox9.Text);
                command.Parameters.AddWithValue("Bank", textBox10.Text);
                command.Parameters.AddWithValue("INN", textBox11.Text);
                command.Parameters.AddWithValue("CPP", textBox12.Text);
                command.Parameters.AddWithValue("OKPO", textBox13.Text);
                await command.ExecuteNonQueryAsync();


            }
            else
            {
                label14.Visible = true;
                label14.Text = "Все поля должны быть заполнены!";
            }

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private async void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            SqlDataReader sqlReader = null;
            SqlCommand command = new SqlCommand("SELECT *FROM [Customers]", sqlConnection);
            try
            {
                sqlReader = await command.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())

                    if (checkBox1.Checked)
                    {
                        listBox1.Items.Add(Convert.ToString(sqlReader["id"]) + "   " + Convert.ToString(sqlReader["Name"]) + " " + Convert.ToString(sqlReader["Address"]) + " " + Convert.ToString(sqlReader["Number"]) + "   " + Convert.ToString(sqlReader["CA"]) + "   " + Convert.ToString(sqlReader["CRA"]) + " " + Convert.ToString(sqlReader["BankIC"]) + " " + Convert.ToString(sqlReader["Bank"]) + " " + Convert.ToString(sqlReader["INN"]) + " " + Convert.ToString(sqlReader["CPP"]) + " " + Convert.ToString(sqlReader["OKPO"]));

                    }
                else
                    
                {
                    listBox1.Items.Add(Convert.ToString(sqlReader["Name"]) + " " + Convert.ToString(sqlReader["Address"]) + " " + Convert.ToString(sqlReader["Number"]) + "   " + Convert.ToString(sqlReader["CA"]) + "   " + Convert.ToString(sqlReader["CRA"]) + " " + Convert.ToString(sqlReader["BankIC"]) + " " + Convert.ToString(sqlReader["Bank"]) + " " + Convert.ToString(sqlReader["INN"]) + " " + Convert.ToString(sqlReader["CPP"]) + " " + Convert.ToString(sqlReader["OKPO"]));

                }
            }


            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }


            listBox2.Items.Clear();
            SqlDataReader sqlReader1 = null;
            SqlCommand cmd = new SqlCommand("SELECT *FROM [Delivers]", sqlConnection);
            try
            {
                sqlReader1 = await cmd.ExecuteReaderAsync();
                while (await sqlReader1.ReadAsync())

                    if (checkBox1.Checked)
                    {
                        listBox2.Items.Add(Convert.ToString(sqlReader1["id"]) + "   " + Convert.ToString(sqlReader1["Name"]) + " " + Convert.ToString(sqlReader1["Address"]) + " " + Convert.ToString(sqlReader1["Number"]) + "   " + Convert.ToString(sqlReader1["CA"]) + "   " + Convert.ToString(sqlReader1["CRA"]) + " " + Convert.ToString(sqlReader1["BankIC"]) + " " + Convert.ToString(sqlReader1["Bank"]) + " " + Convert.ToString(sqlReader1["INN"]) + " " + Convert.ToString(sqlReader1["CPP"]) + " " + Convert.ToString(sqlReader1["OKPO"]));

                    }
                    else
                    {
                        listBox2.Items.Add(Convert.ToString(sqlReader1["Name"]) + " " + Convert.ToString(sqlReader1["Address"]) + " " + Convert.ToString(sqlReader1["Number"]) + "   " + Convert.ToString(sqlReader1["CA"]) + "   " + Convert.ToString(sqlReader1["CRA"]) + " " + Convert.ToString(sqlReader1["BankIC"]) + " " + Convert.ToString(sqlReader1["Bank"]) + " " + Convert.ToString(sqlReader1["INN"]) + " " + Convert.ToString(sqlReader1["CPP"]) + " " + Convert.ToString(sqlReader1["OKPO"]));

                    }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader1 != null)
                    sqlReader1.Close();
            }



            listBox3.Items.Clear();
            SqlDataReader sqlReader2 = null;
            SqlCommand cmd2 = new SqlCommand("SELECT *FROM [Buyers]", sqlConnection);
            try
            {
                sqlReader2 = await cmd2.ExecuteReaderAsync();
                while (await sqlReader2.ReadAsync())
                    if (checkBox1.Checked)
                    {
                        listBox3.Items.Add(Convert.ToString(sqlReader2["id"]) + "   " + Convert.ToString(sqlReader2["Name"]) + " " + Convert.ToString(sqlReader2["Address"]) + " " + Convert.ToString(sqlReader2["Number"]) + "   " + Convert.ToString(sqlReader2["CA"]) + "   " + Convert.ToString(sqlReader2["CRA"]) + " " + Convert.ToString(sqlReader2["BankIC"]) + " " + Convert.ToString(sqlReader2["Bank"]) + " " + Convert.ToString(sqlReader2["INN"]) + " " + Convert.ToString(sqlReader2["CPP"]) + " " + Convert.ToString(sqlReader2["OKPO"]));

                    }
                    else
                    {
                        listBox3.Items.Add(Convert.ToString(sqlReader2["Name"]) + " " + Convert.ToString(sqlReader2["Address"]) + " " + Convert.ToString(sqlReader2["Number"]) + "   " + Convert.ToString(sqlReader2["CA"]) + "   " + Convert.ToString(sqlReader2["CRA"]) + " " + Convert.ToString(sqlReader2["BankIC"]) + " " + Convert.ToString(sqlReader2["Bank"]) + " " + Convert.ToString(sqlReader2["INN"]) + " " + Convert.ToString(sqlReader2["CPP"]) + " " + Convert.ToString(sqlReader2["OKPO"]));

                    }
            }


            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader2 != null)
                    sqlReader2.Close();
            }





            


        }

        private void label25_Click(object sender, EventArgs e)
        {

        }

        private async void button4_Click(object sender, EventArgs e)
        {
            if (label25.Visible)
                label25.Visible = false;

            if (!string.IsNullOrEmpty(textBox14.Text) && !string.IsNullOrWhiteSpace(textBox14.Text) &&
                !string.IsNullOrEmpty(textBox15.Text) && !string.IsNullOrWhiteSpace(textBox15.Text) &&
                !string.IsNullOrEmpty(textBox16.Text) && !string.IsNullOrWhiteSpace(textBox16.Text) &&
                !string.IsNullOrEmpty(textBox17.Text) && !string.IsNullOrWhiteSpace(textBox17.Text) &&
                !string.IsNullOrEmpty(textBox18.Text) && !string.IsNullOrWhiteSpace(textBox18.Text) &&
                !string.IsNullOrEmpty(textBox19.Text) && !string.IsNullOrWhiteSpace(textBox19.Text) &&
                !string.IsNullOrEmpty(textBox20.Text) && !string.IsNullOrWhiteSpace(textBox20.Text) &&
                !string.IsNullOrEmpty(textBox21.Text) && !string.IsNullOrWhiteSpace(textBox21.Text) &&
                !string.IsNullOrEmpty(textBox22.Text) && !string.IsNullOrWhiteSpace(textBox22.Text) &&
                !string.IsNullOrEmpty(textBox23.Text) && !string.IsNullOrWhiteSpace(textBox23.Text) &&
                !string.IsNullOrEmpty(textBox24.Text) && !string.IsNullOrWhiteSpace(textBox24.Text))
            {
                SqlCommand command = new SqlCommand("UPDATE [Customers] SET [Name]=@Name,[Address]=@Address,[Number]=@Number,[CA]=@CA,[CRA]=@CRA,[BankIC]=@BankIC,[Bank]=@Bank,[INN]=@INN,[CPP]=@CPP,[OKPO]=@OKPO WHERE [Id]=@Id", sqlConnection);
                command.Parameters.AddWithValue("Id", textBox24.Text);
                command.Parameters.AddWithValue("Name", textBox14.Text);
                command.Parameters.AddWithValue("Address", textBox15.Text);
                command.Parameters.AddWithValue("Number", textBox16.Text);
                command.Parameters.AddWithValue("CA", textBox17.Text);
                command.Parameters.AddWithValue("CRA", textBox18.Text);
                command.Parameters.AddWithValue("BankIC", textBox19.Text);
                command.Parameters.AddWithValue("Bank", textBox20.Text);
                command.Parameters.AddWithValue("INN", textBox21.Text);
                command.Parameters.AddWithValue("CPP", textBox22.Text);
                command.Parameters.AddWithValue("OKPO", textBox23.Text);

                await command.ExecuteNonQueryAsync();


            }

            else if (!string.IsNullOrEmpty(textBox24.Text) && !string.IsNullOrWhiteSpace(textBox24.Text))
            {

                label25.Visible = true;
                label25.Text = "id должен заполнить!";

            }

            else
            {
                label25.Visible = true;
                label25.Text = "Поля 'имя' и 'цена' должны быть заполнены!";

            }

        }

        private async void button5_Click(object sender, EventArgs e)
        {

            if (label27.Visible)
                label27.Visible = false;


            if (!string.IsNullOrEmpty(textBox25.Text) && !string.IsNullOrWhiteSpace(textBox25.Text))

            {
                SqlCommand command = new SqlCommand("DELETE FROM [Customers] WHERE [Id] = @id", sqlConnection);
                command.Parameters.AddWithValue("id", textBox25.Text);
                await command.ExecuteNonQueryAsync();

            }
            else
            {
                label27.Visible = true;
                label27.Text = "Id должен быть заполнен!";
            }

            MessageBox.Show("Удалено"); 



        }

        
        private void fillByToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.customersTableAdapter1.FillBy(this.databaseDataSet2.Customers);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        
        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {

            if (listBox1.SelectedItem != null && listBox2.SelectedItem != null && listBox3.SelectedItem != null)
            {
                buf = listBox1.SelectedItem.ToString();
                buf1 = listBox2.SelectedItem.ToString();
                buf2 = listBox3.SelectedItem.ToString();
            }
            else
            {
                MessageBox.Show("Выберите данные!");
            }

        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label52_Click(object sender, EventArgs e)
        {

        }


        //columnA
        private void richTextBox14_TextChanged(object sender, EventArgs e)
        {
            cell14 = richTextBox14.Text; 
        }

        private void richTextBox26_TextChanged(object sender, EventArgs e)
        {
            cell26 = richTextBox26.Text;
        }
        private void richTextBox38_TextChanged(object sender, EventArgs e)
        {
            cell38 = richTextBox38.Text;
        }
        private void richTextBox50_TextChanged(object sender, EventArgs e)
        {
            cell50 = richTextBox50.Text;
        }
        private void richTextBox62_TextChanged(object sender, EventArgs e)
        {
            cell62 = richTextBox62.Text;
        }
        private void richTextBox74_TextChanged(object sender, EventArgs e)
        {
            cell74 = richTextBox74.Text;
        }
        private void richTextBox13_TextChanged(object sender, EventArgs e)
        {
            cell13 = richTextBox13.Text;

        }
        private void richTextBox9_TextChanged(object sender, EventArgs e)
        {
            cell9 = richTextBox9.Text;
        }
        //column L 
        private void richTextBox15_TextChanged(object sender, EventArgs e)
        {
            cell15 = richTextBox15.Text;
        }

        private void richTextBox27_TextChanged(object sender, EventArgs e)
        {
            cell27 = richTextBox27.Text;
        }
        private void richTextBox39_TextChanged(object sender, EventArgs e)
        {
            cell39 = richTextBox39.Text;
        }
        private void richTextBox51_TextChanged(object sender, EventArgs e)
        {
            cell51 = richTextBox51.Text;
        }
        private void richTextBox63_TextChanged(object sender, EventArgs e)
        {
            cell63 = richTextBox63.Text;
        }
        private void richTextBox75_TextChanged(object sender, EventArgs e)
        {
            cell75 = richTextBox75.Text;
        }
        private void richTextBox86_TextChanged(object sender, EventArgs e)
        {
            cell86 = richTextBox86.Text;

        }
        private void richTextBox8_TextChanged(object sender, EventArgs e)
        {
            cell8 = richTextBox8.Text;
        }


        //column AF
        private void richTextBox16_TextChanged(object sender, EventArgs e)
        {
            cell16 = richTextBox16.Text;
        }

        private void richTextBox28_TextChanged(object sender, EventArgs e)
        {
            cell28 = richTextBox28.Text;
        }
        private void richTextBox40_TextChanged(object sender, EventArgs e)
        {
            cell40 = richTextBox40.Text;
        }
        private void richTextBox52_TextChanged(object sender, EventArgs e)
        {
            cell52 = richTextBox52.Text;
        }
        private void richTextBox64_TextChanged(object sender, EventArgs e)
        {
            cell64 = richTextBox64.Text;
        }
        private void richTextBox76_TextChanged(object sender, EventArgs e)
        {
            cell76 = richTextBox76.Text;
        }
        private void richTextBox87_TextChanged(object sender, EventArgs e)
        {
            cell87 = richTextBox87.Text;

        }
        private void richTextBox10_TextChanged(object sender, EventArgs e)
        {
            cell10 = richTextBox10.Text;
        }

        //column AF 
        private void richTextBox17_TextChanged(object sender, EventArgs e)
        {
            cell17 = richTextBox17.Text;
        }

        private void richTextBox29_TextChanged(object sender, EventArgs e)
        {
            cell29 = richTextBox29.Text;
        }
        private void richTextBox41_TextChanged(object sender, EventArgs e)
        {
            cell41 = richTextBox41.Text;
        }
        private void richTextBox53_TextChanged(object sender, EventArgs e)
        {
            cell53 = richTextBox53.Text;
        }
        private void richTextBox65_TextChanged(object sender, EventArgs e)
        {
            cell65 = richTextBox65.Text;
        }
        private void richTextBox77_TextChanged(object sender, EventArgs e)
        {
            cell77 = richTextBox77.Text;
        }
        private void richTextBox88_TextChanged(object sender, EventArgs e)
        {
            cell88 = richTextBox88.Text;

        }
        private void richTextBox12_TextChanged(object sender, EventArgs e)
        {
            cell12 = richTextBox12.Text;
        }

        //column BE 
        private void richTextBox18_TextChanged(object sender, EventArgs e)
        {
            cell18 = richTextBox18.Text;
        }

        private void richTextBox30_TextChanged(object sender, EventArgs e)
        {
            cell30 = richTextBox30.Text;
        }
        private void richTextBox42_TextChanged(object sender, EventArgs e)
        {
            cell42 = richTextBox42.Text;
        }
        private void richTextBox54_TextChanged(object sender, EventArgs e)
        {
            cell54 = richTextBox54.Text;
        }
        private void richTextBox66_TextChanged(object sender, EventArgs e)
        {
            cell66 = richTextBox66.Text;
        }
        private void richTextBox78_TextChanged(object sender, EventArgs e)
        {
            cell78 = richTextBox78.Text;
        }
        private void richTextBox89_TextChanged(object sender, EventArgs e)
        {
            cell89 = richTextBox89.Text;

        }
        private void richTextBox11_TextChanged(object sender, EventArgs e)
        {
            cell11 = richTextBox11.Text;
        }


        //column BU 
        private void richTextBox19_TextChanged(object sender, EventArgs e)
        {
            cell19 = richTextBox19.Text;
        }

        private void richTextBox31_TextChanged(object sender, EventArgs e)
        {
            cell31 = richTextBox31.Text;
        }
        private void richTextBox43_TextChanged(object sender, EventArgs e)
        {
            cell43 = richTextBox43.Text;
        }
        private void richTextBox55_TextChanged(object sender, EventArgs e)
        {
            cell55 = richTextBox55.Text;
        }
        private void richTextBox67_TextChanged(object sender, EventArgs e)
        {
            cell67 = richTextBox67.Text;
        }
        private void richTextBox79_TextChanged(object sender, EventArgs e)
        {
            cell79 = richTextBox79.Text;
        }
        private void richTextBox90_TextChanged(object sender, EventArgs e)
        {
            cell90 = richTextBox90.Text;

        }
        private void richTextBox7_TextChanged(object sender, EventArgs e)
        {
            cell7 = richTextBox7.Text;
        }

        //column DH 
        private void richTextBox20_TextChanged(object sender, EventArgs e)
        {
            cell20 = richTextBox20.Text;
        }

        private void richTextBox32_TextChanged(object sender, EventArgs e)
        {
            cell32 = richTextBox32.Text;
        }
        private void richTextBox44_TextChanged(object sender, EventArgs e)
        {
            cell44 = richTextBox44.Text;
        }
        private void richTextBox56_TextChanged(object sender, EventArgs e)
        {
            cell56 = richTextBox56.Text;
        }
        private void richTextBox68_TextChanged(object sender, EventArgs e)
        {
            cell68 = richTextBox68.Text;
        }
        private void richTextBox80_TextChanged(object sender, EventArgs e)
        {
            cell80 = richTextBox80.Text;
        }
        private void richTextBox91_TextChanged(object sender, EventArgs e)
        {
            cell91 = richTextBox91.Text;

        }
        private void richTextBox3_TextChanged(object sender, EventArgs e)
        {
            cell3 = richTextBox13.Text;
        }


        //column DS 
        private void richTextBox21_TextChanged(object sender, EventArgs e)
        {
            cell21 = richTextBox21.Text;
        }

        private void richTextBox33_TextChanged(object sender, EventArgs e)
        {
            cell33 = richTextBox33.Text;
        }
        private void richTextBox45_TextChanged(object sender, EventArgs e)
        {
            cell45 = richTextBox45.Text;
        }
        private void richTextBox57_TextChanged(object sender, EventArgs e)
        {
            cell57 = richTextBox57.Text;
        }
        private void richTextBox69_TextChanged(object sender, EventArgs e)
        {
            cell69 = richTextBox69.Text;
        }
        private void richTextBox81_TextChanged(object sender, EventArgs e)
        {
            cell81 = richTextBox81.Text;
        }
        private void richTextBox92_TextChanged(object sender, EventArgs e)
        {
            cell92 = richTextBox92.Text;

        }
        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {
            cell2 = richTextBox2.Text;
        }


        //column ED
        private void richTextBox22_TextChanged(object sender, EventArgs e)
        {
            cell22 = richTextBox22.Text;
        }

        private void richTextBox34_TextChanged(object sender, EventArgs e)
        {
            cell34 = richTextBox34.Text;
        }
        private void richTextBox46_TextChanged(object sender, EventArgs e)
        {
            cell46 = richTextBox46.Text;
        }
        private void richTextBox58_TextChanged(object sender, EventArgs e)
        {
            cell58 = richTextBox58.Text;
        }
        private void richTextBox70_TextChanged(object sender, EventArgs e)
        {
            cell70 = richTextBox70.Text;
        }
        private void richTextBox82_TextChanged(object sender, EventArgs e)
        {
            cell82 = richTextBox82.Text;
        }
        private void richTextBox93_TextChanged(object sender, EventArgs e)
        {
            cell93 = richTextBox93.Text;

        }
        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            cell1 = richTextBox1.Text;
        }


        //column  EO
        private void richTextBox23_TextChanged(object sender, EventArgs e)
        {
            cell23 = richTextBox23.Text;
        }

        private void richTextBox35_TextChanged(object sender, EventArgs e)
        {
            cell35 = richTextBox35.Text;
        }
        private void richTextBox47_TextChanged(object sender, EventArgs e)
        {
            cell47 = richTextBox47.Text;
        }
        private void richTextBox59_TextChanged(object sender, EventArgs e)
        {
            cell59 = richTextBox59.Text;
        }
        private void richTextBox71_TextChanged(object sender, EventArgs e)
        {
            cell71 = richTextBox71.Text;
        }
        private void richTextBox83_TextChanged(object sender, EventArgs e)
        {
            cell83 = richTextBox83.Text;
        }
        private void richTextBox94_TextChanged(object sender, EventArgs e)
        {
            cell94 = richTextBox94.Text;

        }
        private void richTextBox6_TextChanged(object sender, EventArgs e)
        {
            cell6 = richTextBox6.Text;
        }


        //column FC 
        private void richTextBox24_TextChanged(object sender, EventArgs e)
        {
            cell24 = richTextBox24.Text;
        }

        private void richTextBox36_TextChanged(object sender, EventArgs e)
        {
            cell36 = richTextBox36.Text;
        }
        private void richTextBox48_TextChanged(object sender, EventArgs e)
        {
            cell48 = richTextBox48.Text;
        }
        private void richTextBox60_TextChanged(object sender, EventArgs e)
        {
            cell60 = richTextBox60.Text;
        }
        private void richTextBox72_TextChanged(object sender, EventArgs e)
        {
            cell72 = richTextBox72.Text;
        }
        private void richTextBox84_TextChanged(object sender, EventArgs e)
        {
            cell84 = richTextBox84.Text;
        }
        private void richTextBox95_TextChanged(object sender, EventArgs e)
        {
            cell95 = richTextBox95.Text;

        }
        private void richTextBox5_TextChanged(object sender, EventArgs e)
        {
            cell5 = richTextBox5.Text;
        }
        //column FN 
        private void richTextBox25_TextChanged(object sender, EventArgs e)
        {
            cell25 = richTextBox25.Text;
        }

        private void richTextBox37_TextChanged(object sender, EventArgs e)
        {
            cell37 = richTextBox37.Text;
        }
        private void richTextBox49_TextChanged(object sender, EventArgs e)
        {
            cell49 = richTextBox49.Text;
        }
        private void richTextBox61_TextChanged(object sender, EventArgs e)
        {
            cell61 = richTextBox61.Text;
        }
        private void richTextBox73_TextChanged(object sender, EventArgs e)
        {
            cell73 = richTextBox73.Text;
        }
        private void richTextBox85_TextChanged(object sender, EventArgs e)
        {
            cell85 = richTextBox85.Text;
        }
        private void richTextBox96_TextChanged(object sender, EventArgs e)
        {
            cell96 = richTextBox96.Text;

        }
        private void richTextBox4_TextChanged(object sender, EventArgs e)
        {
            cell4 = richTextBox4.Text;
        }

















        private void textBox70_TextChanged(object sender, EventArgs e)
        {
            Day = textBox70.Text; 
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox72_TextChanged(object sender, EventArgs e)
        {
            TTN = textBox72.Text; 
        }







        private async void button7_Click(object sender, EventArgs e) // Delivers add
        {
        if (label54.Visible)
                label54.Visible = false;



            if (!string.IsNullOrEmpty(textBox35.Text) && !string.IsNullOrWhiteSpace(textBox35.Text) &&
                !string.IsNullOrEmpty(textBox34.Text) && !string.IsNullOrWhiteSpace(textBox34.Text) &&
                !string.IsNullOrEmpty(textBox33.Text) && !string.IsNullOrWhiteSpace(textBox33.Text) &&
                !string.IsNullOrEmpty(textBox32.Text) && !string.IsNullOrWhiteSpace(textBox32.Text) &&
                !string.IsNullOrEmpty(textBox31.Text) && !string.IsNullOrWhiteSpace(textBox31.Text) &&
                !string.IsNullOrEmpty(textBox30.Text) && !string.IsNullOrWhiteSpace(textBox30.Text) &&
                !string.IsNullOrEmpty(textBox29.Text) && !string.IsNullOrWhiteSpace(textBox29.Text) &&
                !string.IsNullOrEmpty(textBox28.Text) && !string.IsNullOrWhiteSpace(textBox28.Text) &&
                !string.IsNullOrEmpty(textBox27.Text) && !string.IsNullOrWhiteSpace(textBox27.Text) &&
                !string.IsNullOrEmpty(textBox26.Text) && !string.IsNullOrWhiteSpace(textBox26.Text))
            {

                SqlCommand cmd = new SqlCommand("INSERT INTO [Delivers] (Name,Address,Number,CA,CRA,BankIC,Bank,INN,CPP,OKPO)VALUES(@Name, @Address,@Number,@CA,@CRA,@BankIC,@Bank,@INN,@CPP,@OKPO)", sqlConnection);

                cmd.Parameters.AddWithValue("Name", textBox35.Text);
                cmd.Parameters.AddWithValue("Address", textBox34.Text);
                cmd.Parameters.AddWithValue("Number", textBox33.Text);
                cmd.Parameters.AddWithValue("CA", textBox32.Text);
                cmd.Parameters.AddWithValue("CRA", textBox31.Text);
                cmd.Parameters.AddWithValue("BankIC", textBox30.Text);
                cmd.Parameters.AddWithValue("Bank", textBox29.Text);
                cmd.Parameters.AddWithValue("INN", textBox28.Text);
                cmd.Parameters.AddWithValue("CPP", textBox27.Text);
                cmd.Parameters.AddWithValue("OKPO", textBox26.Text);
                await cmd.ExecuteNonQueryAsync();


            }
            else
            {
                label54.Visible = true;
                label54.Text = "Все поля должны быть заполнены!";
            }


        }

        private async void button8_Click(object sender, EventArgs e)
        {


            if (label66.Visible)
                label66.Visible = false;

            if (!string.IsNullOrEmpty(textBox36.Text) && !string.IsNullOrWhiteSpace(textBox36.Text) &&
                !string.IsNullOrEmpty(textBox37.Text) && !string.IsNullOrWhiteSpace(textBox37.Text) &&
                !string.IsNullOrEmpty(textBox38.Text) && !string.IsNullOrWhiteSpace(textBox38.Text) &&
                !string.IsNullOrEmpty(textBox39.Text) && !string.IsNullOrWhiteSpace(textBox39.Text) &&
                !string.IsNullOrEmpty(textBox40.Text) && !string.IsNullOrWhiteSpace(textBox40.Text) &&
                !string.IsNullOrEmpty(textBox41.Text) && !string.IsNullOrWhiteSpace(textBox41.Text) &&
                !string.IsNullOrEmpty(textBox42.Text) && !string.IsNullOrWhiteSpace(textBox42.Text) &&
                !string.IsNullOrEmpty(textBox43.Text) && !string.IsNullOrWhiteSpace(textBox43.Text) &&
                !string.IsNullOrEmpty(textBox44.Text) && !string.IsNullOrWhiteSpace(textBox44.Text) &&
                !string.IsNullOrEmpty(textBox45.Text) && !string.IsNullOrWhiteSpace(textBox45.Text) &&
                !string.IsNullOrEmpty(textBox46.Text) && !string.IsNullOrWhiteSpace(textBox46.Text))
            {
                SqlCommand cmd = new SqlCommand("UPDATE [Customers] SET [Name]=@Name,[Address]=@Address,[Number]=@Number,[CA]=@CA,[CRA]=@CRA,[BankIC]=@BankIC,[Bank]=@Bank,[INN]=@INN,[CPP]=@CPP,[OKPO]=@OKPO WHERE [Id]=@Id", sqlConnection);
                cmd.Parameters.AddWithValue("Id", textBox36.Text);
                cmd.Parameters.AddWithValue("Name", textBox37.Text);
                cmd.Parameters.AddWithValue("Address", textBox38.Text);
                cmd.Parameters.AddWithValue("Number", textBox39.Text);
                cmd.Parameters.AddWithValue("CA", textBox40.Text);
                cmd.Parameters.AddWithValue("CRA", textBox41.Text);
                cmd.Parameters.AddWithValue("BankIC", textBox42.Text);
                cmd.Parameters.AddWithValue("Bank", textBox43.Text);
                cmd.Parameters.AddWithValue("INN", textBox44.Text);
                cmd.Parameters.AddWithValue("CPP", textBox45.Text);
                cmd.Parameters.AddWithValue("OKPO", textBox46.Text);

                await cmd.ExecuteNonQueryAsync();


            }

            else if (!string.IsNullOrEmpty(textBox24.Text) && !string.IsNullOrWhiteSpace(textBox24.Text))
            {

                label66.Visible = true;
                label66.Text = "id должен заполнить!";

            }

            else
            {
                label25.Visible = true;
                label25.Text = "Поля 'имя' и 'цена' должны быть заполнены!";

            }



        }

        private async void button9_Click(object sender, EventArgs e)
        {

            //if (label27.Visible)
            //    label27.Visible = false;


            if (!string.IsNullOrEmpty(textBox47.Text) && !string.IsNullOrWhiteSpace(textBox47.Text))
            {
                SqlCommand cmd = new SqlCommand("DELETE FROM [Delivers] WHERE [Id] = @id", sqlConnection);
                cmd.Parameters.AddWithValue("id", textBox47.Text);
                await cmd.ExecuteNonQueryAsync();

            }
            //else
            //{
            //    label27.Visible = true;
            //    label27.Text = "Id должен быть заполнен!";
            //}

            MessageBox.Show("Удалено"); 



            
        }





        private async void button10_Click(object sender, EventArgs e)
        {


            if (label79.Visible)
                label79.Visible = false;



            if (!string.IsNullOrEmpty(textBox57.Text) && !string.IsNullOrWhiteSpace(textBox57.Text) &&
                !string.IsNullOrEmpty(textBox56.Text) && !string.IsNullOrWhiteSpace(textBox56.Text) &&
                !string.IsNullOrEmpty(textBox55.Text) && !string.IsNullOrWhiteSpace(textBox55.Text) &&
                !string.IsNullOrEmpty(textBox54.Text) && !string.IsNullOrWhiteSpace(textBox54.Text) &&
                !string.IsNullOrEmpty(textBox53.Text) && !string.IsNullOrWhiteSpace(textBox53.Text) &&
                !string.IsNullOrEmpty(textBox52.Text) && !string.IsNullOrWhiteSpace(textBox52.Text) &&
                !string.IsNullOrEmpty(textBox51.Text) && !string.IsNullOrWhiteSpace(textBox51.Text) &&
                !string.IsNullOrEmpty(textBox50.Text) && !string.IsNullOrWhiteSpace(textBox50.Text) &&
                !string.IsNullOrEmpty(textBox49.Text) && !string.IsNullOrWhiteSpace(textBox49.Text) &&
                !string.IsNullOrEmpty(textBox48.Text) && !string.IsNullOrWhiteSpace(textBox48.Text))
            {

                SqlCommand cmd2 = new SqlCommand("INSERT INTO [Buyers] (Name,Address,Number,CA,CRA,BankIC,Bank,INN,CPP,OKPO)VALUES(@Name, @Address,@Number,@CA,@CRA,@BankIC,@Bank,@INN,@CPP,@OKPO)", sqlConnection);

                cmd2.Parameters.AddWithValue("Name", textBox57.Text);
                cmd2.Parameters.AddWithValue("Address", textBox56.Text);
                cmd2.Parameters.AddWithValue("Number", textBox55.Text);
                cmd2.Parameters.AddWithValue("CA", textBox54.Text);
                cmd2.Parameters.AddWithValue("CRA", textBox53.Text);
                cmd2.Parameters.AddWithValue("BankIC", textBox52.Text);
                cmd2.Parameters.AddWithValue("Bank", textBox51.Text);
                cmd2.Parameters.AddWithValue("INN", textBox50.Text);
                cmd2.Parameters.AddWithValue("CPP", textBox49.Text);
                cmd2.Parameters.AddWithValue("OKPO", textBox48.Text);
                await cmd2.ExecuteNonQueryAsync();


            }
            else
            {
                label79.Visible = true;
                label79.Text = "Все поля должны быть заполнены!";
            }







        }

        private async void button11_Click(object sender, EventArgs e)
        {

            if (label91.Visible)
                label91.Visible = false;

            if (!string.IsNullOrEmpty(textBox58.Text) && !string.IsNullOrWhiteSpace(textBox58.Text) &&
                !string.IsNullOrEmpty(textBox59.Text) && !string.IsNullOrWhiteSpace(textBox59.Text) &&
                !string.IsNullOrEmpty(textBox60.Text) && !string.IsNullOrWhiteSpace(textBox60.Text) &&
                !string.IsNullOrEmpty(textBox61.Text) && !string.IsNullOrWhiteSpace(textBox61.Text) &&
                !string.IsNullOrEmpty(textBox62.Text) && !string.IsNullOrWhiteSpace(textBox62.Text) &&
                !string.IsNullOrEmpty(textBox63.Text) && !string.IsNullOrWhiteSpace(textBox63.Text) &&
                !string.IsNullOrEmpty(textBox64.Text) && !string.IsNullOrWhiteSpace(textBox64.Text) &&
                !string.IsNullOrEmpty(textBox65.Text) && !string.IsNullOrWhiteSpace(textBox65.Text) &&
                !string.IsNullOrEmpty(textBox66.Text) && !string.IsNullOrWhiteSpace(textBox66.Text) &&
                !string.IsNullOrEmpty(textBox67.Text) && !string.IsNullOrWhiteSpace(textBox67.Text) &&
                !string.IsNullOrEmpty(textBox68.Text) && !string.IsNullOrWhiteSpace(textBox68.Text))
            {
                SqlCommand cmd2 = new SqlCommand("UPDATE [Buyers] SET [Name]=@Name,[Address]=@Address,[Number]=@Number,[CA]=@CA,[CRA]=@CRA,[BankIC]=@BankIC,[Bank]=@Bank,[INN]=@INN,[CPP]=@CPP,[OKPO]=@OKPO WHERE [Id]=@Id", sqlConnection);
                cmd2.Parameters.AddWithValue("Id", textBox58.Text);
                cmd2.Parameters.AddWithValue("Name", textBox59.Text);
                cmd2.Parameters.AddWithValue("Address", textBox60.Text);
                cmd2.Parameters.AddWithValue("Number", textBox61.Text);
                cmd2.Parameters.AddWithValue("CA", textBox62.Text);
                cmd2.Parameters.AddWithValue("CRA", textBox63.Text);
                cmd2.Parameters.AddWithValue("BankIC", textBox64.Text);
                cmd2.Parameters.AddWithValue("Bank", textBox65.Text);
                cmd2.Parameters.AddWithValue("INN", textBox66.Text);
                cmd2.Parameters.AddWithValue("CPP", textBox67.Text);
                cmd2.Parameters.AddWithValue("OKPO", textBox68.Text);

                await cmd2.ExecuteNonQueryAsync();


            }

            else if (!string.IsNullOrEmpty(textBox24.Text) && !string.IsNullOrWhiteSpace(textBox24.Text))
            {

                label91.Visible = true;
                label91.Text = "id должен заполнить!";

            }

            else
            {
                label91.Visible = true;
                label92.Text = "Поля 'имя' и 'цена' должны быть заполнены!";

            }

        }

        private async void button12_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox69.Text) && !string.IsNullOrWhiteSpace(textBox69.Text))
            {
                SqlCommand cmd2 = new SqlCommand("DELETE FROM [Buyers] WHERE [Id] = @id", sqlConnection);
                cmd2.Parameters.AddWithValue("id", textBox69.Text);
                await cmd2.ExecuteNonQueryAsync();

            }
            //else
            //{
            //    label27.Visible = true;
            //    label27.Text = "Id должен быть заполнен!";
            //}

            MessageBox.Show("Удалено"); 
        }

        private void richTextBox97_TextChanged(object sender, EventArgs e)
        {
            driver = richTextBox97.Text; 

        }

        private void textBox71_TextChanged(object sender, EventArgs e)
        {
            Mounth = textBox71.Text; 
        }

        private void textBox73_TextChanged(object sender, EventArgs e)
        {
            Year = textBox73.Text;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
             Dropdone = comboBox1.Text; 
        }

        private void richTextBox98_TextChanged(object sender, EventArgs e)
        {
            Dropdonename = richTextBox98.Text; 
        }

        private void label110_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox99_TextChanged(object sender, EventArgs e)
        {
            Dropallowname = richTextBox99.Text; 
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)

        {
            Dropallowed  = comboBox2.Text; 
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            buhgalter   = comboBox3.Text; 
        }





        private void label28_Click(object sender, EventArgs e)
        {



        }

        private void label53_Click(object sender, EventArgs e)
        {

        }

        private void label114_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox100_TextChanged(object sender, EventArgs e)
        {
            company = richTextBox100.Text; 
        }

        private void richTextBox101_TextChanged(object sender, EventArgs e)
        {
            marka = richTextBox101.Text; 
        }

        private void richTextBox102_TextChanged(object sender, EventArgs e)
        {
            number = richTextBox102.Text; 
        }

        private void textBox74_TextChanged(object sender, EventArgs e)
        {
            names = textBox74.Text; 
        }

        private void textBox75_TextChanged(object sender, EventArgs e)
        {
            places = textBox75.Text; 
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            priem = comboBox4.Text; 
        }

        private void richTextBox103_TextChanged(object sender, EventArgs e)
        {
            sda4a = richTextBox103.Text; 
        }

        private void textBox76_TextChanged(object sender, EventArgs e)
        {
            DayDostavki = textBox76.Text; 
        }

        private void textBox77_TextChanged(object sender, EventArgs e)
        {
            MounthDostavki = textBox77.Text; 
        }

        private void textBox78_TextChanged(object sender, EventArgs e)
        {
            YearDostavki = textBox78.Text; 
        }

        private void richTextBox104_TextChanged(object sender, EventArgs e)
        {
            marshrut = richTextBox104.Text; 
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            pogryzka = comboBox5.Text; 
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            names = comboBox7.Text; 
        }

        private void label127_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox105_TextChanged(object sender, EventArgs e)
        {
            propis1 = richTextBox105.Text; 
        }

        private void richTextBox106_TextChanged(object sender, EventArgs e)
        {
            propis2 = richTextBox106.Text; 
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        





    }
}


    class Excel
    {


        string path = " ";
        _Application excel = new _Excel.Application();
      public  Workbook wb;
        Worksheet ws;
        Worksheet ws1;
        Worksheet ws2;
        Worksheet ws3; 

        public Excel(string path, int Sheet)
        {

            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[1];
            ws1 = wb.Worksheets[2];
            ws2 = wb.Worksheets[3];
            ws3 = wb.Worksheets[4]; 
        }





        public void WriteToCell(string buf, string buf1, string buf2, string cell19, string cell31, string cell43, string cell55, string cell67, string cell79, string cell90, string cell7, string Dropallowed, string Dropdone, string Dropallowname, string Dropdonename, string buhgalter, string driver, string company, string marka, string number, string Day, string Year, string Mounth, string names, string places, string priem, string sda4a, string DayDostavki, string MounthDostavki, string YearDostavki, string marshrut, string pogryzka, string propis1, string propis2 )      //  string cell19, string cell31, string cell43, string cell55, string cell67, string cell79, string cell90, string cell7
        {


        
            ws1.Range["A47"].Value = Dropallowed +  "   " + Dropallowname;
            ws.Range["A55"].Value = Dropdone + "     " +  Dropdonename;
            ws2.Range["A39"].Value = Dropallowed;
            ws2.Range["AG39"].Value = Dropallowname;
            ws2.Range["CC39"].Value = buhgalter;
            ws2.Range["AB42"].Value = Dropdone;
            ws2.Range["BR42"].Value = Dropdonename; 
            ws3.Range["I31"].Value = Dropdone;
            ws3.Range["AH31"].Value = Dropdonename; 



        
            //водитель 

            ws.Range["A56"].Value = "Водитель" + "     " + driver;
            ws.Range["BD56"].Value = "Водитель" + "     " + driver;
            ws.Range["BD19"].Value = driver;
            ws1.Range["BD5"].Value = driver;
            ws1.Range["BD47"].Value = "Водитель" + "     " + driver;
            ws3.Range["L9"].Value = driver;
            ws3.Range["AA36"].Value = driver;
            ws2.Range["FI36"].Value = driver;
            ws2.Range["EB36"].Value = "Водитель";
            ws3.Range["DC31"].Value = driver; 


            ws1.Range["A5"].Value = company;
            ws1.Range["A13"].Value = marka;
            ws1.Range["BP13"].Value = number;

            ws3.Range["N4"].Value = company; 
            ws3.Range["CO4"].Value = marka;
            ws3.Range["EL4"].Value = number; 


            //

            //data


            ws.Range["A47"].Value = Day + "." + Mounth + "." + Year;
            ws.Range["BI8"].Value = Day + "." + Mounth + "." + Year;
            ws.Range["A72"].Value = Day + "." + Mounth + "." + Year + "                          " + driver;
            ws2.Range["FM7"].Value = Day;
            ws2.Range["FS7"].Value = Mounth;
            ws2.Range["FZ7"].Value = Year;



            ws2.Range["BE44"].Value = Day;
           
            ws2.Range["BM44"].Value = Mounth;

            ws2.Range["CG44"].Value = Year;


            ws3.Range["X3"].Value = DayDostavki;
            ws3.Range["AD3"].Value = MounthDostavki;
            ws3.Range["AW3"].Value = YearDostavki; 





            //

            
            ws.Range["A13"].Value = buf; 
            ws2.Range["V9"].Value = buf;
            ws2.Range["V11"].Value = buf1;
            ws.Range["BD13"].Value = buf1;
            ws2.Range["V13"].Value = buf2;
           
          ws.Range["A22"].Value = cell19 + "  " + cell31 + "  " + cell43 + " " + cell55 + " " + cell67 + " " + cell79 + " " + cell90 + " " + cell7; 
            
          //k-vo mest
          ws2.Range["W29"].Value = names; 
          ws2.Range["M31"].Value = places;
            
            //priem sda4a gryza
          ws.Range["A45"].Value = priem;
          ws.Range["BD45"].Value = sda4a;
          ws3.Range["CT14"].Value = sda4a;
          ws3.Range["Q14"].Value = pogryzka; 

           // marshrut
          ws3.Range["FP12"].Value = marshrut;

          ws1.Range["A24"].Value = propis1;
          ws1.Range["A26"].Value = propis2; 

        }
        public void WriteToColumnA( string cell14, string cell26, string cell38,string cell50,string cell62,string cell74,string cell13,string cell9 )
        {

            ws2.Range["A18"].Value = cell14; 
            ws2.Range["A19"].Value = cell26; 
            ws2.Range["A20"].Value = cell38; 
            ws2.Range["A21"].Value = cell50;
            ws2.Range["A22"].Value = cell62;
            ws2.Range["A23"].Value = cell74;
            ws2.Range["A24"].Value = cell13;
            ws2.Range["A25"].Value = cell9; 
        }
        public void WriteToColumnL(string cell15, string cell27, string cell39, string cell51, string cell63, string cell75, string cell86, string cell8)
        {
            ws2.Range["L18"].Value = cell15;
            ws2.Range["L19"].Value = cell27;
            ws2.Range["L20"].Value = cell39;
            ws2.Range["L21"].Value = cell51;
            ws2.Range["L22"].Value = cell63;
            ws2.Range["L23"].Value = cell75;
            ws2.Range["L24"].Value = cell86;
            ws2.Range["L25"].Value = cell8;
        }

        public void WriteToColumnAF(string cell16, string cell28, string cell40, string cell52, string cell64, string cell76, string cell87, string cell10)
        {
            ws2.Range["AF18"].Value = cell16;
            ws2.Range["AF19"].Value = cell28;
            ws2.Range["AF20"].Value = cell40;
            ws2.Range["AF21"].Value = cell52;
            ws2.Range["AF22"].Value = cell64;
            ws2.Range["AF23"].Value = cell76;
            ws2.Range["AF24"].Value = cell87;
            ws2.Range["AF25"].Value = cell10;

        }
        public void WriteToColumnAS(string cell17, string cell29, string cell41, string cell53, string cell65, string cell77, string cell88, string cell12)
        {
            ws2.Range["AS18"].Value = cell17;
            ws2.Range["AS19"].Value = cell29;
            ws2.Range["AS20"].Value = cell41;
            ws2.Range["AS21"].Value = cell53;
            ws2.Range["AS22"].Value = cell65;
            ws2.Range["AS23"].Value = cell77;
            ws2.Range["AS24"].Value = cell88;
            ws2.Range["AS25"].Value = cell12;

        }
        public void WriteToColumnBE(string cell18, string cell30, string cell42, string cell54, string cell66, string cell78, string cell89, string cell11)
    {
            ws2.Range["BE18"].Value = cell18;
            ws2.Range["BE19"].Value = cell30;
            ws2.Range["BE20"].Value = cell42;
            ws2.Range["BE21"].Value = cell54;
            ws2.Range["BE22"].Value = cell66;
            ws2.Range["BE23"].Value = cell78;
            ws2.Range["BE24"].Value = cell89;
            ws2.Range["BE25"].Value = cell11;

    }

        public void WriteToColumnBU(string cell19, string cell31, string cell43, string cell55, string cell67, string cell79, string cell90, string cell7)
        {
            ws2.Range["BU18"].Value = cell19;
            ws2.Range["BU19"].Value = cell31;
            ws2.Range["BU20"].Value = cell43;
            ws2.Range["BU21"].Value = cell55;
            ws2.Range["BU22"].Value = cell67;
            ws2.Range["BU23"].Value = cell79;
            ws2.Range["BU24"].Value = cell90;
            ws2.Range["BU25"].Value = cell7;

            ws3.Range["D23"].Value = cell19;
            ws3.Range["D24"].Value = cell31;
            ws3.Range["D25"].Value = cell43;
            //ws3.Range["D26"].Value = cell55;
            //ws3.Range["D27"].Value = cell67;
            //ws3.Range["D28"].Value = cell79; 


        }
        public void WriteToColumnDH(string cell20, string cell32, string cell44, string cell56, string cell68, string cell80, string cell91, string cell3)
        {
            ws2.Range["DH18"].Value = cell20;
            ws2.Range["DH19"].Value = cell32;
            ws2.Range["DH20"].Value = cell44;
            ws2.Range["DH21"].Value = cell56;
            ws2.Range["DH22"].Value = cell68;
            ws2.Range["DH23"].Value = cell80;
            ws2.Range["DH24"].Value = cell91;
            ws2.Range["DH25"].Value = cell3;

        }
        public void WriteToColumnDS(string cell21, string cell33, string cell45, string cell57, string cell69, string cell81, string cell92, string cell2)
        {
            ws2.Range["DS18"].Value = cell21;
            ws2.Range["DS19"].Value = cell33;
            ws2.Range["DS20"].Value = cell45;
            ws2.Range["DS21"].Value = cell57;
            ws2.Range["DS22"].Value = cell69;
            ws2.Range["DS23"].Value = cell81;
            ws2.Range["DS24"].Value = cell92;
            ws2.Range["DS25"].Value = cell2;

            //
            ws3.Range["BU23"].Value = cell21;
            ws3.Range["BU24"].Value = cell33;
            ws3.Range["BU25"].Value = cell45; 
        }
        public void WriteToColumnED(string cell22, string cell34, string cell46, string cell58, string cell70, string cell82, string cell90, string cell1)
        {
            ws2.Range["ED18"].Value = cell22;
            ws2.Range["ED19"].Value = cell34;
            ws2.Range["ED20"].Value = cell46;
            ws2.Range["ED21"].Value = cell58;
            ws2.Range["ED22"].Value = cell70;
            ws2.Range["ED23"].Value = cell82;
            ws2.Range["ED24"].Value = cell90;
            ws2.Range["ED25"].Value = cell1;

            //

            ws3.Range["CM23"].Value = cell22;
            ws3.Range["CM24"].Value = cell34;
            ws3.Range["CM25"].Value = cell46; 

        }
        public void WriteToColumnEO(string cell23, string cell35, string cell47, string cell59, string cell71, string cell83, string cell91, string cell6)
        {
            ws2.Range["EO18"].Value = cell23;
            ws2.Range["EO19"].Value = cell35;
            ws2.Range["EO20"].Value = cell47;
            ws2.Range["EO21"].Value = cell59;
            ws2.Range["EO22"].Value = cell71;
            ws2.Range["EO23"].Value = cell83;
            ws2.Range["EO24"].Value = cell91;
            ws2.Range["EO25"].Value = cell6;

            double abc = Convert.ToDouble(cell23) + Convert.ToDouble(cell35) + Convert.ToDouble(cell47) + Convert.ToDouble(cell59) + Convert.ToDouble(cell71) + Convert.ToDouble(cell83) + Convert.ToDouble(cell91) + Convert.ToDouble(cell6);


            ws2.Range["DS28"].Value = abc;
            ws.Range["A26"].Value = abc; 
        }
        public void WriteToColumnFC(string cell24, string cell36, string cell48, string cell60, string cell72, string cell84, string cell92, string cell5)
        {
            ws2.Range["FC18"].Value = cell24;
            ws2.Range["FC19"].Value = cell36;
            ws2.Range["FC20"].Value = cell48;
            ws2.Range["FC21"].Value = cell60;
            ws2.Range["FC22"].Value = cell72;
            ws2.Range["FC23"].Value = cell84;
            ws2.Range["FC24"].Value = cell92;
            ws2.Range["FC25"].Value = cell5;

        }
        public void WriteToColumnFN(string cell25, string cell37, string cell49, string cell61, string cell73, string cell85, string cell93, string cell4)
        {
            ws2.Range["FN18"].Value = cell25;
            ws2.Range["FN19"].Value = cell37;
            ws2.Range["FN20"].Value = cell49;
            ws2.Range["FN21"].Value = cell61;
            ws2.Range["FN22"].Value = cell73;
            ws2.Range["FN23"].Value = cell85;
            ws2.Range["FN24"].Value = cell93;
            ws2.Range["FN25"].Value = cell4;

        }

        public void WriteToTTN(string TTN)
    {
        ws.Range["CM8"].Value = TTN;
        ws2.Range["FM6"].Value = TTN;
        ws3.Range["FP2"].Value = TTN; 


    }

    //    public void WriteDriver(string driver)
    //{
      

    //}

        public void WriteToData (string Day, string Mounth, string Year, string driver)
        {
           
    }

       


        //public void Save()
        //{
        //    wb.Save();
        //}

        //public void SaveAs(string path)
        //{
        //    wb.SaveAs(path);
        //}
        public void Close()
        {
            wb.Close();
        }
    }
