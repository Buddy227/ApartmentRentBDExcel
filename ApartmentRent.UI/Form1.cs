using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Security.Cryptography;
using System.Text;
using System.IO;
using System.Reflection;
using System.Collections.Generic;

namespace ApartmentRent.UI
{
    public partial class Form1 : Form
    {
        public Form1(string CurrentUser)
        {
            InitializeComponent();
            usersDataGridView.Hide();
            this.CurrentUser = CurrentUser;
            if (CurrentUser == "admin")
            {
                tabControl1.Enabled = true;
                button8.Enabled = true;
                buttonExcel.Enabled = true;
            }
            else
            {
                tabControl1.Hide();
                buttonExcel.Hide();
                button8.Hide();
                label4.Hide();
                Width = 675;

            }
            label3.Text = "Вы вошли как " + CurrentUser;
        }

        string CurrentUser { get; set; }

        RentRequestDto GetModelFromUI()
        {
            return new RentRequestDto()
            {
                Filled = dateTimePicker1.Value,
                ApartmentDescriptions = listBox1.Items.OfType<ApartmentDescription>().ToList(),

            };
        }
        private void SetModelToUI(RentRequestDto dto)
        {
            button4.Enabled = false;
            dateTimePicker1.Value = dto.Filled;
            listBox1.Items.Clear();
            foreach (var e in dto.ApartmentDescriptions)
            {
                listBox1.Items.Add(e);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            var sfd = new SaveFileDialog() { Filter = "Файлы|*.apartmentrent" };
            var result = sfd.ShowDialog(this);
            if (result == DialogResult.OK)
            {
                var dto = GetModelFromUI();
                RentDtoHelper.WriteToFile(sfd.FileName, dto);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog() { Filter = "Файлы|*.apartmentrent" };
            var result = ofd.ShowDialog(this);
            if (result == DialogResult.OK)
            {
                var dto = RentDtoHelper.LoadFromFile(ofd.FileName);
                SetModelToUI(dto);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var form = new ApartmentDescriptionForm2(new ApartmentDescription());
            var res = form.ShowDialog(this);
            if (res == DialogResult.OK)
            {
                listBox1.Items.Add(form.ad);
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            var ad = listBox1.SelectedItem as ApartmentDescription;
            if (ad == null)
                return;
            var form = new ApartmentDescriptionForm2(ad.Clone());
            var res = form.ShowDialog(this);
            if (res == DialogResult.OK)
            {
                var si = listBox1.SelectedIndex;
                listBox1.Items.Remove(si);
                listBox1.Items.Insert(si, form.ad);
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            var si = listBox1.SelectedIndex;
            listBox1.Items.RemoveAt(si);

        }

        private void listBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex >= 0)
                button4.Enabled = true;
            else button4.Enabled = false;
        }

        public static String encrypt(String input)
        {
            MD5 md5Hasher = MD5.Create();
            byte[] data = md5Hasher.ComputeHash(Encoding.Default.GetBytes(input));
            StringBuilder sBuilder = new StringBuilder();

            for (int i = 0; i < data.Length; i++)
            {
                sBuilder.Append(data[i].ToString("x2"));
            }
            return sBuilder.ToString();

        }


        SqlConnection SqlConnection;

        private async void Form1_Load(object sender, EventArgs e)
        {
            //string dirPath = new FileInfo($"{Assembly.GetAssembly(GetType()).Location}").DirectoryName;
            //string dbName = "Database1.mdf";
            //string connectionString = $@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename={dirPath}\{dbName};Integrated Security=True";
            string path = System.IO.Directory.GetCurrentDirectory();
            string path1 = System.IO.Directory.GetParent(path).ToString();
            string path2 = System.IO.Directory.GetParent(path1).ToString();
            string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=" + path2 + @"\Database1.mdf;Integrated Security=True";

            // TODO: данная строка кода позволяет загрузить данные в таблицу "database1DataSet.Users". При необходимости она может быть перемещена или удалена.
            this.usersTableAdapter.Fill(database1DataSet.Users);
            label20.Text = "заполните поля";

            SqlConnection = new SqlConnection(connectionString);

            await SqlConnection.OpenAsync();

            SqlDataReader sqlReader = null;

            SqlCommand command = new SqlCommand("SELECT * FROM [Users]", SqlConnection);

            try
            {
                sqlReader = await command.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    listBox2.Items.Add("id" + Convert.ToString(sqlReader["Id"])
                        + " | " + Convert.ToString(sqlReader["login"])
                        + " | " + Convert.ToString(sqlReader["fullname"])
                        + " | " + Convert.ToString(sqlReader["phone"])
                        + " | " + Convert.ToString(sqlReader["Company"]));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                {
                    sqlReader.Close();
                }
            }
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (SqlConnection != null && SqlConnection.State != ConnectionState.Closed)
                SqlConnection.Close();
            Application.Exit();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (SqlConnection != null && SqlConnection.State != ConnectionState.Closed)
                SqlConnection.Close();
            Application.Exit();
        }

        private async void buttonAdd_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrWhiteSpace(textBox1.Text) &&
                !string.IsNullOrEmpty(textBox2.Text) && !string.IsNullOrWhiteSpace(textBox2.Text) &&
                !string.IsNullOrEmpty(textBox3.Text) && !string.IsNullOrWhiteSpace(textBox3.Text) &&
                !string.IsNullOrEmpty(textBox4.Text) && !string.IsNullOrWhiteSpace(textBox4.Text) &&
                !string.IsNullOrEmpty(textBox5.Text) && !string.IsNullOrWhiteSpace(textBox5.Text))
            {
                SqlCommand command = new SqlCommand("INSERT INTO [Users] (login, password, fullname, phone, Company)VALUES(@login, @password, @fullname, @phone, @Company)", SqlConnection);

                command.Parameters.AddWithValue("login", textBox1.Text);
                command.Parameters.AddWithValue("password", encrypt(textBox2.Text));
                command.Parameters.AddWithValue("fullname", textBox3.Text);
                command.Parameters.AddWithValue("phone", textBox4.Text);
                command.Parameters.AddWithValue("Company", textBox5.Text);

                await command.ExecuteNonQueryAsync();
                label20.Visible = false;
            }
            else
            {
                label20.Visible = true;
                label20.Text = "Поля должны быть заполнены";
            }
        }

        private async void button8_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();

            SqlDataReader sqlReader = null;
            SqlDataAdapter adapter = new SqlDataAdapter("SELECT * FROM [Users]", SqlConnection);
            SqlCommand command = new SqlCommand("SELECT * FROM [Users]", SqlConnection);
            DataSet ds = new DataSet();
            adapter.Fill(ds);
            usersDataGridView.DataSource = ds.Tables[0];
            try
            {
                sqlReader = await command.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    listBox2.Items.Add("id" + Convert.ToString(sqlReader["Id"])
                        + " | " + Convert.ToString(sqlReader["login"])
                        + " | " + Convert.ToString(sqlReader["fullname"])
                        + " | " + Convert.ToString(sqlReader["phone"])
                        + " | " + Convert.ToString(sqlReader["Company"]));

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                {
                    sqlReader.Close();
                }
            }
        }

        private async void buttonChange_Click(object sender, EventArgs e)
        {
            if (textBox11.Text == "1")
            {
                label21.Text = "Ошибка, нельзя исправить админа!";
                MessageBox.Show("Ты хочешь изменить админа, админ :{");
                return;
            }

            if (label21.Visible)
                label21.Visible = false;

            if (!string.IsNullOrEmpty(textBox6.Text) && !string.IsNullOrWhiteSpace(textBox6.Text) &&
                !string.IsNullOrEmpty(textBox7.Text) && !string.IsNullOrWhiteSpace(textBox7.Text) &&
                !string.IsNullOrEmpty(textBox8.Text) && !string.IsNullOrWhiteSpace(textBox8.Text) &&
                !string.IsNullOrEmpty(textBox9.Text) && !string.IsNullOrWhiteSpace(textBox9.Text) &&
                !string.IsNullOrEmpty(textBox10.Text) && !string.IsNullOrWhiteSpace(textBox10.Text) &&
                !string.IsNullOrEmpty(textBox11.Text) && !string.IsNullOrWhiteSpace(textBox11.Text))
            {
                SqlCommand command = new SqlCommand("UPDATE [Users] SET [login]=@login, [password]=@password, [fullname]=@fullname, [phone]=@phone, [Company]=@Company WHERE [Id]=@Id", SqlConnection);
                command.Parameters.AddWithValue("Company", textBox6.Text);
                command.Parameters.AddWithValue("phone", textBox7.Text);
                command.Parameters.AddWithValue("fullname", textBox8.Text);
                command.Parameters.AddWithValue("password", encrypt(textBox9.Text));
                command.Parameters.AddWithValue("login", textBox10.Text);
                command.Parameters.AddWithValue("Id", textBox11.Text);

                await command.ExecuteNonQueryAsync();
                label21.Visible = false;

            }
            else if (!string.IsNullOrEmpty(textBox11.Text) && !string.IsNullOrWhiteSpace(textBox11.Text))
            {
                label21.Visible = true;
                label21.Text = "Заполните изменения";
            }
            else
            {
                label21.Visible = true;

                label21.Text = "id должен быть заполнен";
            }
        }

        private async void buttonDelete_Click(object sender, EventArgs e)
        {
            if (textBox12.Text == "1")
            {
                MessageBox.Show("Ты хочешь удалить админа :{");
                return;
            }
            if (!string.IsNullOrEmpty(textBox12.Text) && !string.IsNullOrWhiteSpace(textBox12.Text))
            {
                SqlCommand command = new SqlCommand("DELETE FROM [Users] WHERE [Id]=@Id", SqlConnection);

                command.Parameters.AddWithValue("Id", textBox12.Text);
                await command.ExecuteNonQueryAsync();
            }

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Программа сбора данных\n Разработал: Февронин С. Ри-260005");
        }
        private void buttonExcel_Click(object sender, EventArgs e)
        {

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            //Книга.
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            //Таблица
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            ExcelApp.get_Range("A1").Value = "Id";
            ExcelApp.get_Range("B1").Value = "login";
            ExcelApp.get_Range("C1").Value = "password";
            ExcelApp.get_Range("D1").Value = "fullname";
            ExcelApp.get_Range("E1").Value = "phone";
            ExcelApp.get_Range("F1").Value = "Company";
            ExcelApp.get_Range("G1").Value = "Дата создания документа: " + DateTime.Now.ToShortDateString();

            for (int i = 1; i < usersDataGridView.Rows.Count; i++)
            {
                for (int j = 0; j < usersDataGridView.ColumnCount; j++)
                {
                    if (usersDataGridView.Rows[i].Cells[j].GetType().Name == "DataGridViewComboBoxCell")
                    {
                        DataGridViewComboBoxCell dgvcbc = new DataGridViewComboBoxCell();
                        dgvcbc = (DataGridViewComboBoxCell)usersDataGridView.Rows[i].Cells[j];
                        ExcelApp.Cells[i + 1, j + 1] = dgvcbc.EditedFormattedValue;
                    }
                    else
                    {
                        ExcelApp.Cells[i + 1, j + 1] = usersDataGridView.Rows[i].Cells[j].Value;
                    }
                }
            }
            //Вызываем нашу созданную эксельку.
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void usersBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.usersBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.database1DataSet);

        }

        private void usersBindingNavigatorSaveItem_Click_1(object sender, EventArgs e)
        {
            this.Validate();
            this.usersBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.database1DataSet);

        }

        private static Microsoft.Office.Interop.Word.Application wdApp = new Microsoft.Office.Interop.Word.Application();

        static void WordExport(RentRequestDto ad)
        {
            wdApp.Documents.Add();
            Microsoft.Office.Interop.Word.Document doc = wdApp.Documents[1];
            foreach (var e in ad.ApartmentDescriptions)
            {
                wdApp.Selection.TypeText("ФИО получателя: " + e.FullName);
                wdApp.Selection.TypeText(" Телефон: " + e.Phone);
                wdApp.Selection.TypeText(" Адрес: " + e.Address);
                wdApp.Selection.TypeText(" Описание квартиры: " + e.Description);
                wdApp.Selection.TypeText(" Стоимость аренды: " + e.currency);
                wdApp.Selection.TypeText(" Количество комнат: " + e.rooms);
                wdApp.Selection.TypeText(" Есть ли балкон: " + e.balcony);
                wdApp.Selection.TypeText(" Наличие ванны: " + e.bath);
                wdApp.Selection.TypeText(" Наличие душа: " + e.shower);
                wdApp.Selection.TypeText(" Тип жилья: " + e.Type);
                wdApp.Selection.TypeText(" Площадь квартиры: " + e.square + "\n");

            }
            wdApp.Visible = true;
            wdApp = new Microsoft.Office.Interop.Word.Application();
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            var dto = GetModelFromUI();
            WordExport(dto);
        }

    }
}