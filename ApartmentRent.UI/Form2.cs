using System;
using System.Data.SqlClient;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace ApartmentRent.UI
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string login = textBox1.Text;
            string password = Form1.encrypt(textBox2.Text);

            //string dirPath = new FileInfo($"{Assembly.GetAssembly(GetType()).Location}").DirectoryName;
            //string dbName = "Database1.mdf";
            //string connectionString = $@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename={dirPath}\{dbName};Integrated Security=True";

            //string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\Home\Desktop\ApartmentRentFIX - DB + Authorization\ApartmentRent.UI\Database1.mdf;Integrated Security=True";

            string path = System.IO.Directory.GetCurrentDirectory();
            string path1 = System.IO.Directory.GetParent(path).ToString();
            string path2 = System.IO.Directory.GetParent(path1).ToString();

            string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=" + path2 + @"\Database1.mdf;Integrated Security=True";
            SqlConnection SqlConnection = new SqlConnection(connectionString);

            SqlConnection.Open();

            SqlCommand command = new SqlCommand("SELECT * FROM [Users] WHERE login=@login and password=@password", SqlConnection);

            SqlDataReader sqlReader = null;

            try
            {
                command.Parameters.AddWithValue("@login", login);
                command.Parameters.AddWithValue("@password", password);
                sqlReader = command.ExecuteReader();

                if (sqlReader.Read())
                {
                    MessageBox.Show("Вы успешно зашли");
                    sqlReader.Close();
                    this.Hide();
                    Form1 form1 = new Form1(textBox1.Text);
                    form1.ShowDialog();
                    
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
                    MessageBox.Show("Ошибка. Неверный логин или пароль");
                    sqlReader.Close();
                }
            }
        }
    }
}
