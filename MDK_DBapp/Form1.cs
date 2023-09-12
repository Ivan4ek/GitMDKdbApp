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
using System.IO;

namespace MDK_DBapp
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Путь к файлу базы данных Access
            string dbFilePath = @"C:\Users\Sterben\Desktop\dir\database.accdb";

            // Создаем подключение к базе данных
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbFilePath};";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    // Создаем команду для создания таблицы
                    string createTableQuery = "CREATE TABLE MyTable (Image1 OLEObject, Image2 OLEObject, TextColumn Text);";
                    using (OleDbCommand createTableCommand = new OleDbCommand(createTableQuery, connection))
                    {
                        createTableCommand.ExecuteNonQuery();
                    }

                    // Вставляем данные в таблицу
                    string insertDataQuery = "INSERT INTO MyTable (Image1, Image2, TextColumn) VALUES (?, ?, ?);";
                    using (OleDbCommand insertDataCommand = new OleDbCommand(insertDataQuery, connection))
                    {
                        // Параметры для вставки данных
                        insertDataCommand.Parameters.AddWithValue("Image1", @"C:\Users\Sterben\Desktop\dir\1.PNG");
                        insertDataCommand.Parameters.AddWithValue("Image2", @"C:\Users\Sterben\Desktop\dir\2.PNG");
                        insertDataCommand.Parameters.AddWithValue("TextColumn", "1");
                        insertDataCommand.ExecuteNonQuery();

                        insertDataCommand.Parameters.Clear();

                        insertDataCommand.Parameters.AddWithValue("Image1", @"C:\Users\Sterben\Desktop\dir\3.PNG");
                        insertDataCommand.Parameters.AddWithValue("Image2", @"C:\Users\Sterben\Desktop\dir\4.PNG");
                        insertDataCommand.Parameters.AddWithValue("TextColumn", "2");
                        insertDataCommand.ExecuteNonQuery();

                        insertDataCommand.Parameters.Clear();

                        // Повторите этот блок кода для оставшихся строк
                        insertDataCommand.Parameters.AddWithValue("Image1", @"C:\Users\Sterben\Desktop\dir\5.PNG");
                        insertDataCommand.Parameters.AddWithValue("Image2", @"C:\Users\Sterben\Desktop\dir\6.PNG");
                        insertDataCommand.Parameters.AddWithValue("TextColumn", "3");
                        insertDataCommand.ExecuteNonQuery();

                        insertDataCommand.Parameters.Clear();

                        insertDataCommand.Parameters.AddWithValue("Image1", @"C:\Users\Sterben\Desktop\dir\7.PNG");
                        insertDataCommand.Parameters.AddWithValue("Image2", @"C:\Users\Sterben\Desktop\dir\8.PNG");
                        insertDataCommand.Parameters.AddWithValue("TextColumn", "4");
                        insertDataCommand.ExecuteNonQuery();

                        insertDataCommand.Parameters.Clear();

                        insertDataCommand.Parameters.AddWithValue("Image1", @"C:\Users\Sterben\Desktop\dir\9.PNG");
                        insertDataCommand.Parameters.AddWithValue("Image2", @"C:\Users\Sterben\Desktop\dir\10.PNG");
                        insertDataCommand.Parameters.AddWithValue("TextColumn", "5");
                        insertDataCommand.ExecuteNonQuery();
                    }

                    MessageBox.Show("Файл базы данных Access и таблица успешно созданы.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка: " + ex.Message);
                }
            }
        }
    }
}
