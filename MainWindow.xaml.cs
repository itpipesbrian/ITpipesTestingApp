using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data;

namespace ITpipesTestingApplication
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            List<string> lML_Name = new List<string>();
            string sFileLocation;
            using (FileDialog fileDialog = new OpenFileDialog())
            {
                fileDialog.Filter = "Access Database (*.mdb)|*.mdb";

                if (fileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    textBox.Text = sFileLocation = fileDialog.FileName;
                    lML_Name = ConnectToDatabase(sFileLocation);
                    listBox.ItemsSource = lML_Name;
                }
            }
        }

        private List<string> ConnectToDatabase(string pDatabaseLocation)
        {
            List<string> ReturnValue = new List<string>();
            string sAccessConn = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={pDatabaseLocation};";
            try
            {
                using (OleDbConnection oledbConnection = new OleDbConnection(sAccessConn))
                {
                    string sSqlCommand = "SELECT DISTINCT ML.ML_Name FROM ML ORDER BY ML.ML_Name";
                    OleDbCommand oleDbCommand = oledbConnection.CreateCommand();

                    oleDbCommand.Connection = oledbConnection;
                    oleDbCommand.CommandText = sSqlCommand;
                    oledbConnection.Open();
                    OleDbDataReader dataReader = oleDbCommand.ExecuteReader();
                    if (dataReader.HasRows)
                    {
                        while (dataReader.Read())
                        {
                            ReturnValue.Add(dataReader.GetString(0));
                        }

                    }
                    oledbConnection.Close();
                }
            }
            catch(Exception ex)
            {
                System.Windows.MessageBox.Show($"Error in the software: {ex.ToString()}");
            }
            return ReturnValue;
        }
    }
}
