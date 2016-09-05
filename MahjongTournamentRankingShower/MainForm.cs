using MahjongTournamentRankingShower.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace MahjongTournamentRankingShower
{
    public partial class MainForm : Form
    {
        string playersFilePath = string.Empty;
        string scoresFilePath = string.Empty;
        string errorMessage = string.Empty;
        private List<PlayerScore> playersScores = new List<PlayerScore>();
        private List<string[]> sPlayersScores = new List<string[]>();

        public MainForm()
        {
            Thread splashThread = new Thread(new ThreadStart(openSplash));
            splashThread.Start();

            InitializeComponent();
            Thread.Sleep(1000);

            if (!isExcelInstalled())
            {
                MessageBox.Show("Excel not present on your computer.\nPlease get it.");
                Application.Exit();
            }

            if (!ExistsScoresFile())
            {
                MessageBox.Show("Couldn't find scores excel.");
                Application.Exit();
            }

            if (ImportScores() < 0)
            {
                MessageBox.Show(errorMessage);
                Application.Exit();
            }
            else
                splashThread.Abort();

            ShowRanking();
        }

        public void openSplash()
        {
            Application.Run(new SplashForm());
        }

        private bool isExcelInstalled()
        {
            Type officeType = Type.GetTypeFromProgID("Excel.Application");
            if (officeType == null)
            {
                MessageBox.Show("Excel is not present on your computer.");
                return false;
            }
            else
                return true;
        }

        private bool ExistsScoresFile()
        {
            try
            {
                string[] dirs = Directory.GetFiles(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "Score_Tables*");
                if (dirs.Length > 0)
                {
                    scoresFilePath = dirs[dirs.Length - 1];
                    return true;
                }
                return false;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            Activate();
        }

        private int ImportScores()
        {
            bool flagWrongExcel = false;
            errorMessage = string.Empty;
            DataTable dataTable = new DataTable();
            string strConnXlsx = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + scoresFilePath
                + ";Extended Properties=" + '"' + "Excel 12.0 Xml;IMEX=1" + '"';
            string strConnXls = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + scoresFilePath
                + ";Extended Properties=" + '"' + "Excel 8.0;IMEX=1" + '"';
            string sqlExcel;
            string strConn = scoresFilePath.Substring(scoresFilePath.Length - 4).ToLower().Equals("xlsx")
                ? strConnXlsx : strConnXls;
            using (OleDbConnection conn = new OleDbConnection(strConn))
            {
                try
                {
                    conn.Open();
                    var dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    var sheetPlayersTotal = dtSchema.Rows[0].Field<string>("TABLE_NAME");
                    sqlExcel = "SELECT * FROM [" + sheetPlayersTotal + "]";
                    OleDbDataAdapter oleDbdataAdapter = new OleDbDataAdapter(sqlExcel, conn);
                    oleDbdataAdapter.Fill(dataTable);
                }
                catch
                {
                    errorMessage += "\n\tWrong Excel file format.";
                    flagWrongExcel = true;
                }

                if (dataTable == null || dataTable.Rows == null || dataTable.Columns == null)
                {
                    flagWrongExcel = true;
                    errorMessage += "\n\tWrong Excel file format or empty.";
                }
                else if (dataTable.Columns.Count < 6)
                {
                    flagWrongExcel = true;
                    errorMessage += "\n\tThere aren´t enough columns.";
                }
                else if (!dataTable.Columns[0].ColumnName.ToString().ToLower().Equals("id") ||
                    !dataTable.Columns[1].ColumnName.ToString().ToLower().Equals("name") ||
                    !dataTable.Columns[2].ColumnName.ToString().ToLower().Equals("points") ||
                    !dataTable.Columns[3].ColumnName.ToString().ToLower().Equals("score") ||
                    !dataTable.Columns[4].ColumnName.ToString().ToLower().Equals("team") ||
                    !dataTable.Columns[5].ColumnName.ToString().ToLower().Equals("country"))
                {
                    flagWrongExcel = true;
                    errorMessage += "\n\tColumn headers doesn´t match.";
                }

                if (!flagWrongExcel)
                {
                    foreach (DataRow row in dataTable.Rows)
                    {
                        try
                        {
                            if (string.IsNullOrWhiteSpace(row[0].ToString()) || string.IsNullOrWhiteSpace(row[1].ToString()) ||
                                string.IsNullOrWhiteSpace(row[2].ToString()) || string.IsNullOrWhiteSpace(row[3].ToString()) ||
                                string.IsNullOrWhiteSpace(row[4].ToString()) || string.IsNullOrWhiteSpace(row[5].ToString()))
                            {//Nos aseguramos de que no hay ninguna casilla vacía
                                flagWrongExcel = true;
                                AddNewScoreFromExcel(row);
                            }
                            else
                                AddNewScoreFromExcel(row);
                        }
                        catch (Exception)
                        {
                            flagWrongExcel = true;
                            AddNewScoreFromExcel(row);
                        }
                    }
                }
            }
            return flagWrongExcel ? -1 : 1;
        }

        private void AddNewScoreFromExcel(DataRow row)
        {
            playersScores.Add(new PlayerScore(
                row.IsNull(0) || string.IsNullOrWhiteSpace(row[0].ToString()) ? string.Empty : row[0].ToString(),
                row.IsNull(1) || string.IsNullOrWhiteSpace(row[1].ToString()) ? string.Empty : row[1].ToString(),
                row.IsNull(2) || string.IsNullOrWhiteSpace(row[2].ToString()) ? string.Empty : row[2].ToString(),
                row.IsNull(3) || string.IsNullOrWhiteSpace(row[3].ToString()) ? string.Empty : row[3].ToString(),
                row.IsNull(4) || string.IsNullOrWhiteSpace(row[4].ToString()) ? string.Empty : row[4].ToString(),
                row.IsNull(5) || string.IsNullOrWhiteSpace(row[5].ToString()) ? string.Empty : row[5].ToString()
                ));
        }

        private void ShowRanking()
        {
            foreach (PlayerScore ps in playersScores)
            {
                sPlayersScores.Add(new string[] {
                    ps.id.ToString(),
                    ps.name,
                    ps.points.ToString(),
                    ps.score.ToString(),
                    ps.team, ps.country });
            }
            dataGridView.DataSource = ConvertListToDataTable(sPlayersScores);
            dataGridView.Columns[0].HeaderText = "Id";
            dataGridView.Columns[1].HeaderText = "Name";
            dataGridView.Columns[2].HeaderText = "Points";
            dataGridView.Columns[3].HeaderText = "Score";
            dataGridView.Columns[4].HeaderText = "Team";
            dataGridView.Columns[5].HeaderText = "Country";
        }

        private static DataTable ConvertListToDataTable(List<string[]> list)
        {
            // New table.
            DataTable table = new DataTable();

            // Get max columns.
            int columns = 0;
            foreach (var array in list)
            {
                if (array.Length > columns)
                {
                    columns = array.Length;
                }
            }

            // Add columns.
            for (int i = 0; i < columns; i++)
            {
                table.Columns.Add();
            }

            // Add rows.
            foreach (var array in list)
            {
                table.Rows.Add(array);
            }

            return table;
        }
    }
}