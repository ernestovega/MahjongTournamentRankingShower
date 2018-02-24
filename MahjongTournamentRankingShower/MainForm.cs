using MahjongTournamentRankingShower.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace MahjongTournamentRankingShower
{
    public partial class MainForm : Form
    {
        private readonly int SLEEP_TIME = 5000;

        string playersFilePath = string.Empty;
        string scoresFilePath = string.Empty;
        string errorMessage = string.Empty;
        private List<PlayerScore> playersScores = new List<PlayerScore>();
        private List<TeamScore> teamsScores = new List<TeamScore>();
        private List<string[]> sPlayersScores = new List<string[]>();
        private List<string[]> sTeamsScores = new List<string[]>();

        public MainForm()
        {
            Thread splashThread = new Thread(new ThreadStart(openSplash));
            splashThread.Start();

            InitializeComponent();
            Thread.Sleep(1000);

            if (!isExcelInstalled())
            {
                MessageBox.Show(new Form() { TopMost = true }, "Excel not present on your computer.\nPlease get it.");
                Application.Exit();
                return;
            }

            if (!ExistsScoresFile())
            {
                MessageBox.Show(new Form() { TopMost = true }, "Couldn't find scores excel.");
                Application.Exit();
                return;
            }

            if (ImportScores() < 0)
            {
                MessageBox.Show(new Form() { TopMost = true }, errorMessage);
                Application.Exit();
                return;
            }
            else
                splashThread.Abort();


            Thread showerThread = new Thread(new ThreadStart(ShowRanking));
            showerThread.Start();
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
                MessageBox.Show(new Form() { TopMost = true }, "Excel is not present on your computer.");
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
            DataTable dataTablePlayers = new DataTable();
            DataTable dataTableTeams = new DataTable();
            string sqlExcelPlayers;
            string sqlExcelTeams;
            string strConnXlsx = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + scoresFilePath
                + ";Extended Properties=" + '"' + "Excel 12.0 Xml;IMEX=1" + '"';
            string strConnXls = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + scoresFilePath
                + ";Extended Properties=" + '"' + "Excel 8.0;IMEX=1" + '"';
            string strConn = scoresFilePath.Substring(scoresFilePath.Length - 4).ToLower().Equals("xlsx")
                ? strConnXlsx : strConnXls;
            using (OleDbConnection conn = new OleDbConnection(strConn))
            {
                try
                {
                    conn.Open();
                    var dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    var sheetPlayersTotal = dtSchema.Rows[0].Field<string>("TABLE_NAME");
                    var sheetTeamsTotal = dtSchema.Rows[7].Field<string>("TABLE_NAME");
                    sqlExcelPlayers = "SELECT * FROM [" + sheetPlayersTotal + "]";
                    sqlExcelTeams = "SELECT * FROM [" + sheetTeamsTotal + "]";
                    OleDbDataAdapter oleDbdataAdapterPlayers = new OleDbDataAdapter(sqlExcelPlayers, conn);
                    OleDbDataAdapter oleDbdataAdapterTeams = new OleDbDataAdapter(sqlExcelTeams, conn);
                    oleDbdataAdapterPlayers.Fill(dataTablePlayers);
                    oleDbdataAdapterTeams.Fill(dataTableTeams);
                }
                catch(Exception e)
                {
                    errorMessage += "\n\tWrong Excel file format.";
                    flagWrongExcel = true;
                }

                if (dataTablePlayers == null || dataTablePlayers.Rows == null || dataTablePlayers.Columns == null)
                {
                    flagWrongExcel = true;
                    errorMessage += "\n\tWrong Players Excel sheet format or empty.";
                }
                else if (dataTablePlayers.Columns.Count < 6)
                {
                    flagWrongExcel = true;
                    errorMessage += "\n\tThere aren´t enough columns in Players sheet.";
                }
                else if (!dataTablePlayers.Columns[0].ColumnName.ToString().ToLower().Equals("id") ||
                    !dataTablePlayers.Columns[1].ColumnName.ToString().ToLower().Equals("name") ||
                    !dataTablePlayers.Columns[2].ColumnName.ToString().ToLower().Equals("points") ||
                    !dataTablePlayers.Columns[3].ColumnName.ToString().ToLower().Equals("score") ||
                    !dataTablePlayers.Columns[4].ColumnName.ToString().ToLower().Equals("team") ||
                    !dataTablePlayers.Columns[5].ColumnName.ToString().ToLower().Equals("country"))
                {
                    flagWrongExcel = true;
                    errorMessage += "\n\tColumn headers doesn´t match in Players sheet.";
                }

                if (!flagWrongExcel)
                {
                    foreach (DataRow row in dataTablePlayers.Rows)
                    {
                        try
                        {
                            if (string.IsNullOrWhiteSpace(row[0].ToString()) || string.IsNullOrWhiteSpace(row[1].ToString()) ||
                                string.IsNullOrWhiteSpace(row[2].ToString()) || string.IsNullOrWhiteSpace(row[3].ToString()) ||
                                string.IsNullOrWhiteSpace(row[4].ToString()) || string.IsNullOrWhiteSpace(row[5].ToString()))
                            {//Nos aseguramos de que no hay ninguna casilla vacía
                                flagWrongExcel = true;
                                AddNewPlayerScoreFromExcel(row);
                            }
                            else
                                AddNewPlayerScoreFromExcel(row);
                        }
                        catch (Exception)
                        {
                            flagWrongExcel = true;
                            AddNewPlayerScoreFromExcel(row);
                        }
                    }
                }
                if(flagWrongExcel) return -1;

                if (dataTableTeams == null || dataTableTeams.Rows == null || dataTableTeams.Columns == null)
                {
                    flagWrongExcel = true;
                    errorMessage += "\n\tWrong Teams Excel sheet format or empty.";
                }
                else if (dataTableTeams.Columns.Count < 3)
                {
                    flagWrongExcel = true;
                    errorMessage += "\n\tThere aren´t enough columns in Teams sheet.";
                }
                else if (!dataTableTeams.Columns[0].ColumnName.ToString().ToLower().Equals("team") ||
                    !dataTableTeams.Columns[1].ColumnName.ToString().ToLower().Equals("points") ||
                    !dataTableTeams.Columns[2].ColumnName.ToString().ToLower().Equals("score"))
                {
                    flagWrongExcel = true;
                    errorMessage += "\n\tColumn headers doesn´t match in Teams sheet.";
                }

                if (!flagWrongExcel)
                {
                    foreach (DataRow row in dataTableTeams.Rows)
                    {
                        try
                        {
                            if (string.IsNullOrWhiteSpace(row[0].ToString()) || string.IsNullOrWhiteSpace(row[1].ToString()) ||
                                string.IsNullOrWhiteSpace(row[2].ToString()))
                            {//Nos aseguramos de que no hay ninguna casilla vacía
                                flagWrongExcel = true;
                                AddNewTeamScoreFromExcel(row);
                            }
                            else
                                AddNewTeamScoreFromExcel(row);
                        }
                        catch (Exception e)
                        {
                            flagWrongExcel = true;
                            AddNewTeamScoreFromExcel(row);
                        }
                    }
                }
                return flagWrongExcel ? -1 : 1;
            }
        }

        private void AddNewPlayerScoreFromExcel(DataRow row)
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

        private void AddNewTeamScoreFromExcel(DataRow row)
        {
            teamsScores.Add(new TeamScore(
                row.IsNull(0) || string.IsNullOrWhiteSpace(row[0].ToString()) ? string.Empty : row[0].ToString(),
                row.IsNull(1) || string.IsNullOrWhiteSpace(row[1].ToString()) ? string.Empty : row[1].ToString(),
                row.IsNull(2) || string.IsNullOrWhiteSpace(row[2].ToString()) ? string.Empty : row[2].ToString()
                ));
        }

        private void ShowRanking()
        {
            playersScores.OrderBy(x => x.points).ThenBy(x => x.score);
            teamsScores.OrderBy(x => x.points).ThenBy(x => x.score);
            for (int i = 0; i < teamsScores.Count; i++)
            {
                sTeamsScores.Add(new string[] {
                            (i + 1).ToString(),
                            teamsScores[i].team,
                            teamsScores[i].points.ToString(),
                            teamsScores[i].score.ToString() });
            }

            int start = 0, end = 15;
            while (true)
            {
                sPlayersScores.Clear();
                for (int i = start; i < end; i++)
                {
                    sPlayersScores.Add(new string[] {
                        (i + 1).ToString(),
                        playersScores[i].name,
                        playersScores[i].points.ToString(),
                        playersScores[i].score.ToString(),
                        playersScores[i].team });
                }

                dataGridView.Invoke(new MethodInvoker(() => { showPlayers(); }));                
                Thread.Sleep(SLEEP_TIME);

                if(end >= playersScores.Count)
                {
                    dataGridView.Invoke(new MethodInvoker(() => { showTeams(); }));
                    Thread.Sleep(SLEEP_TIME);
                }

                if (end < playersScores.Count)
                {
                    start += 15;
                    end += 15;
                }
                else
                {
                    start = 0;
                    end = 15;
                }
                sPlayersScores.Clear();
            }
        }

        private void showPlayers()
        {
            dataGridView.DataSource = ConvertListToDataTable(sPlayersScores);
            dataGridView.Columns[0].HeaderText = "Position";
            dataGridView.Columns[1].HeaderText = "Name";
            dataGridView.Columns[2].HeaderText = "Points";
            dataGridView.Columns[3].HeaderText = "Score";
            dataGridView.Columns[4].HeaderText = "Team";
            dataGridView.Columns[0].Width = 120;
            dataGridView.Columns[1].Width = 550;
            dataGridView.Columns[2].Width = 100;
            dataGridView.Columns[3].Width = 100;
            dataGridView.Columns[4].Width = 450;
        }

        private void showTeams()
        {
            dataGridView.DataSource = ConvertListToDataTable(sTeamsScores);
            dataGridView.Columns[0].HeaderText = "Position";
            dataGridView.Columns[1].HeaderText = "Team";
            dataGridView.Columns[2].HeaderText = "Points";
            dataGridView.Columns[3].HeaderText = "Score";
            dataGridView.Columns[0].Width = 150;
            dataGridView.Columns[1].Width = 600;
            dataGridView.Columns[2].Width = 150;
            dataGridView.Columns[3].Width = 150;
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