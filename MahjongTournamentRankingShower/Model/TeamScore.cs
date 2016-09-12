using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MahjongTournamentRankingShower.Model
{
    public class TeamScore
    {
        public string team;
        public int points;
        public int score;

        public TeamScore(string team, string points, string score)
        {
            this.team = team;
            this.points = int.Parse(string.IsNullOrEmpty(points) ? "0" : points);
            this.score = int.Parse(string.IsNullOrEmpty(score) ? "0" : score);
        }

        public TeamScore()
        {

        }
    }
}
