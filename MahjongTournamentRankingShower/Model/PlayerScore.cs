using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MahjongTournamentRankingShower.Model
{
    public class PlayerScore
    {
        public int id;
        public string name;
        public int points;
        public int score;
        public string team;
        public string country;

        public PlayerScore(string id, string name, string points, string score, string team, string country)
        {
            this.id = int.Parse(string.IsNullOrEmpty(id) ? "0" : id);
            this.name = name;
            this.points = int.Parse(string.IsNullOrEmpty(points) ? "0" : points); ;
            this.score = int.Parse(string.IsNullOrEmpty(score) ? "0" : score); ;
            this.team = team;
            this.country = country;
        }

        public PlayerScore()
        {

        }
    }
}
