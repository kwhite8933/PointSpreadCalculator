using System;
namespace Games
{ 
    public class Game
    {
        // Constructor
	    public Game()
	    {
             baller();
	    }

        // Member Variables
        private int spread = 0;
        private string awayTeam = "";
        public int spread { get; set; }
        private int homeTeam;

        public int HomeTeam
        {
            get { return homeTeam; }
            set { homeTeam = value; }
        }

        // Member Functions
        private void baller()
        {
            HomeTeam.get();
            HomeTeam.set(1);
        }

    }
}
