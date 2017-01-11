using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace B_Imp
{
    class Move
    {
        public int PTypeMoveMultiplier;
        public string Type1;
        public string Type2;
    }
    class Program
    {
        public static Queue<Move> queue;
        static void Main(string[] args)
        {

        }
        static void InQueueu()
        {
            Move move = queue.Dequeue();
            move.PTypeMoveMultiplier = GetMoveTypeBonus();
        }
        static void B()
        {
            GetMostDamagingMove();
        }

        static void GetMostDamagingMove()
        {
            
        }

        static double GetMoveTypeBonus(string pType,Move move)
        {

        }
    }
}
