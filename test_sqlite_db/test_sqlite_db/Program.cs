using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;

namespace test_sqlite_db
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"C:\Users\CCrowe\Documents\GitHub\Pokemon_Analysis\Pokemon\poke_data.db";
            SQLiteConnection m_dbConnection = new SQLiteConnection("Data Source=" + path + ";Version=3;");
            m_dbConnection.Open();
            string sql = "SELECT * FROM MOVES where pokemon_fk = @number and ACCURACY > 0 AND ATTACK > 0";
            using (SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection))
            {
                command.Parameters.AddWithValue("@number", 1);
                SQLiteDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    var vars = reader.GetValues();
                    int cnt = 0;
                    foreach (var val in vars)
                    {
                        Console.WriteLine(val);
                        Console.WriteLine(reader.GetValue(cnt));
                        Console.WriteLine(cnt);
                        cnt += 1;
                    }
                    int id = Int32.Parse(reader.GetValue(0).ToString());
                    int accuracy = Int32.Parse(reader.GetValue(1).ToString());
                    int attack = Int32.Parse(reader.GetValue(2).ToString());
                    int cat = Int32.Parse(reader.GetString(3));
                    int effect = Int32.Parse(reader.GetValue(4).ToString());
                    int level = Int32.Parse(reader.GetValue(5).ToString());
                    string name = reader.GetValue(6).ToString();
                    int pp = Int32.Parse(reader.GetValue(7).ToString());
                    int pokemon_fk = Int32.Parse(reader.GetValue(9).ToString());
                }
            }
            Console.ReadLine();
        }
    }
}
