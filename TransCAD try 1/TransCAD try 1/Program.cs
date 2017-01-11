using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CaliperForm;

namespace TransCAD_try_1
{
    class Program
    {
        static void Main(string[] args)
        {
        }
    }
    public static void Open_Map() {
    CaliperForm.Connection Conn = new CaliperForm.Connection { MappingServer = "TransCAD" };
    Boolean opened = false;
    try {
        opened = Conn.Open();
        if (opened) {
            // You must declare dk as "dynamic" or the compiler will throw an error
            dynamic dk = Conn.Gisdk;
            string tutorial_folder = dk.Macro("G30 Tutorial Folder") as string;
            // Obtain information about Conn: an array of 
            //  [ "program_path" , "program name" , "program type" , 
            //     build number (integer) , version number (real) , instance number ]
            Object[] program_info = dk.GetProgram();
            string program_name = program_info[1] as string;
            int build_number = (int)program_info[3];
            double version_number = (double)program_info[4];
            // Set the  path that will be used to look for layers used in map files
            string search_path = "c:\\ccdata\\USA (HERE) - 2013 Quarter 4;d:\\ccdata";
            dk.SetSearchPath(search_path);
            // Open a map file in the tutorial folder
            string map_file = tutorial_folder + "BestRout.map";
            var map_options = new OptionsArray();  // you can also use Dictionary or Hashtable 
            map_options["Auto Project"] = "true";
            string data_directory = System.IO.Path.GetDirectoryName(map_file);
            map_options["Force Directory"] = data_directory; // path used to look for layers used in the map 
            string map_name = dk.OpenMap(map_file, map_options) as string;
            if (map_name == null) {
                Console.Out.WriteLine("Cannot open map " + map_file + ". Perhaps some layers cannot be found?");
                return;
            } else {
                Console.Out.WriteLine("map_name = " + map_name);
            }
            // Set the current layer
            dk.SetLayer("County");
            // Get information about the list of layers contained in this map
            dynamic layers = dk.GetMapLayers(map_name, "All");
            dynamic layer_names = layers[0];
            int current_idx = (int)layers[1];
            string current_layer = layers[2] as string;
            dk.CloseMap(map_name);
            Conn.Close();
        }
    } catch (System.Exception error) {
        Console.Out.WriteLine(error.Message);
    }
}
}
