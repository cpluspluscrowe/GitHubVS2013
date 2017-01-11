using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace Pierre_License_Plate_Counts
{
    public class Program
    {
        public static class ODRecords
        {
            public static List<SortedList<DateTime,Trip>> oDList;
            static ODRecords(){
                oDList = new List<SortedList<DateTime, Trip>>();
            }
        }
        public class Trip
        {
            public DateTime Start;
            public DateTime End;
            public string Origin;
            public string Destination;
            public List<Report> OrigDestTrips;
            public List<string> tripStringList;
            public Trip()
            {
                this.OrigDestTrips = new List<Report>();
                this.tripStringList = new List<string>();
            }
            public void AddIncident(Report incident){
                this.OrigDestTrips.Add(incident);
                this.tripStringList.Add(incident.Location.Replace("[NBL]", "[NB]").Replace("[NBR]", "[NB]")
                    .Replace("[SBR]", "[SB]").Replace("[SBL]", "[SB]")
                    .Replace("[EBR]", "[EB]").Replace("[EBL]", "[EB]")
                    .Replace("[WBR]", "[WB]").Replace("[WBL]", "[WB]")
                    );
            }
        }
        public class Car
        {
            public SortedList<DateTime,Report> Incidents;
            public Car(Report report)
            {
                this.Incidents = new SortedList<DateTime,Report>();
                this.Incidents.Add(report.TimeStamp,report);
            }
        }
        public class Report
        {
            public DateTime TimeStamp;
            public string LicensePlate;
            public string Location;
            public Report(DateTime timestamp,string licenseplate,string location){
                this.TimeStamp = timestamp;
                this.LicensePlate = licenseplate;
                this.Location = location;
            }
        }
        public static bool AreListsEqual(List<string> l1, List<string> l2)
        {
            if (l1.Count != l2.Count)
            {
                return false;
            }
            else
            {
                for (int i = 0; i < l1.Count; i++)
                {
                    if (l1.ElementAt(i) == l2.ElementAt(i))
                    {

                    }
                    else
                    {
                        return false;
                    }
                }
            }
            return true;
        }
        public static Excel.Workbook wb;
        public static SortedDictionary<string, List<string>> tripPermutations;
        static void Main(string[] args)
        {
            tripPermutations = new SortedDictionary<string, List<string>>();
            Dictionary<string, Car> carDict = new Dictionary<string, Car>();
            string pathToPierreExcel = @"C:\Users\CCrowe\Documents\Traffic\Carbee\Plate Matching\Pierre SDDOT Plate Matching (Autosaved).xlsm";
            Excel.Application xl = new Excel.Application();
            xl.Visible = true;
            wb = xl.Workbooks.Open(pathToPierreExcel);
            Excel.Worksheet raw = wb.Sheets["Raw Data"];
            for (int i = 2; i <= raw.UsedRange.Rows.Count; i++)//raw.UsedRange.Rows.Count
            {
                Report report = new Report(raw.Range["A" + i.ToString()].Value, raw.Range["B" + i.ToString()].Value, raw.Range["C" + i.ToString()].Value);
                if (carDict.ContainsKey(report.LicensePlate))
                {
                    if (!carDict[report.LicensePlate].Incidents.ContainsKey(report.TimeStamp))
                    {
                        carDict[report.LicensePlate].Incidents.Add(report.TimeStamp, report);
                    }else{
                        Console.WriteLine("Found Repeated Report, time:" + report.TimeStamp + "  Plate:" + report.LicensePlate + " Location:" + report.Location);
                    }
                }
                else
                {
                    carDict.Add(report.LicensePlate,new Car(report));
                }
            }
            Car carEx = carDict["01539A"];
            List<Trip> tripList = new List<Trip>();
            KeyValuePair<DateTime,Report> currentIncident = new KeyValuePair<DateTime,Report>();
            Trip trip = new Trip();
            foreach (KeyValuePair<string, Car> car in carDict)
            {
                if (car.Value.Incidents.First().Value.LicensePlate == "36F640")
                {
                    int p = 5;
                }
                foreach (KeyValuePair<DateTime,Report> incident in car.Value.Incidents)
                {
                    if(incident.Value == car.Value.Incidents.First().Value){
                        trip = new Trip();
                        trip.Origin = car.Value.Incidents.First().Value.Location;
                        trip.AddIncident(incident.Value);
                        currentIncident = incident;
                    }
                    else
                    {
                        var difference = incident.Value.TimeStamp.Subtract(currentIncident.Value.TimeStamp);
                        if (difference.TotalMinutes <= 20)
                        {
                            trip.AddIncident(incident.Value);
                        }
                        else
                        {
                            trip.Destination = trip.OrigDestTrips.Last().Location;
                            tripList.Add(trip);
                            Trip pastTrips = trip;
                            add2Permutation(trip);
                            trip = new Trip();

                            bool foundOldest = false;
                            currentIncident = incident;
                            foreach (var pastTrip in pastTrips.OrigDestTrips)
                            {
                                difference = currentIncident.Value.TimeStamp.Subtract(pastTrip.TimeStamp);
                                if (difference.TotalMinutes <= 20)
                                {
                                    if (foundOldest == false)
                                    {
                                        trip.Origin = pastTrip.Location;
                                        currentIncident = new KeyValuePair<DateTime,Report>(pastTrip.TimeStamp,pastTrip);
                                        trip.AddIncident(pastTrip);
                                        foundOldest = true;
                                    }
                                    else
                                    {
                                        trip.AddIncident(pastTrip);
                                    }
                                }
                            }
                            if (foundOldest == false)
                            {
                                trip.Origin = incident.Value.Location;
                                currentIncident = incident;
                                trip.AddIncident(incident.Value);
                                foundOldest = true;
                            }
                            else
                            {
                                trip.AddIncident(incident.Value);
                            }
                        }
                    }
                }
                trip.Destination = trip.OrigDestTrips.Last().Location;
                tripList.Add(trip);
                //Below, create a list of unique permutations of trips
                add2Permutation(trip);
            }
            foreach (KeyValuePair<string,List<string>> od in tripPermutations)
            {
                SortedList<DateTime, Trip> singlePermuList = new SortedList<DateTime, Trip>();
                foreach(Trip t in tripList){
                    if (AreListsEqual(t.tripStringList,od.Value))
                    {
                        singlePermuList.Add(t.OrigDestTrips.First().TimeStamp,t);
                    }
                }
                ODRecords.oDList.Add(singlePermuList);
            }
            var results1 = from tripAgg in ODRecords.oDList
                           from aggList in tripAgg.Values
                           where aggList.OrigDestTrips.Any(p => p.LicensePlate == "36N804")
                           select aggList;

            createTripSheet("US 14 near Missouri River Bridge [EB]", "SD1804 north of Study [NB]","Bridge (EB) to N of Study (NB)2");
            createTripSheet("SD1804 north of Study [SB]", "US 14 near Missouri River Bridge [WB]", "N or Study (SB) to Bridge (WB)");

            createTripSheet("US 14 near Missouri River Bridge [EB]", "US 14 east of Study Area [EB]", "Bridge (EB) to E of Study (EB)");
            createTripSheet("US 14 east of Study Area [WB]", "US 14 near Missouri River Bridge [WB]", "E of Study (WB) to Bridge (WB)");

            createTripSheet("US 14 near Missouri River Bridge [EB]", "SD34 east of Garfield Ave [EB]", "Bridge (EB) to East of Gar (EB)");
            createTripSheet("SD34 east of Garfield Ave [WB]", "US 14 near Missouri River Bridge [WB]", "E of Gar (WB) to Bridge (WB)");

            /*
            //The below commented out code records the trips.  This was done to validate data correctness
            Excel.Worksheet ex = wb.Sheets["Extra"];
            foreach (var tripListed in tripList)
            {
                int j = 2;
                foreach (var place in tripListed.OrigDestTrips)
                {
                    ex.Cells[row, 1].Value = place.LicensePlate;
                    ex.Cells[row, j].Value = place.TimeStamp;
                    j += 1;
                }
                row += 1;
            }*/
           

            /*
             * The below piece creates the common routes worksheet and fills it in
            Excel.Worksheet CR = wb.Sheets["Common Routes"];
            cnt = 2;
            foreach (var aggList in ODRecords.oDList)
            {
                CR.Range["A" + cnt.ToString()].Value = aggList.Count;
                var first = aggList.First();
                int j = 2;
                foreach (string loc in first.Value.tripStringList)
                {
                    CR.Cells[cnt, j].Value = loc;
                    j += 1;
                }
                cnt += 1;
            }
            */

        }
        public static void add2Permutation(Trip trip)
        {
            string longTripString = "";
            foreach (var loc in trip.tripStringList)
            {
                longTripString += loc.ToString();
            }
            if (!tripPermutations.ContainsKey(longTripString))
            {
                tripPermutations.Add(longTripString, trip.tripStringList);
            }
        }
        public static void createTripSheet(string loc1,string loc2,string name)
        {
            Excel.Worksheet SW = wb.Sheets.Add(After:wb.Sheets[wb.Sheets.Count]);
            //SW.Name = name;
            int counter = 0;
            int row = 3;
            int cnt = 3;
            foreach (var aggList in ODRecords.oDList)
            {
                var exampleItem = aggList.First();
                bool eb = false;
                bool nb = false;
                foreach (string loc in exampleItem.Value.tripStringList)
                {
                    var matches = exampleItem.Value.OrigDestTrips.Where(p => p.LicensePlate == "36N804").ToList();
                    if (matches.Count > 0)
                    {
                        int cx = 0;
                    }
                    if (loc == loc1)
                    {
                        eb = true;
                    }
                    if (loc == loc2)
                    {
                        nb = true;
                    }
                }
                if (eb && nb)
                {
                    foreach (var item in aggList)
                    {
                        int j = 2;
                        cnt += 1;
                        counter += 1;
                        SW.Cells[cnt, 1].Value = item.Value.OrigDestTrips.First().LicensePlate;
                        foreach (var tripItem in item.Value.OrigDestTrips)
                        {
                            SW.Cells[cnt, j].Value = tripItem.TimeStamp + " " + tripItem.Location;
                            if (tripItem.Location.Contains(loc1) || tripItem.Location.Contains(loc2))
                            {
                                SW.Cells[cnt, j].Style = "Good";
                            }
                            j += 1;
                        }
                    }
                }
            }
        }
    }
}
