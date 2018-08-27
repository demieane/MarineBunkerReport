using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;

namespace LearningCsvTools
{
    class Program
    {
        static void Main(string[] args)
        {
            // JOB 3.1
            // Read from csv file with columns

            // Reference path
            // TODO: The file extention is important .txt and .csv are different files

            // PATHS -----------------------------------------------------------------------------------------------------------
            // Path to read data from:
            string path = @"C:\Users\tech006\Source\Repos\LearningCsvTools\LearningCsvTools\BunkerReportDatabaseRead.txt";
            // Path to write data to, Database State:
            string pathWrite = @"C:\Users\tech006\Source\Repos\LearningCsvTools\LearningCsvTools\BunkerReportDatabaseWrite.txt";

            // If the file does not exist, eg. the first creation of the database ever (OPTION_1)
            // The user could decide the name of the database file

            DatabaseInfo Data = new DatabaseInfo(); // Stores information about my Database

            if (!File.Exists(pathWrite))
            {

                Data.State = true; // Created

                // Create the file.
                using (FileStream fs = File.Create(pathWrite))
                {
                    Byte[] info =
                        new UTF8Encoding(true).GetBytes("ID,NAME,SIZE,DATE,DIFFERENCE_REPORTED_HFO,DIFFERENCE_REPORTED_DO,AT_PORT,AT_SEA,NOTES");

                    // Add some information to the file.
                    fs.Write(info, 0, info.Length);
                }

            }
            else
            { Data.State = false; } // Already exists}


            if (File.Exists(path))

            {  
                // The file already exists and i want to retrieve the data it contains
                TextReader textReader = File.OpenText(path);
                
                // create a writer and open the file
                bool UpdateFile = true;
                TextWriter tw = new StreamWriter(pathWrite, UpdateFile);

                var csvWriter = new CsvWriter(tw);
                // Default value
                csvWriter.Configuration.HasHeaderRecord = true;
                if (Data.State == true)
                { csvWriter.NextRecord(); }
                

                var csv = new CsvReader(textReader);

                while (csv.Read())
                {
                    var record = csv.GetRecord<Columns>();
                    Console.WriteLine("=================================");
                    Console.WriteLine("Information Retrieved:  ");
                    Console.WriteLine("FROM: {0}", path);
                    Console.WriteLine("=================================");
                    Console.WriteLine("ID: {0}",record.ID);
                    Console.WriteLine("NAME: {0}",record.NAME);
                    Console.WriteLine("SIZE: {0}",record.SIZE);
                    Console.WriteLine("DATE: {0}",record.DATE);
                    Console.WriteLine("DIFFERENCE_REPORTED_HFO(MT): {0}", record.DIFFERENCE_REPORTED_HFO);
                    Console.WriteLine("DIFFERENCE_REPORTED_DO(MT): {0}", record.DIFFERENCE_REPORTED_DO);
                    Console.WriteLine("PORT: {0}",record.AT_PORT);
                    Console.WriteLine("SEA: {0}",record.AT_SEA);
                    Console.WriteLine("NOTES: {0}", record.NOTES); // The real values of HFO, DO

                    
                    csvWriter.WriteField(record.ID);
                    csvWriter.WriteField(record.NAME);
                    csvWriter.WriteField(record.SIZE);
                    csvWriter.WriteField(record.DATE);
                    csvWriter.WriteField(record.DIFFERENCE_REPORTED_HFO);
                    csvWriter.WriteField(record.DIFFERENCE_REPORTED_DO);
                    csvWriter.WriteField(record.AT_PORT);
                    csvWriter.WriteField(record.AT_SEA);
                    csvWriter.WriteField(record.NOTES);
                    
                    csvWriter.NextRecord();

                    // write a line of text to the file
                    //csvWriter.WriteRecords<Columns>(record);
                    //tw.WriteLine(record.ID);
                    //tw.WriteLine(record.AT_SEA);




                    /*
                     This while loop reads in each row data based on the class Column I have created
                     var record is an implicit type variable that is local
                     TODO: Create class of data type to store this info
                     */
                }
                // close the stream
                tw.Close();
                //else

            }



            // PARSER 
            //string path = "C:\\Users\\tech006\\Source\\Repos\\LearningCsvTools\\LearningCsvTools\\csvSample.csv";
            /*
            string path = @"C:\Users\tech006\Source\Repos\LearningCsvTools\LearningCsvTools\csvSample.csv";
            TextReader reader = File.OpenText(path);
            
                var parser = new CsvHelper.CsvParser(reader);

                while (true)
                {
                    //var row = parser.Read(); // variable with implicit type
                    string[] row = parser.Read(); // variable with explicit type
                    Console.WriteLine("Hello World!");
                    if (row == null)
                    {
                        Console.WriteLine("Hello World before it breaks!");
                        break;
                    }

                    foreach (string value in row)
                        Console.WriteLine("   {0}", value);
                }

            */
            Console.WriteLine("End of main");
        }
        
    }
}

