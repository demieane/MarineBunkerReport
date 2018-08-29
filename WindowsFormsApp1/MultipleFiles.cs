using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel;

/// <summary>
/// Summary description for Class1
/// </summary>
/// 

public class MultipleFiles
{

    // Process all files in the directory passed in, recurse on any directories 
    // that are found, and process the files they contain.
    public static List<string> ProcessDirectory(string targetDirectory, int index)//, out List<string> stringList, out int indexUpdated)
    {
        // Index that iterates over the fileEntries

        Console.WriteLine("File path ID: {0}", index);
        // Process the list of files found in the directory.

        string[] fileEntries = Directory.GetFiles(targetDirectory);
        List<string> stringList = new List<string>();

        foreach (string fileName in fileEntries)
        {

            var result = fileName.Substring(fileName.Length - 4);
            var result_1 = fileName.Substring(fileName.Length - 5);
            var result_2 = fileName.Substring(fileName.Length - 3);
            Console.WriteLine("Last characters: {0}", result);
            Console.WriteLine("Last characters: {0}", result_1);
            Console.WriteLine("Last characters: {0}", result_2);
            if (result == ".xls" || result == ".xlt" || result_1 == ".xlsx" || result_2 == "xls" || result_2 == "xlt")
            {
                ProcessFile(fileName);
                Console.WriteLine("File path ID: {0}", index);
                stringList.Add(fileName);
                //Console.WriteLine(index);
                //Console.WriteLine(stringList[index]);
                index = index + 1;
            }
        }
        //Console.WriteLine(index);

        return stringList;
        
        /*
        // Recurse into subdirectories of this directory.
        string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
        foreach (string subdirectory in subdirectoryEntries)
        {
            ProcessDirectory(subdirectory, index);
        }
        */
    
                
    }


    // Insert logic for processing found files here.
    public static void ProcessFile(string path)
    {
        Console.WriteLine("Processed file '{0}'.", path);
    }


}
