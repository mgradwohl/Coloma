using System;
using System.IO;
using System.Linq;

namespace ColomaAnalysis
{
    class Program
    {
        static void Main(string[] args)
        {
            string folderPath = @"\\iefs\users\mattgr\Coloma\";
            string filepath = @"\\iefs\users\mattgr\Coloma\Analysis\";
            bool first = true;
            foreach (string file in Directory.EnumerateFiles(folderPath, "*.tsv"))
            {
                Console.WriteLine($"Reading {file}");
                string[] lines;
                try
                {
                    lines = File.ReadAllLines(file);
                }
                catch (IOException)
                {
                    Console.WriteLine("Could not open the file, continuing without it");
                    continue;
                }
                Console.WriteLine($"Writing {file}");
                //TODO: Add to a buffer first and then write to the file for faster run time
                //TODO: Delete the old file first, this will just append to the file forever if you don't delete it
                //TODO: Maybe add a timestamp?
                if (first)
                {
                    File.AppendAllLines($"{filepath}dataset.tsv", lines.ToArray());
                    first = false;
                }
                File.AppendAllLines($"{filepath}dataset.tsv", lines.Skip(1).ToArray());
            }
        }
    }
}
