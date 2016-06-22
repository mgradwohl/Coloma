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
            int filesFailed = 0, filesSucceeded = 0;
            var files = Directory.EnumerateFiles(folderPath, "*.tsv").ToArray();
            foreach (string file in files)
            {
                ReadFileLogToConsole(file);
                string[] lines;
                try
                {
                    lines = File.ReadAllLines(file);
                    filesSucceeded++;
                }
                catch (IOException)
                {
                    Console.WriteLine("Could not open the file, continuing without it");
                    filesFailed++;
                    continue;
                }
                WrittenFileLogToConsole(filesSucceeded, filesFailed, file, files.Count());
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

            Console.WriteLine();
            Console.WriteLine($"Files Succeeded: {filesSucceeded}");
            Console.WriteLine($"Files Failed: {filesFailed}");
            Console.WriteLine("Done, thank you. Hit any key to exit");
            Console.ReadKey(true);
        }

        private static void ReadFileLogToConsole(string file)
        {
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine($"Reading {file}");
        }

        private static void WrittenFileLogToConsole(int filesSucceeded, int filesFailed, string file, int fileCount)
        {
            Console.Write("[");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write($"{filesSucceeded} ");
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("|");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write($" {filesFailed}");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine($"] Writing {file}");
            Console.ForegroundColor = ConsoleColor.DarkCyan;
            Console.WriteLine($"Files remaining: {fileCount - (filesSucceeded + filesFailed)}");
        }
    }
}
