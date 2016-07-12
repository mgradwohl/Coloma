using System;
using System.IO;
using System.Linq;

namespace ColomaAnalysis
{
    class Program
    { 
        static void Main(string[] args)
        {
            // Configure this so that only TSVs from our current data schema version are aggregated
            // Aggregating TSVs from unsupported versions results in malformed CSVs where some rows will be shifted
            string targetVersion = "1.0.0.1";

            // Configure this to output to the correct Analysis directory,
            // NOTE: This will destroy other dataset.tsv files by default.
            string filepath = @"\\iefs\users\mattgr\Coloma\Analysis";

            // Configure this to point to the directly containing all the TSV files needing to be aggregated.
            string folderPath = @"\\iefs\users\mattgr\Coloma\";

            DateTime start = DateTime.Now;
            bool first = true;
            int filesFailed = 0, filesSucceeded = 0;
            var files = Directory.EnumerateFiles(folderPath, "*.tsv").ToArray();
            var output = File.Create($"{filepath}dataset.tsv");

            bool firstFile = true; // Only keep the header from the first file.
            foreach (string file in files)
            {
                // The TSV must exactly contain our targetVersion or else skip it.
                if (!file.Contains(targetVersion))
                {
                    filesSucceeded++;
                    SkippedFileLogToConsole(filesSucceeded, filesFailed, file, files.Count(), "File version invalid.");
                    continue;
                }

                ReadFileLogToConsole(file);

                var input = File.OpenRead(file);

                // If we aren't the first file, throw away the header
                if (firstFile == false)
                {
                    string header =
                        "branch	Build	Revision	machineName	deviceId	userName	Logname	level	instanceid	timeCreated	source	message\r\n";
                    
                    // Need to seek 3 more bytes, because the files are UTF-8 meaning they have a 3 byte header.
                    input.Seek(header.Length + 3, SeekOrigin.Current);

                }
                else
                {
                    firstFile = false;
                }

                var buffer = new byte[1024];
                int bytesRead;
                while ((bytesRead = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    output.Write(buffer, 0, bytesRead);
                }
                filesSucceeded++;

                WrittenFileLogToConsole(filesSucceeded, filesFailed, file, files.Count());
                //TODO: Maybe add a timestamp?
            }

            output.Close();
            TimeSpan timeTaken = DateTime.Now - start;
            Console.WriteLine("Took " + timeTaken + "\n");
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

        private static void SkippedFileLogToConsole(int filesSucceeded, int filesFailed, string file, int fileCount, string reason = "")
        {
            Console.Write("[");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write($"{filesSucceeded} ");
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("|");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write($" {filesFailed}");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine($"] Skipping {file}" + " : " + reason);
            Console.ForegroundColor = ConsoleColor.DarkCyan;
            Console.WriteLine($"Files remaining: {fileCount - (filesSucceeded + filesFailed)}");
        }
    }
}
