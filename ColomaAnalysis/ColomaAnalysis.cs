using System;
using System.IO;
using System.Linq;
using Coloma;

namespace ColomaAnalysis
{
    class ColomaAnalysis
    {
        static void Main(string[] args)
        {
            // This is the directory where TSVs are sourced.
            string folderPath = @"D:\source\Coloma\";

            // Get data schema version of Coloma.exe
            // We use this assembly version to verify source TSVs are from the right version of Coloma.exe
            string colomaAssemblyVersion = typeof(Coloma.Coloma).Assembly.GetName().Version.ToString();

            // We also use the assembly version to output the aggregated TSV to the correct directory.
            string filepath = @"D:\source\Coloma\Analysis_" + colomaAssemblyVersion + @"\";

            // It is safe to directly call CreateDirectory to ensure the directory exists
            Directory.CreateDirectory(filepath);

            DateTime start = DateTime.Now;
            int filesFailed = 0, filesSucceeded = 0;
            var files = Directory.EnumerateFiles(folderPath, "*.tsv").ToArray();
            var output = File.Create($"{filepath}dataset.tsv");

            bool firstFile = true; // Only keep the header from the first file.
            foreach (string file in files)
            {
                // The TSV must exactly contain our targetVersion or else skip it.
                if (!file.Contains(colomaAssemblyVersion))
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
