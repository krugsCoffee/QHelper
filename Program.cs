using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using SharpCompress.Archives;
using SharpCompress.Archives.Rar;
using SharpCompress.Common;

namespace QHelper
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("QHelper - Paul G.");

            if (args.Length == 0)
            {
                Console.WriteLine("No file specified. Drag and drop a .zip or .rar file onto this application.");
                return;
            }

            string archivePath = args[0];

            if (!File.Exists(archivePath))
            {
                Console.WriteLine("The specified file does not exist.");
                return;
            }

            string directoryPath = Path.GetDirectoryName(archivePath);
            string baseFolderName = "QHelper";
            string destinationPath = GetUniqueFolderPath(directoryPath, baseFolderName);

            string extension = Path.GetExtension(archivePath).ToLower();

            try
            {
                if (extension == ".zip")
                {
                    ExtractZip(archivePath, destinationPath);
                }

                else if (extension == ".rar")
                {
                    ExtractRar(archivePath, destinationPath);
                }

                else
                {
                    Console.WriteLine("Unsupported file format. Only .zip and .rar are supported.");
                }

                SearchForSolutionOrProject(destinationPath);
            }

            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        static string GetUniqueFolderPath(string baseDirectory, string baseFolderName)
        {
            string folderPath = Path.Combine(baseDirectory, baseFolderName);
            int index = 1;

            while (Directory.Exists(folderPath))
            {
                folderPath = Path.Combine(baseDirectory, $"{baseFolderName} ({index})");
                index++;
            }

            return folderPath;
        }

        static void ExtractZip(string archivePath, string destinationPath)
        {
            Console.WriteLine($"Extracting ZIP file to: {destinationPath}");
            Directory.CreateDirectory(destinationPath);
            ZipFile.ExtractToDirectory(archivePath, destinationPath);
            Console.WriteLine("Extraction complete.");
        }

        static void ExtractRar(string archivePath, string destinationPath)
        {
            Console.WriteLine($"Extracting RAR file to: {destinationPath}");
            Directory.CreateDirectory(destinationPath);

            using (var archive = RarArchive.Open(archivePath))
            {
                foreach (var entry in archive.Entries)
                {
                    if (!entry.IsDirectory)
                    {
                        string destinationFilePath = Path.Combine(destinationPath, entry.Key);
                        Directory.CreateDirectory(Path.GetDirectoryName(destinationFilePath));
                        entry.WriteToFile(destinationFilePath);
                    }
                }
            }

            Console.WriteLine("Extraction complete.");
        }

        static void SearchForSolutionOrProject(string rootPath)
        {
            Console.WriteLine("Searching for .sln or .csproj files...");

            // Search for .sln files in the root folder (prioritize parent directories)
            var solutionFiles = Directory.EnumerateFiles(rootPath, "*.sln", SearchOption.TopDirectoryOnly).ToList();

            if (solutionFiles.Any())
            {
                Console.WriteLine("Found .sln file(s):");
                foreach (var file in solutionFiles)
                {
                    Console.WriteLine(file);
                    OpenWithDefaultProgram(file);
                }
                return; // Stop if .sln files are found
            }

            // If no .sln found, search for .csproj files in the root folder
            var projectFiles = Directory.EnumerateFiles(rootPath, "*.csproj", SearchOption.TopDirectoryOnly).ToList();

            if (projectFiles.Any())
            {
                Console.WriteLine("No .sln file found. Found .csproj file(s):");
                foreach (var file in projectFiles)
                {
                    Console.WriteLine(file);
                    OpenWithDefaultProgram(file);
                }
                return; // Stop if .csproj files are found
            }

            // If no .sln or .csproj found in the root, search in subdirectories
            Console.WriteLine("No .sln or .csproj file found in the root directory. Searching in subdirectories...");

            solutionFiles = Directory.EnumerateFiles(rootPath, "*.sln", SearchOption.AllDirectories).ToList();
            if (solutionFiles.Any())
            {
                Console.WriteLine("Found .sln file(s) in subdirectories:");
                foreach (var file in solutionFiles)
                {
                    Console.WriteLine(file);
                    OpenWithDefaultProgram(file);
                }
                return;
            }

            projectFiles = Directory.EnumerateFiles(rootPath, "*.csproj", SearchOption.AllDirectories).ToList();
            if (projectFiles.Any())
            {
                Console.WriteLine("No .sln file found. Found .csproj file(s) in subdirectories:");
                foreach (var file in projectFiles)
                {
                    Console.WriteLine(file);
                    OpenWithDefaultProgram(file);
                }
                return;
            }

            Console.WriteLine("No .sln or .csproj files found.");
        }

        public static void OpenWithDefaultProgram(string path)
        {
            using Process fileopener = new Process();

            fileopener.StartInfo.FileName = "explorer";
            fileopener.StartInfo.Arguments = "\"" + path + "\"";
            fileopener.Start();
        }
    }
}
