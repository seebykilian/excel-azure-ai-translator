// Importing the NuGet package handling environment variables
using DotNetEnv;
using System.Diagnostics;
using System.IO;

namespace ExcelAzureAiTranslator
{
    // Additional class of the program grouping the general methods
    class Utils
    {
        // Method for loading environment variables based on the execution environment
        private static void LoadEnvironmentVariables()
        {
            try
            {
                // Check if the .env file exists in the current execution environment
                if (File.Exists(".env"))
                {
                    // If it exists, load the environment variables from the .env file
                    Env.Load();
                }
                else
                {
                    // If it doesn't exist, load the environment variables from the .env file in the root directory
                    Env.Load("../../../.env");
                }
            }
            catch
            {
                // If an error occurs during the environment variables loading, generate an error
                ConsoleManager.Error("An error occurred while loading environment variables. Please try again.");
            }
        }

        // Method for getting environment variable value 
        public static string GetEnvironmentVariableValue(string environmentVariableKey)
        {
            try
            {
                // Load environment variables
                LoadEnvironmentVariables();
                // Return the environment variable value got from her key
                return Env.GetString(environmentVariableKey);
            } 
            catch
            {
                // If an error occurs during the environment variable value getting process, generate an error
                ConsoleManager.Error("An error occurred while loading environment variables.");
                // Return null
                return "";
            }
        }

        // Method for verifying if the path provided by the user exists and if the file it refers to is an Excel file
        public static FileInfo CheckFilePathAndType(string specifiedFilePath)
        {
            try
            {
                // Check if the file exists without adding the extension
                if (!File.Exists(specifiedFilePath))
                {
                    // Add the Excel extension and check again
                    specifiedFilePath += ".xlsx";
                    if (!File.Exists(specifiedFilePath))
                    {
                        // If it still doesn't exist, generate an error
                        ConsoleManager.Error("The specified file path doesn't exist or the specified file isn't an Excel file. Please try again specifying a valid Excel file path.");
                    }
                }

                // Check if the file has a valid Excel extension
                if (!Path.GetExtension(specifiedFilePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    // If it doesn't have, generate an error
                    ConsoleManager.Error("The specified file isn't an Excel file. Please try again specifying a valid Excel file path.");
                }

                // Return the file path in the correct format
                return new FileInfo(specifiedFilePath);
            }
            catch
            {
                // If an error occurs during file path and type checking, generate an error
                ConsoleManager.Error("An error occurred while checking the file path and type.");
                // Return a null file path
                return new FileInfo("");
            }
        }

        // Method for closing the program and opening the specified Excel file
        public static void CloseAndOpenFile(FileInfo filePath)
        {
            try
            {
                // Print a message in the console to inform the user to click any key to close the program and open the specified Excel file
                ConsoleManager.Message("Press any key to close Excel Azure AI Translator and open the specified Excel file...");

                // Wait for the user to press any key
                Console.ReadKey();

                // Instantiate ProcessStartInfo
                ProcessStartInfo startInfo = new ProcessStartInfo();
                // Specify that the file should open with the default program associated with its type
                startInfo.UseShellExecute = true;
                // Specify that the file to open is the one provided by the user
                startInfo.FileName = filePath.FullName;
                // Open the specified Excel file
                Process.Start(startInfo);

                // Quit the program and close the command prompt window
                Environment.Exit(0);
            }
            catch
            {
                // If an error occurs during program closing and specified Excel file opening, generate an error
                ConsoleManager.Error("An error occurred while closing the program and opening the specified Excel file. Please try again.");
            }
        }
    } 
} 