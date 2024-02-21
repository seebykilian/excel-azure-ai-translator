using DocumentFormat.OpenXml.Packaging;
using DotNetEnv;

namespace ExcelAzureAiTranslator
{
    class Program
    {
        public static WorkbookPart workbookPart;

        // Method to start the program execution
        static async Task Main()
        {
            // Load environment variables
            LoadEnvironmentVariables();

            // Display copyright information
            ConsoleManager.Copyright();

            // Ask the user for the path of the Excel file to be processed with Excel Azure AI Translator
            string specifiedFilePath = ConsoleManager.AskQuestion("What is the path of the Excel file you want to act on with Excel Azure AI Translator?");
            // Check the specified file path and type
            string filePath = CheckFilePathAndType(specifiedFilePath);

            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, true))
                {
                    workbookPart = document.WorkbookPart;
                    // Ask the user for the worksheet to work on
                    string specifiedWorksheet = ConsoleManager.AskQuestion("Which worksheet do you want to work on?");
                    // Open the specified Excel worksheet
                    ExcelManager.OpenWorksheet(specifiedWorksheet);

                    // Ask the user for the reference column containing the text(s) to be translated
                    string specifiedReferenceColumn = ConsoleManager.AskQuestion("Which reference column containing the text(s) to be translated would you like to use?");
                    // Check the formatting of the specified reference column
                    ExcelManager.CheckReferenceColumnFormatting(specifiedReferenceColumn);
                    // Check the language or language code of the reference column
                    ExcelManager.CheckLanguageOrLanguageCode(ExcelManager.referenceCell.CellValue.Text, true);

                    // Ask the user into which language (language or language code) to translate the cells of the specified reference column
                    string specifiedDestinationLanguage = ConsoleManager.AskQuestion("Into which language (language or language code) do you want to translate the cells of the specified reference column?");
                    // Check the language or language code of the destination column
                    ExcelManager.CheckLanguageOrLanguageCode(specifiedDestinationLanguage, false);

                    // Create the destination column
                    ExcelManager.CreateDestinationColumn(specifiedReferenceColumn);

                    // Set the value of the first cell of the reference column
                    ExcelManager.referenceCell.CellValue.Text = ExcelManager.referenceColumnLanguage;
                    // Set the value of the first cell of the destination column
                    ExcelManager.destinationCell.CellValue.Text = ExcelManager.destinationColumnLanguage;

                    // Translate the cells of the reference column to the destination column
                    await ExcelManager.TranslateReferenceColumnCellsToDestinationColumn(specifiedReferenceColumn);
                }
            }
            catch (IOException)
            {
                // Handle the case where the file is already open by another process
                ConsoleManager.Error("The specified Excel file is already open by another process. Please close it and try again.");
            }
        }

        // Method for loading environment variables based on the execution environment
        public static void LoadEnvironmentVariables()
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
                ConsoleManager.Error("An error occurred while loading environment variables.");
            }
        }

        // Method for verifying if the path provided by the user exists and if the file it refers to is an Excel file
        static string CheckFilePathAndType(string specifiedFilePath)
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

                return specifiedFilePath;
            }
            catch
            {
                // If an error occurs during file path and type checking, generate an error
                ConsoleManager.Error("An error occurred while checking the file path and type.");
                return null;
            }
        }
    }
}