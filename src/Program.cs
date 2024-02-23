using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;

namespace ExcelAzureAiTranslator
{
    // Main class of the program managing its execution
    class Program
    {
        // Method managing the execution of the program by calling the necessary methods
        static async Task Main()
        {
            try
            {
                // Print copyright informations
                ConsoleManager.Copyright();

                // Ask the user for the path of the Excel file to be processed with Excel Azure AI Translator
                string specifiedFilePath = ConsoleManager.AskQuestion("What is the path of the Excel file you want to act on with Excel Azure AI Translator?");
                // Check the specified file path, his type and store his path in a variable  
                FileInfo filePath = Utils.CheckFilePathAndType(specifiedFilePath);

                // Open the specified Excel file
                ExcelManager.OpenFile(filePath);

                // Ask the user for the worksheet to work on
                string specifiedWorksheet = ConsoleManager.AskQuestion("Which worksheet do you want to work on?");
                // Open the specified Excel worksheet
                ExcelManager.OpenWorksheet(specifiedWorksheet);

                // Ask the user for the reference column containing the text(s) to be translated
                string specifiedReferenceColumn = ConsoleManager.AskQuestion("Which reference column containing the text(s) to be translated would you like to use?");
                // Check the formatting of the specified reference column
                ExcelManager.CheckColumnFormatting(specifiedReferenceColumn);

                // Check if the first cell of the reference column isn't null
                if (ExcelManager.firstCellOfReferenceColumn != null)
                {
                    // Check if the first cell of the reference column is a valid language or language code
                    ExcelManager.CheckLanguageOrLanguageCode(ExcelManager.firstCellOfReferenceColumn.Text, true);
                }

                // Ask the user into which language (language or language code) to translate the cells of the specified reference column
                string specifiedDestinationLanguage = ConsoleManager.AskQuestion("Into which language (language or language code) do you want to translate the cells of the specified reference column?");
                // Check if the specified destination language is a valid language or language code
                ExcelManager.CheckLanguageOrLanguageCode(specifiedDestinationLanguage, false);

                // Create the translation destination column
                ExcelManager.CreateDestinationColumn(specifiedReferenceColumn);

                // Check if the first cell of the translation reference column and the first cell of the translation destination column aren't null
                if (ExcelManager.firstCellOfReferenceColumn != null && ExcelManager.firstCellOfDestinationColumn != null)
                {
                    // Set the value of the first cell of the reference column to the translation reference language
                    ExcelManager.firstCellOfReferenceColumn.Value = ExcelManager.referenceLanguage;
                    // Set the horizontal alignment to center for the first cell of the translation reference column
                    ExcelManager.firstCellOfReferenceColumn.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    // Set the value of the first cell of the destination column to the translation destination language
                    ExcelManager.firstCellOfDestinationColumn.Value = ExcelManager.destinationLanguage;
                    // Set the horizontal alignment to center for the first cell of the translation destination column
                    ExcelManager.firstCellOfDestinationColumn.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }

                // Translate the cells of the reference column to the destination column
                await ExcelManager.TranslateReferenceColumnCellsToDestinationColumn(specifiedReferenceColumn);

                // Save changes 
                ExcelManager.SaveFile();

                // Close the program and open the specified Excel file
                Utils.CloseAndOpenFile(filePath);
            }
            catch
            {
                // If an error occurs during the program running, generate an error
                ConsoleManager.Error("An error occurred while running the program. Please try again.");
            }
        }
    }
}