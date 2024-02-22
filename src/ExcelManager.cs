// Importing the NuGet package to manage Excel reading and writing
using OfficeOpenXml;

namespace ExcelAzureAiTranslator
{
    // Additional class of the program grouping Excel methods
    class ExcelManager
    {
        // Static variables to hold Excel package, worksheet, cell references, and language information
        public static ExcelPackage? package;
        public static ExcelWorksheet? worksheet;
        public static ExcelRange? firstCellOfReferenceColumn;
        public static ExcelRange? firstCellOfDestinationColumn;
        public static string? referenceLanguage;
        public static string? referenceLanguageCode;
        public static string? destinationLanguage;
        public static string? destinationLanguageCode;

        // Method for opening an Excel file
        public static void OpenFile(FileInfo filePath)
        {
            try
            {
                // Attempt to open the file and assign it to the package variable
                using (var stream = new FileStream(filePath.FullName, FileMode.Open))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    package = new ExcelPackage(filePath);
                }
            }
            catch (IOException)
            {
                // Handle the case where the file is already open by another process
                ConsoleManager.Error("The specified Excel file is already open by another process. Please close it and try again.");
                // Return a null ExcelPackage
                package = new ExcelPackage("");
            }
        }

        // Method for opening a specific worksheet within the Excel file
        public static void OpenWorksheet(string worksheetName)
        {
            try
            {
                // Attempt to retrieve the specified worksheet and assign it to the worksheet variable
                worksheet = package?.Workbook.Worksheets[worksheetName];
            }
            catch
            {
                // Handle the case where the specified worksheet doesn't exist
                ConsoleManager.Error("The specified worksheet doesn't exist. Please try again specifying a valid worksheet name.");
            }
        }

        // Method for checking the formatting of a column within the worksheet
        public static void CheckColumnFormatting(string columnName)
        {
            try
            {
                // Retrieve the first cell of the specified column
                firstCellOfReferenceColumn = worksheet?.Cells[columnName + "1"];

                // Check if the first cell is empty or null
                if (firstCellOfReferenceColumn?.Value == null || string.IsNullOrEmpty(firstCellOfReferenceColumn.Text))
                {
                    ConsoleManager.Error("The first cell of the specified reference column is empty. Please try again with a correctly formatted reference column.");
                }
            }
            catch
            {
                // Handle any errors that occur during the column formatting check
                ConsoleManager.Error("An error occurred while checking the reference column formatting. Please try again.");
            }
        }

        // Method for checking whether a language or language code is valid
        public static void CheckLanguageOrLanguageCode(string languageOrLanguageCode, bool isFromLanguage)
        {
            try
            {
                // Dictionary containing supported languages and their codes
                Dictionary<string, string> supportedLanguages = new Dictionary<string, string> 
                {
                    { "Afrikaans", "af" }, { "Albanian", "sq" }, { "Amharic", "am" }, { "Arabic", "ar" },
                    { "Armenian", "hy" }, { "Assamese", "as" }, { "Azerbaijani (Latin)", "az" }, { "Bangla", "bn" },
                    { "Bashkir", "ba" }, { "Basque", "eu" }, { "Bhojpuri", "bho" }, { "Bodo", "brx" },
                    { "Bosnian (Latin)", "bs" }, { "Bulgarian", "bg" }, { "Cantonese (Traditional)", "yue" }, { "Catalan", "ca" },
                    { "Chinese (Literary)", "lzh" }, { "Chinese Simplified", "zh-Hans" }, { "Chinese Traditional", "zh-Hant" }, { "chiShona", "sn" },
                    { "Croatian", "hr" }, { "Czech", "cs" }, { "Danish", "da" }, { "Dari", "prs" },
                    { "Divehi", "dv" }, { "Dogri", "doi" }, { "Dutch", "nl" }, { "English", "en" },
                    { "Estonian", "et" }, { "Faroese", "fo" }, { "Fijian", "fj" }, { "Filipino", "fil" },
                    { "Finnish", "fi" }, { "French", "fr" }, { "French (Canada)", "fr-ca" }, { "Galician", "gl" },
                    { "Georgian", "ka" }, { "German", "de" }, { "Greek", "el" }, { "Gujarati", "gu" },
                    { "Haitian Creole", "ht" }, { "Hausa", "ha" }, { "Hebrew", "he" }, { "Hindi", "hi" },
                    { "Hmong Daw (Latin)", "mww" }, { "Hungarian", "hu" }, { "Icelandic", "is" }, { "Igbo", "ig" },
                    { "Indonesian", "id" }, { "Inuinnaqtun", "ikt" }, { "Inuktitut", "iu" }, { "Inuktitut (Latin)", "iu-Latn" },
                    { "Irish", "ga" }, { "Italian", "it" }, { "Japanese", "ja" }, { "Kannada", "kn" },
                    { "Kashmiri", "ks" }, { "Kazakh", "kk" }, { "Khmer", "km" }, { "Kinyarwanda", "rw" },
                    { "Klingon", "tlh-Latn" }, { "Klingon (plqaD)", "tlh-Piqd" }, { "Konkani", "gom" }, { "Korean", "ko" },
                    { "Kurdish (Central)", "ku" }, { "Kurdish (Northern)", "kmr" }, { "Kyrgyz (Cyrillic)", "ky" }, { "Lao", "lo" },
                    { "Latvian", "lv" }, { "Lithuanian", "lt" }, { "Lingala", "ln" }, { "Lower Sorbian", "dsb" },
                    { "Luganda", "lug" }, { "Macedonian", "mk" }, { "Maithili", "mai" }, { "Malagasy", "mg" },
                    { "Malay (Latin)", "ms" }, { "Malayalam", "ml" }, { "Maltese", "mt" }, { "Maori", "mi" },
                    { "Marathi", "mr" }, { "Mongolian (Cyrillic)", "mn-Cyrl" }, { "Mongolian (Traditional)", "mn-Mong" }, { "Myanmar", "my" },
                    { "Nepali", "ne" }, { "Norwegian", "nb" }, { "Nyanja", "nya" }, { "Odia", "or" },
                    { "Pashto", "ps" }, { "Persian", "fa" }, { "Polish", "pl" }, { "Portuguese (Brazil)", "pt" },
                    { "Portuguese (Portugal)", "pt-pt" }, { "Punjabi", "pa" }, { "Queretaro Otomi", "otq" }, { "Romanian", "ro" },
                    { "Rundi", "run" }, { "Russian", "ru" }, { "Samoan (Latin)", "sm" }, { "Serbian (Cyrillic)", "sr-Cyrl" },
                    { "Serbian (Latin)", "sr-Latn" }, { "Sesotho", "st" }, { "Sesotho sa Leboa", "nso" },
                    { "Setswana", "tn" }, { "Sindhi", "sd" }, { "Sinhala", "si" }, { "Slovak", "sk" },
                    { "Slovenian", "sl" }, { "Somali (Arabic)", "so" }, { "Spanish", "es" }, { "Swahili (Latin)", "sw" },
                    { "Swedish", "sv" }, { "Tahitian", "ty" }, { "Tamil", "ta" }, { "Tatar (Latin)", "tt" },
                    { "Telugu", "te" }, { "Thai", "th" }, { "Tibetan", "bo" }, { "Tigrinya", "ti" },
                    { "Tongan", "to" }, { "Turkish", "tr" }, { "Turkmen (Latin)", "tk" }, { "Ukrainian", "uk" },
                    { "Upper Sorbian", "hsb" }, { "Urdu", "ur" }, { "Uyghur (Arabic)", "ug" }, { "Uzbek (Latin)", "uz" },
                    { "Vietnamese", "vi" }, { "Welsh", "cy" }, { "Xhosa", "xh" }, { "Yoruba", "yo" },
                    { "Yucatec Maya", "yua" }, { "Zulu", "zu" }
                };

                // Convert the input to lowercase for case-insensitive comparison
                string lowercaseLanguageOrLanguageCode = languageOrLanguageCode.ToLower();

                // Iterate through the supported languages dictionary
                foreach (var entry in supportedLanguages)
                {
                    // Check if the input matches a language or language code
                    if (lowercaseLanguageOrLanguageCode.Equals(entry.Key.ToLower()) || lowercaseLanguageOrLanguageCode.Equals(entry.Value.ToLower()))
                    {
                        // Assign the language and language code based on whether it's the source or destination language
                        if (isFromLanguage)
                        {
                            referenceLanguage = entry.Key;
                            referenceLanguageCode = entry.Value;
                        }
                        else
                        {
                            destinationLanguage = entry.Key;
                            destinationLanguageCode = entry.Value;
                        }

                        break;
                    }
                }

                // Check if a valid language or language code was found
                if (isFromLanguage && referenceLanguage == null)
                {
                    ConsoleManager.Error("The first cell of the specified reference column doesn't contain a valid language or language code. Please try again with a correctly formatted reference column.");
                }
                else if (!isFromLanguage && destinationLanguage == null)
                {
                    ConsoleManager.Error("The specified destination language isn't a valid language or language code. Please try again specifying a valid destination language or language code.");
                }
            }
            catch
            {
                // Handle any errors that occur during the language or language code check
                ConsoleManager.Error("An error occurred while checking the language or language code. Please try again.");
            }
        }

        // Method for getting the column index from its name
        private static int GetColumnIndexFromName(string columnName)
        {
            try
            {
                // Initialize the variable to hold the column number
                int columnIndex = 0;

                // Loop through each character of the specified string (converting it to uppercase)
                foreach (char c in columnName.ToUpper())
                {
                    // Calculate the column number using the formula based on the position of the letter in the alphabet
                    columnIndex = columnIndex * 26 + (c - 'A' + 1);
                }

                // Return the column number
                return columnIndex;
            }
            catch
            {
                // Handle any errors that occur while getting the column index
                ConsoleManager.Error("An error occurred while getting column index from her name. Please try again.");
                return new int();
            }
        }

        // Method for creating a destination column for translation
        public static void CreateDestinationColumn(string fromColumn)
        {
            try
            {
                // Get the index of the source column
                int fromColumnIndex = GetColumnIndexFromName(fromColumn);

                // Calculate the index of the destination column
                int destinationColumnIndex = fromColumnIndex + 1;

                // Insert a new column next to the source column
                worksheet?.InsertColumn(destinationColumnIndex, 1);

                // Get the reference to the first cell of the destination column
                firstCellOfDestinationColumn = worksheet?.Cells[1, destinationColumnIndex];
            }
            catch
            {
                // Handle any errors that occur during the destination column creation
                ConsoleManager.Error("An error occurred while creating the translation destination column. Please try again.");
            }
        }

        // Method for translating cells from the reference column to the destination column
        public static async Task TranslateReferenceColumnCellsToDestinationColumn(string specifiedReferenceColumn)
        {
            try
            {
                // Get the index of the reference column
                int referenceColumnIndex = GetColumnIndexFromName(specifiedReferenceColumn);

                // Calculate the index of the destination column
                int destinationColumnIndex = referenceColumnIndex + 1;

                // Get the total number of rows in the worksheet
                int? rowCount = worksheet?.Dimension.Rows;

                // Log the start of the translation process
                ConsoleManager.Task("Translation in progress from " + referenceLanguage + " to " + destinationLanguage + ": No cell(s) translated", false);

                // Initialize a variable to track the total number of translated cells
                int totalCellsTranslated = 0;

                // Loop through each row in the worksheet
                for (int row = 2; row <= rowCount; row++)
                {
                    // Get the reference and destination cells for the current row
                    var referenceCell = worksheet?.Cells[row, referenceColumnIndex];
                    var destinationCell = worksheet?.Cells[row, destinationColumnIndex];

                    // Check if the reference cell is not empty
                    if (referenceCell?.Value != null)
                    {
                        // Translate the text in the reference cell and assign it to the destination cell
                        string translatedText = await AzureApiManager.TranslatorAI(referenceCell.Text);

                        if (destinationCell != null)
                        {
                            destinationCell.Value = translatedText;
                        }

                        // Increment the count of translated cells
                        totalCellsTranslated += 1;

                        // Log the progress of the translation
                        ConsoleManager.Task("Translation in progress from " + referenceLanguage + " to " + destinationLanguage + ": " + totalCellsTranslated + " cell(s) translated", true);
                    }
                }

                // Automatically adjusts the width of the reference column based on its content
                worksheet?.Column(referenceColumnIndex).AutoFit();
                // Automatically adjusts the width of the destination column based on its content
                worksheet?.Column(destinationColumnIndex).AutoFit();
            }
            catch
            {
                // Handle any errors that occur during the translation process
                ConsoleManager.Error("An error occurred while translating reference column cells to translation destination column. Please try again.");
            }
        }

        // Method for saving changes to the Excel file
        public static void SaveFile()
        {
            try
            {
                // Save the changes to the Excel package
                package?.Save();
                ConsoleManager.Success("The changes were successfully saved.");
            }
            catch
            {
                // Handle any errors that occur while saving changes
                ConsoleManager.Error("An error occurred while saving changes. Please try again.");
            }
        }
    }
}