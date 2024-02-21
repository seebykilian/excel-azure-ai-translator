using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelAzureAiTranslator
{
    class ExcelManager
    {
        public static WorksheetPart worksheetPart;
        public static SheetData sheetData;
        public static Cell referenceCell;
        public static Cell destinationCell;
        public static string referenceColumnLanguage;
        public static string referenceColumnLanguageCode;
        public static string destinationColumnLanguage;
        public static string destinationColumnLanguageCode;

        public static void OpenWorksheet(string specifiedWorksheet)
        {
            try
            {
                Sheet sheet = Program.workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == specifiedWorksheet);
                worksheetPart = (WorksheetPart)Program.workbookPart.GetPartById(sheet.Id);
                sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            }
            catch
            {
                ConsoleManager.Error("The specified worksheet doesn't exist. Please try again specifying a valid worksheet name.");
            }
        }

        public static void CheckReferenceColumnFormatting(string specifiedReferenceColumn)
        {
            try
            {
                referenceCell = sheetData.Descendants<Cell>().FirstOrDefault(c => string.Compare(c.CellReference.Value, specifiedReferenceColumn + "1", StringComparison.OrdinalIgnoreCase) == 0);

                if (referenceCell == null || referenceCell.CellValue == null)
                {
                    ConsoleManager.Error("The first cell of the specified reference column is empty. Please try again with a correctly formatted reference column.");
                }

                Console.WriteLine(referenceCell.CellValue.Text);
            }
            catch
            {
                ConsoleManager.Error("An error occurred while checking the reference column formatting.");
            }
        }

        public static void CheckLanguageOrLanguageCode(string languageOrLanguageCode, bool isFromLanguage)
        {
            try
            {
                Dictionary<string, string> supportedLanguages = new Dictionary<string, string>
                {
                    { "Afrikaans", "af" }, { "Albanais", "sq" }, { "Amharique", "am" }, { "Arabe", "ar" }, { "Arménien", "hy" },
                    { "Assamais", "as" }, { "Azerbaïdjanais { Latin }", "az" }, { "Bangla", "bn" }, { "Bashkir", "ba" },
                    { "Basque", "eu" }, { "Bhojpouri", "bho" }, { "Bodo", "brx" }, { "Bosniaque { latin }", "bs" },
                    { "Bulgare", "bg" }, { "Cantonais { traditionnel }", "yue" }, { "Catalan", "ca" },
                    { "Chinois { littéraire }", "lzh" }, { "Chinois { simplifié }", "zh-Hans" }, { "Chinois traditionnel", "zh-Hant" },
                    { "chiShona", "sn" }, { "Croate", "hr" }, { "Tchèque", "cs" }, { "Danois", "da" }, { "Dari", "prs" },
                    { "Maldivien", "dv" }, { "Dogri", "doi" }, { "Néerlandais", "nl" }, { "Anglais", "en" }, { "Estonien", "et" },
                    { "Féroïen", "fo" }, { "Fidjien", "fj" }, { "Filipino", "fil" }, { "Finnois", "fi" }, { "Français", "fr" },
                    { "Français { Canada }", "fr-ca" }, { "Galicien", "gl" }, { "Géorgien", "ka" }, { "Allemand", "de" },
                    { "Grec", "el" }, { "Goudjrati", "gu" }, { "Créole haïtien", "ht" }, { "Hausa", "ha" }, { "Hébreu", "he" },
                    { "Hindi", "hi" }, { "Hmong daw { latin }", "mww" }, { "Hongrois", "hu" }, { "Islandais", "is" },
                    { "Igbo", "ig" }, { "Indonésien", "id" }, { "Inuinnaqtun", "ikt" }, { "Inuktitut", "iu" },
                    { "Inuktitut { Latin }", "iu-Latn" }, { "Irlandais", "ga" }, { "Italien", "it" }, { "Japonais", "ja" },
                    { "Kannada", "kn" }, { "Kashmiri", "ks" }, { "Kazakh", "kk" }, { "Khmer", "km" }, { "Kinyarwanda", "rw" },
                    { "Klingon", "tlh-Latn" }, { "Klingon { plqaD }", "tlh-Piqd" }, { "Konkani", "gom" }, { "Coréen", "ko" },
                    { "Kurde { central }", "ku" }, { "Kurde { Nord }", "kmr" }, { "Kirghiz { cyrillique }", "ky" }, { "Lao", "lo" },
                    { "Letton", "lv" }, { "Lituanien", "lt" }, { "Lingala", "ln" }, { "Bas sorabe", "dsb" }, { "Luganda", "lug" },
                    { "Macédonien", "mk" }, { "Maithili", "mai" }, { "Malgache", "mg" }, { "Malais { latin }", "ms" },
                    { "Malayalam", "ml" }, { "Maltais", "mt" }, { "Maori", "mi" }, { "Marathi", "mr" },
                    { "Mongole { cyrillique }", "mn-Cyrl" }, { "Mongol { traditionnel }", "mn-Mong" }, { "Myanmar", "my" },
                    { "Népalais", "ne" }, { "Norvégien", "nb" }, { "Nyanja", "nya" }, { "Odia", "or" }, { "Pachto", "ps" },
                    { "Persan", "fa" }, { "Polonais", "pl" }, { "Portugais { Brésil }", "pt" }, { "Portugais { Portugal }", "pt-pt" },
                    { "Pendjabi", "pa" }, { "Queretaro Otomi", "otq" }, { "Roumain", "ro" }, { "Rundi", "run" }, { "Russe", "ru" },
                    { "Samoan { latin }", "sm" }, { "Serbe { cyrillique }", "sr-Cyrl" }, { "Serbe { latin }", "sr-Latn" },
                    { "Sesotho", "st" }, { "Sotho du Nord", "nso" }, { "Setswana", "tn" }, { "Sindhi", "sd" },
                    { "Cingalais", "si" }, { "Slovaque", "sk" }, { "Slovène", "sl" }, { "Somali { arabe }", "so" },
                    { "Espagnol", "es" }, { "Swahili { latin }", "sw" }, { "Suédois", "sv" }, { "Tahitien", "ty" },
                    { "Tamoul", "ta" }, { "Tatar { latin }", "tt" }, { "Télougou", "te" }, { "Thaï", "th" }, { "Tibétain", "bo" },
                    { "Tigrigna", "ti" }, { "Tonga", "to" }, { "Turc", "tr" }, { "Turkmène { latin }", "tk" }, { "Ukrainien", "uk" },
                    { "Haut sorabe", "hsb" }, { "Ourdou", "ur" }, { "Ouïgour { arabe }", "ug" }, { "Ouzbek { latin }", "uz" },
                    { "Vietnamien", "vi" }, { "Gallois", "cy" }, { "Xhosa", "xh" }, { "Yoruba", "yo" },
                    { "Yucatec Maya", "yua" }, { "Zoulou", "zu" }
                };

                string lowercaseLanguageOrLanguageCode = languageOrLanguageCode.ToLower();

                foreach (var entry in supportedLanguages)
                {
                    if (lowercaseLanguageOrLanguageCode.Equals(entry.Key.ToLower()) || lowercaseLanguageOrLanguageCode.Equals(entry.Value.ToLower()))
                    {
                        if (isFromLanguage)
                        {
                            referenceColumnLanguage = entry.Key;
                            referenceColumnLanguageCode = entry.Value;
                        }
                        else
                        {
                            destinationColumnLanguage = entry.Key;
                            destinationColumnLanguageCode = entry.Value;
                        }

                        break;
                    }
                }

                if (isFromLanguage && referenceColumnLanguage == null)
                {
                    ConsoleManager.Error("The first cell of the specified reference column doesn't contain a valid language or language code. Please try again with a correctly formatted reference column.");
                }
                else if (!isFromLanguage && destinationColumnLanguage == null)
                {
                    ConsoleManager.Error("The specified destination language isn't a valid language or language code. Please try again specifying a valid destination language or language code.");
                }
            }
            catch
            {
                ConsoleManager.Error("An error occurred while checking the language or language code.");
            }
        }

        private static int CellReferenceToColumnIndex(string cellReference)
        {
            int columnNumber = -1;
            for (int i = 0; i < cellReference.Length; ++i)
            {
                if (char.IsLetter(cellReference[i]))
                {
                    columnNumber = columnNumber * 26 + (cellReference[i] - 'A' + 1);
                }
            }
            return columnNumber;
        }

        private static Cell InsertCellInWorksheet(string referenceColumn, string destinationColumn, int destinationColumnNumber)
        {
            Cell referenceCell = null;
            Cell destinationCell = null;
            try
            {
                // Find the row number of the last row containing data
                uint rowIndex = 1;
                Row lastRow = sheetData.Elements<Row>().LastOrDefault();
                if (lastRow != null)
                {
                    rowIndex = lastRow.RowIndex.Value + 1;
                }

                // Find the reference column cells
                referenceCell = sheetData.Descendants<Cell>().FirstOrDefault(c => string.Compare(c.CellReference.Value, referenceColumn + rowIndex.ToString(), StringComparison.OrdinalIgnoreCase) == 0);
                destinationCell = sheetData.Descendants<Cell>().FirstOrDefault(c => string.Compare(c.CellReference.Value, destinationColumn + rowIndex.ToString(), StringComparison.OrdinalIgnoreCase) == 0);

                // If the destination column cell doesn't exist, create it
                if (destinationCell == null)
                {
                    Row row;
                    if (lastRow == null || lastRow.RowIndex != rowIndex)
                    {
                        row = new Row() { RowIndex = rowIndex };
                        sheetData.Append(row);
                    }
                    else
                    {
                        row = lastRow;
                    }

                    string destinationColumnName = destinationColumn + rowIndex.ToString();
                    destinationCell = new Cell() { CellReference = destinationColumnName, DataType = CellValues.String };
                    row.InsertAt(destinationCell, destinationColumnNumber);
                }
            }
            catch
            {
                ConsoleManager.Error("An error occurred while inserting a cell in the worksheet.");
            }
            return destinationCell;
        }

        public static void CreateDestinationColumn(string specifiedReferenceColumn)
        {
            try
            {
                string destinationColumnLetter = specifiedReferenceColumn.Substring(0, 1);
                int destinationColumnNumber = CellReferenceToColumnIndex(specifiedReferenceColumn);

                string destinationColumnLetterPlusOne = ((char)(destinationColumnLetter[0] + 1)).ToString();
                string destinationColumnName = destinationColumnLetterPlusOne + "1";

                destinationCell = InsertCellInWorksheet(specifiedReferenceColumn, destinationColumnName, destinationColumnNumber);

                if (destinationCell == null)
                {
                    ConsoleManager.Error("An error occurred while creating the translation destination column. Please try again.");
                }
            }
            catch
            {
                ConsoleManager.Error("An error occurred while creating the translation destination column.");
            }
        }

        public static async Task TranslateReferenceColumnCellsToDestinationColumn(string specifiedReferenceColumn)
        {
            try
            {
                IEnumerable<Cell> referenceColumnCells = sheetData.Descendants<Cell>().Where(c => c.CellReference.Value.StartsWith(specifiedReferenceColumn));

                foreach (Cell cell in referenceColumnCells)
                {
                    string referenceCellValue = cell.CellValue.Text;
                    string translatedText = await AzureApiManager.TranslatorAI(referenceCellValue);
                    string destinationCellReference = cell.CellReference.Value.Replace(referenceColumnLanguageCode, destinationColumnLanguageCode);

                    Cell destinationCell = sheetData.Descendants<Cell>().FirstOrDefault(c => string.Compare(c.CellReference.Value, destinationCellReference, StringComparison.OrdinalIgnoreCase) == 0);
                    if (destinationCell == null)
                    {
                        destinationCell = InsertCellInWorksheet(referenceColumnLanguageCode, destinationCellReference, CellReferenceToColumnIndex(destinationCellReference));
                    }

                    if (destinationCell != null)
                    {
                        destinationCell.CellValue = new CellValue(translatedText);
                    }
                }

                Program.workbookPart.Workbook.Save();
            }
            catch
            {
                ConsoleManager.Error("An error occurred while translating reference column cells to destination column.");
            }
        }
    }
}