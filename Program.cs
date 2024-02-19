using System.Text;
using Newtonsoft.Json;
using DotNetEnv;
using Excel = Microsoft.Office.Interop.Excel;
using static Program;

class Program
{
    public class TranslationResult
    {
        public DetectedLanguage DetectedLanguage { get; set; }
        public TextResult SourceText { get; set; }
        public Translation[] Translations { get; set; }
    }

    public class DetectedLanguage
    {
        public string Language { get; set; }
        public float Score { get; set; }
    }

    public class TextResult
    {
        public string Text { get; set; }
        public string Script { get; set; }
    }

    public class Translation
    {
        public string Text { get; set; }
        public TextResult Transliteration { get; set; }
        public string To { get; set; }
        public Alignment Alignment { get; set; }
        public SentenceLength SentLen { get; set; }
    }

    public class Alignment
    {
        public string Proj { get; set; }
    }

    public class SentenceLength
    {
        public int[] SrcSentLen { get; set; }
        public int[] TransSentLen { get; set; }
    }

    static async Task Main()
    {
        // Vérifier si le fichier .env existe dans l'environnement d'exécution actuel
        if (File.Exists(".env"))
        {
            // S'il existe, charger les variables d'environnement depuis le fichier .env
            Env.Load();
        }
        else
        {
            // S'il n'existe pas, charger les variables d'environnement depuis le fichier .env du dossier racine
            Env.Load("../../../.env");
        }

        // Restauration de la couleur d'origine
        Console.ForegroundColor = ConsoleColor.White;
        Console.Write(" ___   _  ___  _______  ___      _______  __   __  _______  _______ \n|   | | ||   ||       ||   |    |   _   ||  | |  ||  _    ||       |\n|   |_| ||   ||    _  ||   |    |  |_|  ||  |_|  || | |   ||___    |\n|      _||   ||   |_| ||   |    |       ||       || | |   | ___|   |\n|     |_ |   ||    ___||   |___ |       ||_     _|| |_|   ||___    |\n|    _  ||   ||   |    |       ||   _   |  |   |  |       | ___|   |\n|___| |_||___||___|    |_______||__| |__|  |___|  |_______||_______|\n");
        Console.Write("\nIf you need help formatting your Excel file to translate or using Excel Azure AI Translator, visit the documentation available at the following address: https://github.com/Kiplay03/excel-azure-ai-translator\n");


        Console.WriteLine("\nWhat is the path of the Excel file you want to act on with Excel Azure AI Translator?");
        string specifiedFilePath = Console.ReadLine();

        // Vérifier si le fichier existe sans ajouter l'extension
        if (!File.Exists(specifiedFilePath))
        {
            // Ajouter l'extension ".xlsx" et vérifier à nouveau
            specifiedFilePath += ".xlsx";
            if (!File.Exists(specifiedFilePath))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("The specified file path doesn't exist. Please try again specifying a valid Excel file path.");
                return;
            }
        }

        // Vérifier si le fichier a une extension valide
        if (!Path.GetExtension(specifiedFilePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("The file specified isn't an Excel file. Please try again specifying a valid Excel file path.");
            return;
        }

        // Create an instance of Excel application
        Excel.Application excelApp = new Excel.Application();

        // Open the Excel workbook
        Excel.Workbook workbook = null;
        try
        {
            workbook = excelApp.Workbooks.Open(specifiedFilePath);
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("An error occurred while opening the Excel workbook. Please make sure the file path is correct and try again.");
            workbook.Close(false);
            excelApp.Quit();
            return;
        }

        Console.WriteLine("\nWhich worksheet do you want to work on?");
        string specifiedWorksheet = Console.ReadLine();

        // Get the specified worksheet
        Excel.Worksheet worksheet = null;
        if (workbook != null)
        {
            try
            {
                worksheet = (Excel.Worksheet)workbook.Sheets[specifiedWorksheet];
            }
            catch (Exception)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("The specified worksheet doesn't exist. Please try again specifying a valid worksheet name.");
                workbook.Close(false);
                excelApp.Quit();
                return;
            }
        }

        List<(string, string)> supportedLanguages = new List<(string, string)>
        {
            ("Afrikaans", "af"), ("Albanais", "sq"), ("Amharique", "am"), ("Arabe", "ar"), ("Arménien", "hy"),
            ("Assamais", "as"), ("Azerbaïdjanais (Latin)", "az"), ("Bangla", "bn"), ("Bashkir", "ba"),
            ("Basque", "eu"), ("Bhojpouri", "bho"), ("Bodo", "brx"), ("Bosniaque (latin)", "bs"),
            ("Bulgare", "bg"), ("Cantonais (traditionnel)", "yue"), ("Catalan", "ca"),
            ("Chinois (littéraire)", "lzh"), ("Chinois (simplifié)", "zh-Hans"), ("Chinois traditionnel", "zh-Hant"),
            ("chiShona", "sn"), ("Croate", "hr"), ("Tchèque", "cs"), ("Danois", "da"), ("Dari", "prs"),
            ("Maldivien", "dv"), ("Dogri", "doi"), ("Néerlandais", "nl"), ("Anglais", "en"), ("Estonien", "et"),
            ("Féroïen", "fo"), ("Fidjien", "fj"), ("Filipino", "fil"), ("Finnois", "fi"), ("Français", "fr"),
            ("Français (Canada)", "fr-ca"), ("Galicien", "gl"), ("Géorgien", "ka"), ("Allemand", "de"),
            ("Grec", "el"), ("Goudjrati", "gu"), ("Créole haïtien", "ht"), ("Hausa", "ha"), ("Hébreu", "he"),
            ("Hindi", "hi"), ("Hmong daw (latin)", "mww"), ("Hongrois", "hu"), ("Islandais", "is"),
            ("Igbo", "ig"), ("Indonésien", "id"), ("Inuinnaqtun", "ikt"), ("Inuktitut", "iu"),
            ("Inuktitut (Latin)", "iu-Latn"), ("Irlandais", "ga"), ("Italien", "it"), ("Japonais", "ja"),
            ("Kannada", "kn"), ("Kashmiri", "ks"), ("Kazakh", "kk"), ("Khmer", "km"), ("Kinyarwanda", "rw"),
            ("Klingon", "tlh-Latn"), ("Klingon (plqaD)", "tlh-Piqd"), ("Konkani", "gom"), ("Coréen", "ko"),
            ("Kurde (central)", "ku"), ("Kurde (Nord)", "kmr"), ("Kirghiz (cyrillique)", "ky"), ("Lao", "lo"),
            ("Letton", "lv"), ("Lituanien", "lt"), ("Lingala", "ln"), ("Bas sorabe", "dsb"), ("Luganda", "lug"),
            ("Macédonien", "mk"), ("Maithili", "mai"), ("Malgache", "mg"), ("Malais (latin)", "ms"),
            ("Malayalam", "ml"), ("Maltais", "mt"), ("Maori", "mi"), ("Marathi", "mr"),
            ("Mongole (cyrillique)", "mn-Cyrl"), ("Mongol (traditionnel)", "mn-Mong"), ("Myanmar", "my"),
            ("Népalais", "ne"), ("Norvégien", "nb"), ("Nyanja", "nya"), ("Odia", "or"), ("Pachto", "ps"),
            ("Persan", "fa"), ("Polonais", "pl"), ("Portugais (Brésil)", "pt"), ("Portugais (Portugal)", "pt-pt"),
            ("Pendjabi", "pa"), ("Queretaro Otomi", "otq"), ("Roumain", "ro"), ("Rundi", "run"), ("Russe", "ru"),
            ("Samoan (latin)", "sm"), ("Serbe (cyrillique)", "sr-Cyrl"), ("Serbe (latin)", "sr-Latn"),
            ("Sesotho", "st"), ("Sotho du Nord", "nso"), ("Setswana", "tn"), ("Sindhi", "sd"),
            ("Cingalais", "si"), ("Slovaque", "sk"), ("Slovène", "sl"), ("Somali (arabe)", "so"),
            ("Espagnol", "es"), ("Swahili (latin)", "sw"), ("Suédois", "sv"), ("Tahitien", "ty"),
            ("Tamoul", "ta"), ("Tatar (latin)", "tt"), ("Télougou", "te"), ("Thaï", "th"), ("Tibétain", "bo"),
            ("Tigrigna", "ti"), ("Tonga", "to"), ("Turc", "tr"), ("Turkmène (latin)", "tk"), ("Ukrainien", "uk"),
            ("Haut sorabe", "hsb"), ("Ourdou", "ur"), ("Ouïgour (arabe)", "ug"), ("Ouzbek (latin)", "uz"),
            ("Vietnamien", "vi"), ("Gallois", "cy"), ("Xhosa", "xh"), ("Yoruba", "yo"),
            ("Yucatec Maya", "yua"), ("Zoulou", "zu")
        };

        Console.WriteLine("\nWhich reference column (letter) containing the text(s) to be translated would you like to use?");
        string specifiedReferenceColumn = Console.ReadLine();

        Excel.Range firstCellOfReferenceColumn = null;
        string firstCellOfReferenceColumnValue = null;
        try
        {
            firstCellOfReferenceColumn = (Excel.Range)worksheet.Cells[1, specifiedReferenceColumn];
            firstCellOfReferenceColumnValue = firstCellOfReferenceColumn.Value;
        }
        catch (Exception)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("The specified reference column isn't valid. Please try again specifying a valid reference column letter.");
            workbook.Close(false);
            excelApp.Quit();
            return;
        }

        if (firstCellOfReferenceColumnValue == null)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("The specified reference column isn't formatted correctly. Please try again with a correctly formatted reference column.");
            workbook.Close(false);
            excelApp.Quit();
            return;
        }

        // Initialiser la variable referenceColumnLanguageCode
        string referenceColumnLanguage = null;
        string referenceColumnLanguageCode = null;

        // Convertir la valeur de firstCellOfReferenceColumn en minuscules pour la comparaison
        firstCellOfReferenceColumnValue = firstCellOfReferenceColumnValue.ToLower();

        // Vérifier si la valeur de firstCellOfReferenceColumn correspond à une langue ou à un code langue
        bool isValidFromLanguage = false;
        foreach ((string language, string languageCode) in supportedLanguages)
        {
            string lowercaseLanguage = language.ToLower();
            string lowercaseLanguageCode = languageCode.ToLower();

            if (firstCellOfReferenceColumnValue.Equals(lowercaseLanguage) ||
                firstCellOfReferenceColumnValue.Equals(lowercaseLanguageCode))
            {
                isValidFromLanguage = true;
                firstCellOfReferenceColumn.Value = language;
                referenceColumnLanguage = language;
                referenceColumnLanguageCode = languageCode;
                break;
            }
        }

        if (!isValidFromLanguage)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("The first cell of the specified reference column doesn't contain a valid language or language code. Please try again with a correctly formatted reference column.");
            workbook.Close(false);
            excelApp.Quit();
            return;
        }

        Console.WriteLine("\nInto which language (language or language code) do you want to translate the cells of the specified reference column?");
        string specifiedDestinationLanguage = Console.ReadLine();

        string destinationLanguage = null;
        string destinationLanguageCode = null;

        // Convertir la valeur de firstCellOfReferenceColumn en minuscules pour la comparaison
        specifiedDestinationLanguage = specifiedDestinationLanguage.ToLower();

        // Vérifier si la valeur de firstCellOfReferenceColumn correspond à une langue ou à un code langue
        bool isValidDestinationLanguage = false;
        foreach ((string language, string languageCode) in supportedLanguages)
        {
            string lowercaseLanguage = language.ToLower();
            string lowercaseLanguageCode = languageCode.ToLower();

            if (specifiedDestinationLanguage.Equals(lowercaseLanguage) ||
                specifiedDestinationLanguage.Equals(lowercaseLanguageCode))
            {
                isValidDestinationLanguage = true;
                destinationLanguage = language;
                destinationLanguageCode = languageCode;
                break;
            }
        }

        if (!isValidDestinationLanguage)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("The specified destination language isn't a valid language or language code. Please try again specifying a valid destination language or language code.");
            workbook.Close(false);
            excelApp.Quit();
            return;
        }

        try
        {
            // Parcourir la colonne I de la ligne 2 à la ligne 1280
            for (int i = 2; i <= 1038; i++)
            {
                Excel.Range keyCell = (Excel.Range)worksheet.Cells[i, "B"];
                Excel.Range frCell = (Excel.Range)worksheet.Cells[i, "C"];
                Excel.Range enCell = (Excel.Range)worksheet.Cells[i, "D"];
                Excel.Range esCell = (Excel.Range)worksheet.Cells[i, "E"];
                Excel.Range ptCell = (Excel.Range)worksheet.Cells[i, "F"];
                Excel.Range nlCell = (Excel.Range)worksheet.Cells[i, "G"];

                if (keyCell.Value != null && frCell.Value != null)
                {
                    List<string> results = new List<string>();

                    if (esCell.Value != null)
                    {
                        results.Add(await TranslateCell(frCell.Value, "fr", "pt-pt", "nl", null));
                    }
                    else
                    {
                        results.Add(await TranslateCell(frCell.Value, "fr", "es", "pt-pt", "nl"));
                    }

                    var result = results.First();
                    TranslationResult[] deserializedOutput = JsonConvert.DeserializeObject<TranslationResult[]>(result);

                    foreach (TranslationResult o in deserializedOutput)
                    {
                        foreach (Translation t in o.Translations)
                        {
                            if (t.To == "es")
                            {
                                esCell.Value = t.Text;
                                Console.WriteLine($"[E{i}] -> [{t.To}] {t.Text}");
                            }
                            else if (t.To == "pt-PT")
                            {
                                ptCell.Value = t.Text;
                                Console.WriteLine($"[F{i}] -> [{t.To}] {t.Text}");
                            }
                            else if (t.To == "nl")
                            {
                                nlCell.Value = t.Text;
                                Console.WriteLine($"[G{i}] -> [{t.To}] {t.Text}");
                            }
                        }
                    }
                }
            }
        }
        finally
        {
            // Fermer le classeur Excel et quitter l'application Excel
            workbook.Close(true);
            excelApp.Quit();

            // Libérer les ressources COM
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }
    }
    static async Task<string> TranslateCell(string text, string fromLanguage, string firstLanguage, string secondLanguage, string thirdLanguage)
    {
        string secondLanguageParse = null;
        string thirdLanguageParse = null;

        if (secondLanguage != null)
        {
            secondLanguageParse = $"&to={secondLanguage}";
        }

        if (thirdLanguage != null)
        {
            thirdLanguageParse = $"&to={thirdLanguage}";
        }

        using (var client = new HttpClient())
        using (var request = new HttpRequestMessage())
        {
            object[] body = new object[] { new { Text = text } };
            var requestBody = JsonConvert.SerializeObject(body);

            string requestUri = $"https://api.cognitive.microsofttranslator.com/translate?api-version=3.0&from={fromLanguage}&to={firstLanguage}";

            // Ajouter les langages supplémentaires s'ils sont définis
            if (secondLanguageParse != null)
            {
                requestUri += secondLanguageParse;
            }

            if (thirdLanguageParse != null)
            {
                requestUri += thirdLanguageParse;
            }

            request.Method = HttpMethod.Post;
            request.RequestUri = new Uri(requestUri);
            request.Content = new StringContent(requestBody, Encoding.UTF8, "application/json");
            request.Headers.Add("Ocp-Apim-Subscription-Key", Env.GetString("azureApiKey"));
            request.Headers.Add("Ocp-Apim-Subscription-Region", Env.GetString("azureApiRegion"));

            // Envoyer la requête et obtenir la réponse.
            HttpResponseMessage response = await client.SendAsync(request).ConfigureAwait(false);
            // Lire la réponse sous forme de chaîne.
            string result = await response.Content.ReadAsStringAsync();
            return result;
        }
    }
}
