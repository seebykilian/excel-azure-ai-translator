using System.Text;
using Newtonsoft.Json;
using DotNetEnv;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Runtime.InteropServices;

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

    // Importer les fonctions API Windows nécessaires
    [DllImport("Kernel32")]
    private static extern bool SetConsoleCtrlHandler(EventHandler handler, bool add);

    // Définir le type de délégué pour le gestionnaire d'événements
    private delegate bool EventHandler(CtrlType sig);

    // Définir les types d'événements de fermeture de la console
    private enum CtrlType
    {
        CTRL_C_EVENT = 0,
        CTRL_CLOSE_EVENT = 2,
        CTRL_LOGOFF_EVENT = 5,
        CTRL_SHUTDOWN_EVENT = 6
    }

    static Application application;
    static Workbook workbook;
    static Worksheet worksheet;
    static Excel.Range firstCellOfReferenceColumn;

    static async Task Main()
    {
        LoadEnvironmentVariables();

        // Ajouter le gestionnaire d'événements pour la fermeture de la console
        SetConsoleCtrlHandler(ConsoleCtrlHandler, true);

        GenerateConsoleCopyright();

        string specifiedFilePath = GenerateConsoleQuestion("What is the path of the Excel file you want to act on with Excel Azure AI Translator?");
        CheckFilePathAndType(specifiedFilePath);
        OpenExcelWorkbook(specifiedFilePath);

        string specifiedWorksheet = GenerateConsoleQuestion("Which worksheet do you want to work on?");
        OpenExcelWorksheet(specifiedWorksheet);

        string specifiedReferenceColumn = GenerateConsoleQuestion("Which reference column containing the text(s) to be translated would you like to use?");
        CheckReferenceColumnFormatting(specifiedReferenceColumn);

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

        // Initialiser la variable referenceColumnLanguageCode
        string referenceColumnLanguage = null;
        string referenceColumnLanguageCode = null;

        // Convertir la valeur de firstCellOfReferenceColumn en minuscules pour la comparaison
        string firstCellOfReferenceColumnValue = firstCellOfReferenceColumn.Value.ToLower();

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
                referenceColumnLanguage = language;
                referenceColumnLanguageCode = languageCode;
                break;
            }
        }

        firstCellOfReferenceColumn.Value = referenceColumnLanguage;

        if (!isValidFromLanguage)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("The first cell of the specified reference column doesn't contain a valid language or language code. Please try again with a correctly formatted reference column.");
            workbook.Close(false);
            return;
        }

        Console.WriteLine("\nInto which language (language or language code) do you want to translate the cells of the specified reference column?");
        string specifiedDestinationLanguage = Console.ReadLine();

        string destinationColumnLanguage = null;
        string destinationColumnLanguageCode = null;

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
                destinationColumnLanguage = language;
                destinationColumnLanguageCode = languageCode;
                break;
            }
        }

        if (!isValidDestinationLanguage)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("The specified destination language isn't a valid language or language code. Please try again specifying a valid destination language or language code.");
            workbook.Close(false);
            return;
        }

        int destinationColumn = GetColumnNumber(specifiedReferenceColumn) + 1;

        // Insérer la colonne à droite de la colonne de référence
        Excel.Range destinationColumnPosition = worksheet.Columns[destinationColumn];
        destinationColumnPosition.Insert();

        // Obtenir la première cellule de la colonne de destination et y mettre la valeur
        Excel.Range firstCellOfDestinationColumn = (Excel.Range)worksheet.Cells[1, destinationColumn];

        firstCellOfDestinationColumn.Value = destinationColumnLanguage;

        workbook.Save();
        workbook.Close(false);

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

            // Libérer les ressources COM
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
        }
    }

    // Méthode pour charger les variables d'environnement en fonction de l'environnement d'exécution
    static void LoadEnvironmentVariables()
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
    }

    // Méthode appelée lorsqu'un événement de fermeture de la console est détecté
    private static bool ConsoleCtrlHandler(CtrlType sig)
    {
        // Si un classeur Excel est lancé
        if (workbook != null)
        {
            // Le fermer sans enregistrer
            workbook.Close(false);
        }

        // Si l'application Excel est lancée
        if (application != null)
        {
            // La fermer
            application.Quit();

            // Tuer le processus Excel de manière forcée
            foreach (Process process in Process.GetProcessesByName("Excel"))
            {
                process.Kill();
            }
        }

        // Indiquer que l'événement a été traité et permettre à l'application de se fermer
        return true;
    }

    // Méthode pour générer la section Copyright dans la console
    static void GenerateConsoleCopyright()
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine(" ___   _  ___  _______  ___      _______  __   __  _______  _______ ");
        Console.ForegroundColor = ConsoleColor.DarkRed;
        Console.WriteLine("|   | | ||   ||       ||   |    |   _   ||  | |  ||  _    ||       |");
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("|   |_| ||   ||    _  ||   |    |  |_|  ||  |_|  || | |   ||___    |");
        Console.ForegroundColor = ConsoleColor.DarkRed;
        Console.WriteLine("|      _||   ||   |_| ||   |    |       ||       || | |   | ___|   |");
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("|     |_ |   ||    ___||   |___ |       ||_     _|| |_|   ||___    |");
        Console.ForegroundColor = ConsoleColor.DarkRed;
        Console.WriteLine("|    _  ||   ||   |    |       ||   _   |  |   |  |       | ___|   |");
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("|___| |_||___||___|    |_______||__| |__|  |___|  |_______||_______|");
        Console.ForegroundColor = ConsoleColor.White;
        Console.Write("\nIf you need help formatting your Excel file to translate or using Excel Azure AI Translator, visit the documentation available at the following address: https://github.com/Kiplay03/excel-azure-ai-translator\n");
    }

    // Méthode pour générer une question dans la console et retourner la réponse de l'utilisateur 
    static string GenerateConsoleQuestion(string questionContent)
    {
        // Poser la question dans la console avec le contenu spécifié 
        Console.WriteLine("\n" + questionContent);
        // Enregister dans une variable la réponse de l'utilisateur
        string userResponse = Console.ReadLine();
        // Retourner la variable contenant la réponse
        return userResponse;
    }

    // Méthode pour générer une erreur dans la console avec un message spécifique, fermer un classeur Excel s'il est lancé, arrêter l'application Excel si elle est lancée et arrêter l'exécution du programme
    static void GenerateConsoleError(string errorContent)
    {
        // Mettre la couleur du texte de la console en rouge
        Console.ForegroundColor = ConsoleColor.Red;
        // Retourner dans la console le message d'erreur
        Console.WriteLine(errorContent);

        // Si un classeur Excel est lancé
        if (workbook != null)
        {
            // Le fermer sans enregistrer
            workbook.Close(false);
        }

        // Si l'application Excel est lancée
        if (application != null)
        {
            // La fermer
            application.Quit();

            // Tuer le processus Excel de manière forcée
            foreach (Process process in Process.GetProcessesByName("Excel"))
            {
                process.Kill();
            }
        }

        // Arrêter l'exécution du programme
        Environment.Exit(1);
    }

    // Méthode pour vérifier si le chemin d'accès fourni par l'utilisateur existe et si le fichier vers lequel il renvoie est un fichier Excel
    static void CheckFilePathAndType(string specifiedFilePath)
    {
        // Vérifier si le fichier existe sans ajouter l'extension
        if (!File.Exists(specifiedFilePath))
        {
            // Ajouter l'extension Excel et vérifier de nouveau
            specifiedFilePath += ".xlsx";
            if (!File.Exists(specifiedFilePath))
            {
                // S'il n'existe pas, générer une erreur
                GenerateConsoleError("The specified file path doesn't exist or the specified file isn't an Excel file. Please try again specifying a valid Excel file path.");
            }
        }

        // Vérifier si le fichier a une extension Excel valide
        if (!Path.GetExtension(specifiedFilePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            // S'il n'en a pas, générer une erreur
            GenerateConsoleError("The specified file isn't an Excel file. Please try again specifying a valid Excel file path.");
        }
    }

    // Méthode pour tenter d'ouvrir le classeur Excel fourni par l'utilisateur et retourner une instance de ce dernier
    static void OpenExcelWorkbook(string specifiedFilePath)
    {
        // Définir dans la variable initialisée une instance de l'application Excel 
        application = new Application();

        // Tenter d'ouvrir le classeur Excel fourni par l'utilisateur
        try
        {
            // Définir dans la variable initialisée le classeur Excel fourni par l'utilisateur
            workbook = application.Workbooks.Open(specifiedFilePath);
        }
        catch
        {
            // Si une erreur se produit durant l'ouverture, générer une erreur
            GenerateConsoleError("An error occurred while opening the Excel workbook. Please make sure the file path is correct and try again.");
        }
    }

    // Méthode pour tenter d'ouvrir la feuille de travail Excel fournie par l'utilisateur
    static void OpenExcelWorksheet(string specifiedWorksheet)
    {
        // Tenter d'ouvrir la feuille de travail Excel fournie par l'utilisateur
        try
        {
            // Définir dans la variable initialisée la feuille de travail Excel fournie par l'utilisateur
            worksheet = (Excel.Worksheet)workbook.Sheets[specifiedWorksheet];
        }
        catch
        {
            // Si une erreur se produit durant l'ouverture, générer une erreur
            GenerateConsoleError("The specified worksheet doesn't exist. Please try again specifying a valid worksheet name.");
        }
    }

    // Méthode pour vérifier l'existance de la colonne de référence fournie par l'utilisateur et son formattage
    static void CheckReferenceColumnFormatting(string specifiedReferenceColumn)
    {
        // Tenter de récupérer la première cellule de la colonne de référence fournie par l'utilisateur
        try
        {
            // Récupérer la première cellule de la colonne de référence en convertisant le ou les lettre(s) l'identifiant en nombre
            firstCellOfReferenceColumn = (Excel.Range)worksheet.Cells[1, GetColumnNumber(specifiedReferenceColumn)];
        }
        catch
        {
            // Si une erreur se produit durant la récupération, générer une erreur
            GenerateConsoleError("The specified reference column isn't valid. Please try again specifying a valid reference column letter.");
        }

        // Vérifier si la première cellule de la colonne de référence est vide
        if (firstCellOfReferenceColumn.Value == null)
        {
            // Si elle est vide, générer une erreur
            GenerateConsoleError("The first cell of the specified reference column is empty. Please try again with a correctly formatted reference column.");
        }
    }

    // Méthode pour obtenir le numéro d'une colonne Excel à partir de la lettre ou de la combinaison de lettres fournie par l'utilisateur
    static int GetColumnNumber(string specifiedColumn)
    {
        // Initialiser la variable pour accueillir le numéro de la colonne 
        int columnNumber = 0;

        // Parcourir chaque caractère de la chaîne de caractères spécifiée (en la convertissant en majuscules)
        foreach (char c in specifiedColumn.ToUpper())
        {
            // Calculer le numéro de colonne en utilisant la formule basée sur la position de la lettre dans l'alphabet
            columnNumber = columnNumber * 26 + (c - 'A' + 1);
        }

        // Retourner le numéro de la colonne
        return columnNumber;
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
