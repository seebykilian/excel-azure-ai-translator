namespace ExcelAzureAiTranslator
{
    class ConsoleManager
    {
        // Method for generating the Copyright section in the console
        public static void Copyright()
        {
            try
            {
                // Print the ASCII art copyright message
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
                // Print the documentation link message
                Console.Write("\nIf you need help formatting your Excel file to translate or using Excel Azure AI Translator, visit the documentation available at the following address: https://github.com/Kiplay03/excel-azure-ai-translator\n");
            }
            catch
            {
                // If an error occurs while generating the copyright section, generate an error
                Error("An error occurred while generating the copyright section.");
            }
        }

        // Method to ask a console question and return the user's response
        public static string AskQuestion(string questionContent)
        {
            // Print the question in the console with the specified content
            Console.WriteLine("\n" + questionContent);
            try
            {
                // Read the user's response from the console and store it in a variable
                string userResponse = Console.ReadLine();
                // Return the variable containing the user's response
                return userResponse;
            }
            catch
            {
                // If an error occurs during reading the user's response, generate an error
                Error("An error occurred while reading user response.");
                return null;
            }
        }

        // Method for generating an error in the console with a specific message and stopping the program
        public static void Error(string errorContent)
        {
            try
            {
                // Set console text color to red
                Console.ForegroundColor = ConsoleColor.Red;
                // Print the error message
                Console.WriteLine(errorContent);

                // Exit the program
                Environment.Exit(1);
            }
            catch
            {
                // If an error occurs during error message generation, generate an error
                Console.WriteLine("An error occurred while generating the console error.");
            }
        }
    }
}