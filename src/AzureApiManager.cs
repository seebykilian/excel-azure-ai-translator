// Importing NuGet packages for generate and process an API request 
using Newtonsoft.Json;
using System.Text;

namespace ExcelAzureAiTranslator
{
    // Additional class of the program grouping the Azure API methods
    class AzureApiManager
    {
        // Method for calling the AI Translator to translate text from one language to another
        public static async Task<string> TranslatorAI(string text)
        {
            // Construct the request URI with the reference and destination language codes
            string requestUri = "https://api.cognitive.microsofttranslator.com/translate?api-version=3.0&from=" + ExcelManager.referenceLanguageCode + "&to=" + ExcelManager.destinationLanguageCode;

            // Define the body of the request with the text to translate 
            object[] body = new object[] { new { Text = text } };
            var requestBody = JsonConvert.SerializeObject(body);

            // Set up the HTTP client and the request content
            using (var client = new HttpClient())
            using (var request = new HttpRequestMessage())
            {
                request.Method = HttpMethod.Post;
                request.RequestUri = new Uri(requestUri);
                request.Content = new StringContent(requestBody, Encoding.UTF8, "application/json");

                // Add subscription key and region headers to the request from environment variables
                request.Headers.Add("Ocp-Apim-Subscription-Key", Utils.GetEnvironmentVariableValue("azureApiKey"));
                request.Headers.Add("Ocp-Apim-Subscription-Region", Utils.GetEnvironmentVariableValue("azureApiRegion"));

                // Send the request and get the response
                var response = await client.SendAsync(request);
                var responseBody = await response.Content.ReadAsStringAsync();

                // Deserialize the response
                dynamic result = JsonConvert.DeserializeObject(responseBody) ?? "";
                // Return the translated text
                return result[0].translations[0].text;
            }
        }
    } 
}