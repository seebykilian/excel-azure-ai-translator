using DotNetEnv;
using Newtonsoft.Json;
using System.Text;

namespace ExcelAzureAiTranslator
{
    class AzureApiManager
    {
        public static async Task<string> TranslatorAI(string text)
        {
            string requestUri = "https://api.cognitive.microsofttranslator.com/translate?api-version=3.0&from=" + ExcelManager.referenceColumnLanguageCode + "&to=" + ExcelManager.destinationColumnLanguageCode;

            object[] body = new object[] { new { Text = text } };
            var requestBody = JsonConvert.SerializeObject(body);

            using (var client = new HttpClient())
            using (var request = new HttpRequestMessage())
            {
                request.Method = HttpMethod.Post;
                request.RequestUri = new Uri(requestUri);
                request.Content = new StringContent(requestBody, Encoding.UTF8, "application/json");
                // Add subscription key and region headers to the request
                request.Headers.Add("Ocp-Apim-Subscription-Key", Env.GetString("azureApiKey"));
                request.Headers.Add("Ocp-Apim-Subscription-Region", Env.GetString("azureApiRegion"));

                var response = await client.SendAsync(request);
                var responseBody = await response.Content.ReadAsStringAsync();

                dynamic result = JsonConvert.DeserializeObject(responseBody);
                return result[0].translations[0].text;
            }
        }
    }
}