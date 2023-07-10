using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using System.IO;



namespace EntryNewCompany_FromCustomerMaster
{
    public class Function1
    {
        [FunctionName("Function1")]
        public void Run([TimerTrigger("0 */5 * * * *")]TimerInfo myTimer, ILogger log)//5•ª‚Éˆê‰ñ
        {
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");

            using var fileStream = new FileStream("./phonic-monolith-392109-d6f2255c2bc6.json", FileMode.Open, FileAccess.Read);
            var googleCredential = GoogleCredential.FromStream(fileStream).CreateScoped(SheetsService.Scope.Spreadsheets);
            var sheetsService = new SheetsService(new BaseClientService.Initializer() { HttpClientInitializer = googleCredential });

            var spreadsheetId = "1MjYPr8x8hzd9t-1nR6HgcKicK0gTfzsSq2Q0zzpgPSg";
            var range = "ŒÚ‹qƒ}ƒXƒ^!A2:AK";

            var request = sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);
            var response = request.Execute();
            var values = response.Values;

            foreach (var item in values)
            {
                Console.WriteLine(string.Join(", ", item));
            }

        }

    }
}
