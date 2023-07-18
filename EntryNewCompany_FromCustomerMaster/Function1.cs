using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Collections.Generic;
using static System.Net.Mime.MediaTypeNames;
using Dapper;

namespace EntryNewCompany_FromCustomerMaster
{
    public class Function1
    {

        [FunctionName("Function1")]
        public void Run([TimerTrigger("0 */5 * * * *", RunOnStartup = true)] TimerInfo myTimer, ILogger log)//5���Ɉ��
        {
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");

            using var fileStream = new FileStream("./phonic-monolith-392109-d6f2255c2bc6.json", FileMode.Open, FileAccess.Read);
            var googleCredential = GoogleCredential.FromStream(fileStream).CreateScoped(SheetsService.Scope.Spreadsheets);
            var sheetsService = new SheetsService(new BaseClientService.Initializer() { HttpClientInitializer = googleCredential });
            var spreadsheetId = "1MjYPr8x8hzd9t-1nR6HgcKicK0gTfzsSq2Q0zzpgPSg";
            var range = "�ڋq�}�X�^!A2:AL";
            var request = sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);
            var response = request.Execute();
            var values = response.Values.ToList();

            List<Company> CompanyList = new List<Company>();

            using var connection = new SqliteConnection(values);
            using (var command = connection.CreateCommand())
            {
                connection.Open();
                var sql = "[dbo].[SP_CompanyInitData]";
                var result = connection.Queary<Company>(sql);

                foreach (var value in values)
                {
                    //value.Count�̐���38�ł���Ώo��
                    if (value.Count() == 38)
                    {
                        CompanyList.Add(new Company()
                        {
                            CompanyCode = (string)value[5] ?? "",
                            CompanyName = (string)value[0] ?? "",
                            Furigana = (string)value[15] ?? "",
                            ZipCode = (string)value[17] ?? "",
                            CompanyNameS = (string)value[37] ?? "",
                            PrefName = (string)value[18] ?? "",
                            Address1 = (string)value[19] ?? "",
                            Address2 = (string)value[20] ?? "",
                            Phone = (string)value[24] ?? "",
                            Email = (string)value[25] ?? "",
                            SlackEmail = (string)value[29] ?? "",
                            DockName = (string)value[2] ?? "",
                        });
                    }
                    else
                    {
                        Console.WriteLine("");
                    }
                }
            }
                
            foreach (Company c in CompanyList)
            {
                Console.WriteLine(c.CompanyCode +","+ c.CompanyName +","+ c.Furigana +","+ c.ZipCode +","+
                    c.CompanyNameS +","+c.PrefName +","+ c.Address1 +","+ c.Address2 +","+
                    c.Phone +","+c.Email +","+ c.SlackEmail+","+ c.DockName);
                
               
            }

           
        }
    }

    public class Company
    {
        public string CompanyCode { get; set; }
        public string CompanyName { get; set; } 
        public string Furigana { get; set; }
        public string CompanyNameS { get; set; }//����
        public string ZipCode { get; set; }//�X�֔ԍ�
        public string PrefName { get; set; }//�s���{��
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string Phone { get; set; }
        public string Email { get; set; }
        public string SlackEmail { get; set; }
        public string DockName { get; set; }

       
    }
}


