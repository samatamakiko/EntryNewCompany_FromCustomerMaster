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
using Microsoft.AspNetCore.Mvc.ModelBinding.Validation;


namespace EntryNewCompany_FromCustomerMaster
{
    public class Function1
    {

        [FunctionName("Function1")]
        public void Run([TimerTrigger("0 */5 * * * *", RunOnStartup = true)] TimerInfo myTimer, ILogger log)//5分に一回
        {
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");

            using var fileStream = new FileStream("./phonic-monolith-392109-d6f2255c2bc6.json", FileMode.Open, FileAccess.Read);
            var googleCredential = GoogleCredential.FromStream(fileStream).CreateScoped(SheetsService.Scope.Spreadsheets);
            var sheetsService = new SheetsService(new BaseClientService.Initializer() { HttpClientInitializer = googleCredential });
            var spreadsheetId = "1MjYPr8x8hzd9t-1nR6HgcKicK0gTfzsSq2Q0zzpgPSg";
            var range = "顧客マスタ!A431:AL432";
            var request = sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);
            var response = request.Execute();
            var values = response.Values.ToList();

            List<Company> CompanyList = new List<Company>();
            foreach (var value in values)
            {
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

            }

            using (var connection = new SqlConnection(CompanyList))
            {
                var result = connection.Query<Company>(
                    "MyStoredProcedure",
                    commandType: CommandType.StoredProcedure
                ).ToList();


                foreach (Company c in CompanyList)
                {
                    //杉戸→1100、八潮→1000に倉庫コード変換
                    if (c.DockName == "Sugito")
                    {
                        c.DockName = c.DockName.Replace("Sugito", "1100");
                    }
                    else if (c.DockName == "Yashio")
                    {
                        c.DockName = c.DockName.Replace("Yashio", "1000");
                    }

                    //必須項目に値が入っていなかったらbreakする
                    if (c.CompanyCode == "")
                    {
                        break;
                    }
                    else if (c.CompanyName == "")
                    {
                        break;
                    }
                    else if (c.Furigana == "")
                    {
                        break;
                    }
                    else if (c.ZipCode == "")
                    {
                        break;
                    }
                    else if (c.CompanyNameS == "")
                    {
                        break;
                    }
                    else if (c.PrefName == "")
                    {
                        break;
                    }
                    else if (c.Address1 == "")
                    {
                        break;
                    }
                    else if (c.DockName == "")
                    {
                        break;
                    }
                    else
                    {
                        Console.WriteLine(c.CompanyCode + "," + c.CompanyName + "," + c.Furigana + "," + c.ZipCode + "," +
                        c.CompanyNameS + "," + c.PrefName + "," + c.Address1 + "," + c.Address2 + "," +
                        c.Phone + "," + c.Email + "," + c.SlackEmail + "," + c.DockName);
                    }



                }
            }
        }
    }

    public class Company
    {
        public string CompanyCode { get; set; }
        public string CompanyName { get; set; } 
        public string Furigana { get; set; }
        public string CompanyNameS { get; set; }//略称
        public string ZipCode { get; set; }//郵便番号
        public string PrefName { get; set; }//都道府県
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string Phone { get; set; }
        public string Email { get; set; }
        public string SlackEmail { get; set; }
        public string DockName { get; set; }

       
    }
}


