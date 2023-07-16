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
            var range = "�ڋq�}�X�^!A2:AK5";

            var request = sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);
            var response = request.Execute();
            var values = response.Values.ToList();

           
            List<Company> CompanyList = new List<Company>();
            foreach (var value in values)
            {
                CompanyList.Add(new Company() {
                    CompanyCode = Int32.Parse((string)value[5]),
                    CompanyName = (string)value[0],
                    CompanyKana = (string)value[15],
                    PostCode = (string)value[17],
                    //CompanyAbbreviation = (string)value[16],//�V��������
                    Prefectures = (string)value[18],
                    ComponyAddress1 = (string)value[19],
                    ComponyAddress2 = (string)value[20],
                    TelephoneNumber = (string)value[24],
                    CompanyEmail = (string)value[25],
                    SlackEmail = (string)value[29],
                    DockName = (string)value[2],
                });
            }

         
            foreach (Company c in CompanyList)
            {
                Console.WriteLine(c.CompanyCode +","+ c.CompanyName +","+ c.CompanyKana +","+ c.PostCode +","+
                    c.Prefectures +","+ c.ComponyAddress1 +","+ c.ComponyAddress2 +","+ c.TelephoneNumber +","+
                    c.CompanyEmail +","+ c.SlackEmail +","+ c.DockName);
                
               
            }

           
        }
    }

    public class Company
    {
        public int CompanyCode { get; set; }
        public string CompanyName { get; set; } 
        public string CompanyKana { get; set; }
        public string CompanyAbbreviation { get; set; }//����
        public string PostCode { get; set; }//�X�֔ԍ�
        public string Prefectures { get; set; }//�s���{��
        public string ComponyAddress1 { get; set; }
        public string ComponyAddress2 { get; set; }
        public string TelephoneNumber { get; set; }
        public string CompanyEmail { get; set; }
        public string SlackEmail { get; set; }
        public string DockName { get; set; }

       
    }
}


