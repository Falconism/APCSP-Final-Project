using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace APCSP_Final_Project
{
    class Clerk
    {
        static string[] Scopes = {SheetsService.Scope.Spreadsheets};
        UserCredential credential;
        SheetsService service;

        public void connect()
        {
            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets,Scopes, "user", CancellationToken.None, new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }
        }
        public void createService(string AppName)
        {
            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = AppName,
            });
        }
        public string createSheet()
        {
            var createRequest = service.Spreadsheets.Create(new Spreadsheet()
            {
                Properties = new SpreadsheetProperties()
                {
                    Title = "APCSP Create Task"
                }
            });

            var response = createRequest.Execute();
            return response.SpreadsheetId;
        }
        public void writeSheet(string id, List<Card> data)
        {
            var newData = convertToObj(data);
            ValueRange dataDetails = new ValueRange();
            dataDetails.Values = newData;
            var request = service.Spreadsheets.Values.Update(dataDetails, id, "A1");
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            request.Execute();
        }
        public List<IList<Object>> convertToObj(List<Card> cards)
        {
            List<IList<Object>> returnedData = new List<IList<Object>>();
            foreach (Card card in cards)
            {
                List<Object> row = new List<Object>();
                row.Add(card.Name);
                row.Add(card.ID);
                row.Add(card.Artist);
                row.Add(card.FlavorText);
                returnedData.Add(row);
            }
            return returnedData;
        }
    }
}