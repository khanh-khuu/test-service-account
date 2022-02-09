using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Hosting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.JSInterop;

namespace TestServiceAccount.Pages
{
    public partial class Index
    {
        [Inject] public IWebHostEnvironment _env { get; set; }
        [Inject] public IJSRuntime _JS { get; set; }

        public readonly string ClientId = "904796907236-ngv4pq5g9lq9b884cfudlhm27i9oj2u0.apps.googleusercontent.com";

        public readonly string ClientSecret = "p3aSigQU0Vc_Zx9xvJbb2HM4";

        public string Value { get; set; }
        public bool IsBusy { get; set; } = false;
        public CancellationTokenSource ctSource { get; set; }

        ElementReference TextAreaRef;

        void ScrollToEnd()
        {
            _JS.InvokeVoidAsync("scrollToEnd", new object[] { TextAreaRef });
        }

        public async Task GetGoogleSheets()
        {
            if (IsBusy)
            {
                ctSource.Cancel();
                IsBusy = false;
                return;
            }
            IsBusy = true;
            var certificate = new X509Certificate2(Path.Combine(_env.ContentRootPath, "horizon-k-6ad5b2140c73.p12"), "notasecret", X509KeyStorageFlags.Exportable);

            const string user = "horizon@horizon-k.iam.gserviceaccount.com";

            var serviceAccountCredentialInitializer = new ServiceAccountCredential.Initializer(user)
            {
                Scopes = new[] { "https://spreadsheets.google.com/feeds" }
            }.FromCertificate(certificate);

            var credential = new ServiceAccountCredential(serviceAccountCredentialInitializer);

            if (!credential.RequestAccessTokenAsync(CancellationToken.None).Result)
                throw new InvalidOperationException("Access token request failed.");

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "Test App",
            });

            // Define request parameters.
            var spreadsheetId = "1lIsBk7jEBvrFd7-me5hJlftXPZ8wzg8_TCOg29A748k";
            var range = "Sheet4!A1:E";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Prints the names and majors of students in a sample spreadsheet:
            // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
            ValueRange response = await request.ExecuteAsync();

            ctSource = new CancellationTokenSource();
            
            Value = null;
            
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    // Print columns A and E, which correspond to indices 0 and 4.
                    Value += $"{row[0]} | {row[1]} | {row[2]} | {row[3]} | {row[4]}";
                }
            }
            else
            {
                Value = "No data found.";
            }
            IsBusy = false;
        }

        public async Task WriteGoogleSheets()
        {
            IsBusy = true;
            var certificate = new X509Certificate2(Path.Combine(_env.ContentRootPath, "horizon-k-6ad5b2140c73.p12"), "notasecret", X509KeyStorageFlags.Exportable);

            const string user = "horizon@horizon-k.iam.gserviceaccount.com";

            var serviceAccountCredentialInitializer = new ServiceAccountCredential.Initializer(user)
            {
                Scopes = new[] { "https://spreadsheets.google.com/feeds" }
            }.FromCertificate(certificate);

            var credential = new ServiceAccountCredential(serviceAccountCredentialInitializer);

            if (!credential.RequestAccessTokenAsync(CancellationToken.None).Result)
                throw new InvalidOperationException("Access token request failed.");

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "Test App",
            });

            // Define request parameters.
            var spreadsheetId = "1lIsBk7jEBvrFd7-me5hJlftXPZ8wzg8_TCOg29A748k";
            var range = "Sheet4!M1";
            IList<IList<object>> values = new List<IList<object>>();
            var row = new List<object>();
            row.Add(DateTimeOffset.Now.ToString());

            values.Add(row);

            ValueRange body = new ValueRange
            {
                Values = values,
            };
            SpreadsheetsResource.ValuesResource.UpdateRequest request =
                    service.Spreadsheets.Values.Update(body, spreadsheetId, range);

            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            await request.ExecuteAsync();
            IsBusy = false;
        }
    }
}
