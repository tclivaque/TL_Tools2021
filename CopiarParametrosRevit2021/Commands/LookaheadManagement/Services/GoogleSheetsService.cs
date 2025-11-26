using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace TL_Tools2021.Commands.LookaheadManagement.Services
{
    public class GoogleSheetsService
    {
        private static readonly string CREDENTIALS_PATH = Path.Combine(
            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
            "revitsheetsintegration-89c34b39c2ae.json");

        public SheetsService Service { get; private set; }

        public GoogleSheetsService()
        {
            InitializeService();
        }

        private void InitializeService()
        {
            try
            {
                GoogleCredential credential;
                using (var stream = new FileStream(CREDENTIALS_PATH, FileMode.Open, FileAccess.Read))
                {
#pragma warning disable CS0618
                    credential = GoogleCredential.FromStream(stream)
                        .CreateScoped(new[] { SheetsService.Scope.SpreadsheetsReadonly });
#pragma warning restore CS0618
                }

                Service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "RevitLookAheadAddin"
                });
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al inicializar credenciales: {ex.Message}");
            }
        }

        public IList<IList<object>> ReadData(string spreadsheetId, string range)
        {
            try
            {
                var response = Service.Spreadsheets.Values.Get(spreadsheetId, range).Execute();
                return response.Values ?? new List<IList<object>>();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error leyendo Google Sheets: {ex.Message}");
            }
        }
    }
}
