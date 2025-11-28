using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace CopiarParametrosRevit2021.Helpers
{
    public class GoogleSheetsService
    {
        private SheetsService _service;

        private static readonly string CREDENTIALS_FILE = "revitsheetsintegration-89c34b39c2ae.json";

        public GoogleSheetsService()
        {
            InitializeService();
        }

        private void InitializeService()
        {
            try
            {
                string assemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string credentialPath = Path.Combine(assemblyPath, CREDENTIALS_FILE);

                if (!File.Exists(credentialPath))
                {
                    throw new FileNotFoundException($"No se encontró el archivo de credenciales en: {credentialPath}");
                }

                GoogleCredential credential;
                using (var stream = new FileStream(credentialPath, FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleCredential.FromStream(stream)
                        .CreateScoped(new[] { SheetsService.Scope.Spreadsheets });
                }

                _service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "RevitPluginColors"
                });
            }
            catch (Exception ex)
            {
                throw new Exception($"Error de autenticación con Google: {ex.Message}");
            }
        }

        public List<string> ObtenerCategoriasDesdeSheet(string spreadsheetId, string sheetName)
        {
            try
            {
                string range = $"'{sheetName}'!A:B";
                SpreadsheetsResource.ValuesResource.GetRequest request =
                    _service.Spreadsheets.Values.Get(spreadsheetId, range);

                ValueRange response = request.Execute();
                IList<IList<object>> values = response.Values;

                if (values != null && values.Count > 0)
                {
                    foreach (var row in values)
                    {
                        if (row.Count < 2) continue;

                        string key = row[0]?.ToString()?.Trim().ToUpper();

                        if (key == "CATEGORIAS")
                        {
                            string contenidoCeldaB = row[1]?.ToString();

                            if (string.IsNullOrWhiteSpace(contenidoCeldaB))
                                return new List<string>();

                            return contenidoCeldaB
                                .Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(c => c.Trim())
                                .Where(c => !string.IsNullOrWhiteSpace(c))
                                .ToList();
                        }
                    }
                }

                return new List<string>();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error leyendo Google Sheets: {ex.Message}");
            }
        }
    }
}
