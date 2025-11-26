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

        // Ajusta el nombre si tu archivo JSON es distinto
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
                    // Usamos Scope completo para permitir lectura y escritura en todo el plugin
                    credential = GoogleCredential.FromStream(stream)
                        .CreateScoped(new[] { SheetsService.Scope.Spreadsheets });
                }

                _service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "RevitPluginUnified"
                });
            }
            catch (Exception ex)
            {
                throw new Exception($"Error de autenticación con Google: {ex.Message}");
            }
        }

        /// <summary>
        /// Método genérico para leer cualquier rango (Usado por Lookahead y otros)
        /// </summary>
        public IList<IList<object>> ReadData(string spreadsheetId, string range)
        {
            try
            {
                var request = _service.Spreadsheets.Values.Get(spreadsheetId, range);
                var response = request.Execute();
                return response.Values ?? new List<IList<object>>();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error leyendo datos de Google Sheets: {ex.Message}");
            }
        }

        /// <summary>
        /// Método específico para leer la lista blanca de categorías (Usado por Auditoría/Colores)
        /// </summary>
        public List<string> ObtenerCategoriasDesdeSheet(string spreadsheetId, string sheetName)
        {
            try
            {
                // Reutilizamos ReadData para no repetir lógica
                string range = $"'{sheetName}'!A:B";
                var values = ReadData(spreadsheetId, range);

                if (values != null && values.Count > 0)
                {
                    foreach (var row in values)
                    {
                        if (row.Count < 2) continue;

                        string key = row[0]?.ToString()?.Trim().ToUpper();
                        if (key == "CATEGORIAS")
                        {
                            string contenidoCeldaB = row[1]?.ToString();
                            if (string.IsNullOrWhiteSpace(contenidoCeldaB)) return new List<string>();

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
                throw new Exception($"Error obteniendo categorías: {ex.Message}");
            }
        }
    }
}