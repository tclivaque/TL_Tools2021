using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

public static class GoogleSheetsHelper
{
    /// <summary>
    /// Descarga y procesa datos de Google Sheets publicado como CSV
    /// </summary>
    /// <param name="url">URL del CSV publicado</param>
    /// <param name="columnIndex1">Índice de la primera columna a extraer (base 0)</param>
    /// <param name="columnIndex2">Índice de la segunda columna a extraer (base 0)</param>
    /// <returns>Tupla con las dos listas de datos extraídas</returns>
    public static (List<string> column1, List<string> column2) GetColumnsFromSheet(string url, int columnIndex1, int columnIndex2)
    {
        try
        {
            // Descargar CSV
            string csvData = DownloadCSV(url);

            if (string.IsNullOrEmpty(csvData))
            {
                return (new List<string>(), new List<string>());
            }

            // Parsear CSV
            List<List<string>> matrix = ParseCSV(csvData);

            // Reemplazar ";" por "," en todos los datos
            matrix = matrix.Select(row =>
                row.Select(cell => cell.Replace(';', ',')).ToList()
            ).ToList();

            // Transponer matriz
            List<List<string>> transposed = TransposeMatrix(matrix);

            // Extraer columnas específicas
            List<string> col1 = (columnIndex1 < transposed.Count) ? transposed[columnIndex1] : new List<string>();
            List<string> col2 = (columnIndex2 < transposed.Count) ? transposed[columnIndex2] : new List<string>();

            return (col1, col2);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error al obtener datos de Google Sheets: {ex.Message}", ex);
        }
    }

    private static string DownloadCSV(string url)
    {
        using (WebClient client = new WebClient())
        {
            client.Encoding = Encoding.UTF8;
            return client.DownloadString(url);
        }
    }

    private static List<List<string>> ParseCSV(string csvData)
    {
        List<List<string>> result = new List<List<string>>();

        string[] lines = csvData.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

        foreach (string line in lines)
        {
            if (string.IsNullOrWhiteSpace(line))
                continue;

            List<string> cells = line.Split(',')
                .Select(cell => cell.Replace('\n', ' ').Replace('\r', ' ').Trim())
                .ToList();

            result.Add(cells);
        }

        return result;
    }

    private static List<List<string>> TransposeMatrix(List<List<string>> matrix)
    {
        if (matrix == null || matrix.Count == 0)
            return new List<List<string>>();

        int maxColumns = matrix.Max(row => row.Count);
        List<List<string>> transposed = new List<List<string>>();

        for (int col = 0; col < maxColumns; col++)
        {
            List<string> newRow = new List<string>();

            foreach (var row in matrix)
            {
                if (col < row.Count)
                    newRow.Add(row[col]);
                else
                    newRow.Add("");
            }

            transposed.Add(newRow);
        }

        return transposed;
    }
}