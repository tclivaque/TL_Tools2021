using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Linq;

namespace CopiarParametrosRevit2021.Commands.LookaheadManagement
{
    [Transaction(TransactionMode.Manual)]
    public class MembreteLookaheadCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // 1. PROCESAMIENTO DEL NOMBRE DEL ARCHIVO
                string fileName = doc.Title;

                // Quitamos la extensión si existe
                if (Path.HasExtension(fileName))
                    fileName = Path.GetFileNameWithoutExtension(fileName);

                string[] parts = fileName.Split('-');

                // Validación de seguridad
                if (parts.Length < 5)
                {
                    TaskDialog.Show("Error",
                        "El nombre del archivo no tiene el formato esperado.\n" +
                        "Ej: 200114-CCC02-MO-EE-046900");
                    return Result.Failed;
                }

                // Extraer Especialidad (Índice 3) -> Ej: "EE"
                string especialidad = parts[3];

                // Extraer Sector/Activo (Índice 4) -> Ej: "046900"
                string rawSector = parts[4];
                string codigoActivo = "";

                // LÓGICA DE EXTRACCIÓN DEL SECTOR
                if (rawSector.StartsWith("000") && rawSector.Length >= 6)
                {
                    codigoActivo = rawSector.Substring(3, 3); // 000410 -> 410
                }
                else if (rawSector.StartsWith("0") && rawSector.Length >= 4)
                {
                    codigoActivo = rawSector.Substring(1, 3); // 046900 -> 469
                }
                else
                {
                    codigoActivo = rawSector.Length >= 3 ?
                        rawSector.Substring(rawSector.Length - 3) : rawSector;
                }

                // 2. CÁLCULO DE FECHAS Y SEMANA
                DateTime hoy = DateTime.Today;

                // Calcular próximo lunes
                int diasParaLunes = ((int)DayOfWeek.Monday - (int)hoy.DayOfWeek + 7) % 7;
                DateTime fechaEntrega = hoy.AddDays(diasParaLunes);

                // Calcular número de semana (Inicio: 09/12/2024)
                DateTime inicioSemanas = new DateTime(2024, 12, 9);
                int semanaEntrega = (int)((fechaEntrega - inicioSemanas).TotalDays / 7) + 1;

                // 3. BÚSQUEDA DEL PLANO OBJETIVO ("-LPS-S")
                ViewSheet targetSheet = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .FirstOrDefault(s => s.SheetNumber.Contains("-LPS-S"));

                if (targetSheet == null)
                {
                    TaskDialog.Show("Error",
                        "No se encontró ningún plano que contenga '-LPS-S' en su número.");
                    return Result.Failed;
                }

                // Buscar el Titleblock dentro de ese plano
                FamilyInstance titleblock = new FilteredElementCollector(doc, targetSheet.Id)
                    .OfCategory(BuiltInCategory.OST_TitleBlocks)
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .FirstOrDefault();

                if (titleblock == null)
                {
                    TaskDialog.Show("Error",
                        $"El plano {targetSheet.SheetNumber} no tiene un membrete colocado.");
                    return Result.Failed;
                }

                // 4. TRANSACCIÓN DE ESCRITURA
                using (Transaction t = new Transaction(doc, "Actualizar Membrete LookAhead"))
                {
                    t.Start();

                    // --- A. ACTUALIZAR DATOS DEL PLANO ---
                    targetSheet.Name = $"PLANO DE LOOK AHEAD PLANNING - SEMANA {semanaEntrega}";

                    // Nuevo número: 200114-CCC02-PL-[ESPECIALIDAD]-[ACTIVO]-LPS-S[SEMANA]
                    string newSheetNumber =
                        $"200114-CCC02-PL-{especialidad}-{codigoActivo}-LPS-S{semanaEntrega}";

                    if (!targetSheet.SheetNumber.Equals(newSheetNumber))
                    {
                        try
                        {
                            targetSheet.SheetNumber = newSheetNumber;
                        }
                        catch { }
                    }

                    // Fecha de Emisión
                    Parameter dateParam = targetSheet.get_Parameter(BuiltInParameter.SHEET_ISSUE_DATE);
                    if (dateParam != null)
                        dateParam.Set(fechaEntrega.ToString("dd/MM/yyyy"));

                    // --- B. ACTUALIZAR DATOS DEL MEMBRETE ---
                    SetParameter(titleblock, "NOMBRE DEL PROYECTO", "I.E. FRANCISCO BOLOGNESI CERVANTES");
                    SetParameter(titleblock, "DISCIPLINA", GetDisciplinaName(especialidad));
                    SetParameter(titleblock, "ESCALA", "INDICADA");
                    SetParameter(titleblock, "RV", "C0");

                    // Apagar visibilidades
                    SetParameter(titleblock, "ESCALA GRAFICA", 0);
                    SetParameter(titleblock, "NORTE REFERENTE", 0);
                    SetParameter(titleblock, "EMITIDO PARA CONSTRUCCIÓN", 0);

                    // --- C. CHECKBOXES DE ACTIVOS ---
                    foreach (Parameter param in titleblock.Parameters)
                    {
                        if (param.Definition.Name.StartsWith("ACTIVO "))
                        {
                            // Si contiene el código activo, marcar check (1), sino (0)
                            param.Set(param.Definition.Name.Contains(codigoActivo) ? 1 : 0);
                        }
                    }

                    t.Commit();
                }

                TaskDialog.Show("Éxito",
                    $"Membrete Actualizado:\nPlano: {targetSheet.SheetNumber}\nSemana: {semanaEntrega}");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private void SetParameter(Element element, string paramName, object value)
        {
            Parameter param = element.LookupParameter(paramName);

            if (param != null && !param.IsReadOnly)
            {
                if (value is string s)
                    param.Set(s);
                else if (value is int i)
                    param.Set(i);
                else if (value is double d)
                    param.Set(d);
            }
        }

        private string GetDisciplinaName(string sigla)
        {
            switch (sigla.ToUpper())
            {
                case "AR":
                    return "ARQUITECTURA";
                case "EE":
                    return "INSTALACIONES ELÉCTRICAS";
                case "ES":
                    return "ESTRUCTURAS";
                case "SA":
                    return "INSTALACIONES SANITARIAS";
                case "ME":
                    return "INSTALACIONES MECÁNICAS";
                default:
                    return "ARQUITECTURA";
            }
        }
    }
}
