using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using System;
using System.Reflection;
using System.IO;

public class App : IExternalApplication
{
    public Result OnStartup(UIControlledApplication application)
    {
        // Nombre de la pestaña y los paneles
        string tabName = "TL_Tools2021";
        string panelAvanceValorizacion = "Planos de Avance y Valorización";
        string panelCopiar = "Copiar Valores de Parámetros";
        string panelRevision = "Revisión de Parámetros";
        string panelIncidenciasSDI = "Incidencias/SDI";
        string panelLookahead = "Lookahead";

        try
        {
            // Crear pestaña
            try { application.CreateRibbonTab(tabName); } catch { /* La pestaña ya existe */ }

            // PANEL 1: Planos de Avance y Valorización
            RibbonPanel panelAvance = null;
            foreach (RibbonPanel p in application.GetRibbonPanels(tabName))
            {
                if (p.Name == panelAvanceValorizacion)
                {
                    panelAvance = p;
                    break;
                }
            }
            if (panelAvance == null)
                panelAvance = application.CreateRibbonPanel(tabName, panelAvanceValorizacion);

            // PANEL 2: Copiar Valores de Parámetros
            RibbonPanel panelCopiarValores = null;
            foreach (RibbonPanel p in application.GetRibbonPanels(tabName))
            {
                if (p.Name == panelCopiar)
                {
                    panelCopiarValores = p;
                    break;
                }
            }
            if (panelCopiarValores == null)
                panelCopiarValores = application.CreateRibbonPanel(tabName, panelCopiar);

            // PANEL 3: Revisión de Parámetros
            RibbonPanel panelRev = null;
            foreach (RibbonPanel p in application.GetRibbonPanels(tabName))
            {
                if (p.Name == panelRevision)
                {
                    panelRev = p;
                    break;
                }
            }
            if (panelRev == null)
                panelRev = application.CreateRibbonPanel(tabName, panelRevision);

            // PANEL 4: INCIDENCIAS / SDI
            RibbonPanel panelSDI = null;
            foreach (RibbonPanel p in application.GetRibbonPanels(tabName))
            {
                if (p.Name == panelIncidenciasSDI)
                {
                    panelSDI = p;
                    break;
                }
            }
            if (panelSDI == null)
                panelSDI = application.CreateRibbonPanel(tabName, panelIncidenciasSDI);

            // PANEL 5: LOOKAHEAD
            RibbonPanel panelLookaheadRibbon = null;
            foreach (RibbonPanel p in application.GetRibbonPanels(tabName))
            {
                if (p.Name == panelLookahead)
                {
                    panelLookaheadRibbon = p;
                    break;
                }
            }
            if (panelLookaheadRibbon == null)
                panelLookaheadRibbon = application.CreateRibbonPanel(tabName, panelLookahead);

            // Ruta al ensamblado actual
            string path = Assembly.GetExecutingAssembly().Location;

            // ============ PANEL 1: PLANOS DE AVANCE Y VALORIZACIÓN ============

            // BOTÓN 1: Valorización
            PushButtonData buttonDataValorizacion = new PushButtonData("btnValorizacion",
                                                                      "Valorización",
                                                                      path,
                                                                      "ValorizacionCommand");
            PushButton buttonValorizacion = panelAvance.AddItem(buttonDataValorizacion) as PushButton;
            buttonValorizacion.ToolTip = "Asigna valores de valorización a elementos seleccionados";
            buttonValorizacion.LongDescription = "Establece EJECUTADO=True, SEMANA VALORIZADA=calculada, EN CURSO=False, RESTRICCION=False en los elementos seleccionados.";

            // BOTÓN 2: En curso
            PushButtonData buttonDataEnCurso = new PushButtonData("btnEnCurso",
                                                                 "En curso",
                                                                 path,
                                                                 "EnCursoCommand");
            PushButton buttonEnCurso = panelAvance.AddItem(buttonDataEnCurso) as PushButton;
            buttonEnCurso.ToolTip = "Marca elementos como EN CURSO";
            buttonEnCurso.LongDescription = "Establece EN CURSO=True, EJECUTADO=False, SEMANA DE EJECUCION=vacío en los elementos seleccionados.";

            // BOTÓN 3: CARs Enchape
            PushButtonData buttonDataCars = new PushButtonData("btnCars",
                                                              "CARs\nEnchape",
                                                              path,
                                                              "CarsCommand");
            PushButton buttonCars = panelAvance.AddItem(buttonDataCars) as PushButton;
            buttonCars.ToolTip = "Asigna valores CARS ENCHAPE con semana de ejecución calculada";
            buttonCars.LongDescription = "Establece EJECUTADO=True, SEMANA DE EJECUCION=SEM XX, EN CURSO=False, RESTRICCION=False en los elementos seleccionados.";

            // BOTÓN 4: Tarrajeo
            PushButtonData buttonDataTarrajeo = new PushButtonData("btnTarrajeo",
                                                                  "Tarrajeo",
                                                                  path,
                                                                  "TarrajeoCommand");
            PushButton buttonTarrajeo = panelAvance.AddItem(buttonDataTarrajeo) as PushButton;
            buttonTarrajeo.ToolTip = "Asigna MET_SUP = TR_EJECUTADO";
            buttonTarrajeo.LongDescription = "Establece el parámetro MET_SUP con el valor TR_EJECUTADO en los elementos seleccionados.";

            // BOTÓN 5: Solaqueo
            PushButtonData buttonDataSolaqueo = new PushButtonData("btnSolaqueo",
                                                                  "Solaqueo",
                                                                  path,
                                                                  "SolaqueoCommand");
            PushButton buttonSolaqueo = panelAvance.AddItem(buttonDataSolaqueo) as PushButton;
            buttonSolaqueo.ToolTip = "Asigna MET_SUP = SL_EJECUTADO";
            buttonSolaqueo.LongDescription = "Establece el parámetro MET_SUP con el valor SL_EJECUTADO en los elementos seleccionados.";

            // BOTÓN 6: Restaurar
            PushButtonData buttonDataRestaurar = new PushButtonData("btnRestaurar",
                                                                   "Restaurar",
                                                                   path,
                                                                   "RestaurarCommand");
            PushButton buttonRestaurar = panelAvance.AddItem(buttonDataRestaurar) as PushButton;
            buttonRestaurar.ToolTip = "Restaura parámetros a valores por defecto";
            buttonRestaurar.LongDescription = "Limpia todos los parámetros modificados por las otras herramientas, regresándolos a sus valores por defecto.";

            // ============ PANEL 2: COPIAR VALORES DE PARÁMETROS ============

            // BOTÓN 7: ⚙️Copiar
            PushButtonData buttonDataConfigCopiar = new PushButtonData("btnConfigParametrosCopiar",
                                                                       "⚙️Copiar",
                                                                       path,
                                                                       "ConfigurarParametrosCopiarCommand");
            PushButton buttonConfigCopiar = panelCopiarValores.AddItem(buttonDataConfigCopiar) as PushButton;
            buttonConfigCopiar.ToolTip = "Configura qué parámetros se copiarán";
            buttonConfigCopiar.LongDescription = "Define la lista de parámetros que se copiarán de un elemento fuente a destino.";

            // BOTÓN 8: Copiar Valores
            PushButtonData buttonDataCopiar = new PushButtonData("btnCopiar",
                                                                "Copiar\nValores",
                                                                path,
                                                                "CopiarParametrosConfiguradosCommand");
            PushButton buttonCopiar = panelCopiarValores.AddItem(buttonDataCopiar) as PushButton;
            buttonCopiar.ToolTip = "Copia los parámetros configurados previamente";
            buttonCopiar.LongDescription = "Copia los parámetros definidos en la configuración de un elemento fuente a elementos destino.";

            // ============ PANEL 3: REVISIÓN DE PARÁMETROS ============

            // BOTÓN 9: Aplicar colores
            PushButtonData buttonDataAplicar = new PushButtonData("btnAplicarColores",
                                                                 "Aplicar\ncolores",
                                                                 path,
                                                                 "EjecutarOverrideCommand");
            PushButton buttonAplicar = panelRev.AddItem(buttonDataAplicar) as PushButton;
            buttonAplicar.ToolTip = "Aplica colores a elementos según parámetro configurado";
            buttonAplicar.LongDescription = "Colorea todos los elementos visibles en la vista actual según el valor del parámetro configurado.";

            // BOTÓN 10: Reset
            PushButtonData buttonDataReset = new PushButtonData("btnReset",
                                                               "Reset",
                                                               path,
                                                               "ResetOverridesCommand");
            PushButton buttonReset = panelRev.AddItem(buttonDataReset) as PushButton;
            buttonReset.ToolTip = "Restablece overrides gráficos de todos los elementos";
            buttonReset.LongDescription = "Elimina todas las sobreescrituras gráficas aplicadas a los elementos en la vista actual.";

            // ============ PANEL 4: INCIDENCIAS / SDI ============

            // BOTÓN 12: Dar formato
            PushButtonData buttonDataDarFormato = new PushButtonData("btnDarFormato",
                                                                    "Dar formato",
                                                                    path,
                                                                    "DarFormatoCommand");
            PushButton buttonDarFormato = panelSDI.AddItem(buttonDataDarFormato) as PushButton;
            buttonDarFormato.ToolTip = "Estandariza formato de SDI e Incidencias";
            buttonDarFormato.LongDescription = "Importa datos de Google Sheets y estandariza los parámetros NUMERO DE SDI, TIPO DE MODIFICACION y MODIFICADO en todos los elementos del modelo.";

            // ============ PANEL 5: LOOKAHEAD ============

            // BOTÓN 13: Procesar Lookahead (Asignar + Membrete)
            PushButtonData buttonDataProcesarLookahead = new PushButtonData("btnProcesarLookahead",
                                                                            "Procesar\nLookahead",
                                                                            path,
                                                                            "CopiarParametrosRevit2021.Commands.LookaheadManagement.ProcesarLookaheadCommand");
            PushButton buttonProcesarLookahead = panelLookaheadRibbon.AddItem(buttonDataProcesarLookahead) as PushButton;
            buttonProcesarLookahead.ToolTip = "Procesa Lookahead completo: Asigna semanas y actualiza membrete";
            buttonProcesarLookahead.LongDescription = "Ejecuta el proceso completo de Lookahead:\n1. Lee datos de Google Sheets y asigna semanas a elementos\n2. Actualiza automáticamente el membrete del plano LPS-S";

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Error", ex.Message);
            return Result.Failed;
        }
    }

    public Result OnShutdown(UIControlledApplication application)
    {
        return Result.Succeeded;
    }
}