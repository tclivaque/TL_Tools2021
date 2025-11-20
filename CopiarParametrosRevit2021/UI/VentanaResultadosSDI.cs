using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Linq;

public class VentanaResultadosSDI : Window
{
    public VentanaResultadosSDI(SDIProcessor.ResultadosProcesamiento resultados)
    {
        InitializeComponent(resultados);
    }

    private void InitializeComponent(SDIProcessor.ResultadosProcesamiento resultados)
    {
        this.Title = "Resultados - Dar Formato SDI/Incidencias";
        this.Width = 1000;
        this.Height = 600;
        this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
        this.ResizeMode = ResizeMode.CanResize;

        Grid gridPrincipal = new Grid();
        gridPrincipal.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        gridPrincipal.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        gridPrincipal.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        // RESUMEN (superior)
        TextBlock txtResumen = new TextBlock();
        txtResumen.Text = resultados.GenerarResumen();
        txtResumen.Margin = new Thickness(10);
        txtResumen.FontWeight = FontWeights.Bold;
        txtResumen.Background = System.Windows.Media.Brushes.LightYellow;
        txtResumen.Padding = new Thickness(10);
        Grid.SetRow(txtResumen, 0);

        // TABLAS (centro con scroll)
        ScrollViewer scrollViewer = new ScrollViewer();
        scrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
        scrollViewer.Margin = new Thickness(10);

        StackPanel panelTablas = new StackPanel();

        // Tabla 1: SDI Estandarizados
        if (resultados.ElementosSDI.Count > 0)
        {
            panelTablas.Children.Add(CrearTitulo("SDI Estandarizados"));
            panelTablas.Children.Add(CrearTablaAntesDespues(resultados.ElementosSDI));
        }

        // Tabla 2: Incidencias Estandarizadas
        if (resultados.ElementosINCIDENCIA.Count > 0)
        {
            panelTablas.Children.Add(CrearTitulo("Incidencias Estandarizadas"));
            panelTablas.Children.Add(CrearTablaAntesDespues(resultados.ElementosINCIDENCIA));
        }

        // Tabla 3: SDI Ya Correctos
        if (resultados.ElementosYaCorrectos_SDI.Count > 0)
        {
            panelTablas.Children.Add(CrearTitulo("SDI Ya Correctos (Actualizados)"));
            panelTablas.Children.Add(CrearTablaAntesDespues(resultados.ElementosYaCorrectos_SDI));
        }

        // Tabla 4: Incidencias Ya Correctas
        if (resultados.ElementosYaCorrectos_INC.Count > 0)
        {
            panelTablas.Children.Add(CrearTitulo("Incidencias Ya Correctas (Actualizadas)"));
            panelTablas.Children.Add(CrearTablaAntesDespues(resultados.ElementosYaCorrectos_INC));
        }

        // Elementos con problemas
        if (resultados.ElementosSDI_NoEncontrados.Count > 0)
        {
            panelTablas.Children.Add(CrearTituloAdvertencia("SDI No Encontrados en Diccionario"));
            panelTablas.Children.Add(CrearListaElementos(resultados.ElementosSDI_NoEncontrados));
        }

        if (resultados.ElementosConflictoSDI_INC.Count > 0)
        {
            panelTablas.Children.Add(CrearTituloAdvertencia("Conflictos SDI + Incidencia"));
            panelTablas.Children.Add(CrearListaElementos(resultados.ElementosConflictoSDI_INC));
        }

        if (resultados.ElementosINC_FormatoInvalido.Count > 0)
        {
            panelTablas.Children.Add(CrearTituloAdvertencia("Incidencias con Formato Inválido"));
            panelTablas.Children.Add(CrearListaElementos(resultados.ElementosINC_FormatoInvalido));
        }

        scrollViewer.Content = panelTablas;
        Grid.SetRow(scrollViewer, 1);

        // BOTÓN CERRAR (inferior)
        Button btnCerrar = new Button();
        btnCerrar.Content = "Cerrar";
        btnCerrar.Width = 100;
        btnCerrar.Height = 35;
        btnCerrar.Margin = new Thickness(10);
        btnCerrar.HorizontalAlignment = HorizontalAlignment.Right;
        btnCerrar.Click += (s, e) => this.Close();
        Grid.SetRow(btnCerrar, 2);

        gridPrincipal.Children.Add(txtResumen);
        gridPrincipal.Children.Add(scrollViewer);
        gridPrincipal.Children.Add(btnCerrar);

        this.Content = gridPrincipal;
    }

    private TextBlock CrearTitulo(string texto)
    {
        TextBlock titulo = new TextBlock();
        titulo.Text = texto;
        titulo.FontSize = 14;
        titulo.FontWeight = FontWeights.Bold;
        titulo.Margin = new Thickness(0, 15, 0, 5);
        titulo.Foreground = System.Windows.Media.Brushes.DarkBlue;
        return titulo;
    }

    private TextBlock CrearTituloAdvertencia(string texto)
    {
        TextBlock titulo = new TextBlock();
        titulo.Text = texto;
        titulo.FontSize = 14;
        titulo.FontWeight = FontWeights.Bold;
        titulo.Margin = new Thickness(0, 15, 0, 5);
        titulo.Foreground = System.Windows.Media.Brushes.DarkRed;
        return titulo;
    }

    private DataGrid CrearTablaAntesDespues(List<SDIProcessor.ElementoAntesDespues> elementos)
    {
        DataGrid tabla = new DataGrid();
        tabla.AutoGenerateColumns = false;
        tabla.IsReadOnly = true;
        tabla.GridLinesVisibility = DataGridGridLinesVisibility.All;
        tabla.HeadersVisibility = DataGridHeadersVisibility.Column;
        tabla.MaxHeight = 300;
        tabla.Margin = new Thickness(0, 0, 0, 10);

        // Columnas
        tabla.Columns.Add(new DataGridTextColumn
        {
            Header = "ID Elemento",
            Binding = new System.Windows.Data.Binding("ElementoId"),
            Width = 100
        });

        tabla.Columns.Add(new DataGridTextColumn
        {
            Header = "NUMERO DE SDI (Antes)",
            Binding = new System.Windows.Data.Binding("NumeroSDI_Antes"),
            Width = 180
        });

        tabla.Columns.Add(new DataGridTextColumn
        {
            Header = "NUMERO DE SDI (Después)",
            Binding = new System.Windows.Data.Binding("NumeroSDI_Despues"),
            Width = 180
        });

        tabla.Columns.Add(new DataGridTextColumn
        {
            Header = "TIPO MODIFICACION (Antes)",
            Binding = new System.Windows.Data.Binding("TipoMod_Antes"),
            Width = 220
        });

        tabla.Columns.Add(new DataGridTextColumn
        {
            Header = "TIPO MODIFICACION (Después)",
            Binding = new System.Windows.Data.Binding("TipoMod_Despues"),
            Width = 220
        });

        // Preparar datos para la tabla
        var datos = elementos.Select(e => new
        {
            ElementoId = e.Elemento.Id.IntegerValue.ToString(),
            NumeroSDI_Antes = string.IsNullOrEmpty(e.NumeroSDI_Antes) ? "[vacío]" : e.NumeroSDI_Antes,
            NumeroSDI_Despues = string.IsNullOrEmpty(e.NumeroSDI_Despues) ? "[vacío]" : e.NumeroSDI_Despues,
            TipoMod_Antes = string.IsNullOrEmpty(e.TipoMod_Antes) ? "[vacío]" : e.TipoMod_Antes,
            TipoMod_Despues = string.IsNullOrEmpty(e.TipoMod_Despues) ? "[vacío]" : e.TipoMod_Despues
        }).ToList();

        tabla.ItemsSource = datos;
        return tabla;
    }

    private DataGrid CrearListaElementos(List<Autodesk.Revit.DB.Element> elementos)
    {
        DataGrid tabla = new DataGrid();
        tabla.AutoGenerateColumns = false;
        tabla.IsReadOnly = true;
        tabla.GridLinesVisibility = DataGridGridLinesVisibility.All;
        tabla.HeadersVisibility = DataGridHeadersVisibility.Column;
        tabla.MaxHeight = 200;
        tabla.Margin = new Thickness(0, 0, 0, 10);

        tabla.Columns.Add(new DataGridTextColumn
        {
            Header = "ID Elemento",
            Binding = new System.Windows.Data.Binding("ElementoId"),
            Width = 150
        });

        tabla.Columns.Add(new DataGridTextColumn
        {
            Header = "Categoría",
            Binding = new System.Windows.Data.Binding("Categoria"),
            Width = 200
        });

        tabla.Columns.Add(new DataGridTextColumn
        {
            Header = "NUMERO DE SDI",
            Binding = new System.Windows.Data.Binding("NumeroSDI"),
            Width = 200
        });

        tabla.Columns.Add(new DataGridTextColumn
        {
            Header = "TIPO MODIFICACION",
            Binding = new System.Windows.Data.Binding("TipoMod"),
            Width = 300
        });

        var datos = elementos.Select(e =>
        {
            var paramNumero = e.LookupParameter("NUMERO DE SDI");
            var paramTipo = e.LookupParameter("TIPO DE MODIFICACION");

            return new
            {
                ElementoId = e.Id.IntegerValue.ToString(),
                Categoria = e.Category?.Name ?? "[Sin categoría]",
                NumeroSDI = paramNumero?.AsString() ?? "[vacío]",
                TipoMod = paramTipo?.AsString() ?? "[vacío]"
            };
        }).ToList();

        tabla.ItemsSource = datos;
        return tabla;
    }
}