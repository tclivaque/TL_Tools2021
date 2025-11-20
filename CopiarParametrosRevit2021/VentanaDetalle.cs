using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

public class VentanaDetalle : Window
{
    public VentanaDetalle(List<ElementId> elementIds, string valorParametro, Document doc)
    {
        InitializeComponent(elementIds, valorParametro, doc);
    }

    private void InitializeComponent(List<ElementId> elementIds, string valorParametro, Document doc)
    {
        this.Title = $"Detalle: {valorParametro}";
        this.Width = 600;
        this.Height = 400;
        this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
        this.Topmost = true;

        System.Windows.Controls.Grid gridPrincipal = new System.Windows.Controls.Grid();
        gridPrincipal.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        gridPrincipal.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        // Título
        TextBlock titulo = new TextBlock();
        titulo.Text = $"Elementos con valor: {valorParametro} ({elementIds.Count} elementos)";
        titulo.FontSize = 14;
        titulo.FontWeight = FontWeights.Bold;
        titulo.Margin = new Thickness(10);
        System.Windows.Controls.Grid.SetRow(titulo, 0);

        // DataGrid
        DataGrid dataGrid = new DataGrid();
        dataGrid.IsReadOnly = true;
        dataGrid.AutoGenerateColumns = false;
        dataGrid.Margin = new Thickness(10);
        System.Windows.Controls.Grid.SetRow(dataGrid, 1);

        // Columnas
        dataGrid.Columns.Add(new DataGridTextColumn
        {
            Header = "Categoría",
            Binding = new System.Windows.Data.Binding("Categoria"),
            Width = new DataGridLength(1, DataGridLengthUnitType.Star)
        });

        dataGrid.Columns.Add(new DataGridTextColumn
        {
            Header = "Nombre de Tipo",
            Binding = new System.Windows.Data.Binding("NombreTipo"),
            Width = new DataGridLength(2, DataGridLengthUnitType.Star)
        });

        dataGrid.Columns.Add(new DataGridTextColumn
        {
            Header = "ID",
            Binding = new System.Windows.Data.Binding("Id"),
            Width = new DataGridLength(0.7, DataGridLengthUnitType.Star)
        });

        // Recopilar datos de elementos
        List<ElementoInfo> datosElementos = new List<ElementoInfo>();

        foreach (ElementId elemId in elementIds)
        {
            try
            {
                Element elem = doc.GetElement(elemId);
                if (elem != null)
                {
                    string categoria = elem.Category != null ? elem.Category.Name : "Sin categoría";
                    string nombreTipo = "N/A";

                    // Obtener nombre de tipo
                    ElementId typeId = elem.GetTypeId();
                    if (typeId != null && typeId != ElementId.InvalidElementId)
                    {
                        ElementType elemType = doc.GetElement(typeId) as ElementType;
                        if (elemType != null)
                        {
                            nombreTipo = elemType.Name;
                        }
                    }

                    datosElementos.Add(new ElementoInfo
                    {
                        Categoria = categoria,
                        NombreTipo = nombreTipo,
                        Id = elemId.IntegerValue.ToString()
                    });
                }
            }
            catch
            {
                // Continuar con el siguiente elemento
            }
        }

        // Ordenar por Categoría, luego Nombre de Tipo, luego ID
        var datosOrdenados = datosElementos
            .OrderBy(e => e.Categoria)
            .ThenBy(e => e.NombreTipo)
            .ThenBy(e => int.Parse(e.Id))
            .ToList();

        dataGrid.ItemsSource = datosOrdenados;

        gridPrincipal.Children.Add(titulo);
        gridPrincipal.Children.Add(dataGrid);

        this.Content = gridPrincipal;
    }
}

public class ElementoInfo
{
    public string Categoria { get; set; }
    public string NombreTipo { get; set; }
    public string Id { get; set; }
}