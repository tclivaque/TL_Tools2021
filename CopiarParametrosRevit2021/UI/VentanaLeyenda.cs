using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using RevitColor = Autodesk.Revit.DB.Color;
using WpfColor = System.Windows.Media.Color;

public class VentanaLeyenda : Window
{
    private static VentanaLeyenda _instancia;
    private LeyendaEventHandler _eventHandler;
    private ExternalEvent _externalEvent;
    private OverrideCommandEventHandler _overrideEventHandler;
    private ExternalEvent _overrideExternalEvent;
    private StackPanel _panelLeyenda;
    private Button _btnMostrarTodos;
    private System.Windows.Controls.Grid _menuLateral;
    private bool _menuExpanded = true;
    private string _parametroActivo = "";
    private UIApplication _uiApp;

    private static string ConfigFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "TL_Tools2021",
        "parametro_config.txt");

    // Lista de par metros disponibles
    private readonly string[] parametros = new string[]
    {
        "Assembly Code",
        "CODIGO DE ELEMENTO",
        "ID DE ELEMENTO",
        "ACTIVO",
        "PARTIDA",
        "AMBIENTE",
        "UNIDAD",
        "NIVEL",
        "SUBPARTIDA",
        "MODULO",
        "EJES",
        "SISTEMA",
        "DESCRIPTION",
        "EMPRESA",
        "ELEMENTO"
    };

    private VentanaLeyenda()
    {
        InitializeComponent();
    }

    public static VentanaLeyenda ObtenerInstancia()
    {
        if (_instancia == null || !_instancia.IsLoaded)
        {
            _instancia = new VentanaLeyenda();
        }
        return _instancia;
    }

    private void InitializeComponent()
    {
        this.Title = "Leyenda de Colores";
        this.Width = 650;  // Aumentado para el menú lateral
        this.Height = 500;
        this.WindowStartupLocation = WindowStartupLocation.Manual;
        this.Left = SystemParameters.PrimaryScreenWidth - this.Width - 50;
        this.Top = 100;
        this.ShowInTaskbar = true;
        this.ResizeMode = ResizeMode.CanResize;

        // Forzar ventana al frente siempre
        this.Topmost = true;
        this.Activated += (s, e) => { this.Topmost = true; };
        this.Deactivated += (s, e) =>
        {
            System.Threading.Tasks.Task.Delay(10).ContinueWith(_ =>
            {
                this.Dispatcher.Invoke(() => { this.Topmost = true; });
            });
        };

        // Grid principal con 2 columnas (menú lateral + contenido principal)
        System.Windows.Controls.Grid gridPrincipal = new System.Windows.Controls.Grid();
        gridPrincipal.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });  // Menú lateral
        gridPrincipal.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });  // Contenido

        // === MENÚ LATERAL ===
        _menuLateral = CrearMenuLateral();
        System.Windows.Controls.Grid.SetColumn(_menuLateral, 0);
        gridPrincipal.Children.Add(_menuLateral);

        // === CONTENIDO PRINCIPAL (Leyenda) ===
        System.Windows.Controls.Grid gridLeyenda = new System.Windows.Controls.Grid();
        gridLeyenda.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        gridLeyenda.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        System.Windows.Controls.Grid.SetColumn(gridLeyenda, 1);

        ScrollViewer scrollViewer = new ScrollViewer();
        scrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;

        _panelLeyenda = new StackPanel();
        _panelLeyenda.Margin = new Thickness(10);
        scrollViewer.Content = _panelLeyenda;
        System.Windows.Controls.Grid.SetRow(scrollViewer, 0);

        _btnMostrarTodos = new Button();
        _btnMostrarTodos.Content = "Mostrar Todos";
        _btnMostrarTodos.Height = 35;
        _btnMostrarTodos.Margin = new Thickness(10);
        _btnMostrarTodos.Click += BtnMostrarTodos_Click;
        System.Windows.Controls.Grid.SetRow(_btnMostrarTodos, 1);

        gridLeyenda.Children.Add(scrollViewer);
        gridLeyenda.Children.Add(_btnMostrarTodos);
        gridPrincipal.Children.Add(gridLeyenda);

        this.Content = gridPrincipal;

        // Leer parámetro activo
        _parametroActivo = LeerParametroGuardado();
    }

    private System.Windows.Controls.Grid CrearMenuLateral()
    {
        System.Windows.Controls.Grid menuGrid = new System.Windows.Controls.Grid();
        menuGrid.Background = new SolidColorBrush(WpfColor.FromRgb(245, 245, 245));
        menuGrid.Width = 200;

        StackPanel stackMenu = new StackPanel();
        stackMenu.Margin = new Thickness(5);

        // Botón hamburger
        Button btnToggle = new Button();
        btnToggle.Content = "☰";
        btnToggle.FontSize = 20;
        btnToggle.Height = 40;
        btnToggle.Margin = new Thickness(0, 0, 0, 10);
        btnToggle.Click += (s, e) =>
        {
            _menuExpanded = !_menuExpanded;
            menuGrid.Width = _menuExpanded ? 200 : 40;
            foreach (var child in stackMenu.Children)
            {
                if (child is Button btn && btn != btnToggle)
                {
                    btn.Visibility = _menuExpanded ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
                }
            }
        };
        stackMenu.Children.Add(btnToggle);

        // Agregar botones de parámetros
        foreach (string param in parametros)
        {
            Button btnParam = new Button();
            btnParam.Content = param;
            btnParam.Height = 35;
            btnParam.Margin = new Thickness(0, 2, 0, 2);
            btnParam.HorizontalContentAlignment = HorizontalAlignment.Left;
            btnParam.Padding = new Thickness(5);
            btnParam.Tag = param;

            // Color inicial
            if (param == _parametroActivo)
            {
                btnParam.Background = new SolidColorBrush(WpfColor.FromRgb(144, 238, 144));  // Verde pastel
            }
            else
            {
                btnParam.Background = new SolidColorBrush(WpfColor.FromRgb(220, 220, 220));  // Gris claro
            }

            btnParam.Click += BtnParametro_Click;
            stackMenu.Children.Add(btnParam);
        }

        menuGrid.Children.Add(stackMenu);
        return menuGrid;
    }

    private void BtnParametro_Click(object sender, RoutedEventArgs e)
    {
        Button btnClicked = sender as Button;
        if (btnClicked == null) return;

        string parametroSeleccionado = btnClicked.Tag.ToString();

        // Actualizar colores de todos los botones
        StackPanel stackMenu = (_menuLateral.Children[0] as StackPanel);
        foreach (var child in stackMenu.Children)
        {
            if (child is Button btn && btn.Tag is string)
            {
                if (btn.Tag.ToString() == parametroSeleccionado)
                {
                    btn.Background = new SolidColorBrush(WpfColor.FromRgb(144, 238, 144));  // Verde pastel
                }
                else
                {
                    btn.Background = new SolidColorBrush(WpfColor.FromRgb(220, 220, 220));  // Gris claro
                }
            }
        }

        // Guardar parámetro seleccionado
        _parametroActivo = parametroSeleccionado;
        GuardarParametro(parametroSeleccionado);

        // Ejecutar el comando con el nuevo parámetro
        if (_uiApp != null)
        {
            EjecutarComandoConParametro(_uiApp, parametroSeleccionado);
        }
    }

    private void GuardarParametro(string parametro)
    {
        try
        {
            string directory = Path.GetDirectoryName(ConfigFilePath);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }
            File.WriteAllText(ConfigFilePath, parametro);
        }
        catch { }
    }

    private string LeerParametroGuardado()
    {
        try
        {
            if (File.Exists(ConfigFilePath))
            {
                return File.ReadAllText(ConfigFilePath).Trim();
            }
        }
        catch { }
        return parametros[0];  // Parámetro por defecto
    }

    private void EjecutarComandoConParametro(UIApplication app, string parametro)
    {
        // Ejecutar el comando usando el ExternalEvent
        if (_overrideExternalEvent != null)
        {
            _overrideExternalEvent.Raise();
        }
    }

    public void InicializarEventHandler(UIApplication app)
    {
        _uiApp = app;
        if (_eventHandler == null)
        {
            _eventHandler = new LeyendaEventHandler();
            _externalEvent = ExternalEvent.Create(_eventHandler);
        }
        if (_overrideEventHandler == null)
        {
            _overrideEventHandler = new OverrideCommandEventHandler();
            _overrideExternalEvent = ExternalEvent.Create(_overrideEventHandler);
        }
    }

    public void ActualizarLeyenda(Dictionary<string, List<ElementId>> elementosPorValor,
                                  List<ElementId> elementosSinValor,
                                  Dictionary<string, RevitColor> coloresPorValor,
                                  View vistaActiva)
    {
        _panelLeyenda.Children.Clear();

        if (_eventHandler != null)
        {
            _eventHandler.ElementosPorValor = elementosPorValor;
            _eventHandler.ElementosSinValor = elementosSinValor;
            _eventHandler.VistaActiva = vistaActiva;
        }

        TextBlock titulo = new TextBlock();
        titulo.Text = "Valores encontrados:";
        titulo.FontWeight = FontWeights.Bold;
        titulo.Margin = new Thickness(0, 0, 0, 10);
        _panelLeyenda.Children.Add(titulo);

        var valoresOrdenados = elementosPorValor.OrderByDescending(kvp => kvp.Value.Count);

        foreach (var kvp in valoresOrdenados)
        {
            string valor = kvp.Key;
            int cantidad = kvp.Value.Count;
            RevitColor colorRevit = coloresPorValor.ContainsKey(valor) ? coloresPorValor[valor] : new RevitColor(128, 128, 128);

            AgregarItemLeyenda(valor, cantidad, colorRevit);
        }

        if (elementosSinValor != null && elementosSinValor.Count > 0)
        {
            AgregarItemLeyenda("[SIN VALOR]", elementosSinValor.Count, new RevitColor(255, 0, 0));
        }
    }

    private void AgregarItemLeyenda(string valor, int cantidad, RevitColor colorRevit)
    {
        WpfColor colorWpf = WpfColor.FromRgb(colorRevit.Red, colorRevit.Green, colorRevit.Blue);

        // Panel horizontal que contiene el botón de aislar y el botón de detalle
        StackPanel panelHorizontal = new StackPanel();
        panelHorizontal.Orientation = Orientation.Horizontal;
        panelHorizontal.Margin = new Thickness(0, 3, 0, 3);

        // Botón para aislar/ocultar elementos (izquierda)
        Button btnAislar = new Button();
        btnAislar.Background = Brushes.Transparent;
        btnAislar.BorderThickness = new Thickness(1);
        btnAislar.BorderBrush = Brushes.LightGray;
        btnAislar.HorizontalContentAlignment = HorizontalAlignment.Left;
        btnAislar.Padding = new Thickness(5);
        btnAislar.Cursor = Cursors.Hand;
        btnAislar.Width = 280;
        btnAislar.Tag = valor;

        // Click izquierdo: aislar
        btnAislar.Click += (s, e) => ItemLeyenda_Click(valor, false);

        // Click derecho: ocultar
        btnAislar.MouseRightButtonDown += (s, e) =>
        {
            ItemLeyenda_Click(valor, true);
            e.Handled = true;
        };

        StackPanel contenidoAislar = new StackPanel();
        contenidoAislar.Orientation = Orientation.Horizontal;

        Border cuadroColor = new Border();
        cuadroColor.Width = 22;
        cuadroColor.Height = 22;
        cuadroColor.Background = new SolidColorBrush(colorWpf);
        cuadroColor.BorderBrush = Brushes.Black;
        cuadroColor.BorderThickness = new Thickness(1);
        cuadroColor.Margin = new Thickness(0, 0, 10, 0);

        TextBlock texto = new TextBlock();
        texto.Text = $"{valor} ({cantidad})";
        texto.VerticalAlignment = VerticalAlignment.Center;
        texto.TextTrimming = TextTrimming.CharacterEllipsis;
        texto.MaxWidth = 220;

        contenidoAislar.Children.Add(cuadroColor);
        contenidoAislar.Children.Add(texto);
        btnAislar.Content = contenidoAislar;

        // Botón "Detalle" (derecha)
        Button btnDetalle = new Button();
        btnDetalle.Content = "Detalle";
        btnDetalle.Width = 70;
        btnDetalle.Height = 28;
        btnDetalle.Margin = new Thickness(5, 0, 0, 0);
        btnDetalle.Background = new SolidColorBrush(WpfColor.FromRgb(220, 220, 220));
        btnDetalle.Click += (s, e) => MostrarDetalle(valor);

        panelHorizontal.Children.Add(btnAislar);
        panelHorizontal.Children.Add(btnDetalle);

        _panelLeyenda.Children.Add(panelHorizontal);
    }

    private void MostrarDetalle(string valor)
    {
        if (_eventHandler == null) return;

        List<ElementId> elementos = new List<ElementId>();

        if (valor == "[SIN VALOR]" && _eventHandler.ElementosSinValor != null)
        {
            elementos = _eventHandler.ElementosSinValor;
        }
        else if (_eventHandler.ElementosPorValor != null && _eventHandler.ElementosPorValor.ContainsKey(valor))
        {
            elementos = _eventHandler.ElementosPorValor[valor];
        }

        if (elementos.Count > 0 && _eventHandler.VistaActiva != null)
        {
            VentanaDetalle ventanaDetalle = new VentanaDetalle(elementos, valor, _eventHandler.VistaActiva.Document);
            ventanaDetalle.Show();
        }
    }

    private void ItemLeyenda_Click(string valor, bool ocultar)
    {
        if (_eventHandler != null && _externalEvent != null)
        {
            _eventHandler.ValorSeleccionado = valor;
            _eventHandler.MostrarTodos = false;
            _eventHandler.OcultarElementos = ocultar;
            _externalEvent.Raise();
        }
    }

    private void BtnMostrarTodos_Click(object sender, RoutedEventArgs e)
    {
        if (_eventHandler != null && _externalEvent != null)
        {
            _eventHandler.MostrarTodos = true;
            _eventHandler.OcultarElementos = false;
            _externalEvent.Raise();
        }
    }

    protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
    {
        e.Cancel = true;
        this.Hide();
    }

    public new void Close()
    {
        this.Hide();
    }
}
