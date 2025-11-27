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
using System.Windows.Shapes;
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

    private System.Windows.Controls.TextBox _txtParametroInput;

    private static string ConfigFilePath = System.IO.Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "TL_Tools2021",
        "parametro_config.txt");

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
        this.Width = 800;
        this.Height = 900;
        this.WindowStartupLocation = WindowStartupLocation.Manual;
        this.Left = SystemParameters.PrimaryScreenWidth - this.Width - 50;
        this.Top = 100;
        this.ShowInTaskbar = true;
        this.ResizeMode = ResizeMode.CanResize;
        this.Background = new SolidColorBrush(WpfColor.FromRgb(250, 250, 250));

        this.Topmost = true;
        this.Activated += (s, e) => { this.Topmost = true; };
        this.Deactivated += (s, e) =>
        {
            System.Threading.Tasks.Task.Delay(10).ContinueWith(_ =>
            {
                this.Dispatcher.Invoke(() => { this.Topmost = true; });
            });
        };

        System.Windows.Controls.Grid gridPrincipal = new System.Windows.Controls.Grid();
        gridPrincipal.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        gridPrincipal.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        // === MENÚ LATERAL ===
        _menuLateral = CrearMenuLateral();
        System.Windows.Controls.Grid.SetColumn(_menuLateral, 0);
        gridPrincipal.Children.Add(_menuLateral);

        // === CONTENIDO PRINCIPAL ===
        System.Windows.Controls.Grid gridLeyenda = new System.Windows.Controls.Grid();
        gridLeyenda.Margin = new Thickness(10);
        gridLeyenda.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        gridLeyenda.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        System.Windows.Controls.Grid.SetColumn(gridLeyenda, 1);

        ScrollViewer scrollViewer = new ScrollViewer();
        scrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;

        _panelLeyenda = new StackPanel();
        _panelLeyenda.Margin = new Thickness(0, 0, 10, 0);
        _panelLeyenda.HorizontalAlignment = HorizontalAlignment.Left;
        scrollViewer.Content = _panelLeyenda;
        System.Windows.Controls.Grid.SetRow(scrollViewer, 0);

        _btnMostrarTodos = new Button();
        _btnMostrarTodos.Content = "Mostrar Todos";
        _btnMostrarTodos.Height = 40;
        _btnMostrarTodos.Margin = new Thickness(0, 10, 0, 0);
        _btnMostrarTodos.Style = CrearEstiloBotonRedondeado();
        _btnMostrarTodos.HorizontalContentAlignment = HorizontalAlignment.Center;
        _btnMostrarTodos.Background = Brushes.LightGray;
        _btnMostrarTodos.Click += BtnMostrarTodos_Click;
        System.Windows.Controls.Grid.SetRow(_btnMostrarTodos, 1);

        gridLeyenda.Children.Add(scrollViewer);
        gridLeyenda.Children.Add(_btnMostrarTodos);
        gridPrincipal.Children.Add(gridLeyenda);

        this.Content = gridPrincipal;

        _parametroActivo = LeerParametroGuardado();

        if (!parametros.Contains(_parametroActivo))
        {
            _txtParametroInput.Text = _parametroActivo;
        }

        ActualizarEstiloBotonesMenu();
    }

    private Style CrearEstiloBotonRedondeado(double cornerRadius = 10)
    {
        Style style = new Style(typeof(Button));
        ControlTemplate template = new ControlTemplate(typeof(Button));
        FrameworkElementFactory borderFactory = new FrameworkElementFactory(typeof(Border));

        borderFactory.Name = "Border";
        borderFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(cornerRadius));
        borderFactory.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Button.BackgroundProperty));
        borderFactory.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(Button.BorderBrushProperty));
        borderFactory.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(Button.BorderThicknessProperty));

        FrameworkElementFactory contentFactory = new FrameworkElementFactory(typeof(ContentPresenter));

        // CORRECCIÓN AQUÍ: Especificamos System.Windows.Controls.Control explícitamente
        contentFactory.SetValue(ContentPresenter.HorizontalAlignmentProperty, new TemplateBindingExtension(System.Windows.Controls.Control.HorizontalContentAlignmentProperty));
        contentFactory.SetValue(ContentPresenter.VerticalAlignmentProperty, new TemplateBindingExtension(System.Windows.Controls.Control.VerticalContentAlignmentProperty));

        contentFactory.SetValue(ContentPresenter.MarginProperty, new Thickness(5, 0, 5, 0));

        borderFactory.AppendChild(contentFactory);
        template.VisualTree = borderFactory;

        style.Setters.Add(new Setter(Button.TemplateProperty, template));
        return style;
    }

    private System.Windows.Controls.Grid CrearMenuLateral()
    {
        System.Windows.Controls.Grid menuGrid = new System.Windows.Controls.Grid();
        menuGrid.Background = new SolidColorBrush(WpfColor.FromRgb(240, 240, 240));

        menuGrid.Width = 176;

        ScrollViewer scrollMenu = new ScrollViewer();
        scrollMenu.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;

        StackPanel stackMenu = new StackPanel();
        stackMenu.Margin = new Thickness(10);

        Button btnToggle = new Button();
        btnToggle.Content = "☰";
        btnToggle.FontSize = 18;
        btnToggle.Width = 40;
        btnToggle.Height = 40;
        btnToggle.Margin = new Thickness(0, 0, 0, 15);
        btnToggle.HorizontalAlignment = HorizontalAlignment.Center;
        btnToggle.HorizontalContentAlignment = HorizontalAlignment.Center;
        btnToggle.Style = CrearEstiloBotonRedondeado(8);
        btnToggle.Background = Brushes.White;
        btnToggle.BorderBrush = Brushes.LightGray;
        btnToggle.BorderThickness = new Thickness(1);

        btnToggle.Click += (s, e) =>
        {
            _menuExpanded = !_menuExpanded;
            menuGrid.Width = _menuExpanded ? 176 : 60;

            foreach (var child in stackMenu.Children)
            {
                if (child is Button btn && btn != btnToggle)
                {
                    btn.Visibility = _menuExpanded ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
                }
                else if (child is Border brd)
                {
                    brd.Visibility = _menuExpanded ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
                }
                else if (child is TextBlock txt)
                {
                    txt.Visibility = _menuExpanded ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
                }
            }
        };
        stackMenu.Children.Add(btnToggle);

        foreach (string param in parametros)
        {
            Button btnParam = new Button();
            btnParam.Content = param;
            btnParam.Height = 35;
            btnParam.Margin = new Thickness(0, 3, 0, 3);
            btnParam.HorizontalAlignment = HorizontalAlignment.Stretch;
            btnParam.HorizontalContentAlignment = HorizontalAlignment.Center;
            btnParam.Tag = param;
            btnParam.Style = CrearEstiloBotonRedondeado(8);
            btnParam.Click += BtnParametro_Click;
            stackMenu.Children.Add(btnParam);
        }

        TextBlock sep = new TextBlock();
        sep.Text = "──────────────";
        sep.Foreground = Brushes.Gray;
        sep.HorizontalAlignment = HorizontalAlignment.Center;
        sep.Margin = new Thickness(0, 10, 0, 5);
        stackMenu.Children.Add(sep);

        TextBlock lblCustom = new TextBlock();
        lblCustom.Text = "Personalizado:";
        lblCustom.FontSize = 11;
        lblCustom.Foreground = Brushes.DarkGray;
        lblCustom.Margin = new Thickness(2, 0, 0, 2);
        stackMenu.Children.Add(lblCustom);

        Border customPanel = new Border();
        StackPanel stackCustom = new StackPanel();

        _txtParametroInput = new System.Windows.Controls.TextBox();
        _txtParametroInput.Height = 30;
        _txtParametroInput.Margin = new Thickness(0, 0, 0, 5);
        _txtParametroInput.VerticalContentAlignment = VerticalAlignment.Center;
        _txtParametroInput.Padding = new Thickness(5, 0, 0, 0);

        _txtParametroInput.Resources.Add(typeof(Border), new Style(typeof(Border))
        {
            Setters = { new Setter(Border.CornerRadiusProperty, new CornerRadius(5)) }
        });

        Button btnAplicar = new Button();
        btnAplicar.Content = "Aplicar";
        btnAplicar.Height = 30;
        btnAplicar.HorizontalContentAlignment = HorizontalAlignment.Center;
        btnAplicar.Background = Brushes.White;
        btnAplicar.BorderBrush = Brushes.LightGray;
        btnAplicar.BorderThickness = new Thickness(1);
        btnAplicar.Style = CrearEstiloBotonRedondeado(8);
        btnAplicar.Click += BtnAplicarPersonalizado_Click;

        stackCustom.Children.Add(_txtParametroInput);
        stackCustom.Children.Add(btnAplicar);

        customPanel.Child = stackCustom;
        stackMenu.Children.Add(customPanel);

        scrollMenu.Content = stackMenu;
        menuGrid.Children.Add(scrollMenu);
        return menuGrid;
    }

    private void BtnParametro_Click(object sender, RoutedEventArgs e)
    {
        Button btnClicked = sender as Button;
        if (btnClicked == null) return;

        string parametroSeleccionado = btnClicked.Tag.ToString();
        _parametroActivo = parametroSeleccionado;

        _txtParametroInput.Text = "";

        ActualizarEstiloBotonesMenu();
        GuardarParametro(parametroSeleccionado);
        EjecutarComandoExterno();
    }

    private void BtnAplicarPersonalizado_Click(object sender, RoutedEventArgs e)
    {
        string paramCustom = _txtParametroInput.Text.Trim();
        if (string.IsNullOrEmpty(paramCustom)) return;

        _parametroActivo = paramCustom;

        ActualizarEstiloBotonesMenu();

        GuardarParametro(paramCustom);
        EjecutarComandoExterno();
    }

    private void ActualizarEstiloBotonesMenu()
    {
        if (!(_menuLateral.Children[0] is ScrollViewer sv) || !(sv.Content is StackPanel stackMenu)) return;

        SolidColorBrush activeColor = new SolidColorBrush(WpfColor.FromRgb(126, 224, 99));
        SolidColorBrush inactiveColor = new SolidColorBrush(WpfColor.FromRgb(225, 225, 225));

        foreach (var child in stackMenu.Children)
        {
            if (child is Button btn && btn.Tag is string)
            {
                if (btn.Tag.ToString().Equals(_parametroActivo, StringComparison.OrdinalIgnoreCase))
                {
                    btn.Background = activeColor;
                    btn.Foreground = Brushes.White;
                    btn.FontWeight = FontWeights.Bold;
                }
                else
                {
                    btn.Background = inactiveColor;
                    btn.Foreground = Brushes.Black;
                    btn.FontWeight = FontWeights.Normal;
                }
            }
        }
    }

    private void GuardarParametro(string parametro)
    {
        try
        {
            string directory = System.IO.Path.GetDirectoryName(ConfigFilePath);
            if (!Directory.Exists(directory)) Directory.CreateDirectory(directory);
            File.WriteAllText(ConfigFilePath, parametro);
        }
        catch { }
    }

    private string LeerParametroGuardado()
    {
        try
        {
            if (File.Exists(ConfigFilePath)) return File.ReadAllText(ConfigFilePath).Trim();
        }
        catch { }
        return parametros[0];
    }

    private void EjecutarComandoExterno()
    {
        if (_uiApp != null && _overrideExternalEvent != null)
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
        titulo.Text = $"Valores: {_parametroActivo}";
        titulo.FontSize = 14;
        titulo.FontWeight = FontWeights.Bold;
        titulo.Margin = new Thickness(0, 0, 0, 15);
        _panelLeyenda.Children.Add(titulo);

        var valoresOrdenados = elementosPorValor.OrderBy(kvp => kvp.Key);

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

        StackPanel panelHorizontal = new StackPanel();
        panelHorizontal.Orientation = Orientation.Horizontal;
        panelHorizontal.Margin = new Thickness(0, 4, 0, 4);
        panelHorizontal.HorizontalAlignment = HorizontalAlignment.Left;

        Button btnItem = new Button();
        btnItem.Style = CrearEstiloBotonRedondeado(6);
        btnItem.Background = Brushes.White;
        btnItem.BorderThickness = new Thickness(1);
        btnItem.BorderBrush = new SolidColorBrush(WpfColor.FromRgb(220, 220, 220));
        btnItem.HorizontalContentAlignment = HorizontalAlignment.Left;
        btnItem.Padding = new Thickness(5);
        btnItem.Cursor = Cursors.Hand;
        btnItem.ToolTip = valor;

        btnItem.Width = 300;

        btnItem.Tag = valor;

        btnItem.PreviewMouseDown += (s, e) =>
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                ItemLeyenda_Accion(valor, accion: "aislar");
            }
            else if (e.ChangedButton == MouseButton.Right)
            {
                ItemLeyenda_Accion(valor, accion: "ocultar");
                e.Handled = true;
            }
            else if (e.ChangedButton == MouseButton.Middle)
            {
                ItemLeyenda_Accion(valor, accion: "seleccionar");
                e.Handled = true;
            }
        };

        StackPanel contenidoItem = new StackPanel();
        contenidoItem.Orientation = Orientation.Horizontal;
        contenidoItem.HorizontalAlignment = HorizontalAlignment.Left;

        Border cuadroColor = new Border();
        cuadroColor.Width = 20;
        cuadroColor.Height = 20;
        cuadroColor.CornerRadius = new CornerRadius(4);
        cuadroColor.Background = new SolidColorBrush(colorWpf);
        cuadroColor.BorderBrush = Brushes.Gray;
        cuadroColor.BorderThickness = new Thickness(1);
        cuadroColor.Margin = new Thickness(0, 0, 10, 0);

        TextBlock texto = new TextBlock();
        texto.Text = $"{valor} ({cantidad})";
        texto.VerticalAlignment = VerticalAlignment.Center;
        texto.FontSize = 12;
        texto.TextTrimming = TextTrimming.CharacterEllipsis;
        texto.MaxWidth = 230;
        texto.TextAlignment = TextAlignment.Left;

        contenidoItem.Children.Add(cuadroColor);
        contenidoItem.Children.Add(texto);
        btnItem.Content = contenidoItem;

        Button btnDetalle = new Button();
        btnDetalle.Content = "...";
        btnDetalle.ToolTip = "Ver lista detallada";
        btnDetalle.Width = 30;
        btnDetalle.Height = 30;
        btnDetalle.Margin = new Thickness(5, 0, 0, 0);
        btnDetalle.Style = CrearEstiloBotonRedondeado(15);
        btnDetalle.Background = new SolidColorBrush(WpfColor.FromRgb(240, 240, 240));
        btnDetalle.Click += (s, e) => MostrarDetalle(valor);

        panelHorizontal.Children.Add(btnItem);
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

    private void ItemLeyenda_Accion(string valor, string accion)
    {
        if (_eventHandler != null && _externalEvent != null)
        {
            _eventHandler.ValorSeleccionado = valor;
            _eventHandler.MostrarTodos = false;
            _eventHandler.OcultarElementos = false;
            _eventHandler.SeleccionarElementos = false;

            switch (accion)
            {
                case "ocultar":
                    _eventHandler.OcultarElementos = true;
                    break;
                case "seleccionar":
                    _eventHandler.SeleccionarElementos = true;
                    break;
                case "aislar":
                default:
                    break;
            }

            _externalEvent.Raise();
        }
    }

    private void BtnMostrarTodos_Click(object sender, RoutedEventArgs e)
    {
        if (_eventHandler != null && _externalEvent != null)
        {
            _eventHandler.MostrarTodos = true;
            _eventHandler.OcultarElementos = false;
            _eventHandler.SeleccionarElementos = false;
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