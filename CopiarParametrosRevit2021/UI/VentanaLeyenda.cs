using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using RevitColor = Autodesk.Revit.DB.Color;
using WpfColor = System.Windows.Media.Color;

public class VentanaLeyenda : Window
{
    private static VentanaLeyenda _instancia;
    private LeyendaEventHandler _eventHandler;
    private ExternalEvent _externalEvent;
    private StackPanel _panelLeyenda;
    private Button _btnMostrarTodos;

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
        this.Width = 420;
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

        System.Windows.Controls.Grid gridPrincipal = new System.Windows.Controls.Grid();
        gridPrincipal.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        gridPrincipal.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

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

        gridPrincipal.Children.Add(scrollViewer);
        gridPrincipal.Children.Add(_btnMostrarTodos);

        this.Content = gridPrincipal;
    }

    public void InicializarEventHandler(UIApplication app)
    {
        if (_eventHandler == null)
        {
            _eventHandler = new LeyendaEventHandler();
            _externalEvent = ExternalEvent.Create(_eventHandler);
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

        // Botón para aislar elementos (izquierda)
        Button btnAislar = new Button();
        btnAislar.Background = Brushes.Transparent;
        btnAislar.BorderThickness = new Thickness(1);
        btnAislar.BorderBrush = Brushes.LightGray;
        btnAislar.HorizontalContentAlignment = HorizontalAlignment.Left;
        btnAislar.Padding = new Thickness(5);
        btnAislar.Cursor = System.Windows.Input.Cursors.Hand;
        btnAislar.Width = 280;
        btnAislar.Click += (s, e) => ItemLeyenda_Click(valor);

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

    private void ItemLeyenda_Click(string valor)
    {
        if (_eventHandler != null && _externalEvent != null)
        {
            _eventHandler.ValorSeleccionado = valor;
            _eventHandler.MostrarTodos = false;
            _externalEvent.Raise();
        }
    }

    private void BtnMostrarTodos_Click(object sender, RoutedEventArgs e)
    {
        if (_eventHandler != null && _externalEvent != null)
        {
            _eventHandler.MostrarTodos = true;
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