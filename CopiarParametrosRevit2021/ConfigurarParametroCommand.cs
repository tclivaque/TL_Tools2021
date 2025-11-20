using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.IO;
using WinForms = System.Windows.Forms;

[Transaction(TransactionMode.Manual)]
public class ConfigurarParametroCommand : IExternalCommand
{
    private static string ConfigFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "TL_Tools2021",
        "parametro_config.txt");

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        try
        {
            // Leer parámetro guardado anteriormente (si existe)
            string parametroActual = LeerParametroGuardado();

            // Mostrar formulario directamente
            FormularioParametro formulario = new FormularioParametro(parametroActual);

            if (formulario.ShowDialog() == WinForms.DialogResult.OK)
            {
                string nuevoParametro = formulario.NombreParametro;

                if (!string.IsNullOrWhiteSpace(nuevoParametro))
                {
                    GuardarParametro(nuevoParametro);
                    return Result.Succeeded;
                }
            }

            return Result.Cancelled;
        }
        catch (Exception ex)
        {
            message = $"Error al configurar parámetro: {ex.Message}";
            return Result.Failed;
        }
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
        return string.Empty;
    }

    private void GuardarParametro(string nombreParametro)
    {
        try
        {
            string directorio = Path.GetDirectoryName(ConfigFilePath);
            if (!Directory.Exists(directorio))
            {
                Directory.CreateDirectory(directorio);
            }
            File.WriteAllText(ConfigFilePath, nombreParametro);
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Error", $"No se pudo guardar la configuración: {ex.Message}");
        }
    }
}

// Formulario simple para entrada de texto
public class FormularioParametro : WinForms.Form
{
    private WinForms.TextBox textBox;
    private WinForms.Button btnOk;
    private WinForms.Button btnCancel;
    private WinForms.Label label;

    public string NombreParametro { get; private set; }

    public FormularioParametro(string valorActual)
    {
        InitializeComponent();
        textBox.Text = valorActual;
    }

    private void InitializeComponent()
    {
        this.Text = "Configurar Parámetro";
        this.Width = 400;
        this.Height = 150;
        this.StartPosition = WinForms.FormStartPosition.CenterScreen;
        this.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        label = new WinForms.Label();
        label.Text = "Nombre del parámetro:";
        label.Location = new System.Drawing.Point(20, 20);
        label.AutoSize = true;

        textBox = new WinForms.TextBox();
        textBox.Location = new System.Drawing.Point(20, 45);
        textBox.Width = 340;

        btnOk = new WinForms.Button();
        btnOk.Text = "Aceptar";
        btnOk.Location = new System.Drawing.Point(180, 80);
        btnOk.Click += (s, e) => { NombreParametro = textBox.Text; this.DialogResult = WinForms.DialogResult.OK; };

        btnCancel = new WinForms.Button();
        btnCancel.Text = "Cancelar";
        btnCancel.Location = new System.Drawing.Point(270, 80);
        btnCancel.Click += (s, e) => { this.DialogResult = WinForms.DialogResult.Cancel; };

        this.Controls.Add(label);
        this.Controls.Add(textBox);
        this.Controls.Add(btnOk);
        this.Controls.Add(btnCancel);

        this.AcceptButton = btnOk;
        this.CancelButton = btnCancel;
    }
}