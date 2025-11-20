using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Windows.Forms;
using WinForms = System.Windows.Forms;

[Transaction(TransactionMode.Manual)]
public class ConfigurarParametrosCopiarCommand : IExternalCommand
{
    private static string ConfigFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "TL_Tools2021",
        "parametros_copiar_config.txt");

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        try
        {
            // Leer parámetros guardados anteriormente
            string parametrosActuales = LeerParametrosGuardados();

            // Mostrar formulario
            FormularioParametrosCopiar formulario = new FormularioParametrosCopiar(parametrosActuales);

            if (formulario.ShowDialog() == WinForms.DialogResult.OK)
            {
                string nuevosParametros = formulario.ParametrosConfig;

                if (!string.IsNullOrWhiteSpace(nuevosParametros))
                {
                    GuardarParametros(nuevosParametros);
                    return Result.Succeeded;
                }
            }

            return Result.Cancelled;
        }
        catch (Exception ex)
        {
            message = $"Error al configurar parámetros: {ex.Message}";
            return Result.Failed;
        }
    }

    private string LeerParametrosGuardados()
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

    private void GuardarParametros(string parametros)
    {
        try
        {
            string directorio = Path.GetDirectoryName(ConfigFilePath);
            if (!Directory.Exists(directorio))
            {
                Directory.CreateDirectory(directorio);
            }
            File.WriteAllText(ConfigFilePath, parametros);
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Error", $"No se pudo guardar la configuración: {ex.Message}");
        }
    }
}

public class FormularioParametrosCopiar : WinForms.Form
{
    private WinForms.TextBox textBox;
    private WinForms.Button btnOk;
    private WinForms.Button btnCancel;
    private WinForms.Label label;
    private WinForms.Label labelEjemplo;

    public string ParametrosConfig { get; private set; }

    public FormularioParametrosCopiar(string valorActual)
    {
        InitializeComponent();
        textBox.Text = valorActual;
    }

    private void InitializeComponent()
    {
        this.Text = "Configurar Parámetros a Copiar";
        this.Width = 450;
        this.Height = 180;
        this.StartPosition = WinForms.FormStartPosition.CenterScreen;
        this.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        label = new WinForms.Label();
        label.Text = "Nombres de parámetros a copiar (separados por comas):";
        label.Location = new System.Drawing.Point(20, 20);
        label.Width = 400;
        label.Height = 20;

        labelEjemplo = new WinForms.Label();
        labelEjemplo.Text = "Ejemplo: AMBIENTE,NIVEL,ACTIVO,SISTEMA";
        labelEjemplo.Location = new System.Drawing.Point(20, 40);
        labelEjemplo.Width = 400;
        labelEjemplo.Height = 20;
        labelEjemplo.ForeColor = System.Drawing.Color.Gray;

        textBox = new WinForms.TextBox();
        textBox.Location = new System.Drawing.Point(20, 65);
        textBox.Width = 400;

        btnOk = new WinForms.Button();
        btnOk.Text = "Aceptar";
        btnOk.Location = new System.Drawing.Point(230, 105);
        btnOk.Click += (s, e) => { ParametrosConfig = textBox.Text; this.DialogResult = WinForms.DialogResult.OK; };

        btnCancel = new WinForms.Button();
        btnCancel.Text = "Cancelar";
        btnCancel.Location = new System.Drawing.Point(320, 105);
        btnCancel.Click += (s, e) => { this.DialogResult = WinForms.DialogResult.Cancel; };

        this.Controls.Add(label);
        this.Controls.Add(labelEjemplo);
        this.Controls.Add(textBox);
        this.Controls.Add(btnOk);
        this.Controls.Add(btnCancel);

        this.AcceptButton = btnOk;
        this.CancelButton = btnCancel;
    }
}