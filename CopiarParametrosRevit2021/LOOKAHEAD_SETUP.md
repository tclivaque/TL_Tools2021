# Configuración del Módulo LOOKAHEAD

## Requisitos Previos

Para que el módulo LOOKAHEAD funcione correctamente, necesitas configurar las credenciales de Google Sheets.

## Pasos de Instalación

### 1. Archivo de Credenciales

El módulo LOOKAHEAD requiere un archivo de credenciales de Google Service Account para acceder a Google Sheets.

**Ubicación del archivo:**
```
CopiarParametrosRevit2021/revitsheetsintegration-89c34b39c2ae.json
```

**IMPORTANTE:** Por seguridad, este archivo NO se incluye en el repositorio. Debes obtenerlo del administrador del proyecto o crearlo tú mismo.

### 2. Formato del Archivo de Credenciales

El archivo debe tener el siguiente formato JSON:

```json
{
  "type": "service_account",
  "project_id": "revitsheetsintegration",
  "private_key_id": "...",
  "private_key": "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n",
  "client_email": "revit-service@revitsheetsintegration.iam.gserviceaccount.com",
  "client_id": "...",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "...",
  "universe_domain": "googleapis.com"
}
```

### 3. Instalación de Paquetes NuGet

El proyecto requiere los siguientes paquetes NuGet:

- `Google.Apis` (v1.68.0)
- `Google.Apis.Auth` (v1.68.0)
- `Google.Apis.Core` (v1.68.0)
- `Google.Apis.Sheets.v4` (v1.68.0.3421)
- `Newtonsoft.Json` (v13.0.1)

Instala estos paquetes usando NuGet Package Manager en Visual Studio:

```
Install-Package Google.Apis -Version 1.68.0
Install-Package Google.Apis.Auth -Version 1.68.0
Install-Package Google.Apis.Core -Version 1.68.0
Install-Package Google.Apis.Sheets.v4 -Version 1.68.0.3421
Install-Package Newtonsoft.Json -Version 13.0.1
```

### 4. Configuración de Google Sheets

El módulo está configurado para acceder a:

- **Spreadsheet ID:** `1DPSRZDrqZkCxaHQrIIaz5NSf5m3tLJcggvAx9k8x9SA`
- **Hojas:**
  - `LOOKAHEAD`: Contiene los datos de planificación
  - `CONFIG_ACTIVIDADES`: Contiene las reglas de configuración

Asegúrate de que el service account tenga acceso de lectura a este spreadsheet.

## Uso

Una vez configurado, el módulo LOOKAHEAD estará disponible en Revit en la pestaña **TL_Tools2021** > Panel **LOOKAHEAD**:

- **Asignar:** Asigna Look Ahead a elementos desde Google Sheets
- **Membrete:** Actualiza el membrete del plano LPS-S automáticamente

## Solución de Problemas

### Error: "No se encontró el archivo de credenciales"

Verifica que el archivo `revitsheetsintegration-89c34b39c2ae.json` esté en la carpeta raíz del proyecto.

### Error: "No se pudo autenticar con Google Sheets"

- Verifica que el archivo JSON tenga el formato correcto
- Asegúrate de que el service account tenga permisos en el spreadsheet
- Verifica que la conexión a Internet esté activa

## Seguridad

**¡NUNCA subas el archivo de credenciales al repositorio!**

El archivo `.gitignore` está configurado para ignorar automáticamente archivos de credenciales.
