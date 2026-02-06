# Teams Transcript Downloader

🎯 **Extensión de Chrome para descargar transcripciones de Microsoft Stream/SharePoint en formato JSON**

![Version](https://img.shields.io/badge/version-1.0.0-blue)
![Chrome](https://img.shields.io/badge/chrome-compatible-green)
![License](https://img.shields.io/badge/license-MIT-purple)

## ✨ Características

- 📝 Extrae transcripciones completas de grabaciones de Microsoft Teams/Stream
- 💾 Descarga en formato JSON estructurado
- 📅 Nomenclatura automática: `AAMMDD_PROYECTO_Transcripcion.json`
- 🎨 Interfaz moderna con diseño premium
- ⚡ Detección automática de transcripciones

## 📦 Instalación

### Método 1: Carga como extensión sin empaquetar (Desarrollo)

1. **Abre Chrome** y navega a `chrome://extensions/`

2. **Activa el Modo desarrollador** (esquina superior derecha)

3. **Haz clic en "Cargar descomprimida"** 

4. **Selecciona la carpeta** `Chrome_Teams` (esta carpeta)

5. ¡Listo! La extensión aparecerá en tu barra de herramientas

### Crear iconos PNG (Requerido)

Antes de instalar, necesitas crear los iconos PNG desde el SVG. Puedes usar cualquier herramienta de conversión:

**Opción A: Usando un editor online**
1. Abre `icons/icon.svg` en un editor SVG online (como [svgtopng.com](https://svgtopng.com))
2. Exporta en tamaños: 16x16, 32x32, 48x48, 128x128
3. Guarda como `icon16.png`, `icon32.png`, `icon48.png`, `icon128.png` en la carpeta `icons/`

**Opción B: Usando ImageMagick (línea de comandos)**
```bash
cd icons
magick icon.svg -resize 16x16 icon16.png
magick icon.svg -resize 32x32 icon32.png
magick icon.svg -resize 48x48 icon48.png
magick icon.svg -resize 128x128 icon128.png
```

## 🚀 Uso

1. **Abre una grabación** de Microsoft Teams en SharePoint/Stream

2. **Asegúrate de que la transcripción esté visible** en el panel lateral

3. **Haz clic en el icono** de la extensión en la barra de herramientas

4. **Ingresa el nombre del proyecto** (ej: "Reunion_PERTE_EPSAR")

5. **Haz clic en "Descargar Transcripción"**

6. El archivo se guardará como: `260206_Reunion_PERTE_EPSAR_Transcripcion.json`

## 📄 Formato del archivo JSON

```json
{
  "metadata": {
    "projectName": "Reunion_PERTE_EPSAR",
    "exportDate": "2026-02-06T11:30:00.000Z",
    "fileName": "260206_Reunion_PERTE_EPSAR_Transcripcion.json",
    "source": "https://...",
    "duration": "01:23:45",
    "totalEntries": 150,
    "speakers": ["Usuario 1", "Usuario 2"]
  },
  "transcript": [
    {
      "index": 1,
      "timestamp": "00:00:05",
      "speaker": "Usuario 1",
      "text": "Buenos días a todos..."
    },
    ...
  ]
}
```

## 🔧 Estructura del proyecto

```
Chrome_Teams/
├── manifest.json      # Configuración de la extensión
├── popup.html         # Interfaz del popup
├── popup.js           # Lógica del popup
├── content.js         # Script de contenido
├── styles.css         # Estilos CSS
├── icons/
│   ├── icon.svg       # Icono fuente
│   ├── icon16.png     # 16x16
│   ├── icon32.png     # 32x32
│   ├── icon48.png     # 48x48
│   └── icon128.png    # 128x128
└── README.md          # Este archivo
```

## 🐛 Solución de problemas

### "No se encontró transcripción"
- Asegúrate de que el panel de transcripción esté visible en la página
- Algunos videos pueden no tener transcripción disponible
- Espera a que la página cargue completamente

### "Página no compatible"
- La extensión solo funciona en páginas de SharePoint y Microsoft Stream
- Verifica que la URL contenga `.sharepoint.com` o `.microsoft.com`

## 📋 Permisos requeridos

- `activeTab`: Para acceder al contenido de la pestaña actual
- `scripting`: Para inyectar scripts de extracción
- Host permissions para SharePoint y Microsoft

## 📝 Licencia

MIT License - Siéntete libre de usar y modificar.

---

Desarrollado con ❤️ para facilitar la gestión de transcripciones de reuniones.
