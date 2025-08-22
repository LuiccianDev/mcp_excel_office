<div align="center">
  <h1>MCP Office Excel Server</h1>
  <p>
    <em>Potente servidor para la manipulación programática de documentos Excel (.xlsx) mediante MCP</em>
  </p>

[![Python Version](https://img.shields.io/badge/python-3.11%2B-blue.svg)](https://www.python.org/downloads/)
[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![MCP Compatible](https://img.shields.io/badge/MCP-Compatible-brightgreen)](https://modelcontextprotocol.io)
[![Ruff](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/ruff/main/assets/badge/v2.json)](https://github.com/astral-sh/ruff)
</div>

## 📖 Descripción

Servidor MCP (Model Context Protocol) para integración con Microsoft Excel que permite crear, modificar y gestionar archivos de Excel de manera programática a través de herramientas MCP estandarizadas.

## 📋 Tabla de Contenidos

- [✨ Características Principales](#-características-principales)
- [🚀 Instalación](#-instalación)
  - [Requisitos Previos](#-requisitos-previos)
  - [Instalación con UV (Recomendado)](#-instalación-con-uv-recomendado)
  - [Instalación con pip](#-instalación-con-pip)
  - [Entorno Virtual (Opcional)](#-entorno-virtual-opcional)
- [⚙️ Configuración](#️-configuración)
- [🚀 Uso Rápido](#-uso-rápido)
- [📚 Uso Avanzado](#-uso-avanzado)
- [🧪 Testing](#-testing)
- [🧩 Estructura del Proyecto](#-estructura-del-proyecto)
- [🔧 Herramientas de Desarrollo](#-herramientas-de-desarrollo)
- [🤝 Contribuyendo](#-contribuyendo)
- [📄 Licencia](#-licencia)

## ✨ Características Principales

- **Procesamiento de Hojas de Cálculo**: Creación, lectura y modificación de archivos Excel (.xlsx)
- **Operaciones de Formato**: Aplicación de estilos, formatos y fórmulas
- **Integración MCP**: Compatible con el Modelo de Contexto de Protocolo para integración con otros servicios
- **Alto Rendimiento**: Optimizado para manejar archivos grandes de manera eficiente
- **Seguro**: Validación de acceso a archivos y manejo de errores robusto

## 🚀 Instalación

### 📋 Requisitos Previos

- Python 3.11 o superior
- [UV](https://github.com/astral-sh/uv) (recomendado) o pip
- Git (para clonar el repositorio)

### 🔄 Clonar el Repositorio

```bash
git clone https://github.com/tu-usuario/mcp_excel_office.git
cd mcp_excel_office
```

### ⚡ Instalación con UV (Recomendado)

1. **Instalar dependencias básicas**:
   ```bash
   uv sync
   ```

2. **Modo desarrollo** (incluye dependencias de desarrollo y testing):
   ```bash
   uv sync --dev
   ```

3. **Modo producción** (solo dependencias necesarias):
   ```bash
   uv sync --production
   ```

### 🐍 Instalación con pip

1. **Instalar el paquete**:
   ```bash
   pip install .
   ```

2. **Modo desarrollo** (instalación editable):
   ```bash
   pip install -e ".[dev]"
   ```

### 🌐 Entorno Virtual (Opcional)

Se recomienda usar un entorno virtual para aislar las dependencias:

```bash
# Crear entorno virtual
uv venv

# Activar en Windows
.venv\Scripts\activate

# Instalar dependencias
uv sync

uv sync --all--groups
# Desactivar entorno virtual
deactivate
```

### 🏗️ Construir el Módulo

Para crear un paquete instalable del proyecto usando `uv build`:

```bash
# Construir el paquete
uv build

# Instalar desde el paquete construido
uv pip install dist/mcp_excel_office-*.whl

```

Para ver todas las opciones disponibles:
```bash
uv build --help
```

#### Formatear todo el código:
```bash
# Usando uv directamente
uv run pre-commit run --all-files
```

#### Formatear solo el código modificado:
```bash
# Usando uv directamente
uv run pre-commit run <file>
```

### Testing

Ejecuta las pruebas unitarias con:

```bash
uv run pytest
```

## ⚙️ Configuración

### 🔧 Configuración del Entorno de Desarrollo

#### VS Code
El proyecto incluye configuración automática para VS Code que:
- Formatea automáticamente al guardar
- Usa black como formateador
- Organiza imports automáticamente

#### Configuración de MCP (Model Context Protocol)


1. Construir con UV
Para crear un paquete instalable:

```bash
# Construir el paquete
uv build

# Instalar el paquete localmente
uv pip install dist/mcp_excel_office-*.whl


```


2. Claude Desktop Git
Añade la siguiente configuración a tu `mcp_config.json` para la integración con Git:

```json
{
    "mcpServers": {
        "officeExcel": {
            "command": "uv",
            "args": ["run", "mcp-office-excel"]
        }
    }
}
```

#### DXT Pack
Para empaquetar el proyecto con DXT:

```bash
# Empaquetar el proyecto
dxt pack
```

Para más información sobre DXT, visita: [DXT en GitHub](https://github.com/anthropics/dxt)


## 🗂 Estructura del Proyecto

```text
mcp-office-excel/
├── mcp_excel/               # Código principal del servidor
│   ├── core/                # Funcionalidades principales
│   ├── tools/               # Herramientas MCP
│   ├── utils/               # Utilidades
│   └── exceptions/          # Excepciones personalizadas
├── tests/                   # Pruebas unitarias
├── format.py                # Script de formateo
├── pyproject.toml           # Configuración del proyecto
├── TOOLS.md                 # Documentación de herramientas MCP
└── README.md                # Documentación principal
```

## 🔧 Herramientas MCP Disponibles

### Operaciones de Libro de Trabajo
- `create_workbook`: Crea un nuevo libro de Excel
- `create_worksheet`: Añade una nueva hoja a un libro existente
- `get_workbook_metadata`: Obtiene metadatos del libro

### Operaciones de Datos
- `write_data_to_excel`: Escribe datos en una hoja de cálculo
- `read_data_from_excel`: Lee datos de una hoja de cálculo

### Operaciones de Formato
- `format_range`: Aplica formato a un rango de celdas
- `set_column_width`: Ajusta el ancho de columnas
- `set_row_height`: Ajusta la altura de filas

Para una documentación detallada de todas las herramientas MCP, consulte [TOOLS.md](TOOLS.md).

## 🌟 Características MCP

### Protocolo de Contexto
- Integración con el ecosistema MCP
- Interfaz estandarizada para operaciones de Excel
- Manejo de errores consistente

### Seguridad
- Validación de rutas de archivos
- Manejo seguro de memoria
- Protección contra inyección de fórmulas

## 📄 Licencia

Este proyecto está bajo la licencia MIT. Ver el archivo [LICENSE](LICENSE) para más detalles.

## 🤝 Contribuir

Las contribuciones son bienvenidas. Por favor, lea las [pautas de contribución](CONTRIBUTING.md) antes de enviar cambios.

### 📋 Pautas de Contribución

1. Haz un fork del repositorio
2. Crea una rama para tu característica (`git checkout -b feature/AmazingFeature`)
3. Haz commit de tus cambios (`git commit -m 'Add some AmazingFeature'`)
4. Haz push a la rama (`git push origin feature/AmazingFeature`)
5. Abre un Pull Request

### 📋 Pautas de Código

- Sigue el estilo de código existente
- Incluye pruebas para nuevas funcionalidades
- Actualiza la documentación según sea necesario
- Asegúrate de que todas las pruebas pasen

## 🐛 Reportar Errores

Si encuentras algún error o tienes sugerencias, por favor [abre un issue](https://github.com/tu-usuario/mcp_excel_office/issues) en GitHub.

## 📄 Licencia

Este proyecto está bajo la Licencia MIT. Consulta el archivo [LICENSE](LICENSE) para más información.

---

<div align="center">
  <p>Creado con por LuiccianDev</p>
  <p>
    <a href="https://github.com/tu-usuario/mcp_excel_office">GitHub</a> |
    <a href="https://modelcontextprotocol.io">MCP</a> |
    <a href="https://pypi.org/project/mcp-excel">PyPI</a>
  </p>
</div>
