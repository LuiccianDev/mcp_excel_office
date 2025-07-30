<div align="center">
  <h1> MCP Office Excel Server</h1>
  <p>
    <em>Potente servidor para la manipulación programática de documentos Excel (.xlsx) mediante MCP</em>
  </p>

[![Python Version](https://img.shields.io/badge/python-3.11%2B-blue.svg)](https://www.python.org/downloads/)
[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![MCP Compatible](https://img.shields.io/badge/MCP-Compatible-brightgreen)](https://modelcontextprotocol.io)
</Div>
Servidor MCP (Model Context Protocol) para integración con Microsoft Excel que permite crear, modificar y gestionar archivos de Excel de manera programática a través de herramientas MCP estandarizadas.

## 📋 Tabla de Contenidos

- [✨ Características Principales](#-características-principales)
- [🚀 Instalación](#-instalación)
- [⚙️ Configuración](#️-configuración)
- [📚 Uso](#-uso)
- [🧪 Testing](#-testing)
- [🤝 Contribuyendo](#-contribuyendo)

## ✨ Características Principales

- **Procesamiento de Hojas de Cálculo**: Creación, lectura y modificación de archivos Excel (.xlsx)
- **Operaciones de Formato**: Aplicación de estilos, formatos y fórmulas
- **Integración MCP**: Compatible con el Modelo de Contexto de Protocolo para integración con otros servicios
- **Alto Rendimiento**: Optimizado para manejar archivos grandes de manera eficiente
- **Seguro**: Validación de acceso a archivos y manejo de errores robusto

## 🚀 Instalación

### 📋 Requisitos Previos
- Python 3.11 o superior
- Gestor de paquetes UV (recomendado) o pip

### ⚡ Instalación con UV (Recomendado)

```bash
# Instalar dependencias usando uv
uv sync

# Instalar en modo desarrollo (incluye dependencias de desarrollo)
uv sync --dev

# Instalar en modo producción (solo dependencias necesarias)
uv sync --production
```

### 🐍 Instalación con pip

```bash
# Instalar dependencias
pip install .

# Instalar en modo desarrollo
pip install -e ".[dev]"
```

### 🛠️ Scripts de Ayuda

```bash
# Usar el script helper para instalación
python scripts.py install
```



Este proyecto usa `black` como formateador de código Python.

#### Formatear todo el código:
```bash
# Usando uv directamente
uv run black .

# Usando el script helper
python scripts.py format

# Usando el script específico
python format.py
```

#### Verificar formato del código:
```bash
# Usando uv directamente
uv run black --check .

# Usando el script helper
python scripts.py check

# Usando el script específico
python format.py --check
```

## 🚀 Ejecutar el servidor
```bash
# Usando uv
uv run python main.py

# Usando el script helper
python scripts.py run
```

## ⚙️ Configuración

### 🔧 VS Code
El proyecto incluye configuración automática para VS Code que:
- Formatea automáticamente al guardar
- Usa black como formateador
- Organiza imports automáticamente

### ⚡ Configuración de Black
La configuración de black está en `pyproject.toml`:
- Longitud de línea: 88 caracteres
- Target Python: 3.11+
- Excluye directorios estándar (cache, git, etc.)

## 🛠️ Scripts Disponibles

### Básicos
- `python scripts.py format` - Formatear código automáticamente
- `python scripts.py check` - Verificar formato del código
- `python scripts.py run` - Iniciar el servidor MCP
- `python scripts.py install` - Instalar dependencias

### Herramientas MCP
- `mcp_server_excel` - Inicia el servidor MCP para Excel
  ```bash
  mcp_server_excel
  ```

### Pruebas
```bash
# Ejecutar pruebas unitarias
pytest

# Ejecutar pruebas con cobertura
pytest --cov=mcp_excel_server tests/
```

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
