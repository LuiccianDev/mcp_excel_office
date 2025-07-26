# MCP Office Excel Server 2

Servidor MCP para integración con Microsoft Excel que permite crear, modificar y gestionar archivos de Excel a través de herramientas MCP.

## Instalación

```bash
# Instalar dependencias usando uv
uv sync

# O usar el script helper
python scripts.py install
```

## Desarrollo

### Formateo de Código

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

### Ejecutar el servidor:
```bash
# Usando uv
uv run python main.py

# Usando el script helper
python scripts.py run
```

## Configuración

### VS Code
El proyecto incluye configuración automática para VS Code que:
- Formatea automáticamente al guardar
- Usa black como formateador
- Organiza imports automáticamente

### Black Configuration
La configuración de black está en `pyproject.toml`:
- Longitud de línea: 88 caracteres
- Target Python: 3.11+
- Excluye directorios estándar (cache, git, etc.)

## Scripts Disponibles

- `python scripts.py format` - Formatear código
- `python scripts.py check` - Verificar formato
- `python scripts.py run` - Ejecutar servidor
- `python scripts.py install` - Instalar dependencias

## Estructura del Proyecto

```
mcp-office-excel-2/
├── mcp_excel_server/     # Código principal del servidor
│   ├── core/            # Funcionalidades principales
│   ├── tools/           # Herramientas MCP
│   ├── utils/           # Utilidades
│   └── exceptions/      # Excepciones personalizadas
├── scripts.py           # Scripts de desarrollo
├── format.py           # Script de formateo
├── pyproject.toml      # Configuración del proyecto
└── README.md           # Este archivo
```
