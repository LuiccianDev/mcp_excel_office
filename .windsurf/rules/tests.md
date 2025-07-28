---
trigger: manual
---

# Reglas de Desarrollo - Proyecto Python MCP Excel

## Estructura del Proyecto
```text
project/
│
├── src/
│    └── mcp_excel/
│    │     └── core/
│    │     │    └── [módulos del proyecto].py
│    │     └── tools/
│    │     │    └── [módulos del proyecto].py
│    │     └── utils/
│    │     │     └── [módulos del proyecto].py
│    │     └── config/
│    │     │     └── mcp_config.py
│    │     └── exceptions/
│    │     │     └── [módulos del proyecto].py
│    │     └── validators/
│    │     │    └── [módulos del proyecto].py
│    │     └── main.py
│    └── mcp_server.py
├── tests/
│     └── test_[modulo].py
├── requirements.txt
├── pyproyect.toml
├── uv.lock
├── .python-version
└── README.md
```

---

## Reglas de Estructura y Organización

### Organización de Módulos
- **`core/`**: Lógica principal y funcionalidades centrales del proyecto
- **`tools/`**: Herramientas y utilidades específicas para el manejo de Excel/MCP
- **`utils/`**: Funciones auxiliares y helpers generales
- **`config/`**: Configuraciones del proyecto (archivos de configuración MCP)
- **`exceptions/`**: Definición de excepciones personalizadas
- **`validators/`**: Funciones de validación de datos y parámetros

### Convenciones de Nomenclatura
- **Archivos**: `snake_case.py` (ejemplo: `excel_processor.py`)
- **Clases**: `PascalCase` (ejemplo: `ExcelProcessor`)
- **Funciones y variables**: `snake_case` (ejemplo: `process_excel_file`)
- **Constantes**: `UPPER_SNAKE_CASE` (ejemplo: `MAX_FILE_SIZE`)

---

## Reglas para el Entorno de Pruebas

### Framework y Estructura de Tests
1. **Usar Pytest como framework de testing**
2. **Todas las pruebas deben residir dentro del directorio `tests/`**
3. **Todos los archivos de prueba deben comenzar con `test_`** (ejemplo: `test_utils.py`)
4. **Todos los métodos de prueba deben comenzar con `test_`** para que `pytest` los detecte

### 🔧 Configuración de Importaciones
5. **Asegurar compatibilidad de importaciones** desde el directorio raíz del proyecto agregando al inicio de cada archivo de prueba:
    ```python
    import sys
    import os
    sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
    ```

### Buenas Prácticas en Tests
6. **Evitar el uso de `print()`** en los tests. En su lugar, usar `assert` para verificaciones automáticas
7. **Los tests deben ser pequeños, claros y enfocados** - Un test = una responsabilidad
8. **Usar `assert` para todas las condiciones**
   ```python
   assert func(3) == 9
   assert len(result) > 0
   assert isinstance(response, dict)
   ```

### Estructura Recomendada de un Test
```python
import pytest
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.mcp_excel.utils.example import function_to_test

@pytest.mark.unit
def test_function_name():
    # Arrange (Preparación)
    input_data = "test_data"
    expected_result = "expected_output"

    # Act (Acción)
    result = function_to_test(input_data)

    # Assert (Verificación)
    assert result == expected_result

@pytest.mark.asyncio
async def test_async_function():
    # Para funciones asíncronas
    result = await async_function_to_test()
    assert result is not None

@pytest.mark.slow
def test_slow_operation():
    # Para tests que tardan mucho
    pass
```

### Uso de Markers
- **`@pytest.mark.unit`**: Para tests unitarios
- **`@pytest.mark.integration`**: Para tests de integración
- **`@pytest.mark.slow`**: Para tests que tardan mucho
- **`@pytest.mark.asyncio`**: Para tests asíncronos
- **`@pytest.mark.excel`**: Para tests específicos de Excel

---

## Reglas de Calidad de Código

### Documentación
- **Todas las funciones públicas deben tener docstrings**
- **Usar formato Google Style para docstrings**:
  ```python
  def process_excel(file_path: str) -> dict:
      """Procesa un archivo Excel y retorna datos estructurados.

      Args:
          file_path (str): Ruta al archivo Excel

      Returns:
          dict: Datos procesados del archivo Excel

      Raises:
          FileNotFoundError: Si el archivo no existe
          ValueError: Si el archivo no es válido
      """
  ```

### Manejo de Errores
- **Usar excepciones personalizadas** definidas en `exceptions/`
- **Validar entradas** usando funciones del módulo `validators/`
- **Proporcionar mensajes de error descriptivos**

### Estilo de Código
- **Seguir PEP 8** para el estilo de código Python
- **Usar type hints** para todas las funciones públicas
- **Limionar líneas a 88 caracteres** (compatible con Black formatter)
- **Usar f-strings** para formateo de cadenas

---

## Reglas de Desarrollo

### Control de Versiones
- **Commits pequeños y descriptivos**
- **Usar mensajes de commit en español** siguiendo el formato:
  ```
  tipo: descripción breve

  Descripción más detallada si es necesaria
  ```
  Tipos: `feat`, `fix`, `docs`, `test`, `refactor`, `style`

### Testing antes de Commit
- **Ejecutar todos los tests** antes de hacer commit:
  ```bash
  # Ejecutar todos los tests
  pytest

  # Ejecutar tests con cobertura detallada
  pytest --cov=src/mcp_excel --cov-report=term-missing

  # Ejecutar solo tests unitarios
  pytest -m unit

  # Ejecutar tests excluyendo los lentos
  pytest -m "not slow"

  # Ejecutar tests específicos de un módulo
  pytest tests/test_utils.py
  ```
- **Asegurar cobertura mínima del 80%** (configurado automáticamente)
- **No hacer commit de código que rompa tests existentes**

### Dependencias
- **Documentar todas las dependencias** en `requirements.txt`
- **Usar versiones específicas** para dependencias críticas
- **Revisar y actualizar dependencias regularmente**
