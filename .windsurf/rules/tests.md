---
trigger: manual
---

---

```text
project/
│
├── mcp_word_server/
│ └── core
│ │   └── [módulos del proyecto].py
│ └── tools
│ │   └── [módulos del proyecto].py
│ └── utils
│     └── [módulos del proyecto].py
├── tests/
│ └── test_[modulo].py
├── requirements.txt
└── README.md
```
## ✅ Reglas para el entorno de pruebas

1. **Usar Pytest como framework de testing.**
2. **Todas las pruebas deben residir dentro del directorio `tests/`.**
3. **Todos los archivos de prueba deben comenzar con `test_`** (ejemplo: `test_utils.py`).
4. **Todos los métodos de prueba deben comenzar con `test_`** para que `pytest` los detecte.
5. **Asegurar compatibilidad de importaciones** desde el directorio raíz del proyecto agregando al inicio de cada archivo de prueba:

    ```python
    import sys
    import os

    sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
    ```

6. **Evitar el uso de `print()`** en los tests. En su lugar, usar `assert` para verificaciones automáticas.

---

## 🧪 Estilo y buenas prácticas en tests

- 🧩 **Los tests deben ser pequeños, claros y enfocados.**  
  Un test = una responsabilidad.

- ✅ **Usar `assert` para todas las condiciones.**  
  Ejemplo:
  ```python
  assert func(3) == 9