FROM python:3.13-slim

WORKDIR /app

# Instala uv globalmente
RUN pip install uv

# Copia el código fuente
COPY . /app

# Instala dependencias globalmente
RUN pip install -e ".[dev]"

# Comando de entrada
CMD ["uv", "run", "mcp_excel"]
