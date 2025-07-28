---
trigger: always_on
---

# Standardized Development Guidelines for MCP Projects in Python

To ensure a controlled, maintainable, and Cascade-compatible development environment, all projects implementing the Model Context Protocol (MCP) in Python must follow these guidelines.

## Language and Style

### Project Language
- All code must be written in **Python**, using the latest stable version (Python ≥3.10 recommended)

### Naming Conventions
- **snake_case** for variables, functions, and methods
- **PascalCase** for classes
- **UPPER_SNAKE_CASE** for constants and environment keys

### Core Principles
- **Type Hints**: Consistent and mandatory throughout the entire codebase. Avoid using `Any` unless absolutely necessary
- **Readability over Premature Optimization**: Code must prioritize clarity and understanding over micro-optimizations that lack real impact
- **Modular Structure**: Split the project into logical modules with well-defined responsibilities. Avoid monolithic files

## Type Hints

All functions and class attributes must include **explicit type annotations**.

### Guidelines
- Avoid `Any` unless no meaningful alternative exists
- Use modern Python type syntax:
  - `list[str]`, `dict[str, int]`, `tuple[...]`
  - `str | None` instead of `Optional[str]`

### Preferred Patterns
- Use `TypedDict` or `pydantic.BaseModel` for structured data
- Use `Literal` and `Enum` for controlled values
- Use `Annotated[...]` when additional metadata is needed
- Use `Final` for constants and immutable declarations
- Use `NewType` to differentiate similar base types with different meanings
- Always specify return types explicitly, even when `None`

### Example
```python
from typing import Final, Literal, NewType
from pydantic import BaseModel

UserId = NewType('UserId', str)
API_VERSION: Final[str] = "1.0.0"

class UserConfig(BaseModel):
    name: str
    role: Literal["admin", "user", "guest"]
    active: bool

def get_user_status(user_id: UserId) -> str | None:
    """Get the current status of a user."""
    # Implementation here
    return None
```

## Comments and Documentation

### Informative Comments
- Comments must be relevant and meaningful
- Avoid generic notes like `# put something here` or `# TODO` without context
- Focus on **why** a block exists, especially in dynamic or context-heavy logic
- Explain **why** something is done, not just **what** is done—especially in context-handling sections

### Docstrings
All public functions must include docstrings with:
- Description of the function's purpose
- Parameters and their types
- Return types and possible values
- Exceptions that may be raised

### Example
```python
def process_context_data(context: dict[str, Any], user_id: UserId) -> ProcessedContext:
    """
    Process raw context data for a specific user.

    This function validates and transforms context data to ensure it meets
    the MCP protocol requirements and user-specific constraints.

    Args:
        context: Raw context data from the client
        user_id: Unique identifier for the requesting user

    Returns:
        ProcessedContext: Validated and transformed context data

    Raises:
        ValidationError: If context data fails validation
        UserNotFoundError: If user_id is invalid
    """
    # Implementation here
    pass
```

## Variables and Naming

### Descriptive Identifiers
- Variable names should clearly communicate their purpose
- Avoid obscure or excessive abbreviations
- Names must convey semantic meaning clearly
- Avoid abbreviations unless industry-standard

### Built-in Safety
- Do not use names that overwrite Python's built-in functions or types (e.g., `list`, `type`, `id`, etc.)

### Constants
- Constants must be written in `UPPER_SNAKE_CASE`
- Keep constants at the top of modules
- Separate constants clearly from regular logic

### Example
```python
# Good
MAX_RETRY_ATTEMPTS: Final[int] = 3
DEFAULT_TIMEOUT_SECONDS: Final[float] = 30.0
user_session_data: dict[str, Any] = {}

# Bad
max_retries = 3  # Not a constant declaration
userData = {}    # Not snake_case
list = []        # Shadows built-in
```

---

## CHECKLIST - Your code MUST include:
- [ ] Complete type hints on everything
- [ ] Pydantic models for data validation
- [ ] Custom exceptions with clear names
- [ ] Comprehensive docstrings
- [ ] Error handling with try/catch
- [ ] Constants properly declared
- [ ] Basic unit tests
- [ ] Modular structure (separate logical concerns)

---

## COMMON MISTAKES TO AVOID:
- Missing return type annotations
- Using generic variable names (data, item, etc.)
- No input validation
- Missing error handling
- No docstrings
- Using built-in names (list, type, id)
- No tests provided

---

## Final Note

This guideline promotes a readable, traceable, and scalable development pattern for Python projects using MCP and Cascade. It is intended for high-resilience systems that interface with AI agents and future-facing orchestration frameworks.

Adherence to these standards is mandatory for all MCP Python projects to ensure consistency, maintainability, and reliability across the entire ecosystem.
