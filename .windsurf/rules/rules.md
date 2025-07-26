---
trigger: always_on
---

Standardized Development Guidelines for MCP Projects in Python
To ensure a controlled, maintainable, and Cascade-compatible development environment, all projects implementing the Model Context Protocol (MCP) in Python must follow these guidelines:

Language and Style
Project Language: All code must be written in Python (latest stable version).

Naming Conventions:
 - snake_case for variables, functions, and methods
 - PascalCase for classes
 - UPPER_SNAKE_CASE for constants and environment keys
 - Type Hints:
   - Consistent and mandatory throughout the entire - - codebase. Avoid using Any unless absolutely necessary.
 - Readability over Premature Optimization:
   - Code must prioritize clarity and understanding over micro-optimizations that lack real impact.
 - Modular Structure:
   - Split the project into logical modules with well-defined responsibilities. Avoid monolithic files.

Comments and Documentation
 - Informative Comments:
   - Comments must be relevant and meaningful. Avoid generic notes like # put something here or # TODO without context.
 - Code Intention:
   - Explain why something is done, not just what is done—especially in context-handling sections.
 - Docstrings:
   - All public functions must include docstrings with description, parameters, and return types.

Variables and Naming
 - Descriptive Identifiers:
   - Variable names should clearly communicate their purpose. Avoid obscure or excessive abbreviations.
 - Avoid Shadowing Built-ins:
   - Do not use names that overwrite Python’s built-in functions or types (e.g., list, type, id, etc.).
 - Constant Declaration:
   - Constants must be written in UPPER_SNAKE_CASE and separated clearly from regular logic.

Testing and Safety
 - Test Coverage:
   - Unit tests must be included for all significant logic, especially context, memory, and error-handling routines.
 - Error Handling:
   -Use custom exceptions where appropriate. Do not silently ignore important errors.
 - Input Validation:
   -All external inputs must be validated before processing, especially when dealing with dynamic context files.
   - Use Pydantic for input validation
   - Handle exceptions with custom error types when necessary

Final Note
- This guide promotes a readable, traceable, and scalable development pattern for Python projects using MCP and Cascade.
- It is designed to grow with complex architectures and foster long-term technical collaboration.