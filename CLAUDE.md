# Project Context: .NET Backend

## Tech Stack
- **Language/Framework**: .NET (C#)
- **Testing Framework**: xUnit
- **Mocking Library**: NSubstitute
- **MCP Server**: Built using the latest ModelContextProtocol NuGet package

## Coding Standards
- Use file-scoped namespaces.
- Follow standard .NET naming conventions (PascalCase for methods/classes).
- When writing tests, use the naming pattern: `MethodName_StateUnderTest_ExpectedBehavior`.
- Always ensure a `.gitignore` file exists in the project root. If creating a new project, include a C#/.NET-appropriate `.gitignore`.

## Test Patterns
- Always use **NSubstitute** for mocking dependencies (e.g., `Substitute.For<T>()`).
- Use xUnit attributes (`[Fact]`, `[Theory]`, `[InlineData]`).
- Assertions should use standard xUnit `Assert` methods.

## Build & Test Commands
- **Build**: `dotnet build`
- **Run Tests**: `dotnet test`
- **Clean**: `dotnet clean`