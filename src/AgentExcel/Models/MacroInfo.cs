namespace AgentExcel.Models;

/// <summary>
/// Information about a VBA macro (procedure) inside an Excel workbook.
/// </summary>
/// <param name="Module">The module/component name containing the macro.</param>
/// <param name="ModuleType">The type of the module (e.g. 'Module', 'Class', 'Document', 'Form').</param>
/// <param name="Name">The name of the macro.</param>
/// <param name="Type">The type of procedure ('Sub' or 'Function').</param>
/// <param name="Parameters">The list of parameter declarations (e.g. 'arg1 As String').</param>
public record MacroInfo(string Module, string ModuleType, string Name, string Type, List<string> Parameters);
