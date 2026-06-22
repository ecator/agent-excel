using System.Collections.Generic;

using AgentExcel.Models;
using AgentExcel.Services;

using NUnit.Framework;

namespace AgentExcel.Tests.Services;

[TestFixture]
[Category("Unit")]
public class MacroServiceTests
{
    [Test]
    public void ParseVbaProcedures_WithSimpleSubAndFunction_ParsesSuccessfully()
    {
        // Arrange
        string vbaCode = @"
Public Sub SayHello()
    MsgBox ""Hello""
End Sub

Function AddNumbers(a As Integer, b As Integer) As Integer
    AddNumbers = a + b
End Function
";

        // Act
        var result = MacroService.ParseVbaProcedures(vbaCode, "Module1", "Module");

        // Assert
        Assert.That(result, Has.Count.EqualTo(2));

        var helloMacro = result[0];
        Assert.That(helloMacro.Name, Is.EqualTo("SayHello"));
        Assert.That(helloMacro.Type, Is.EqualTo("Sub"));
        Assert.That(helloMacro.Module, Is.EqualTo("Module1"));
        Assert.That(helloMacro.ModuleType, Is.EqualTo("Module"));
        Assert.That(helloMacro.Parameters, Is.Empty);

        var addMacro = result[1];
        Assert.That(addMacro.Name, Is.EqualTo("AddNumbers"));
        Assert.That(addMacro.Type, Is.EqualTo("Function"));
        Assert.That(addMacro.Parameters, Is.EqualTo(new List<string> { "a As Integer", "b As Integer" }));
    }

    [Test]
    public void ParseVbaProcedures_WithLineContinuation_ParsesSuccessfully()
    {
        // Arrange
        string vbaCode = @"
Sub LongMacro(a As String, _
              b As Integer, _
              Optional c As Boolean = False)
    ' Do something
End Sub
";

        // Act
        var result = MacroService.ParseVbaProcedures(vbaCode, "Module1", "Module");

        // Assert
        Assert.That(result, Has.Count.EqualTo(1));
        var macro = result[0];
        Assert.That(macro.Name, Is.EqualTo("LongMacro"));
        Assert.That(macro.Type, Is.EqualTo("Sub"));
        Assert.That(macro.Parameters, Is.EqualTo(new List<string> { "a As String", "b As Integer", "Optional c As Boolean = False" }));
    }
}
