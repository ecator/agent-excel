using System;
using System.Collections.Generic;
using System.IO;

using AgentExcel.Models;
using AgentExcel.Models.Requests;
using AgentExcel.Providers;
using AgentExcel.Services;

using NUnit.Framework;

using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel.Tests.Services;

[TestFixture]
[Category("Unit")]
public class MacroServiceTests : BaseTests
{
    private FakeExcelConnectionProvider? _provider;
    private MacroService? _service;

    [SetUp]
    public void Setup()
    {
        _provider = new FakeExcelConnectionProvider();
        _service = new MacroService(_provider);
    }

    [TearDown]
    public void Teardown()
    {
        _provider?.Dispose();
        _provider = null;
    }

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

    [Test]
    public void ExportModule_WithNullOrEmptyModuleName_ThrowsArgumentException()
    {
        // Act & Assert
        Assert.That(() => _service!.ExportModule("Book1", null!, "output.bas"), Throws.ArgumentException);
        Assert.That(() => _service!.ExportModule("Book1", string.Empty, "output.bas"), Throws.ArgumentException);
        Assert.That(() => _service!.ExportModule("Book1", "   ", "output.bas"), Throws.ArgumentException);
    }

    [Test]
    public void ExportModule_WithNullOrEmptyOutputFile_ThrowsArgumentException()
    {
        // Act & Assert
        Assert.That(() => _service!.ExportModule("Book1", "Module1", null!), Throws.ArgumentException);
        Assert.That(() => _service!.ExportModule("Book1", "Module1", string.Empty), Throws.ArgumentException);
        Assert.That(() => _service!.ExportModule("Book1", "Module1", "   "), Throws.ArgumentException);
    }

    [Test]
    public void ExportModule_WithNonExistentDirectory_ThrowsDirectoryNotFoundException()
    {
        // Arrange
        string nonExistentPath = Path.Combine(TestDataTempPath, "non_existent_folder_abc", "module.bas");

        // Act & Assert
        Assert.That(() => _service!.ExportModule("Book1", "Module1", nonExistentPath), Throws.TypeOf<DirectoryNotFoundException>());
    }

    [Test]
    public void ImportModule_WithNullOrEmptyFilePath_ThrowsArgumentException()
    {
        // Act & Assert
        Assert.That(() => _service!.ImportModule("Book1", null!), Throws.ArgumentException);
        Assert.That(() => _service!.ImportModule("Book1", string.Empty), Throws.ArgumentException);
        Assert.That(() => _service!.ImportModule("Book1", "   "), Throws.ArgumentException);
    }

    [Test]
    public void ImportModule_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        // Arrange
        string nonExistentPath = Path.Combine(TestDataTempPath, "non_existent_file_xyz.bas");

        // Act & Assert
        Assert.That(() => _service!.ImportModule("Book1", nonExistentPath), Throws.TypeOf<FileNotFoundException>());
    }

    [Test]
    public void DeleteModule_WithNullOrEmptyModuleName_ThrowsArgumentException()
    {
        // Act & Assert
        Assert.That(() => _service!.DeleteModule("Book1", null!), Throws.ArgumentException);
        Assert.That(() => _service!.DeleteModule("Book1", string.Empty), Throws.ArgumentException);
        Assert.That(() => _service!.DeleteModule("Book1", "   "), Throws.ArgumentException);
    }

    [Test]
    public void GetModuleNameFromFile_WithAttributeVbName_ExtractsModuleName()
    {
        // Arrange
        string testFile = Path.Combine(TestDataTempPath, "temp_macro_with_attr.bas");
        File.WriteAllText(testFile, "Attribute VB_Name = \"CustomModule\"\r\nSub Test()\r\nEnd Sub");

        try
        {
            // Act
            string moduleName = MacroService.GetModuleNameFromFile(testFile);

            // Assert
            Assert.That(moduleName, Is.EqualTo("CustomModule"));
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }

    [Test]
    public void GetModuleNameFromFile_WithoutAttributeVbName_ReturnsFileNameWithoutExtension()
    {
        // Arrange
        string testFile = Path.Combine(TestDataTempPath, "PlainModule.bas");
        File.WriteAllText(testFile, "Sub Test()\r\nEnd Sub");

        try
        {
            // Act
            string moduleName = MacroService.GetModuleNameFromFile(testFile);

            // Assert
            Assert.That(moduleName, Is.EqualTo("PlainModule"));
        }
        finally
        {
            if (File.Exists(testFile)) File.Delete(testFile);
        }
    }

    private class FakeExcelConnectionProvider : IExcelConnectionProvider
    {
        public Excel.Application? GetApp(bool createNew = false) => null;
        public void SafeReleaseComObject(object? obj) { }
        public void NotifyActivity() { }
        public void Dispose() { }
    }
}
