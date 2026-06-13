using System;
using System.IO;

[assembly: NUnit.Framework.NonParallelizable]

namespace AgentExcel.Tests;

/// <summary>
/// Base class for all test classes, providing shared utility methods and setup logic.
/// </summary>
public abstract class BaseTests
{
    /// <summary>
    /// Gets the absolute path to the TestData directory.
    /// </summary>
    protected string TestDataPath { get; }

    /// <summary>
    /// Gets the absolute path to the TestData/temp directory.
    /// </summary>
    protected string TestDataTempPath { get; }

    protected BaseTests()
    {
        string baseDir = AppContext.BaseDirectory;
        string testDataDir = FindTestDataDir(baseDir);

        TestDataPath = testDataDir;
        TestDataTempPath = Path.Combine(testDataDir, "temp");

        // Ensure temp directory exists
        if (!Directory.Exists(TestDataTempPath))
        {
            Directory.CreateDirectory(TestDataTempPath);
        }
    }

    private static string FindTestDataDir(string startDir)
    {
        string current = startDir;
        while (!string.IsNullOrEmpty(current))
        {
            // 1. Try finding 'tests/TestData' under current directory
            string candidate1 = Path.Combine(current, "tests", "TestData");
            if (Directory.Exists(candidate1))
            {
                return Path.GetFullPath(candidate1);
            }

            // 2. Try finding 'TestData' directly if we are already inside 'tests'
            string candidate2 = Path.Combine(current, "TestData");
            if (Directory.Exists(candidate2))
            {
                string parentName = Path.GetFileName(current);
                if (string.Equals(parentName, "tests", StringComparison.OrdinalIgnoreCase))
                {
                    return Path.GetFullPath(candidate2);
                }
            }

            string? parent = Directory.GetParent(current)?.FullName;
            if (parent == current || string.IsNullOrEmpty(parent))
            {
                break;
            }
            current = parent;
        }

        // Fallback: search relative to AppContext.BaseDirectory directly if the search failed
        string fallback = Path.GetFullPath(Path.Combine(startDir, "..", "..", "..", "TestData"));
        if (Directory.Exists(fallback))
        {
            return fallback;
        }

        throw new DirectoryNotFoundException($"Could not find tests/TestData directory starting from {startDir}");
    }
}

