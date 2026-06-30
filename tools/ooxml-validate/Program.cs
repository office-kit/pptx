// Secondary OOXML oracle.
//
// xmllint + the ECMA-376 XSDs check that our XML is *structurally* well-formed
// against the schema. They do NOT check the semantic OOXML rules layered on top
// of the grammar (attribute co-constraints, relationship targeting, value
// ranges, part-level invariants) — exactly the rules PowerPoint enforces when it
// decides a file is "corrupt" and offers to repair it.
//
// This program runs Microsoft's own OpenXmlValidator over each generated .pptx,
// prints every ValidationErrorInfo it finds (Id, Description, owning part,
// XPath), and exits non-zero if any file has errors. It is the same engine that
// backs the Open XML SDK Productivity Tool referenced in CLAUDE.md / README.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

if (args.Length == 0)
{
    Console.Error.WriteLine("usage: ooxml-validate <file-or-directory> [more...]");
    return 2;
}

// Accept either explicit .pptx paths or a directory to scan, so CI can pass
// `samples/out` and a developer can pass a single deck.
var files = args
    .SelectMany(arg => Directory.Exists(arg)
        ? Directory.EnumerateFiles(arg, "*.pptx", SearchOption.AllDirectories)
        : new[] { arg })
    .OrderBy(path => path, StringComparer.Ordinal)
    .ToList();

if (files.Count == 0)
{
    Console.Error.WriteLine("ooxml-validate: no .pptx files matched the given paths");
    return 2;
}

// Target the newest format version. The validator only flags features as
// unsupported relative to the *targeted* version, so validating against the
// latest band keeps post-2007 features (transitions, animations, newer chart
// pieces) from showing up as false positives — we want genuine rule
// violations, not version-targeting noise.
var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
var totalErrors = 0;

foreach (var file in files)
{
    if (!File.Exists(file))
    {
        Console.Error.WriteLine($"ooxml-validate: file not found: {file}");
        totalErrors++;
        continue;
    }

    try
    {
        using var doc = PresentationDocument.Open(file, false);
        var errors = validator.Validate(doc).ToList();

        if (errors.Count == 0)
        {
            Console.WriteLine($"OK    {file}");
            continue;
        }

        Console.WriteLine($"FAIL  {file}  ({errors.Count} error(s))");
        foreach (var error in errors)
        {
            Console.WriteLine($"  [{error.Id}] {error.Description}");
            Console.WriteLine($"      part:  {error.Part?.Uri}");
            Console.WriteLine($"      xpath: {error.Path?.XPath}");
        }
        totalErrors += errors.Count;
    }
    catch (Exception ex)
    {
        // A part the SDK cannot even load (wrong root element, malformed XML)
        // throws from Validate() instead of surfacing as a ValidationErrorInfo.
        // Report it as a failure for this file and keep going, so one bad deck
        // doesn't hide the validation status of every deck after it.
        Console.WriteLine($"FAIL  {file}  (could not be loaded by the Open XML SDK)");
        Console.WriteLine($"  [load-error] {ex.Message}");
        totalErrors++;
    }
}

Console.WriteLine();
Console.WriteLine(totalErrors == 0
    ? $"ooxml-validate: all {files.Count} file(s) passed OpenXmlValidator ({FileFormatVersions.Microsoft365})"
    : $"ooxml-validate: {totalErrors} validation error(s) across {files.Count} file(s)");

return totalErrors == 0 ? 0 : 1;
