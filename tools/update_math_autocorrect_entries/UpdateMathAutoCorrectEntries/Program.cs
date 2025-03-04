using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using static System.Net.Mime.MediaTypeNames;

namespace MathAutoCorrectInstaller
{
    internal class Program
    {
        enum Mode { Skip, Interactive, Force }

        static void Main(string[] args)
        {
            Console.WriteLine("Math AutoCorrect Installer");

            Mode mode = Mode.Skip; // Default mode
            bool showHelp = false;

            foreach (string arg in args)
            {
                if (arg.Equals("-i", StringComparison.OrdinalIgnoreCase) ||
                    arg.Equals("--interactive", StringComparison.OrdinalIgnoreCase))
                {
                    mode = Mode.Interactive;
                }
                else if (arg.Equals("-f", StringComparison.OrdinalIgnoreCase) ||
                         arg.Equals("--force", StringComparison.OrdinalIgnoreCase))
                {
                    mode = Mode.Force;
                }
                else if (arg.Equals("-h", StringComparison.OrdinalIgnoreCase) ||
                         arg.Equals("--help", StringComparison.OrdinalIgnoreCase))
                {
                    showHelp = true;
                }
                else
                {
                    Console.WriteLine($"** Error - Unknown argument: {arg}");
                    showHelp = true;
                }
            }

            if (showHelp)
            {
                ShowHelp();
                return;
            }

            switch (mode)
            {
                case Mode.Skip:
                    Console.WriteLine("Adding new Math AutoCorrect entries (skipping existing entries)...\n");
                    break;
                case Mode.Interactive:
                    Console.WriteLine("Adding Math AutoCorrect entries (interactive mode - will prompt before overwriting existing entries)...\n");
                    break;
                case Mode.Force:
                    Console.WriteLine("Adding and updating Math AutoCorrect entries (force overwrite mode)...\n");
                    break;
            }

            Microsoft.Office.Interop.Word.Application wordApp = null;

            try
            {
                wordApp = new Microsoft.Office.Interop.Word.Application();
                wordApp.Visible = false;
                Console.WriteLine("Connected to Word application.");

                // Get existing entries to check for duplicates and to use as templates
                var existingEntries = new Dictionary<string, string>();
                Console.WriteLine("Retrieving existing Math AutoCorrect entries...");

                foreach (OMathAutoCorrectEntry entry in wordApp.OMathAutoCorrect.Entries)
                {
                    // Store the name (without the backslash) and value
                    var name = entry.Name;
                    if (name.StartsWith("\\"))
                    {
                        name = name.Substring(1);
                    }

                    existingEntries[name] = entry.Value;
                }
                Console.WriteLine($"Found {existingEntries.Count} existing Math AutoCorrect entries.");

                // Get the entries to add
                var entriesToAdd = GetProposedEntries();
                Console.WriteLine($"Found {entriesToAdd.Count} Math AutoCorrect entries to precess.");

                // Add the new entries to Word
                int addedCount = 0;
                int skippedCount = 0;
                foreach (var entry in entriesToAdd)
                {
                    try
                    {
                        string value = null;

                        // If the symbol should be derived from an existing entry
                        if (!string.IsNullOrEmpty(entry.ExistingEntry))
                        {
                            string existingName = entry.ExistingEntry.TrimStart('\\').TrimStart('`').TrimEnd('`');

                            if (existingEntries.ContainsKey(existingName))
                            {
                                value = existingEntries[existingName];
                                Console.WriteLine($"Using value from existing entry '{existingName}' for '{entry.Name}'");
                            }
                            else
                            {
                                Console.WriteLine($"** Warning: Referenced existing entry '{existingName}' not found for '{entry.Name}'");
                                // If we can't find the existing entry, use the symbol as fallback
                                value = entry.Symbol;
                            }
                        }
                        else
                        {
                            // Use the symbol directly if no existing entry is referenced
                            value = entry.Symbol;
                        }

                        // Ensure the entry name starts with a backslash
                        string name = entry.Name;
                        if (!name.StartsWith("\\"))
                        {
                            name = "\\" + name;
                        }

                        // Check if the entry already exists
                        bool entryExists = existingEntries.ContainsKey(entry.Name.TrimStart('\\'));
                        bool shouldAdd = true;

                        if (entryExists)
                        {
                            string existingValue = existingEntries[entry.Name.TrimStart('\\')];
                            bool valuesDiffer = (existingValue != value);

                            switch (mode)
                            {
                                case Mode.Skip:
                                    // Skip existing entries
                                    Console.WriteLine($"Skipping existing entry: {entry.Name}");
                                    shouldAdd = false;
                                    break;

                                case Mode.Interactive:
                                    // Only prompt if values are different
                                    if (valuesDiffer)
                                    {
                                        Console.WriteLine($"Entry '{entry.Name}' already exists with a different value.");
                                        Console.Write("Do you want to overwrite it? (y/n): ");
                                        string response = Console.ReadLine()?.Trim().ToLower() ?? "";
                                        shouldAdd = (response == "y" || response == "yes");
                                        
                                        if (!shouldAdd)
                                        {
                                            Console.WriteLine($"Skipping entry: {entry.Name}");
                                        }
                                    }
                                    else
                                    {
                                        // If values are the same, no need to overwrite
                                        Console.WriteLine($"Entry '{entry.Name}' already exists with the same value. Skipping.");
                                        shouldAdd = false;
                                    }
                                    break;

                                case Mode.Force:
                                    // Always overwrite in force mode
                                    Console.WriteLine($"Overwriting existing entry: {entry.Name}");
                                    shouldAdd = true;
                                    break;
                            }
                        }

                        if (!shouldAdd)
                        {
                            skippedCount++;
                            continue;
                        }

                        if (name.Length > 1 && value != null && value.Length > 0)
                        {
                            // If entry exists, remove it first
                            if (entryExists)
                            {
                                try
                                {
                                    wordApp.OMathAutoCorrect.Entries[name].Delete();
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"** Error removing existing entry {name}: {ex.Message}");
                                }
                            }

                            // Add the entry
                            wordApp.OMathAutoCorrect.Entries.Add(Name: name, Value: value);

                            addedCount++;
                            Console.WriteLine($"Added: {name}");
                        }
                        else
                        {
                            Console.WriteLine($"** Error - skipping {name}: Name or value doesn't meet length requirements");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"** Failed to add {entry.Name}: {ex.Message}");
                    }
                }

                switch (mode)
                {
                    case Mode.Skip:
                        Console.WriteLine($"\nSuccessfully added {addedCount} new Math AutoCorrect entries (skipped {skippedCount} existing entries).");
                        break;
                    case Mode.Interactive:
                        Console.WriteLine($"\nSuccessfully added {addedCount} Math AutoCorrect entries in interactive mode.");
                        break;
                    case Mode.Force:
                        Console.WriteLine($"\nSuccessfully added/updated {addedCount} Math AutoCorrect entries (force mode).");
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                // Clean up Word application
                if (wordApp != null)
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp);
                }

                Console.WriteLine("\nPress any key to exit...");
                Console.ReadKey();
            }
        }


        static void ShowHelp()
        {
            Console.WriteLine("Math AutoCorrect Installer - Command Line Options");
            Console.WriteLine("================================================");
            Console.WriteLine("Usage: MathAutoCorrectInstaller.exe [options]");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine("  -i, --interactive    Interactive mode - prompt before overwriting entries");
            Console.WriteLine("  -f, --force          Force overwrite existing entries without asking");
            Console.WriteLine("  -h, --help           Show this help message");
            Console.WriteLine();
            Console.WriteLine("Modes:");
            Console.WriteLine("  Default (no options)  Skip any existing entries automatically");
            Console.WriteLine("  Interactive           Ask before overwriting existing entries with different values");
            Console.WriteLine("  Force                 Overwrite all existing entries without asking");
            Console.WriteLine();
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        static List<AutoCorrectEntry> GetProposedEntries()
        {
            // Create a list of all proposed entries directly
            var entries = new List<AutoCorrectEntry>
            {
                // Currency Symbols
                new AutoCorrectEntry { Name = "cents", Symbol = "¢", ExistingEntry = "" },

                // Numbers - Repeating decimal
                new AutoCorrectEntry { Name = "repeat", Symbol = "‾", ExistingEntry = "\\overbar" },
                new AutoCorrectEntry { Name = "repeating", Symbol = "‾", ExistingEntry = "\\overbar" },
                new AutoCorrectEntry { Name = "vinculum", Symbol = "‾", ExistingEntry = "\\overbar" },

                // Basic Symbols and Operations
                new AutoCorrectEntry { Name = "infinity", Symbol = "∞", ExistingEntry = "\\infty" },
                new AutoCorrectEntry { Name = "2root", Symbol = "√", ExistingEntry = "\\sqrt" },
                new AutoCorrectEntry { Name = "3root", Symbol = "∛", ExistingEntry = "\\cbrt" },
                new AutoCorrectEntry { Name = "4root", Symbol = "∜", ExistingEntry = "\\qdrt" },
                new AutoCorrectEntry { Name = "comp", Symbol = "∘", ExistingEntry = "\\circ" },
                new AutoCorrectEntry { Name = "deg", Symbol = "°", ExistingEntry = "\\degree" },
                new AutoCorrectEntry { Name = "rad", Symbol = "㎭", ExistingEntry = "" },
                new AutoCorrectEntry { Name = "join", Symbol = "⋈", ExistingEntry = "\\bowtie" },
                new AutoCorrectEntry { Name = "qed", Symbol = "∎", ExistingEntry = "" },
                new AutoCorrectEntry { Name = "endproof", Symbol = "∎", ExistingEntry = "" },

                // Geometry and Measurement
                new AutoCorrectEntry { Name = "circle", Symbol = "◯", ExistingEntry = "" },
                new AutoCorrectEntry { Name = "circledot", Symbol = "⊙", ExistingEntry = "\\odot" },
                new AutoCorrectEntry { Name = "line", Symbol = "⃡", ExistingEntry = "\\tvec" },
                new AutoCorrectEntry { Name = "seg", Symbol = "‾", ExistingEntry = "\\overbar" },
                new AutoCorrectEntry { Name = "measangle", Symbol = "∡", ExistingEntry = "" },
                new AutoCorrectEntry { Name = "rightangle", Symbol = "∟", ExistingEntry = "" },
                new AutoCorrectEntry { Name = "triangle", Symbol = "△", ExistingEntry = "" },
                new AutoCorrectEntry { Name = "parallelogram", Symbol = "▱", ExistingEntry = "\\underline" },
                new AutoCorrectEntry { Name = "notparallel", Symbol = "∦", ExistingEntry = "" },
                new AutoCorrectEntry { Name = "ray", Symbol = "⃗", ExistingEntry = "\\vec" },
                new AutoCorrectEntry { Name = "arc", Symbol = "⌢", ExistingEntry = "\\overparen" },

                // Inequalities and Relations
                new AutoCorrectEntry { Name = "nlt", Symbol = "≮", ExistingEntry = "/<" },
                new AutoCorrectEntry { Name = "notlt", Symbol = "≮", ExistingEntry = "/<" },
                new AutoCorrectEntry { Name = "ngt", Symbol = "≯", ExistingEntry = "/>" },
                new AutoCorrectEntry { Name = "notgt", Symbol = "≯", ExistingEntry = "/>" },
                new AutoCorrectEntry { Name = "nleq", Symbol = "≰", ExistingEntry = "/\\le" },
                new AutoCorrectEntry { Name = "notle", Symbol = "≰", ExistingEntry = "/\\le" },
                new AutoCorrectEntry { Name = "nge", Symbol = "≱", ExistingEntry = "/\\ge" },
                new AutoCorrectEntry { Name = "notge", Symbol = "≱", ExistingEntry = "/\\ge" },
                new AutoCorrectEntry { Name = "ngeq", Symbol = "≱", ExistingEntry = "/\\ge" },
                new AutoCorrectEntry { Name = "notgeq", Symbol = "≱", ExistingEntry = "/\\ge" },
                new AutoCorrectEntry { Name = "not", Symbol = "¬", ExistingEntry = "\\neg" },
                new AutoCorrectEntry { Name = "muchgreater", Symbol = "≫", ExistingEntry = "\\gg" },
                new AutoCorrectEntry { Name = "muchless", Symbol = "≪", ExistingEntry = "\\ll" },
                new AutoCorrectEntry { Name = "notapprox", Symbol = "≉", ExistingEntry = "/\\approx" },
                new AutoCorrectEntry { Name = "notcong", Symbol = "≢", ExistingEntry = "/\\cong" },

                // Calculus
                new AutoCorrectEntry { Name = "doubleint", Symbol = "∬", ExistingEntry = "\\iint" },
                new AutoCorrectEntry { Name = "tripleint", Symbol = "∭", ExistingEntry = "\\iiint" },
                new AutoCorrectEntry { Name = "dprime", Symbol = "″", ExistingEntry = "\\pprime" },
                new AutoCorrectEntry { Name = "doubleprime", Symbol = "″", ExistingEntry = "\\pprime" },
                new AutoCorrectEntry { Name = "tprime", Symbol = "‴", ExistingEntry = "\\ppprime" },
                new AutoCorrectEntry { Name = "tripleprime", Symbol = "‴", ExistingEntry = "\\ppprime" },
                new AutoCorrectEntry { Name = "qprime", Symbol = "⁗", ExistingEntry = "\\pppprime" },
                new AutoCorrectEntry { Name = "quadprime", Symbol = "⁗", ExistingEntry = "\\pppprime" },
                new AutoCorrectEntry { Name = "grad", Symbol = "∇", ExistingEntry = "\\nabla" },
                new AutoCorrectEntry { Name = "laplace", Symbol = "∆", ExistingEntry = "\\inc" },

                // Set Theory
                new AutoCorrectEntry { Name = "union", Symbol = "∪", ExistingEntry = "\\cup" },
                new AutoCorrectEntry { Name = "Union", Symbol = "⋃", ExistingEntry = "\\bigcup" },
                new AutoCorrectEntry { Name = "intersection", Symbol = "∩", ExistingEntry = "\\cap" },
                new AutoCorrectEntry { Name = "Intersection", Symbol = "⋂", ExistingEntry = "\\bigcap" },
                new AutoCorrectEntry { Name = "notsubset", Symbol = "⊄", ExistingEntry = "/\\subset" },
                new AutoCorrectEntry { Name = "notsuperset", Symbol = "⊅", ExistingEntry = "/\\superset" },
                new AutoCorrectEntry { Name = "notsubseteq", Symbol = "⊈", ExistingEntry = "/\\subseteq" },
                new AutoCorrectEntry { Name = "notsuperseteq", Symbol = "⊉", ExistingEntry = "/\\superseteq" },
                new AutoCorrectEntry { Name = "subsetnoteq", Symbol = "⊊", ExistingEntry = "/\\subseteq" },
                new AutoCorrectEntry { Name = "supersetnoteq", Symbol = "⊋", ExistingEntry = "/\\superseteq" },
                new AutoCorrectEntry { Name = "belongs", Symbol = "∈", ExistingEntry = "\\in" },
                new AutoCorrectEntry { Name = "element", Symbol = "∈", ExistingEntry = "\\in" },
                new AutoCorrectEntry { Name = "contains", Symbol = "∋", ExistingEntry = "\\ni" },
                new AutoCorrectEntry { Name = "owns", Symbol = "∋", ExistingEntry = "\\ni" },
                new AutoCorrectEntry { Name = "powerset", Symbol = "℘", ExistingEntry = "\\wp" },
                new AutoCorrectEntry { Name = "complement", Symbol = "∁", ExistingEntry = "" },

                // Number Theory
                new AutoCorrectEntry { Name = "divide", Symbol = "∣", ExistingEntry = "\\mid" },
                new AutoCorrectEntry { Name = "notdivide", Symbol = "∤", ExistingEntry = "" },

                // Logic/Boolean Operations
                new AutoCorrectEntry { Name = "and", Symbol = "∧", ExistingEntry = "\\wedge" },
                new AutoCorrectEntry { Name = "land", Symbol = "∧", ExistingEntry = "\\wedge" },
                new AutoCorrectEntry { Name = "or", Symbol = "∨", ExistingEntry = "\\vee" },
                new AutoCorrectEntry { Name = "lor", Symbol = "∨", ExistingEntry = "\\vee" },
                new AutoCorrectEntry { Name = "nand", Symbol = "⊼", ExistingEntry = "" },
                new AutoCorrectEntry { Name = "nor", Symbol = "⊽", ExistingEntry = "" },
                new AutoCorrectEntry { Name = "xor", Symbol = "⊕", ExistingEntry = "\\oplus" },
                new AutoCorrectEntry { Name = "xnor", Symbol = "⊙", ExistingEntry = "\\odot" },

                // Logic Proofs and Reasoning
                new AutoCorrectEntry { Name = "proves", Symbol = "⊢", ExistingEntry = "\\vdash" },
                new AutoCorrectEntry { Name = "tautology", Symbol = "⊤", ExistingEntry = "\\top" },
                new AutoCorrectEntry { Name = "false", Symbol = "⊥", ExistingEntry = "\\bot" },
                new AutoCorrectEntry { Name = "contradiction", Symbol = "⊥", ExistingEntry = "\\bot" },
                new AutoCorrectEntry { Name = "implication", Symbol = "→", ExistingEntry = "\\rightarrow" },
                new AutoCorrectEntry { Name = "implies", Symbol = "→", ExistingEntry = "\\rightarrow" },
                new AutoCorrectEntry { Name = "biconditional", Symbol = "↔", ExistingEntry = "\\leftrightarrow" },
                new AutoCorrectEntry { Name = "Implication", Symbol = "⇒", ExistingEntry = "\\rightarrow" },
                new AutoCorrectEntry { Name = "Implies", Symbol = "⇒", ExistingEntry = "\\rightarrow" },
                new AutoCorrectEntry { Name = "Biconditional", Symbol = "⇔", ExistingEntry = "\\leftrightarrow" },
                new AutoCorrectEntry { Name = "forces", Symbol = "⊩", ExistingEntry = "" },
                new AutoCorrectEntry { Name = "entailment", Symbol = "⊨", ExistingEntry = "\\models" },
                new AutoCorrectEntry { Name = "true", Symbol = "⊨", ExistingEntry = "\\models" },
                new AutoCorrectEntry { Name = "foreach", Symbol = "∀", ExistingEntry = "\\forall" },
                new AutoCorrectEntry { Name = "forsome", Symbol = "∃", ExistingEntry = "\\exists" },

                // Statistics and Probability
                new AutoCorrectEntry { Name = "stddev", Symbol = "σ", ExistingEntry = "\\sigma" },
                new AutoCorrectEntry { Name = "mean", Symbol = "μ", ExistingEntry = "\\mu" },
                new AutoCorrectEntry { Name = "corr", Symbol = "ρ", ExistingEntry = "\\rho" },
                new AutoCorrectEntry { Name = "expect", Symbol = "𝔼", ExistingEntry = "\\doubleE" },
                new AutoCorrectEntry { Name = "prob", Symbol = "ℙ", ExistingEntry = "\\doubleP" },

                // Matrix Operations
                new AutoCorrectEntry { Name = "kron", Symbol = "⊗", ExistingEntry = "\\otimes" },
                new AutoCorrectEntry { Name = "hadamard", Symbol = "⊙", ExistingEntry = "\\odot" },
                new AutoCorrectEntry { Name = "adjoint", Symbol = "†", ExistingEntry = "" },
                new AutoCorrectEntry { Name = "identity", Symbol = "𝐈", ExistingEntry = "" },
                new AutoCorrectEntry { Name = "directsum", Symbol = "⊕", ExistingEntry = "\\oplus" }
            };

            return entries;
        }
    }

    class AutoCorrectEntry
    {
        public string Name { get; set; }
        public string Symbol { get; set; }
        public string ExistingEntry { get; set; }
    }
}