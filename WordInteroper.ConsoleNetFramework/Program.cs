using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using CSharpFunctionalExtensions;
using WordInteroper.Interfaces;
using WordInteroper.Models;

namespace WordInteroper.ConsoleNetFramework
{
    public class SetCheckBox : IWordCheckBox
    {
        public SetCheckBox(string title, string tag, bool value)
        {
            Title = title;
            Tag = tag;
            Checked = value;
        }

        public string Title { get; }
        public string Tag { get; }
        public bool Checked { get; }
    }

    internal class Program
    {
        private static readonly List<TokenReplacement> TokenReplacements = new List<TokenReplacement>
        {
            new TokenReplacement {Token = "__DATE__", Replacement = DateTime.Today.ToString(CultureInfo.InvariantCulture)}
        };

        private static readonly SetCheckBox[] CheckboxValues =
        {
            new SetCheckBox(title: "Checkbox-Title", tag: "Checkbox-tag", value: true),
        };
        
        private static void Main(string[] args)
        {
            Console.WriteLine("Enter filepath to docx document.");
            string path = Console.ReadLine();
            using (WordEdit wordEdit = WordEdit.OpenEdit(path, LoadMode.KeepAlive))
            {
                Result replaceTokensResult = wordEdit.ReplaceTokens(TokenReplacements);

                if (replaceTokensResult.IsFailure)
                {
                    Console.WriteLine(replaceTokensResult.Error);
                    Console.WriteLine("Press any key to close the program. No changes will be made.");
                    Console.ReadKey();
                    return;
                }

                Result setCheckboxResult = wordEdit.SetCheckboxes(CheckboxValues);
                if (setCheckboxResult.IsFailure)
                {
                    Console.WriteLine(setCheckboxResult.Error);
                    Console.WriteLine("Press any key to close the program. No changes will be made.");
                }

                string fileName = Path.GetFileNameWithoutExtension(wordEdit.File.FullName);
                string newPath = Path.Combine(wordEdit.File.Directory.FullName, $"{fileName}.pdf");
                Result saveAsPdfResult = wordEdit.ExportAsPdf(newPath);

                string outputMessage = saveAsPdfResult.IsSuccess
                    ? $"Pdf saved at: {newPath}"
                    : saveAsPdfResult.Error;

                Console.WriteLine(outputMessage);
            }

            Console.WriteLine("All done. Press any key to quit");
        }
    }
}