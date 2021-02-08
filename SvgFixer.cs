using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace VisioSvgFixer
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Visio SVG Fixer");
            Console.WriteLine("by Aaron Dinnage");
            Console.WriteLine();

            if (args.Length < 1)
            {
                Usage();
                return;
            }

            var files = ProcessArguments(args);
            if (files == null || files.Count == 0)
            {
                Usage();
                return;
            }

            FixFiles(files);
        }

        static void Usage()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("      VisioSvgFixer.exe <file_or_folder_paths>...");
            Console.WriteLine();
            Console.WriteLine("All parameters will be considered a file or folder path to be processed.");
            Console.WriteLine("* Files will be individually processed.");
            Console.WriteLine("* Folders will have all .SVG files inside them processed (excludes sub-folders).");
            Console.WriteLine();
        }

        static List<string> ProcessArguments(string[] args)
        {
            var files = new List<string>();

            foreach (string path in args)
            {
                if (Directory.Exists(path))
                {
                    var folderFiles = Directory.GetFiles(path, "*.svg");
                    files.AddRange(folderFiles);

                    Console.WriteLine("Folder: {0} ({1} files)", path, folderFiles.Length);
                }
                else if (File.Exists(path))
                {
                    files.Add(path);

                    Console.WriteLine("File: {0}", path);
                }
                else
                {
                    Console.WriteLine("File or Folder not fount: {0}", path);
                }
            }

            Console.WriteLine();

            return files;
        }

        static void FixFiles(List<string> files)
        {
            foreach (var file in files)
            {
                Console.WriteLine("Processing file: {0} ...", file);
                bool updated = false;

                var xDoc = XDocument.Load(file, LoadOptions.None);
                var ns = xDoc.Root.GetDefaultNamespace();
                var xmlNs = xDoc.Root.GetNamespaceOfPrefix("xml");

                // Visio exports an xml:space="preserve" attribute on the <svg> element preventing reformatting to optimise file size
                var spaceAttribute = xDoc.Root.Attribute(xmlNs + "space");
                if (spaceAttribute != null && String.Equals(spaceAttribute.Value, "preserve", StringComparison.OrdinalIgnoreCase))
                {
                    spaceAttribute.SetValue("default");
                    updated = true;
                }

                // Remove all comments
                var comments = xDoc.DescendantNodes().OfType<XComment>();
                if (comments.Count() != 0)
                {
                    comments.Remove();
                    updated = true;
                }

                // Remove all "title" and "desc" nodes.
                RecursiveRemove(xDoc.Root.Elements(), ns, "title", ref updated);
                RecursiveRemove(xDoc.Root.Elements(), ns, "desc", ref updated);

                if (updated)
                {
                    Console.WriteLine("Saving file: {0}", file);
                    xDoc.Save(file, SaveOptions.DisableFormatting); // SaveOptions.DisableFormatting removes all whitespace
                }
            }
        }

        static void RecursiveRemove(IEnumerable<XElement> subElements, XNamespace ns, string tag, ref bool updated)
        {
            foreach (var subElement in subElements)
            {
                var element = subElement.Element(ns + tag);
                if (element != null)
                {
                    element.Remove();
                    updated = true;
                }

                RecursiveRemove(subElement.Elements(), ns, tag, ref updated);
            }
        }

    }
}
