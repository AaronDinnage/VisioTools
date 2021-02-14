using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

                // Strip whitespace first
                string originalXml = File.ReadAllText(file);
                string outputXml = originalXml;
                outputXml = Regex.Replace(outputXml, @"\n[\t]*", "\n"); // Remove whitespace at the start of a line
                outputXml = Regex.Replace(outputXml, @">[\s\r\n]*<", "><"); // Remove white space between tags
                if (!String.Equals(originalXml, outputXml))
                {
                    File.WriteAllText(file, outputXml);
                    updated = true;
                }

                var xDoc = ReadFile(file);
                var ns = xDoc.Root.GetDefaultNamespace();

                // Visio exports an xml:space="preserve" attribute on the <svg> element
                var xmlNs = xDoc.Root.GetNamespaceOfPrefix("xml");
                var spaceAttribute = xDoc.Root.Attribute(xmlNs + "space");
                if (spaceAttribute != null && String.Equals(spaceAttribute.Value, "preserve", StringComparison.OrdinalIgnoreCase))
                {
                    spaceAttribute.Remove();
                    WriteFile(file, xDoc);

                    // Reload ...
                    xDoc = ReadFile(file);
                    ns = xDoc.Root.GetDefaultNamespace();
                    xmlNs = null;

                    updated = true;
                }

                var docType = xDoc.DocumentType;
                if (docType != null)
                {
                    docType.Remove();
                    updated = true;
                }

                var comments = xDoc.DescendantNodes().OfType<XComment>();
                if (comments.Count() != 0)
                {
                    comments.Remove();
                    updated = true;
                }

                // XML Events namespace
                var evNs = xDoc.Root.Attributes().Where(x => x.IsNamespaceDeclaration && x.Name.LocalName == "ev").FirstOrDefault();
                if (evNs != null)
                {
                    evNs.Remove();
                    updated = true;
                }

                var rootDimensions = xDoc.Root.Attributes().Where(x => x.Name == "width" || x.Name == "height");
                if (rootDimensions.Count() != 0)
                {
                    rootDimensions.Remove();
                    updated = true;
                }

                var titleNodes = xDoc.Root.DescendantNodes().OfType<XElement>().Where(x => x.Name == ns + "title");
                if (titleNodes.Count() != 0)
                {
                    titleNodes.Remove();
                    updated = true;
                }

                var descNodes = xDoc.Root.DescendantNodes().OfType<XElement>().Where(x => x.Name == ns + "desc");
                if (descNodes.Count() != 0)
                {
                    descNodes.Remove();
                    updated = true;
                }

                var idAttributes = xDoc.Root.DescendantNodes().OfType<XElement>().Where(x => x.Attribute("id") != null);
                if (idAttributes.Count() != 0)
                {
                    idAttributes.Attributes("id").Remove();
                    updated = true;
                }

                if (updated)
                {
                    Console.WriteLine("Saving file: {0}", file);
                    WriteFile(file, xDoc);
                }
            }
        }

        private static void WriteFile(string file, XDocument xDoc)
        {
            var xmlWriterSettings = new XmlWriterSettings() { Indent = false, OmitXmlDeclaration = true };
            using (var xmlWriter = XmlWriter.Create(file, xmlWriterSettings))
                xDoc.Save(xmlWriter);
        }

        private static XDocument ReadFile(string file)
        {
            var readerSettings = new XmlReaderSettings() { IgnoreWhitespace = true };
            using (var xmlReader = XmlReader.Create(file, readerSettings))
                return XDocument.Load(file, LoadOptions.None);
        }
    }
}
