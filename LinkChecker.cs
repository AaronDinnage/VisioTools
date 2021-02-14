using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace VisioLinkChecker
{
    class Program
    {
        public class VisioPage
        {
            public string File;
            public string Page;
            public string Id;
            public string DisplayName;
        }

        public class LinkDetails
        {
            public string File;
            public string Page;
            public string Description;
            public string Url;
            public bool Broken;
        }

        enum AppMode
        {
            None = 0,
            Check,
            Update,
        }

        enum UpdateType
        {
            ObjectText,
            LinkText,
            LinkUrl,
        }

        static AppMode mode = AppMode.None;

        static void Main(string[] args)
        {
            Console.WriteLine("Visio Link Checker");
            Console.WriteLine("by Aaron Dinnage");
            Console.WriteLine();

            if (args.Length < 2)
            {
                Usage();
                return;
            }

            var files = ProcessArguments(args);
            if (mode == AppMode.None || files == null || files.Count == 0)
            {
                Usage();
                return;
            }

            switch (mode)
            {
                case AppMode.Check:
                    CheckLinks(files);
                    break;

                case AppMode.Update:
                    UpdateLinks(files);
                    break;
            }
        }

        static void Usage()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("      VisioLinkChecker.exe {check|update} <file_or_folder_paths>...");
            Console.WriteLine();
            Console.WriteLine("The first parameter is either 'check' or 'update' to set the mode of the app.");
            Console.WriteLine("In Update mode a local file named 'LinkUpdates.csv' must be present with a 'find,replace' format inside.");
            Console.WriteLine("All other parameters will be considered a file or folder path to be processed.");
            Console.WriteLine("* Files will be individually processed.");
            Console.WriteLine("* Folders will have all .VSDX files inside them processed (excludes sub-folders).");
            Console.WriteLine();
        }

        static List<string> ProcessArguments(string[] args)
        {
           switch (args[0].ToLowerInvariant())
            {
                case "check":
                case "c":
                    mode = AppMode.Check;
                    break;
                
                case "update":
                case "u":
                    mode = AppMode.Update;
                    break;
                
                default:
                    mode = AppMode.None;
                    return null;
            }

            var files = new List<string>();

            for (int i = 1; i < args.Length; i++)
            {
                string path = args[i];

                if (Directory.Exists(path))
                {
                    var folderFiles = Directory.GetFiles(path, "*.vsdx");
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

        static void CheckLinks(List<string> files)
        {
            Console.WriteLine("Checking links ...");
            Console.WriteLine();

            var pages = new List<VisioPage>();
            var links = new List<LinkDetails>();

            Console.WriteLine("Extracting hyperlinks from files ({0})...", files.Count);
            foreach (var file in files)
                ExtractLinks(file, pages, links);
            Console.WriteLine();

            var urls = new List<string>();
            foreach (var link in links)
                if (!urls.Contains(link.Url))
                    urls.Add(link.Url);

            Console.WriteLine("Processing unique hyperlinks ({0}) ...", urls.Count);
            Parallel.ForEach(urls, new ParallelOptions { MaxDegreeOfParallelism = 8 },
            url =>
            {
                if (!CheckUrl(url))
                {
                    Console.WriteLine("Broken: {0}", url);

                    foreach (var link in links.Where(x => String.Equals(x.Url, url, StringComparison.OrdinalIgnoreCase)))
                        link.Broken = true;
                }
            });
            Console.WriteLine();

            int brokenLinks = links.Where(x => x.Broken).Count();
            if (brokenLinks == 0)
            {
                Console.WriteLine("No broken links found!");
                return;
            }

            Console.WriteLine("Broken Links by File and Page ({0}):", brokenLinks);
            foreach (var file in files)
                foreach (var link in links.Where(x => x.File == file && x.Broken))
                    Console.WriteLine("{0}, {1}, {2}, {3}", link.File, GetPageDisplayName(pages, link.File, link.Page), link.Description, link.Url);
        }

        static void UpdateLinks(List<string> files)
        {
            Console.WriteLine("Updating links ...");
            Console.WriteLine();

            var updates = new List<Tuple<UpdateType, string, string>>();

            using (var streamReader = new StreamReader("LinkUpdates.csv"))
            {
                string line;
                while ((line = streamReader.ReadLine()) != null)
                {
                    if (String.IsNullOrWhiteSpace(line))
                        continue;

                    var components = line.Split(',');
                    if (components.Length != 3)
                    {
                        Console.WriteLine("Error reading CSV, expecting three values per line (type, find, replace).");
                        continue;
                    }

                    UpdateType updateType;
                    if (!Enum.TryParse<UpdateType>(components[0], true, out updateType))
                    {
                        Console.WriteLine("Error reading type in CSV, expecting three values per line (type, find, replace).");
                        continue;
                    }

                    var tuple = new Tuple<UpdateType, string, string>(updateType, components[1], components[2]);
                    updates.Add(tuple);
                }
            }

            if (updates.Count() == 0)
            {
                Console.WriteLine("No updates in CSV");
                return;
            }

            var objectTextUpdates = updates.Where(update => update.Item1 == UpdateType.ObjectText).Select(x => new Tuple<string, string>(x.Item2, x.Item3)).ToList();
            var linkTextUpdates = updates.Where(update => update.Item1 == UpdateType.LinkText).Select(x => new Tuple<string, string>(x.Item2, x.Item3)).ToList();
            var linkUrlUpdates = updates.Where(update => update.Item1 == UpdateType.LinkUrl).Select(x => new Tuple<string, string>(x.Item2, x.Item3)).ToList();

            foreach (var file in files)
            {
                Console.WriteLine("Processing file: {0} ...", file);

                // Open for Read/Write access, we are going to update files in this mode! :)
                using (var fileStream = new FileStream(file, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                using (var zipArchive = new ZipArchive(fileStream, ZipArchiveMode.Update, true))
                {
                    foreach (var entry in zipArchive.Entries)
                    {
                        if (Path.GetDirectoryName(entry.FullName) == "visio\\pages" && entry.Name != "pages.xml") // page?.xml files
                        {
                            using (var entryStream = entry.Open())
                            {
                                var xDoc = XDocument.Load(entryStream);
                                var ns = xDoc.Root.GetDefaultNamespace();

                                bool updated = false;

                                if (objectTextUpdates.Count() > 0)
                                {
                                    var textElements = xDoc.Root.DescendantNodes().OfType<XElement>()
                                        .Where(node => node.Name == ns + "Text");

                                    foreach (var textUpdate in objectTextUpdates)
                                    {
                                        var textMatches = textElements.Where(x => x.Value == textUpdate.Item1 + "\n");
                                        foreach (var match in textMatches)
                                        {
                                            match.Value = textUpdate.Item2;
                                            updated = true;
                                        }
                                    }
                                }

                                if (linkTextUpdates.Count() > 0 || linkUrlUpdates.Count() > 0)
                                {
                                    var linkElements = xDoc.Root.DescendantNodes().OfType<XElement>()
                                        .Where(node => node.Name == ns + "Section").Where(node => node.Attribute("N").Value == "Hyperlink")
                                        .Elements(ns + "Row");

                                    foreach (var linkElement in linkElements)
                                    {
                                        var cells = linkElement.Elements(ns + "Cell");

                                        string description = cells.First(x => x.Attribute("N").Value == "Description").Attribute("V").Value;
                                        string address = cells.First(x => x.Attribute("N").Value == "Address").Attribute("V").Value;
                                        string extraInfo = cells.First(x => x.Attribute("N").Value == "ExtraInfo").Attribute("V").Value;
                                        string url = address;
                                        if (!String.IsNullOrWhiteSpace(extraInfo))
                                            url += "?" + extraInfo;

                                        var linkUrlUpdate = linkUrlUpdates.SingleOrDefault(x=>x.Item1 == url);
                                        if (linkUrlUpdate != null)
                                        {
                                            var newUrlParts = linkUrlUpdate.Item2.Split('?');

                                            string newAddress = newUrlParts[0];
                                            string newExtraInfo = newUrlParts.Length == 2 ? newUrlParts[0] : String.Empty;

                                            cells.First(x => x.Attribute("N").Value == "Address").Attribute("V").Value = newAddress;
                                            cells.First(x => x.Attribute("N").Value == "ExtraInfo").Attribute("V").Value = newExtraInfo;

                                            updated = true;
                                        }

                                        var linkTextUpdate = linkTextUpdates.SingleOrDefault(x=>x.Item1 == description);
                                        if (linkTextUpdate != null)
                                        {
                                            cells.First(x => x.Attribute("N").Value == "Description").Attribute("V").Value = linkTextUpdate.Item2;

                                            updated = true;
                                        }

                                        // Ensure all links open in a New Window ...
                                        var newWindowV = cells.First(x => x.Attribute("N").Value == "NewWindow").Attribute("V");
                                        if (newWindowV.Value == "0")
                                        {
                                            newWindowV.Value = "1";
                                            updated = true;
                                        }
                                    }
                                }

                                // Remove all alt-tags
                                var altTexts = xDoc.Root.DescendantNodes().OfType<XElement>()
                                    .Where(node => node.Name == ns + "Section")
                                    .Where(node => node.Attribute("N").Value == "User")
                                    .Where(node => node.Elements().Count() == 1 && node.Elements().First().Name == ns + "Row" && node.Elements().First().Attribute("N").Value == "visAltText");
                                if (altTexts.Count() > 0)
                                {
                                    altTexts.Remove();
                                    updated = true;
                                }

                                if (updated)
                                {
                                    Console.WriteLine("Saving file page: {0}", entry.FullName);
                                    entryStream.Seek(0, SeekOrigin.Begin);
                                    xDoc.Save(entryStream, SaveOptions.DisableFormatting);
                                    entryStream.SetLength(entryStream.Position);
                                }
                            }
                        }
                    }
                }
            }
        }

        static void ExtractLinks(string file, List<VisioPage> pages, List<LinkDetails> links)
        {
            Console.WriteLine("Processing file: {0} ...", file);

            using (var fileStream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var zipArchive = new ZipArchive(fileStream, ZipArchiveMode.Read, true))
            {
                foreach (var entry in zipArchive.Entries)
                {
                    if (Path.GetDirectoryName(entry.FullName) == "visio\\pages" && entry.Name == "pages.xml")
                    {
                        // Process the pages.xml to get the display names and rID (rels ID)

                        using (var entryStream = entry.Open())
                        {
                            var xDoc = XDocument.Load(entryStream, LoadOptions.PreserveWhitespace);

                            var ns = xDoc.Root.GetDefaultNamespace();
                            var nsR = xDoc.Root.GetNamespaceOfPrefix("r");

                            var pageElements = xDoc.Root.Elements(ns + "Page");
                            foreach (var pageElement in pageElements)
                            {
                                string displayName = pageElement.Attribute("Name").Value;
                                
                                var relElement = pageElement.Element(ns + "Rel");
                                string rid = relElement.Attribute(nsR + "id").Value;

                                var page = pages.FirstOrDefault(x => x.Id == rid);
                                if (page == null)
                                {
                                    page = new VisioPage() { File = file, Id = rid };
                                    pages.Add(page);
                                }

                                page.DisplayName = displayName;
                            }
                        }
                    }
                    else if (Path.GetDirectoryName(entry.FullName) == "visio\\pages\\_rels" && entry.Name == "pages.xml.rels")
                    {
                        // Process the _rels/pages.xml.rels to get the rID to page mapping
                        // <Relationship Target="page1.xml" Type="http://schemas.microsoft.com/visio/2010/relationships/page" Id="rId1"/>

                        using (var entryStream = entry.Open())
                        {
                            var xDoc = XDocument.Load(entryStream, LoadOptions.PreserveWhitespace);
                            var ns = xDoc.Root.GetDefaultNamespace();
                            var relationships = xDoc.Root.Elements(ns + "Relationship");
                            foreach (var relationship in relationships)
                            {
                                string target = relationship.Attribute("Target").Value;
                                string rid = relationship.Attribute("Id").Value;

                                var page = pages.FirstOrDefault(x => x.Id == rid);
                                if (page == null)
                                {
                                    page = new VisioPage() { File = file, Id = rid };
                                    pages.Add(page);
                                }

                                page.Page = target;
                            }
                        }
                    }
                    else if (Path.GetDirectoryName(entry.FullName) == "visio\\pages") // page?.xml
                    {
                        using (var entryStream = entry.Open())
                        {
                            var xDoc = XDocument.Load(entryStream, LoadOptions.PreserveWhitespace);
                            var ns = xDoc.Root.GetDefaultNamespace();

                            var linkElements = xDoc.Root.Element(ns + "Shapes")
                                .Elements(ns + "Shape")
                                .Elements(ns + "Section").Where(x => x.Attribute("N").Value == "Hyperlink")
                                .Elements(ns + "Row");

                            foreach (var linkElement in linkElements)
                            {
                                var cells = linkElement.Elements(ns + "Cell");

                                string url = cells.First(x => x.Attribute("N").Value == "Address").Attribute("V").Value;
                                string description = cells.First(x => x.Attribute("N").Value == "Description").Attribute("V").Value;
                                string extraInfo = cells.First(x => x.Attribute("N").Value == "ExtraInfo").Attribute("V").Value;

                                if (!String.IsNullOrWhiteSpace(extraInfo))
                                    url += "?" + extraInfo;

                                links.Add(new LinkDetails()
                                {
                                    Broken      = false,
                                    Description = description,
                                    File        = file,
                                    Page        = entry.Name,
                                    Url         = url
                                });
                            }
                        }
                    }
                }
            }
        }

        static bool CheckUrl(string url)
        {
            //Console.WriteLine("Checking {0} ...", url);
            try
            {
                var request = WebRequest.Create(url) as HttpWebRequest;
                request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:59.0) Gecko/20100101 Firefox/59.0";

                request.Method = "GET";
                //request.Method = "HEAD";

                using (var response = request.GetResponse() as HttpWebResponse)
                {
                    //Console.WriteLine(response.StatusCode);
                    return (response.StatusCode == HttpStatusCode.OK);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + " - " + url);
                return false;
            }
        }

        static string GetPageDisplayName(List<VisioPage> pages, string file, string page)
        {
            var visioPage = pages.FirstOrDefault(x => x.File == file && x.Page == page);
            if (visioPage == null)
                return page;

            return visioPage.DisplayName;
        }

    }
}
