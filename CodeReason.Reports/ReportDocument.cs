/************************************************************************
 * Copyright: Hans Wolff
 *
 * License:  This software abides by the LGPL license terms. For further
 *           licensing information please see the top level LICENSE.txt 
 *           file found in the root directory of CodeReason Reports.
 *
 * Author:   Hans Wolff
 *
 ************************************************************************/

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Text;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Xps.Packaging;
using System.Windows.Xps.Serialization;
using CodeReason.Reports.Document;

namespace CodeReason.Reports
{
    /// <summary>
    ///     Contains a complete report template without data
    /// </summary>
    public class ReportDocument
    {
        /// <summary>
        ///     Gets or sets the page header height
        /// </summary>
        public double PageHeaderHeight { get; set; }

        /// <summary>
        ///     Gets or sets the page footer height
        /// </summary>
        public double PageFooterHeight { get; set; }

        /// <summary>
        ///     Gets the original page height of the FlowDocument
        /// </summary>
        public double PageHeight { get; private set; }

        /// <summary>
        ///     Gets the original page width of the FlowDocument
        /// </summary>
        public double PageWidth { get; private set; }

        /// <summary>
        ///     Gets or sets the optional report name
        /// </summary>
        public string ReportName { get; set; }

        /// <summary>
        ///     Gets or sets the optional report title
        /// </summary>
        public string ReportTitle { get; set; }

        /// <summary>
        ///     XAML image path
        /// </summary>
        public string XamlImagePath { get; set; }

        /// <summary>
        ///     XAML report data
        /// </summary>
        public string XamlData { get; set; }

        /// <summary>
        ///     Gets or sets the compression option which is used to create XPS files
        /// </summary>
        public CompressionOption XpsCompressionOption { get; set; }

        /// <summary>
        ///     Fire event after a page has been completed
        /// </summary>
        /// <param name="ea">GetPageCompletedEventArgs</param>
        public void FireEventGetPageCompleted(GetPageCompletedEventArgs ea)
        {
            if (GetPageCompleted != null) GetPageCompleted(this, ea);
        }

        /// <summary>
        ///     Fire event after a data row has been bound
        /// </summary>
        /// <param name="ea">DataRowBoundEventArgs</param>
        public void FireEventDataRowBoundEventArgs(DataRowBoundEventArgs ea)
        {
            if (DataRowBound != null) DataRowBound(this, ea);
        }

        /// <summary>
        ///     Creates a flow document of the report data
        /// </summary>
        /// <returns></returns>
        /// <exception cref="ArgumentException">XAML data does not represent a FlowDocument</exception>
        /// <exception cref="ArgumentException">Flow document must have a specified page height</exception>
        /// <exception cref="ArgumentException">Flow document must have a specified page width</exception>
        /// <exception cref="ArgumentException">"Flow document must have only one ReportProperties section, but it has {0}"</exception>
        public FlowDocument CreateFlowDocument(Dictionary<string, ImageSource> imageSources)
        {
            var mem = new MemoryStream();
            var buf = Encoding.UTF8.GetBytes(XamlData);
            mem.Write(buf, 0, buf.Length);
            mem.Position = 0;
            var res = XamlReader.Load(mem) as FlowDocument;
            if (res == null) throw new ArgumentException("XAML data does not represent a FlowDocument");

            if (res.PageHeight == double.NaN)
                throw new ArgumentException("Flow document must have a specified page height");
            if (res.PageWidth == double.NaN)
                throw new ArgumentException("Flow document must have a specified page width");

            // remember original values
            PageHeight = res.PageHeight;
            PageWidth = res.PageWidth;

            // search report properties
            var walker = new DocumentWalker();
            var headers = walker.Walk<SectionReportHeader>(res);
            var footers = walker.Walk<SectionReportFooter>(res);
            var properties = walker.Walk<ReportProperties>(res);
            if (properties.Count > 0)
            {
                if (properties.Count > 1)
                    throw new ArgumentException(
                        string.Format("Flow document must have only one ReportProperties section, but it has {0}",
                            properties.Count));
                var prop = properties[0];
                if (prop.ReportName != null) 
                    ReportName = prop.ReportName;
                if (prop.ReportTitle != null) 
                    ReportTitle = prop.ReportTitle;

                // remove properties section from FlowDocument
                var parent = prop.Parent;
                if (parent is FlowDocument)
                {
                    ((FlowDocument)parent).Blocks.Remove(prop);
                    parent = null;
                }
                if (parent is Section)
                {
                    ((Section)parent).Blocks.Remove(prop);
                }
            }

            if (headers.Count > 0)
                PageHeaderHeight = headers[0].PageHeaderHeight;
            if (footers.Count > 0)
                PageFooterHeight = footers[0].PageFooterHeight;

            // make height smaller to have enough space for page header and page footer
            res.PageHeight = PageHeight - PageHeight * (PageHeaderHeight + PageFooterHeight) / 100d;

            // search image objects
            var images = new List<Image>();
            walker.Tag = images;
            walker.VisualVisited += WalkerVisualVisited;
            walker.Walk(res);

            // load all images
            foreach (var image in images)
            {
                if (ImageProcessing != null) ImageProcessing(this, new ImageEventArgs(this, image));
                try
                {
                    if (image.Tag is string)
                    {
                        var img = image.Tag.ToString();
                        if (imageSources.ContainsKey(img))
                            image.Source = imageSources[img];
                        else
                            image.Source =
                                new BitmapImage(new Uri("file:///" + Path.Combine(XamlImagePath, img)));
                    }
                }
                catch (Exception ex)
                {
                    // fire event on exception and check for Handled = true after each invoke
                    if (ImageError != null)
                    {
                        var handled = false;
                        lock (ImageError)
                        {
                            var eventArgs = new ImageErrorEventArgs(ex, this, image);
                            foreach (var ed in ImageError.GetInvocationList())
                            {
                                ed.DynamicInvoke(this, eventArgs);
                                if (eventArgs.Handled)
                                {
                                    handled = true;
                                    break;
                                }
                            }
                        }
                        if (!handled) throw;
                    }
                    else throw;
                }
                if (ImageProcessed != null) ImageProcessed(this, new ImageEventArgs(this, image));
                // TODO: find a better way to specify file names
            }

            return res;
        }

        private void WalkerVisualVisited(object sender, object visitedObject, bool start)
        {
            if (!(visitedObject is Image)) return;

            var walker = sender as DocumentWalker;
            if (walker == null) return;

            var list = walker.Tag as List<Image>;
            if (list == null) return;

            list.Add((Image)visitedObject);
        }

        /// <summary>
        ///     Helper method to create page header or footer from flow document template
        /// </summary>
        /// <param name="data">report data</param>
        /// <returns></returns>
        public XpsDocument CreateXpsDocument(ReportData data)
        {
            var ms = new MemoryStream();
            var pkg = Package.Open(ms, FileMode.Create, FileAccess.ReadWrite);
            var pack = "pack://report.xps";
            PackageStore.RemovePackage(new Uri(pack));
            PackageStore.AddPackage(new Uri(pack), pkg);
            var doc = new XpsDocument(pkg, CompressionOption.NotCompressed, pack);
            var rsm = new XpsSerializationManager(new XpsPackagingPolicy(doc), false);

            var rp = new ReportPaginator(this, data);
            rsm.SaveAsXaml(rp);
            return doc;
        }

        /// <summary>
        ///     Helper method to create page header or footer from flow document template
        /// </summary>
        /// <param name="data">enumerable report data</param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException">data</exception>
        public XpsDocument CreateXpsDocument(IEnumerable<ReportData> data)
        {
            if (data == null) throw new ArgumentNullException("data");
            var count = 0;
            ReportData firstData = null;
            foreach (var rd in data)
            {
                if (firstData == null) firstData = rd;
                count++;
            }
            if (count == 1)
                return CreateXpsDocument(firstData);
            // we have only one ReportData object -> use the normal ReportPaginator instead

            var ms = new MemoryStream();
            var pkg = Package.Open(ms, FileMode.Create, FileAccess.ReadWrite);
            var pack = "pack://report.xps";
            PackageStore.RemovePackage(new Uri(pack));
            PackageStore.AddPackage(new Uri(pack), pkg);
            var doc = new XpsDocument(pkg, CompressionOption.NotCompressed, pack);
            var rsm = new XpsSerializationManager(new XpsPackagingPolicy(doc), false);

            var rp = new MultipleReportPaginator(this, data);
            rsm.SaveAsXaml(rp);
            return doc;
        }

        /// <summary>
        ///     Helper method to create page header or footer from flow document template
        /// </summary>
        /// <param name="data">report data</param>
        /// <param name="fileName">file to save XPS to</param>
        /// <returns></returns>
        public XpsDocument CreateXpsDocument(ReportData data, string fileName)
        {
            var pkg = Package.Open(fileName, FileMode.Create, FileAccess.ReadWrite);
            var pack = "pack://report.xps";
            PackageStore.RemovePackage(new Uri(pack));
            PackageStore.AddPackage(new Uri(pack), pkg);
            var doc = new XpsDocument(pkg, XpsCompressionOption, pack);
            var rsm = new XpsSerializationManager(new XpsPackagingPolicy(doc), false);

            var rp = new ReportPaginator(this, data);
            rsm.SaveAsXaml(rp);
            rsm.Commit();
            pkg.Close();
            return new XpsDocument(fileName, FileAccess.Read);
        }

        /// <summary>
        ///     Helper method to create page header or footer from flow document template
        /// </summary>
        /// <param name="data">enumerable report data</param>
        /// <param name="fileName">file to save XPS to</param>
        /// <returns></returns>
        public XpsDocument CreateXpsDocument(IEnumerable<ReportData> data, string fileName)
        {
            if (data == null) throw new ArgumentNullException("data");
            var count = 0;
            ReportData firstData = null;
            foreach (var rd in data)
            {
                if (firstData == null) firstData = rd;
                count++;
            }
            if (count == 1)
                return CreateXpsDocument(firstData);
            // we have only one ReportData object -> use the normal ReportPaginator instead

            var pkg = Package.Open(fileName, FileMode.Create, FileAccess.ReadWrite);
            var pack = "pack://report.xps";
            PackageStore.RemovePackage(new Uri(pack));
            PackageStore.AddPackage(new Uri(pack), pkg);
            var doc = new XpsDocument(pkg, XpsCompressionOption, pack);
            var rsm = new XpsSerializationManager(new XpsPackagingPolicy(doc), false);

            var rp = new MultipleReportPaginator(this, data);
            rsm.SaveAsXaml(rp);
            rsm.Commit();
            pkg.Close();
            return new XpsDocument(fileName, FileAccess.Read);
        }

        #region Events

        /// <summary>
        ///     Event occurs after a data row is bound
        /// </summary>
        public event EventHandler<DataRowBoundEventArgs> DataRowBound;

        /// <summary>
        ///     Event occurs after a page has been completed
        /// </summary>
        public event GetPageCompletedEventHandler GetPageCompleted;

        /// <summary>
        ///     Event occurs if an exception has encountered while loading the BitmapSource
        /// </summary>
        public event EventHandler<ImageErrorEventArgs> ImageError;

        /// <summary>
        ///     Event occurs before an image is being processed
        /// </summary>
        public event EventHandler<ImageEventArgs> ImageProcessing;

        /// <summary>
        ///     Event occurs after an image has being processed
        /// </summary>
        public event EventHandler<ImageEventArgs> ImageProcessed;

        #endregion
    }
}