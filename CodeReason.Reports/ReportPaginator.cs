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
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using CodeReason.Reports.Document;
using CodeReason.Reports.Interfaces;

namespace CodeReason.Reports
{
    /// <summary>
    ///     Creates all pages of a report
    /// </summary>
    public class ReportPaginator : DocumentPaginator
    {
        private int _pageCount;

        private Size _pageSize = Size.Empty;

        /// <summary>
        ///     Constructor
        /// </summary>
        /// <param name="report">report document</param>
        /// <param name="data">report data</param>
        /// <exception cref="ArgumentException">Flow document must have a specified page height</exception>
        /// <exception cref="ArgumentException">Flow document must have a specified page width</exception>
        /// <exception cref="ArgumentException">Flow document can have only one report header section</exception>
        /// <exception cref="ArgumentException">Flow document can have only one report footer section</exception>
        public ReportPaginator(ReportDocument report, ReportData data)
        {
            _report = report;
            _data = data;

            _flowDocument = report.CreateFlowDocument(data.ReportImages);
            _pageSize = new Size(_flowDocument.PageWidth, _flowDocument.PageHeight);

            if (_flowDocument.PageHeight == double.NaN)
                throw new ArgumentException("Flow document must have a specified page height");
            if (_flowDocument.PageWidth == double.NaN)
                throw new ArgumentException("Flow document must have a specified page width");

            _dynamicCache = new ReportPaginatorDynamicCache(_flowDocument);
            var listPageHeaders = _dynamicCache.GetFlowDocumentVisualListByType(typeof(SectionReportHeader));
            if (listPageHeaders.Count > 1)
                throw new ArgumentException("Flow document can have only one report header section");
            if (listPageHeaders.Count == 1) _blockPageHeader = (SectionReportHeader)listPageHeaders[0];
            var listPageFooters = _dynamicCache.GetFlowDocumentVisualListByType(typeof(SectionReportFooter));
            if (listPageFooters.Count > 1)
                throw new ArgumentException("Flow document can have only one report footer section");
            if (listPageFooters.Count == 1) _blockPageFooter = (SectionReportFooter)listPageFooters[0];

            _paginator = ((IDocumentPaginatorSource)_flowDocument).DocumentPaginator;

            // remove header and footer in our working copy
            var block = _flowDocument.Blocks.FirstBlock;
            while (block != null)
            {
                var thisBlock = block;
                block = block.NextBlock;
                if ((thisBlock == _blockPageHeader) || (thisBlock == _blockPageFooter))
                    _flowDocument.Blocks.Remove(thisBlock);
            }

            // get report context values
            _reportContextValues = _dynamicCache.GetFlowDocumentVisualListByInterface(typeof(IInlineContextValue));

            FillData();
        }

        /// <summary>
        ///     Determines if the current page count is valid
        /// </summary>
        public override bool IsPageCountValid
        {
            get { return _paginator.IsPageCountValid; }
        }

        /// <summary>
        ///     Gets the total page count
        /// </summary>
        public override int PageCount
        {
            get { return _pageCount; }
        }

        /// <summary>
        ///     Gets or sets the page size
        /// </summary>
        public override Size PageSize
        {
            get { return _pageSize; }
            set { _pageSize = value; }
        }

        /// <summary>
        ///     Gets the paginator source
        /// </summary>
        public override IDocumentPaginatorSource Source
        {
            get { return _paginator.Source; }
        }

        protected void RememberAggregateValue(Dictionary<string, List<object>> aggregateValues, string aggregateGroups,
            object value)
        {
            if (string.IsNullOrEmpty(aggregateGroups)) return;

            var aggregateGroupParts = aggregateGroups.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);

            // remember value for aggregate functions
            List<object> aggregateValueList;
            foreach (var aggregateGroup in aggregateGroupParts)
            {
                var trimmedGroup = aggregateGroup.Trim();
                if (string.IsNullOrEmpty(trimmedGroup)) continue;
                if (!aggregateValues.TryGetValue(trimmedGroup, out aggregateValueList))
                {
                    aggregateValueList = new List<object>();
                    aggregateValues[trimmedGroup] = aggregateValueList;
                }
                aggregateValueList.Add(value);
            }
        }

        /// <summary>
        ///     Fill charts with data
        /// </summary>
        /// <param name="charts">list of charts</param>
        /// <exception cref="InvalidProgramException">window.Content is not a FrameworkElement</exception>
        protected virtual void FillCharts(ArrayList charts)
        {
            Window window = null;

            // fill charts
            foreach (IChart chart in charts)
            {
                if (chart == null) continue;
                var chartCanvas = chart as Canvas;
                if (string.IsNullOrEmpty(chart.TableName)) continue;
                if (string.IsNullOrEmpty(chart.TableColumns)) continue;

                var table = _data.GetDataTableByName(chart.TableName);
                if (table == null) continue;

                if (chartCanvas != null)
                {
                    // HACK: this here is REALLY dirty!!!
                    var newChart = (IChart)chart.Clone();
                    if (window == null)
                    {
                        window = new Window();
                        window.WindowStyle = WindowStyle.None;
                        window.BorderThickness = new Thickness(0);
                        window.ShowInTaskbar = false;
                        window.Left = 30000;
                        window.Top = 30000;
                        window.Show();
                    }
                    window.Width = chartCanvas.Width + 2 * SystemParameters.BorderWidth;
                    window.Height = chartCanvas.Height + 2 * SystemParameters.BorderWidth;
                    window.Content = newChart;

                    newChart.DataColumns = null;

                    newChart.DataView = table.DefaultView;
                    newChart.DataColumns = chart.TableColumns.Split(',', ';');
                    newChart.UpdateChart();

                    var windowContent = window.Content as FrameworkElement;
                    if (windowContent == null)
                        throw new InvalidProgramException("window.Content is not a FrameworkElement");
                    var bitmap = new RenderTargetBitmap((int)(windowContent.RenderSize.Width * 600d / 96d),
                        (int)(windowContent.RenderSize.Height * 600d / 96d), 600d, 600d, PixelFormats.Pbgra32);
                    bitmap.Render(window);
                    chartCanvas.Children.Add(new Image { Source = bitmap });
                }
                else
                {
                    chart.DataColumns = null;

                    chart.DataView = table.DefaultView;
                    chart.DataColumns = chart.TableColumns.Split(',', ';');
                    chart.UpdateChart();
                }
            }

            if (window != null) window.Close();
        }

        /// <summary>
        ///     Fills document with data
        /// </summary>
        /// <exception cref="InvalidDataException">ReportTableRow must have a TableRowGroup as parent</exception>
        protected virtual void FillData()
        {
            var blockDocumentValues = _dynamicCache.GetFlowDocumentVisualListByInterface(typeof(IInlineDocumentValue));
            // walker.Walk<IInlineDocumentValue>(_flowDocument);
            var blockTableRows = _dynamicCache.GetFlowDocumentVisualListByInterface(typeof(ITableRowForDataTable));
            var conditionalRows = _dynamicCache.GetFlowDocumentVisualListByInterface(typeof (ITableRowConditional));
            // walker.Walk<TableRowForDataTable>(_flowDocument);
            var blockAggregateValues = _dynamicCache.GetFlowDocumentVisualListByType(typeof(InlineAggregateValue));
            // walker.Walk<InlineAggregateValue>(_flowDocument);
            var charts = _dynamicCache.GetFlowDocumentVisualListByInterface(typeof(IChart));
            // walker.Walk<IChart>(_flowDocument);
            var dynamicHeaderTableRows =
                _dynamicCache.GetFlowDocumentVisualListByInterface(typeof(ITableRowForDynamicHeader));
            var dynamicDataTableRows =
                _dynamicCache.GetFlowDocumentVisualListByInterface(typeof(ITableRowForDynamicDataTable));
            var documentConditions = _dynamicCache.GetFlowDocumentVisualListByInterface(typeof(IDocumentCondition));

            var blocks = new List<Block>();
            if (_blockPageHeader != null) blocks.Add(_blockPageHeader);
            if (_blockPageFooter != null) blocks.Add(_blockPageFooter);

            var walker = new DocumentWalker();
            blockDocumentValues.AddRange(walker.TraverseBlockCollection<IInlineDocumentValue>(blocks));

            var aggregateValues = new Dictionary<string, List<object>>();

            FillCharts(charts);

            // hide conditional text blocks
            foreach (IDocumentCondition dc in documentConditions)
            {
                if (dc == null) continue;
                dc.PerformRenderUpdate(_data);
            }

            // fill report values
            foreach (IInlineDocumentValue dv in blockDocumentValues)
            {
                if (dv == null) continue;
                object obj;
                if ((dv.PropertyName != null) && (_data.ReportDocumentValues.TryGetValue(dv.PropertyName, out obj)))
                {
                    dv.Value = obj;
                    RememberAggregateValue(aggregateValues, dv.AggregateGroup, obj);
                }
                else
                {
                    if ((_data.ShowUnknownValues) && (dv.Value == null))
                        dv.Value = "[" + (dv.PropertyName ?? "NULL") + "]";
                    RememberAggregateValue(aggregateValues, dv.AggregateGroup, null);
                }
            }

            // fill dynamic tables
            foreach (ITableRowForDynamicDataTable iTableRow in dynamicDataTableRows)
            {
                var tableRow = iTableRow as TableRow;
                if (tableRow == null) continue;

                var tableGroup = tableRow.Parent as TableRowGroup;
                if (tableGroup == null) continue;

                TableRow currentRow;

                var table = _data.GetDataTableByName(iTableRow.TableName);

                for (var i = 0; i < table.Rows.Count; i++)
                {
                    currentRow = new TableRow();

                    var dataRow = table.Rows[i];
                    for (var j = 0; j < table.Columns.Count; j++)
                    {
                        var value = dataRow[j].ToString();
                        currentRow.Cells.Add(new TableCell(new Paragraph(new Run(value))));
                    }
                    tableGroup.Rows.Add(currentRow);
                }
            }

            foreach (ITableRowForDynamicHeader iTableRow in dynamicHeaderTableRows)
            {
                var tableRow = iTableRow as TableRow;
                if (tableRow == null) continue;

                var table = _data.GetDataTableByName(iTableRow.TableName);

                foreach (DataRow row in table.Rows)
                {
                    var value = row[0].ToString();
                    var tableCell = new TableCell(new Paragraph(new Run(value)));
                    tableRow.Cells.Add(tableCell);
                }
            }

            foreach (ITableRowConditional row in conditionalRows)
            {
                var tableRow = row as TableRow;
                if (tableRow == null) continue;

                var rowGroup = tableRow.Parent as TableRowGroup;
                if (rowGroup == null) continue;

                //if (row.Visible == false)
                //{
                //    rowGroup.Rows.Remove(tableRow);
                //}

                var table = _data.GetDataTableByName(row.TableName);
                if (table == null)
                {
                    rowGroup.Rows.Remove(tableRow);
                }
            }

            // group table row groups
            var groupedRows = new Dictionary<TableRowGroup, List<TableRow>>();
            var tableNames = new Dictionary<TableRowGroup, string>();
            foreach (TableRow tableRow in blockTableRows)
            {
                var rowGroup = tableRow.Parent as TableRowGroup;
                if (rowGroup == null) continue;

                var iTableRow = tableRow as ITableRowForDataTable;
                if ((iTableRow != null) && (iTableRow.TableName != null))
                {
                    string tableName;
                    if (tableNames.TryGetValue(rowGroup, out tableName))
                    {
                        if (tableName != iTableRow.TableName.Trim().ToLowerInvariant())
                            throw new ReportingException(
                                "TableRowGroup cannot be mapped to different DataTables in TableRowForDataTable");
                    }
                    else tableNames[rowGroup] = iTableRow.TableName.Trim().ToLowerInvariant();
                }

                List<TableRow> rows;
                if (!groupedRows.TryGetValue(rowGroup, out rows))
                {
                    rows = new List<TableRow>();
                    groupedRows[rowGroup] = rows;
                }
                rows.Add(tableRow);
            }

            // fill tables
            foreach (var groupedRow in groupedRows)
            {
                var rowGroup = groupedRow.Key;

                var iTableRow = groupedRow.Value[0] as ITableRowForDataTable;
                if (iTableRow == null) continue;

                var table = _data.GetDataTableByName(iTableRow.TableName);
                if (table == null)
                {
                    if (_data.ShowUnknownValues)
                    {
                        // show unknown values
                        foreach (var tableRow in groupedRow.Value)
                            foreach (var cell in tableRow.Cells)
                            {
                                var localWalker = new DocumentWalker();
                                var tableCells = localWalker.TraverseBlockCollection<ITableCellValue>(cell.Blocks);
                                foreach (var cv in tableCells)
                                {
                                    var dv = cv as IPropertyValue;
                                    if (dv == null) continue;
                                    dv.Value = "[" + dv.PropertyName + "]";
                                    RememberAggregateValue(aggregateValues, cv.AggregateGroup, null);
                                }
                            }
                    }
                    else continue;
                }
                else
                {
                    var listNewRows = new List<TableRow>();
                    TableRow newTableRow;

                    // clone XAML rows
                    var clonedRows = new List<string>();
                    foreach (var row in rowGroup.Rows)
                    {
                        var reportTableRow = row as TableRowForDataTable;
                        if (reportTableRow == null) clonedRows.Add(null);
                        clonedRows.Add(XamlWriter.Save(reportTableRow));
                    }

                    for (var i = 0; i < table.Rows.Count; i++)
                    {
                        var dataRow = table.Rows[i];

                        for (var j = 0; j < rowGroup.Rows.Count; j++)
                        {
                            var row = rowGroup.Rows[j];

                            var reportTableRow = row as TableRowForDataTable;
                            if (reportTableRow == null)
                            {
                                // clone regular row
                                listNewRows.Add(XamlHelper.CloneTableRow(row));
                            }
                            else
                            {
                                // clone ReportTableRows
                                newTableRow = (TableRow)XamlHelper.LoadXamlFromString(clonedRows[j]);

                                foreach (var cell in newTableRow.Cells)
                                {
                                    var localWalker = new DocumentWalker();
                                    var newCells = localWalker.TraverseBlockCollection<ITableCellValue>(cell.Blocks);
                                    foreach (var cv in newCells)
                                    {
                                        var dv = cv as IPropertyValue;
                                        if (dv == null) continue;
                                        try
                                        {
                                            var obj = dataRow[dv.PropertyName];
                                            if (obj == DBNull.Value) obj = null;
                                            dv.Value = obj;

                                            RememberAggregateValue(aggregateValues, cv.AggregateGroup, obj);
                                        }
                                        catch
                                        {
                                            if (_data.ShowUnknownValues) dv.Value = "[" + dv.PropertyName + "]";
                                            else dv.Value = "";
                                            RememberAggregateValue(aggregateValues, cv.AggregateGroup, null);
                                        }
                                    }
                                }
                                listNewRows.Add(newTableRow);

                                // fire event
                                _report.FireEventDataRowBoundEventArgs(new DataRowBoundEventArgs(_report, dataRow)
                                {
                                    TableName = dataRow.Table.TableName,
                                    TableRow = newTableRow
                                });
                            }
                        }
                    }
                    rowGroup.Rows.Clear();
                    foreach (var row in listNewRows) rowGroup.Rows.Add(row);
                }
            }

            // fill aggregate values
            foreach (InlineAggregateValue av in blockAggregateValues)
            {
                if (string.IsNullOrEmpty(av.AggregateGroup)) continue;

                var aggregateGroups = av.AggregateGroup.Split(new[] { ',', ';', ' ' },
                    StringSplitOptions.RemoveEmptyEntries);

                foreach (var group in aggregateGroups)
                {
                    if (!aggregateValues.ContainsKey(group))
                    {
                        av.Text = av.EmptyValue;
                        break;
                    }
                }
                av.Text = av.ComputeAndFormat(aggregateValues);
            }
        }

        /// <summary>
        ///     Clones a visual block
        /// </summary>
        /// <param name="block">block to be cloned</param>
        /// <param name="pageNumber">current page number</param>
        /// <returns>cloned block</returns>
        /// <exception cref="InvalidProgramException">Error cloning XAML block</exception>
        private ContainerVisual CloneVisualBlock(Block block, int pageNumber)
        {
            var tmpDoc = new FlowDocument();
            tmpDoc.ColumnWidth = double.PositiveInfinity;
            tmpDoc.PageHeight = _report.PageHeight;
            tmpDoc.PageWidth = _report.PageWidth;
            tmpDoc.PagePadding = new Thickness(0);

            var xaml = XamlWriter.Save(block);
            var newBlock = XamlReader.Parse(xaml) as Block;
            if (newBlock == null) throw new InvalidProgramException("Error cloning XAML block");
            tmpDoc.Blocks.Add(newBlock);

            var walkerBlock = new DocumentWalker();
            var blockValues = new ArrayList();
            blockValues.AddRange(walkerBlock.Walk<IInlineContextValue>(tmpDoc));

            // fill context values
            FillContextValues(blockValues, pageNumber);

            var dp = ((IDocumentPaginatorSource)tmpDoc).DocumentPaginator.GetPage(0);
            return (ContainerVisual)dp.Visual;
        }

        protected virtual void FillContextValues(ArrayList list, int pageNumber)
        {
            // fill context values
            foreach (IInlineContextValue cv in list)
            {
                if (cv == null) continue;
                var reportContextValueType = ReportPaginatorStaticCache.GetReportContextValueTypeByName(cv.PropertyName);
                if (reportContextValueType == null)
                {
                    if (_data.ShowUnknownValues)
                        cv.Value = "<" + (cv.PropertyName ?? "NULL") + ">";
                    else cv.Value = "";
                }
                else
                {
                    switch (reportContextValueType.Value)
                    {
                        case ReportContextValueType.PageNumber:
                            cv.Value = pageNumber;
                            break;
                        case ReportContextValueType.PageCount:
                            cv.Value = _pageCount;
                            break;
                        case ReportContextValueType.ReportName:
                            cv.Value = _report.ReportName;
                            break;
                        case ReportContextValueType.ReportTitle:
                            cv.Value = _report.ReportTitle;
                            break;
                    }
                }
            }
        }

        /// <summary>
        ///     This is most important method, modifies the original
        /// </summary>
        /// <param name="pageNumber">page number</param>
        /// <returns></returns>
        public override DocumentPage GetPage(int pageNumber)
        {
            for (var i = 0; i < 2; i++) // do it twice because filling context values could change the page count
            {
                // compute page count
                if (pageNumber == 0)
                {
                    _paginator.ComputePageCount();
                    _pageCount = _paginator.PageCount;
                }

                // fill context values
                FillContextValues(_reportContextValues, pageNumber + 1);
            }

            var page = _paginator.GetPage(pageNumber);
            if (page == DocumentPage.Missing) return DocumentPage.Missing; // page missing

            _pageSize = page.Size;

            // add header block
            var newPage = new ContainerVisual();

            if (_blockPageHeader != null)
            {
                var v = CloneVisualBlock(_blockPageHeader, pageNumber + 1);
                v.Offset = new Vector(0, 0);
                newPage.Children.Add(v);
            }

            // TODO: process ReportContextValues

            // add content page
            var smallerPage = new ContainerVisual();
            smallerPage.Offset = new Vector(0, _report.PageHeaderHeight / 100d * _report.PageHeight);
            smallerPage.Children.Add(page.Visual);
            newPage.Children.Add(smallerPage);

            // add footer block
            if (_blockPageFooter != null)
            {
                var v = CloneVisualBlock(_blockPageFooter, pageNumber + 1);
                v.Offset = new Vector(0, _report.PageHeight - _report.PageFooterHeight / 100d * _report.PageHeight);
                newPage.Children.Add(v);
            }

            // create modified BleedBox
            var bleedBox = new Rect(page.BleedBox.Left, page.BleedBox.Top, page.BleedBox.Width,
                _report.PageHeight - (page.Size.Height - page.BleedBox.Size.Height));

            // create modified ContentBox
            var contentBox = new Rect(page.ContentBox.Left, page.ContentBox.Top, page.ContentBox.Width,
                _report.PageHeight - (page.Size.Height - page.ContentBox.Size.Height));

            var dp = new DocumentPage(newPage, new Size(_report.PageWidth, _report.PageHeight), bleedBox, contentBox);
            _report.FireEventGetPageCompleted(new GetPageCompletedEventArgs(page, pageNumber, null, false, null));
            return dp;
        }

        // ReSharper disable InconsistentNaming
        /// <summary>
        ///     Reference to a original flowdoc paginator
        /// </summary>
        protected DocumentPaginator _paginator;

        protected FlowDocument _flowDocument;
        protected ReportDocument _report;
        protected ReportData _data;
        protected Block _blockPageHeader;
        protected Block _blockPageFooter;
        protected ArrayList _reportContextValues;
        protected ReportPaginatorDynamicCache _dynamicCache;
        // ReSharper restore InconsistentNaming
    }
}