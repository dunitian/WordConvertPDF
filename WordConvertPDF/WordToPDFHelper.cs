using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.IO;

namespace WordConvertPDF
{
    public static class WordToPDFHelper
    {
        /// <summary>
        /// Word转换成PDF
        /// </summary>
        /// <param name="inputPath">载入路径</param>
        /// <param name="outputPath">保存路径</param>
        /// <param name="startPage">初始页码（默认为第一页[0]）</param>
        /// <param name="endPage">结束页码（默认为最后一页）</param>
        public static bool WordToPDF(string inputPath, string outputPath, int startPage = 0, int endPage = 0)
        {
            bool b = true;

            #region 初始化
            //初始化一个application
            Application wordApplication = new Application();
            //初始化一个document
            Document wordDocument = null;
            #endregion

            #region 参数设置~~我去累死宝宝了~~
            //word路径
            object wordPath = Path.GetFullPath(inputPath);

            //输出路径
            string pdfPath = Path.GetFullPath(outputPath);

            //导出格式为PDF
            WdExportFormat wdExportFormat = WdExportFormat.wdExportFormatPDF;

            //导出大文件
            WdExportOptimizeFor wdExportOptimizeFor = WdExportOptimizeFor.wdExportOptimizeForPrint;

            //导出整个文档
            WdExportRange wdExportRange = WdExportRange.wdExportAllDocument;

            //开始页码
            int startIndex = startPage;

            //结束页码
            int endIndex = endPage;

            //导出不带标记的文档（这个可以改）
            WdExportItem wdExportItem = WdExportItem.wdExportDocumentContent;

            //包含word属性
            bool includeDocProps = true;

            //导出书签
            WdExportCreateBookmarks paramCreateBookmarks = WdExportCreateBookmarks.wdExportCreateWordBookmarks;

            //默认值
            object paramMissing = Type.Missing;

            #endregion

            #region 转换
            try
            {
                //打开word
                wordDocument = wordApplication.Documents.Open(ref wordPath, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing);
                //转换成指定格式
                if (wordDocument != null)
                {
                    wordDocument.ExportAsFixedFormat(pdfPath, wdExportFormat, false, wdExportOptimizeFor, wdExportRange, startIndex, endIndex, wdExportItem, includeDocProps, true, paramCreateBookmarks, true, true, false, ref paramMissing);
                }
            }
            catch (Exception ex)
            {
                b = false;
            }
            finally
            {
                //关闭
                if (wordDocument != null)
                {
                    wordDocument.Close(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordDocument = null;
                }

                //退出
                if (wordApplication != null)
                {
                    wordApplication.Quit(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordApplication = null;
                }
            }

            return b;
            #endregion
        }

        /// <summary>
        /// Word转换成PDF（带日记）
        /// </summary>
        /// <param name="inputPath">载入路径</param>
        /// <param name="outputPath">保存路径</param>
        /// <param name="log">转换日记</param>
        /// <param name="startPage">初始页码（默认为第一页[0]）</param>
        /// <param name="endPage">结束页码（默认为最后一页）</param>
        public static void WordToPDFCreateLog(string inputPath, string outputPath, out string log, int startPage = 0, int endPage = 0)
        {
            log = "success";

            #region 初始化
            //初始化一个application
            Application wordApplication = new Application();
            //初始化一个document
            Document wordDocument = null;
            #endregion

            #region 参数设置~~我去累死宝宝了~~
            //word路径
            object wordPath = Path.GetFullPath(inputPath);

            //输出路径
            string pdfPath = Path.GetFullPath(outputPath);

            //导出格式为PDF
            WdExportFormat wdExportFormat = WdExportFormat.wdExportFormatPDF;

            //导出大文件
            WdExportOptimizeFor wdExportOptimizeFor = WdExportOptimizeFor.wdExportOptimizeForPrint;

            //导出整个文档
            WdExportRange wdExportRange = WdExportRange.wdExportAllDocument;

            //开始页码
            int startIndex = startPage;

            //结束页码
            int endIndex = endPage;

            //导出不带标记的文档（这个可以改）
            WdExportItem wdExportItem = WdExportItem.wdExportDocumentContent;

            //包含word属性
            bool includeDocProps = true;

            //导出书签
            WdExportCreateBookmarks paramCreateBookmarks = WdExportCreateBookmarks.wdExportCreateWordBookmarks;

            //默认值
            object paramMissing = Type.Missing;

            #endregion

            #region 转换
            try
            {
                //打开word
                wordDocument = wordApplication.Documents.Open(ref wordPath, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing, ref paramMissing);
                //转换成指定格式
                if (wordDocument != null)
                {
                    wordDocument.ExportAsFixedFormat(pdfPath, wdExportFormat, false, wdExportOptimizeFor, wdExportRange, startIndex, endIndex, wdExportItem, includeDocProps, true, paramCreateBookmarks, true, true, false, ref paramMissing);
                }
            }
            catch (Exception ex)
            {
                if (ex != null) { log = ex.ToString(); }
            }
            finally
            {
                //关闭
                if (wordDocument != null)
                {
                    wordDocument.Close(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordDocument = null;
                }

                //退出
                if (wordApplication != null)
                {
                    wordApplication.Quit(ref paramMissing, ref paramMissing, ref paramMissing);
                    wordApplication = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            #endregion
        }
    }
}
