
using NPOI.OpenXml4Net.Util;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            string wordPath = "C:\\Users\\ZJW\\Desktop\\test.docx";
            FileStream fs = new FileStream(wordPath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            XWPFDocument docx = new XWPFDocument(fs);//打开07（.docx）以上的版本的文档
            List<XWPFSDT> contentControls = new List<XWPFSDT>();

            //遍历word中内容控件
            foreach (var element in docx.BodyElements)
            {
                ///段落只包含 内容控件
                if (element.ElementType == BodyElementType.CONTENTCONTROL)
                {
                    XWPFSDT sdt = (XWPFSDT)element;

                
                    contentControls.Add(sdt);
                    int i = docx.BodyElements.IndexOf(element);
                }
                ///段落包含文字和内容控件
                if (element.ElementType == BodyElementType.PARAGRAPH)
                {
                    var para = element as XWPFParagraph;

                    foreach (var run in para.IRuns) {

                        if (run is XWPFSDT sdt)
                        {
                            if (sdt.Content is XWPFSDTContent abc)
                            {
                                if (abc._sdtRun.Items[0] is CT_R cT_R)
                                {
                                    if (cT_R.Items[0] is CT_Text cT_Text)
                                    {
                                       // cT_Text.Value = "test";
                                    }
                                }
                            }
                        }
                    }
                }
                ///表格包括内容控件
                if (element.ElementType == BodyElementType.TABLE) {
                    var wPFTable = element as XWPFTable;

                    foreach (var row in wPFTable.Rows) {

                        foreach (var cell in row.GetTableCells()) {
                            var paras = cell.Paragraphs;

                            foreach (var para in paras)
                            {
                                foreach (var run in para.IRuns)
                                {
                                    if (run is XWPFSDT sdt)
                                    {
                                        if (sdt.Content is XWPFSDTContent abc)
                                        {
                                            if (abc._sdtRun.Items[0] is CT_R cT_R)
                                            {
                                                if (cT_R.Items[0] is CT_Text cT_Text)
                                                {
                                                    // cT_Text.Value = "test";
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                        }
                    }


                }

            }

            foreach (var para in docx.Paragraphs)
            {
                if (para.ElementType == BodyElementType.CONTENTCONTROL)
                {

                }
                var runs = para.Runs;
            }



            FileStream output = new FileStream(wordPath, FileMode.Create);
            docx.Write(output);
            fs.Close();
            fs.Dispose();
            output.Close();
            output.Dispose();
        }
    }
}
