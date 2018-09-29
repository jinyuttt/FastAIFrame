using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace FastStream.office
{
  public  class WordTool
    {
        public void WordCreate(string filePath,string content)
        {
            XWPFDocument MyDoc = new XWPFDocument();
          
            CT_SectPr m_SectPr =new CT_SectPr();       //实例一个尺寸类的实例
            m_SectPr.pgSz.w = 16838;        //设置宽度（这里是一个ulong类型）
            m_SectPr.pgSz.h = 11906;        //设置高度（这里是一个ulong类型）
            MyDoc.Document.body.sectPr = m_SectPr;          //设置页面的尺;
           var data= MyDoc.CreateParagraph();
            // 向新文档中添加段落

            data.Alignment = ParagraphAlignment.CENTER;
            // 向该段落中添加文字
            XWPFRun titleRun = data.CreateRun();
            titleRun.SetText(content);
           //处理doc后，生成新的文件，写入doc ，生成word完成。 
            FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write);
            MyDoc.Write(file);
            file.Close();

        }
        public void TemplateContent(string templete,string fileDir,Dictionary<string,string> content)
        {
           // XWPFDocument document = null;
            string filePath=Path.Combine(fileDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".docx");
            try
            {
               
                using (FileStream stream = File.OpenRead(templete))
                {
                    XWPFDocument doc = new XWPFDocument(stream);
                    //遍历段落
                    foreach (var para in doc.Paragraphs)
                    {
                        ReplaceKey(para,content);
                    }
                    //遍历表格
                    var tables = doc.Tables;
                    foreach (var table in tables)
                    {
                        foreach (var row in table.Rows)
                        {
                            foreach (var cell in row.GetTableCells())
                            {
                                foreach (var para in cell.Paragraphs)
                                {
                                    ReplaceKey(para,content);
                                }
                            }
                        }
                    }
                    using (MemoryStream ms = new MemoryStream())
                    {

                        doc.Write(ms);
                        using (FileStream fsWrite = new FileStream(filePath, FileMode.OpenOrCreate))
                        {
                            fsWrite.Write(ms.ToArray(), 0, ms.ToArray().Length);
                        };
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("文件{0}打开失败，错误：{1}", new string[] { templete, ex.ToString() }));
            }
        }
        private void ReplaceKey(XWPFParagraph para,Dictionary<string,string> editData)
        {

           
           foreach(KeyValuePair<string,string> kv in editData)
            {
                para.ReplaceText(kv.Key, kv.Value);
            }
        }
        private void ReplaceKey(XWPFParagraph para)
        {

            string text = para.ParagraphText;//段落文本
            var runs = para.Runs;
            string styleid = para.Style;
            for (int i = 0; i < runs.Count; i++)
            {
                var run = runs[i];
                text = run.ToString();
                runs[i].SetText(text + 2, 0);
            }
        }
    }
}
