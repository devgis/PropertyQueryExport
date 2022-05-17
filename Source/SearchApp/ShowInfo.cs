using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.OpenXmlFormats.Dml.WordProcessing;
using Microsoft.Office.Interop.Word;
using MSWord = Microsoft.Office.Interop.Word;

namespace SearchApp
{
    public partial class ShowInfo : Form
    {
        private TuItem tuitem = null;
        Object Nothing = System.Reflection.Missing.Value;
        private string picPath = System.Windows.Forms.Application.StartupPath+"\\Pics\\"; //图片路径
        public ShowInfo(TuItem tuItem )
        {
            InitializeComponent();
            this.tuitem = tuItem;
        }

        private void ShowInfo_Load(object sender, EventArgs e)
        {
            tbCEC.Text = tuitem.CEC;
            tbPH.Text = string.Empty;
            tb编号.Text = tuitem.编号;
            tb地点.Text = tuitem.地点;
            tb地下水深度.Text = tuitem.地下水深度;
            tb发生分类名.Text = string.Empty;
            tb海拔.Text = tuitem.海拔;
            tb经度.Text = tuitem.X;
            tb全钾.Text = tuitem.全钾;
            tb全磷.Text = tuitem.全磷;
            tb全铁.Text = tuitem.全铁;
            tb容量.Text = string.Empty;
            tb水质.Text = tuitem.水质;
            tb速磷.Text = tuitem.速效磷;
            tb速效氮.Text = tuitem.速效氮;
            tb速效钾.Text = tuitem.速效钾;
            tb速效磷.Text = tuitem.速效磷;
            tb土地利用类型.Text = tuitem.土地利用类;
            tb土纲.Text = tuitem.土纲;
            tb土类.Text = tuitem.土类;
            tb土壤类型.Text = tuitem.土壤类型;
            tb土系.Text = tuitem.土系;
            tb土族.Text = tuitem.土族;
            tb土族2.Text = tuitem.土族;
            tb维度.Text = tuitem.Y;
            tb亚纲.Text = tuitem.亚纲;
            tb亚类.Text = tuitem.亚类;
            tb游离态氧化铁.Text = tuitem.游离态氧化;
            tb有机碳.Text = tuitem.有机碳;
            tb有机质.Text = tuitem.有机质;
            tb有效土层厚度.Text = tuitem.有效土层厚;
            tb粘土矿物含量.Text = tuitem.黏土矿物类;
            tb植被类型.Text = tuitem.植被类型;
            tb质地.Text = tuitem.质地;

            //加载图片
            try
            {
                pictureBox1.Image = Image.FromFile(picPath + tuitem.编号 + "\\" + tuitem.编号.Replace("-","") + "01.jpg");
            }
            catch
            { }
            try
            {
                pictureBox2.Image = Image.FromFile(picPath + tuitem.编号 + "\\" + tuitem.编号.Replace("-", "") + "02.jpg");
            }
            catch
            { }
            try
            {
                pictureBox3.Image = Image.FromFile(picPath + tuitem.编号 + "\\" + tuitem.编号.Replace("-", "") + "03.jpg");
            }
            catch
            { }
            
            
            
        }

        private void btExportWord_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Word文件|*.docx";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                //开始写入word
                
                //Directory.CreateDirectory("C:/CNSI"); //创建文件所在目录
                
                //创建Word文档
                Microsoft.Office.Interop.Word.Application WordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
                Microsoft.Office.Interop.Word.Document WordDoc = WordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);

                WordDoc.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument;//激活页面内容的编辑

                //添加页眉
                WordApp.ActiveWindow.View.Type = WdViewType.wdOutlineView;
                WordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekPrimaryHeader;
                WordApp.ActiveWindow.ActivePane.Selection.InsertAfter("[页眉内容]");
                WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;//设置右对齐
                WordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;//跳出页眉设置

                WordApp.Selection.ParagraphFormat.LineSpacing = 15f;//设置文档的行间距

                //移动焦点并换行
                object count = 14;
                object WdLine = Microsoft.Office.Interop.Word.WdUnits.wdLine;//换一行;
                WordApp.Selection.MoveDown(ref WdLine, ref count, ref Nothing);//移动焦点
                WordApp.Selection.TypeParagraph();//插入段落

                

                AddTitle(WordApp,WordDoc,"基本信息");
                AddSubTitle(WordApp, WordDoc, "一，土系名称");
                //土纲	亚纲	土类	亚类	土族	地点	地点

                AddRow(WordApp, WordDoc, string.Format("土纲:{0}", tuitem.土纲));
                AddRow(WordApp, WordDoc, string.Format("亚纲:{0}", tuitem.亚纲));
                AddRow(WordApp, WordDoc, string.Format("土类:{0}", tuitem.土类));
                AddRow(WordApp, WordDoc, string.Format("亚类:{0}", tuitem.亚类));
                AddRow(WordApp, WordDoc, string.Format("土族:{0}", tuitem.土族));
                AddRow(WordApp, WordDoc, string.Format("地点:{0}", tuitem.地点));
                //AddRow(WordApp, WordDoc, string.Format("地点:{0}", tuitem.CEC));


                AddSubTitle(WordApp, WordDoc, "二，物理性质");

                //质地	黏土矿物类型	质地	粒径
                AddRow(WordApp, WordDoc, string.Format("质地:{0}", tuitem.质地));
                AddRow(WordApp, WordDoc, string.Format("黏土矿物类型:{0}", tuitem.黏土矿物类));
                AddRow(WordApp, WordDoc, string.Format("质地:{0}", tuitem.质地));
                AddRow(WordApp, WordDoc, string.Format("粒径:{0}", tuitem.粒径));


                AddSubTitle(WordApp, WordDoc, "三，化学性质");
                //有机质	有机碳	全氮	速效氮	全磷	速效磷	全钾	速效钾	CEC	全铁	游离态氧化铁	全铁
                AddRow(WordApp, WordDoc, string.Format("有机质:{0}", tuitem.有机质));
                AddRow(WordApp, WordDoc, string.Format("有机碳:{0}", tuitem.有机碳));
                AddRow(WordApp, WordDoc, string.Format("全氮:{0}", tuitem.全氮));
                AddRow(WordApp, WordDoc, string.Format("速效氮:{0}", tuitem.速效氮));
                AddRow(WordApp, WordDoc, string.Format("全磷:{0}", tuitem.全磷));
                AddRow(WordApp, WordDoc, string.Format("速效磷:{0}", tuitem.速效磷));
                AddRow(WordApp, WordDoc, string.Format("全钾:{0}", tuitem.全钾));
                AddRow(WordApp, WordDoc, string.Format("速效钾:{0}", tuitem.速效钾));
                AddRow(WordApp, WordDoc, string.Format("CEC:{0}", tuitem.CEC));
                AddRow(WordApp, WordDoc, string.Format("全铁:{0}", tuitem.全铁));
                AddRow(WordApp, WordDoc, string.Format("游离态氧化铁:{0}", tuitem.游离态氧化));
                AddRow(WordApp, WordDoc, string.Format("全铁:{0}", tuitem.全铁1));


                AddSubTitle(WordApp, WordDoc, "四，土系背景环境");
                //X	Y	市	地点	海拔（m）	土壤类型	土地利用类型	植被类型	有效土层厚度（cm）	地下水深度（m）	水质
                AddRow(WordApp, WordDoc, string.Format("经度:{0}", tuitem.X));
                AddRow(WordApp, WordDoc, string.Format("纬度:{0}", tuitem.Y));
                AddRow(WordApp, WordDoc, string.Format("地点:{0}", tuitem.地点));
                AddRow(WordApp, WordDoc, string.Format("海拔（m）:{0}", tuitem.海拔));
                AddRow(WordApp, WordDoc, string.Format("土壤类型:{0}", tuitem.土壤类型));
                AddRow(WordApp, WordDoc, string.Format("土地利用类型:{0}", tuitem.土地利用类));
                AddRow(WordApp, WordDoc, string.Format("植被类型:{0}", tuitem.植被类型));
                AddRow(WordApp, WordDoc, string.Format("有效土层厚度（cm）:{0}", tuitem.有效土层厚));
                AddRow(WordApp, WordDoc, string.Format("地下水深度（m）:{0}", tuitem.地下水深度));
                AddRow(WordApp, WordDoc, string.Format("水质:{0}", tuitem.水质));

                AddSubTitle(WordApp, WordDoc, "五，土系景观照");

                /*
                tbPH.Text = string.Empty;
                tb编号.Text = tuitem.编号;
                tb地点.Text = tuitem.地点;
                tb地下水深度.Text = tuitem.地下水深度;
                tb发生分类名.Text = string.Empty;
                tb海拔.Text = tuitem.海拔;
                tb经度.Text = tuitem.X;
                tb全钾.Text = tuitem.全钾;
                tb全磷.Text = tuitem.全磷;
                tb全铁.Text = tuitem.全铁;
                tb容量.Text = string.Empty;
                tb水质.Text = tuitem.水质;
                tb速磷.Text = tuitem.速效磷;
                tb速效氮.Text = tuitem.速效氮;
                tb速效钾.Text = tuitem.速效钾;
                tb速效磷.Text = tuitem.速效磷;
                tb土地利用类型.Text = tuitem.土地利用类;
                tb土纲.Text = tuitem.土纲;
                tb土类.Text = tuitem.土类;
                tb土壤类型.Text = tuitem.土壤类型;
                tb土系.Text = tuitem.土系;
                tb土族.Text = tuitem.土族;
                tb土族2.Text = tuitem.土族;
                tb维度.Text = tuitem.Y;
                tb亚纲.Text = tuitem.亚纲;
                tb亚类.Text = tuitem.亚类;
                tb游离态氧化铁.Text = tuitem.游离态氧化;
                tb有机碳.Text = tuitem.有机碳;
                tb有机质.Text = tuitem.有机质;
                tb有效土层厚度.Text = tuitem.有效土层厚;
                tb粘土矿物含量.Text = tuitem.黏土矿物类;
                tb植被类型.Text = tuitem.植被类型;
                tb质地.Text = tuitem.质地;
                 * */

                AddRow(WordApp, WordDoc, "1,土系剖面照");
                AddPic(WordApp, WordDoc, picPath + tuitem.编号 + "\\" + tuitem.编号.Replace("-", "") + "01.jpg", "土系剖面照");
                AddRow(WordApp, WordDoc, "2,地貌景观照");
                AddPic(WordApp, WordDoc, picPath + tuitem.编号 + "\\" + tuitem.编号.Replace("-", "") + "02.jpg", "地貌景观照");
                AddRow(WordApp, WordDoc, "3,地面覆盖照");
                AddPic(WordApp, WordDoc, picPath + tuitem.编号 + "\\" + tuitem.编号.Replace("-", "") + "03.jpg", "地面覆盖照");

                object filename=sfd.FileName;
                WordDoc.SaveAs(ref filename, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
                WordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
                WordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
                MessageBox.Show("导出成功！");
            }
        }

        private void AddTitle(Microsoft.Office.Interop.Word.Application oWordDoc, Microsoft.Office.Interop.Word.Document oWordApp, string title)
        {
            oWordApp.ActiveWindow.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            oWordApp.ActiveWindow.Selection.Font.Name = "宋体";
            oWordApp.ActiveWindow.Selection.Font.Size = 21f;
            oWordApp.ActiveWindow.Selection.Font.Scaling = 100;
            oWordApp.ActiveWindow.Selection.Text = title;
            oWordApp.ActiveWindow.Selection.TypeText(title);
            oWordApp.ActiveWindow.Selection.TypeParagraph();//另起一段
        }

        private void AddSubTitle(Microsoft.Office.Interop.Word.Application oWordDoc, Microsoft.Office.Interop.Word.Document oWordApp, string title)
        {
            oWordApp.ActiveWindow.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oWordApp.ActiveWindow.Selection.Font.Name = "宋体";
            oWordApp.ActiveWindow.Selection.Font.Size = 18f;
            oWordApp.ActiveWindow.Selection.Font.Scaling = 100;
            oWordApp.ActiveWindow.Selection.Text = title;
            oWordApp.ActiveWindow.Selection.TypeText(title);
            oWordApp.ActiveWindow.Selection.TypeParagraph();//另起一段
        }
        private void AddRow(Microsoft.Office.Interop.Word.Application oWordDoc, Microsoft.Office.Interop.Word.Document oWordApp, string content)
        {
            oWordApp.ActiveWindow.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oWordApp.ActiveWindow.Selection.Font.Name = "宋体";
            oWordApp.ActiveWindow.Selection.Font.Size = 15f;
            oWordApp.ActiveWindow.Selection.Font.Scaling = 100;
            oWordApp.ActiveWindow.Selection.Text = content;
            oWordApp.ActiveWindow.Selection.TypeText(content);
            oWordApp.ActiveWindow.Selection.TypeParagraph();//另起一段
        }

        private void AddPic(Microsoft.Office.Interop.Word.Application wordApp, Microsoft.Office.Interop.Word.Document wordDoc, string filePath, string title)
        {
            //定义要向文档中插入图片的位置
            object range = wordDoc.Paragraphs.Last.Range;
            //定义该图片是否为外部链接
            object linkToFile = false;//默认
            //定义插入的图片是否随word一起保存
            object saveWithDocument = true;
            //向word中写入图片
            wordDoc.InlineShapes.AddPicture(filePath, ref Nothing, ref Nothing, ref Nothing);

            object unite = Microsoft.Office.Interop.Word.WdUnits.wdStory;
            wordApp.Selection.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;//居中显示图片
            wordDoc.InlineShapes[1].Height = 130;
            wordDoc.InlineShapes[1].Width = 200;
            wordDoc.Content.InsertAfter("\n");
            wordApp.Selection.EndKey(ref unite, ref Nothing);
            wordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            wordApp.Selection.Font.Size = 10;//字体大小
            wordApp.Selection.TypeText(title+"\n");
            //object LinkToFile = false;
            //object SaveWithDocument = true;
            //object Anchor = WordDoc.Application.Selection.Range;
            
            //WordDoc.Application.ActiveDocument.InlineShapes.AddPicture(filePath, ref LinkToFile, ref SaveWithDocument, ref Anchor);
            //WordDoc.Application.ActiveDocument.InlineShapes[1].Width = 100f;//图片宽度
            //WordDoc.Application.ActiveDocument.InlineShapes[1].Height = 100f;//图片高度
            ////将图片设置为四周环绕型
            ////Microsoft.Office.Interop.Word.Shape s = WordDoc.Application.ActiveDocument.InlineShapes[1].ConvertToShape();
            ////s.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapInline;
            //WordApp.ActiveWindow.Selection.TypeParagraph();//另起一段
        }
    }
}
