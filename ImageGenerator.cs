using SharpDX;
using SharpDX.Direct2D1;
using SharpDX.DirectWrite;
using SharpDX.WIC;
using System.Collections.Generic;
using System.Text;
using SharpDX.DXGI;
using DWriteFactory = SharpDX.DirectWrite.Factory;
using System;
using SharpDX.IO;
using System.IO;
using Nova.MediaItem;
using System.Threading;
using System.Diagnostics;

namespace Nova.Rss
{
    public class Character
    {
        public string content;
        public float width;
        public float height;
        public bool isTitle;
        public System.Drawing.Font font;
        public System.Drawing.Color color;

        public Character(string content, System.Drawing.Font font, System.Drawing.Color color)
        {
            this.content = content;
            this.font = font;
            this.width = font.Size * content.Length;
            this.height = font.Size * 4 / 3;
            this.color = color;
        }
        public Character(char content, System.Drawing.Font font, System.Drawing.Color color)
        {
            this.content = content.ToString();
            this.font = font;
            this.width = font.Size;
            this.height = font.Size * 4 / 3;
            this.color = color;
        }
    }

    /// <summary>
    /// 格式化段落
    /// </summary>
    public class Paragraph
    {
        public System.Drawing.Font font;
        public RssBodyType type;
        public List<Character> list = new List<Character>();

        public Paragraph(List<Character> list, RssBodyType type)
        {
            this.list = list;
            this.type = type;
            this.font = list[0].font;
        }
    }

    /// <summary>
    /// rss 结构类型
    /// </summary>
    public enum RssBodyType
    {
        Title,
        Time,
        Body
    }

    /// <summary>
    /// 字符行
    /// </summary>
    public class Line
    {
        public List<Block> content = new List<Block>();

        public void Add(Block block)
        {
            this.content.Add(block);
        }
    }

    /// <summary>
    /// 字符块
    /// </summary>
    public class Block
    {
        public System.Drawing.Font font;
        public string content;
        public float left;
        public float top;
        public float right;
        public float bottom;
        public float width;
        public float height;
        public RssBodyType type;

        public Block(float lf, float tp, float rh, float bm, RssBodyType type = RssBodyType.Title)
        {
            this.right = rh;
            this.top = tp;
            this.left = lf;
            this.bottom = bm;
            this.type = type;
            this.width = this.right - this.left;
            this.height = this.bottom - this.top;
        }

        public void Add(Character character, float characterWidth)
        {
            this.content += character.content;
            this.right += characterWidth;
            if (this.bottom - this.top < character.height)
            {
                this.bottom = this.top + character.height;
            }
            this.width = this.right - this.left;
            this.height = this.bottom - this.top;
        }
    }

    /// <summary>
    /// 字符页
    /// </summary>
    public class Page
    {
        public List<Line> lines = new List<Line>();

        public void Add(Line line)
        {
            this.lines.Add(line);
        }
    }

    public class ImageGenerator
    {
        //图像的宽度
        public int pageWidth;
        //图像的高度
        public int pageHeight;
        //图片的保存路径
        public string path;
        //图片背景色
        public System.Drawing.Color backgroundColor;

        //全部图像生成完成
        public delegate void GenerateCompleteDelegate(string path);
        public event GenerateCompleteDelegate GenerateCompleteEvent;

        //单一图像生成完成
        public delegate void SingleGenerateCompleteDelegate(string fileNmae);
        public event SingleGenerateCompleteDelegate SingleGenerateCompleteEvent;

        public ImageGenerator(int width, int height, string path, System.Drawing.Color backgroundColor)
        {
            this.pageWidth = width;
            this.pageHeight = height;
            this.path = path;
            this.backgroundColor = backgroundColor;
        }

        /// <summary>
        /// 从rss解析到每一页中每一行的数据
        /// </summary>
        /// <param name="channel"></param>
        /// <returns></returns>
        private List<Page> GetPageListFromRss(RssInfo rssInfo, RssMedia rssMedia)
        {
            List<Paragraph> paragraphs = new List<Paragraph>();

            foreach (RssItemInfo item in rssInfo.RssItem)
            {
                if (item != null)
                {
                    if (rssMedia.IsShowRssTitle)
                    {
                        paragraphs.AddRange(GetCharacter(GetCleanString(item.Title),
                                          rssMedia.RssTitleProp.TextFont,
                                          rssMedia.RssTitleProp.TextColor,
                                          RssBodyType.Title));
                    }
                    if (rssMedia.IsShowRssPublishTime)
                    {
                        paragraphs.AddRange(GetCharacter(GetCleanString(item.PubDate),
                                            rssMedia.RssPublishTimeProp.TextFont,
                                            rssMedia.RssPublishTimeProp.TextColor,
                                            RssBodyType.Time));
                    }
                    if (rssMedia.IsShowRssBody)
                    {
                        if (item.Description != null)
                        {
                            foreach (var v in item.Description)
                            {
                                paragraphs.AddRange(GetCharacter(GetCleanString(v.content),
                                                                 rssMedia.RssBodyProp.TextFont,
                                                                 rssMedia.RssPublishTimeProp.TextColor,
                                                                 RssBodyType.Body));
                            }
                        }
                    }
                }
            }
            List<Page> pageList = GetPageList(paragraphs, pageWidth, pageHeight, rssMedia);
            paragraphs.Clear();
            paragraphs = null;
            return pageList;
        }

        private string GetCleanString(string rawString)
        {
            string cleanData = "";
            if (rawString != null)
            {
                cleanData = rawString.Replace("”", "\"")
                                     .Replace("“", "\"")
                                     .Replace("’", "\'")
                                     .Replace("‘", "\'")
                                     .Replace("（", "(")
                                     .Replace("）", ")");
            }
            return cleanData;
        }

        //从Rss中取得格式化字符段落
        private List<Paragraph> GetCharacter(string text, System.Drawing.Font font, System.Drawing.Color color, RssBodyType type)
        {
            List<Paragraph> paragraphs = new List<Paragraph>();
            List<Character> characters = new List<Character>();

            if (string.IsNullOrEmpty(text))
            {
                return null;
            }

            char[] cr = text.ToCharArray();
            StringBuilder strbuilder = new StringBuilder();

            foreach (char c in cr)
            {
                //拼接成英文单词，防止单词越行                    
                if ((c >= 'a' && c <= 'z') || ((c >= 'A' && c <= 'Z')))
                {
                    strbuilder.Append(c);
                }
                else
                {
                    if (strbuilder.Length != 0)
                    {
                        characters.Add(new Character(strbuilder.ToString(), font, color));
                        strbuilder = new StringBuilder();
                    }
                    characters.Add(new Character(c, font, color));
                }
            }

            if (strbuilder.Length != 0)
            {
                characters.Add(new Character(strbuilder.ToString(), font, color));
            }

            Paragraph paragraph = new Paragraph(characters, type);
            paragraph.font = font;
            paragraphs.Add(paragraph);

            return paragraphs;
        }

        /// <summary>
        /// 由字符段落获得图像页面
        /// </summary>
        /// <param name="paragraphList"></param>
        /// <param name="pageWidth"></param>
        /// <param name="pageHeight"></param>
        /// <returns></returns>
        private List<Page> GetPageList(List<Paragraph> paragraphList, int pageWidth, int pageHeight, RssMedia rssMedia)
        {
            List<Page> pageList = new List<Page>();
            Page page = new Page();
            Line line = new Line();
            Block block = new Block(0, 0, 0, 0);

            if (rssMedia.DisplayType == RssMedia.RssDisplayType.ScrollLeftToRight || rssMedia.DisplayType == RssMedia.RssDisplayType.ScrollRightToLeft)
            {
                foreach (Paragraph paragraph in paragraphList)
                {
                    if (block.content == null)
                    {
                        line = new Line();
                        block = new Block(0, pageHeight / 2 - paragraph.font.Height, 0, pageHeight / 2);
                        block.type = paragraph.type;
                        block.font = paragraph.font;
                    }
                    else
                    {
                        line.Add(block);
                        block = new Block(block.right, pageHeight / 2 - paragraph.font.Height, block.right, pageHeight / 2);
                        block.type = paragraph.type;
                        block.font = paragraph.font;
                    }

                    float characterWidth = 0.0f;

                    foreach (Character character in paragraph.list)
                    {
                        characterWidth = GetCharacterWidth(character);
                        //该行还能容下该字符
                        if (block.right + characterWidth < pageWidth)
                        {
                            block.Add(character, characterWidth);
                        }
                        //该行容不下该字符
                        else
                        {
                            if (block.content != null)
                            {
                                line.Add(block);
                                page.Add(line);
                                pageList.Add(page);
                                page = new Page();
                                line = new Line();
                                block = new Block(0, pageHeight / 2 - paragraph.font.Height, 0, pageHeight / 2, paragraph.type);
                                block.font = character.font;

                                block.Add(character, characterWidth);
                            }
                        }
                    }
                }
                line.Add(block);
                page.Add(line);
                pageList.Add(page);
            }
            else
            {
                foreach (Paragraph paragraph in paragraphList)
                {
                    if (block.content != null)
                    {
                        line.Add(block);
                        page.Add(line);
                    }

                    if (block.bottom + paragraph.font.Size * 4 / 3 < pageHeight)
                    {
                        block = new Block(0, block.bottom, 0, block.bottom + paragraph.font.Size * 4 / 3, paragraph.type);
                        line = new Line();
                    }
                    else
                    {
                        line.Add(block);
                        pageList.Add(page);
                        block = new Block(0, 0, 0, 0, paragraph.type);
                        line = new Line();
                        page = new Page();
                    }

                    block.font = paragraph.font;
                    float characterWidth = 0.0f;

                    foreach (Character character in paragraph.list)
                    {
                        characterWidth = GetCharacterWidth(character);
                        //该行还能容下该字符
                        if (block.right + characterWidth < pageWidth)
                        {
                            block.font = character.font;
                            block.Add(character, characterWidth);
                        }
                        //该行容不下该字符
                        else
                        {
                            //该页能够容下该行
                            if (block.bottom + character.height < pageHeight)
                            {
                                line.Add(block);
                                page.Add(line);
                                block = new Block(0, block.bottom, 0, block.bottom + character.height, paragraph.type);
                                block.font = character.font;
                                block.Add(character, characterWidth);
                                line = new Line();
                            }
                            //该页容不下该行
                            else
                            {
                                line.Add(block);
                                page.Add(line);
                                pageList.Add(page);
                                page = new Page();
                                line = new Line();
                                block = new Block(0, 0, 0, 0, paragraph.type);
                                block.font = character.font;
                                block.Add(character, characterWidth);
                            }
                        }
                    }
                }
                line.Add(block);
                page.Add(line);
                pageList.Add(page);
            }
            return pageList;
        }

        /// <summary>
        /// 取得字符对应的宽度
        /// </summary>
        /// <param name="character"></param>
        /// <returns></returns>
        private float GetCharacterWidth(Character character)
        {
            DWriteFactory dwFactory = new DWriteFactory(SharpDX.DirectWrite.FactoryType.Shared);
            Vector2 offset = new Vector2(0.0f, 0.0f);
            TextFormat textFormat = new TextFormat(dwFactory,
                                                   character.font.FontFamily.Name,
                                                   character.font.Bold ? SharpDX.DirectWrite.FontWeight.Bold : SharpDX.DirectWrite.FontWeight.Regular,
                                                   character.font.Italic ? SharpDX.DirectWrite.FontStyle.Italic : SharpDX.DirectWrite.FontStyle.Normal,
                                                   character.font.Size);
            TextLayout textLayout = new TextLayout(dwFactory,
                                                   "好" + character.content,
                                                   textFormat,
                                                   this.pageWidth,
                                                   this.pageHeight);
            textLayout.SetUnderline(character.font.Underline, new TextRange(0, character.content.Length));
            textLayout.SetFontStyle(character.font.Italic ? FontStyle.Italic : FontStyle.Normal, new TextRange(0, character.content.Length));
            textLayout.SetFontWeight(character.font.Bold ? FontWeight.Bold : FontWeight.Normal, new TextRange(0, character.content.Length));
            ClusterMetrics[] clusterMetrics = textLayout.GetClusterMetrics();

            float result = 0.0f;
            float prefixCharWidth = clusterMetrics[0].Width;
            foreach (ClusterMetrics v in clusterMetrics)
            {
                result += v.Width;
            }

            dwFactory.Dispose();
            textFormat.Dispose();
            textLayout.Dispose();
            dwFactory = null;
            textFormat = null;
            textLayout = null;

            return result - prefixCharWidth;
        }

        /// <summary>
        /// 从rss生成图像
        /// </summary>
        /// <param name="channel">从rss源获得的数据</param>
        /// <param name="path">保存图像的路径</param>
        private void GetImageFromRss(object obj)
        {
            ImageObj image = (ImageObj)obj;
            RssInfo rssInfo = image.rssInfo;
            RssMedia rssMedia = image.rssMedia;
            string fileName = "";

            if (rssInfo == null || rssMedia == null)
            {
                return;
            }

            List<Page> pages = GetPageListFromRss(rssInfo, rssMedia);
            var wicFactory = new ImagingFactory();
            var d2dFactory = new SharpDX.Direct2D1.Factory();
            var dwFactory = new SharpDX.DirectWrite.Factory();
            var renderTargetProperties = new RenderTargetProperties(RenderTargetType.Default,
                                                                    new SharpDX.Direct2D1.PixelFormat(Format.Unknown, SharpDX.Direct2D1.AlphaMode.Unknown),
                                                                    0,
                                                                    0,
                                                                    RenderTargetUsage.None, FeatureLevel.Level_DEFAULT);

            var wicBitmap = new SharpDX.WIC.Bitmap(wicFactory,
                                               pageWidth,
                                               pageHeight,
                                               SharpDX.WIC.PixelFormat.Format32bppBGR,
                                               BitmapCreateCacheOption.CacheOnLoad);
            var d2dRenderTarget = new WicRenderTarget(d2dFactory,
                                                      wicBitmap,
                                                      renderTargetProperties);

            d2dRenderTarget.AntialiasMode = AntialiasMode.PerPrimitive;
            var solidColorBrush = new SolidColorBrush(d2dRenderTarget,
                                                      new SharpDX.Color(this.backgroundColor.R,
                                                                        this.backgroundColor.G,
                                                                        this.backgroundColor.B,
                                                                        this.backgroundColor.A));
            var textBodyBrush = new SolidColorBrush(d2dRenderTarget,
                                                    new SharpDX.Color(rssMedia.RssBodyProp.
                                                                      TextColor.R,
                                                                      rssMedia.RssBodyProp.TextColor.G,
                                                                      rssMedia.RssBodyProp.TextColor.B,
                                                                      rssMedia.RssBodyProp.TextColor.A));
            var titleColorBrush = new SolidColorBrush(d2dRenderTarget,
                                                      new SharpDX.Color(rssMedia.RssTitleProp.TextColor.R,
                                                                        rssMedia.RssTitleProp.TextColor.G,
                                                                        rssMedia.RssTitleProp.TextColor.B,
                                                                        rssMedia.RssTitleProp.TextColor.A));
            var publishDateColorBrush = new SolidColorBrush(d2dRenderTarget,
                                                            new SharpDX.Color(rssMedia.RssPublishTimeProp.TextColor.R,
                                                                              rssMedia.RssPublishTimeProp.TextColor.G,
                                                                              rssMedia.RssPublishTimeProp.TextColor.B,
                                                                              rssMedia.RssPublishTimeProp.TextColor.A));
            TextLayout textLayout;
            try
            {
                int count = 0;

                foreach (Page page in pages)
                {
                    d2dRenderTarget.BeginDraw();
                    d2dRenderTarget.Clear(new SharpDX.Color(this.backgroundColor.R,
                                                                            this.backgroundColor.G,
                                                                            this.backgroundColor.B,
                                                                            this.backgroundColor.A));

                    foreach (Line line in page.lines)
                    {
                        foreach (Block block in line.content)
                        {
                            TextFormat textFormat = new TextFormat(dwFactory,
                                                               block.font.FontFamily.Name,
                                                               block.font.Size);
                            textLayout = new TextLayout(dwFactory,
                                                    block.content,
                                                    textFormat,
                                                    block.width,
                                                    block.height);
                            switch (block.type)
                            {
                                case RssBodyType.Title:
                                    textLayout.SetUnderline(rssMedia.RssTitleProp.TextFont.Underline,
                                                            new TextRange(0, block.content.Length));
                                    textLayout.SetFontStyle(rssMedia.RssTitleProp.TextFont.Italic ? FontStyle.Italic : FontStyle.Normal,
                                                            new TextRange(0, block.content.Length));
                                    textLayout.SetFontWeight(rssMedia.RssTitleProp.TextFont.Bold ? FontWeight.Bold : FontWeight.Normal,
                                                             new TextRange(0, block.content.Length));
                                    d2dRenderTarget.DrawTextLayout(new Vector2(block.left, block.top), textLayout, titleColorBrush);
                                    break;
                                case RssBodyType.Time:
                                    textLayout.SetUnderline(rssMedia.RssPublishTimeProp.TextFont.Underline,
                                                            new TextRange(0, block.content.Length));
                                    textLayout.SetFontStyle(rssMedia.RssPublishTimeProp.TextFont.Italic ? FontStyle.Italic : FontStyle.Normal,
                                                            new TextRange(0, block.content.Length));
                                    textLayout.SetFontWeight(rssMedia.RssPublishTimeProp.TextFont.Bold ? FontWeight.Bold : FontWeight.Normal,
                                                             new TextRange(0, block.content.Length));
                                    d2dRenderTarget.DrawTextLayout(new Vector2(block.left, block.top), textLayout, publishDateColorBrush);
                                    break;
                                case RssBodyType.Body:
                                    textLayout.SetUnderline(rssMedia.RssBodyProp.TextFont.Underline,
                                                            new TextRange(0, block.content.Length));
                                    textLayout.SetFontStyle(rssMedia.RssBodyProp.TextFont.Italic ? FontStyle.Italic : FontStyle.Normal,
                                                            new TextRange(0, block.content.Length));
                                    textLayout.SetFontWeight(rssMedia.RssBodyProp.TextFont.Bold ? FontWeight.Bold : FontWeight.Normal,
                                                             new TextRange(0, block.content.Length));
                                    d2dRenderTarget.DrawTextLayout(new Vector2(block.left, block.top), textLayout, textBodyBrush);
                                    break;
                                default:
                                    break;
                            }
                            textFormat.Dispose();
                            textFormat = null;
                        }
                    }

                    d2dRenderTarget.EndDraw();

                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                    fileName = string.Format("{0}{1}.jpg", path, DateTime.Now.Ticks.ToString());
                    var stream = new WICStream(wicFactory, fileName, NativeFileAccess.Write);
                    var encoder = new PngBitmapEncoder(wicFactory);
                    encoder.Initialize(stream);
                    var bitmapFrameEncode = new BitmapFrameEncode(encoder);
                    bitmapFrameEncode.Initialize();
                    bitmapFrameEncode.SetSize(pageWidth, pageHeight);
                    var pixelFormatGuid = SharpDX.WIC.PixelFormat.FormatDontCare;
                    bitmapFrameEncode.SetPixelFormat(ref pixelFormatGuid);
                    bitmapFrameEncode.WriteSource(wicBitmap);
                    bitmapFrameEncode.Commit();
                    encoder.Commit();
                    bitmapFrameEncode.Dispose();
                    encoder.Dispose();
                    stream.Dispose();

                    Console.WriteLine("*********image count is : " + count++);
                    //发送单个图片生成事件
                    if (SingleGenerateCompleteEvent != null)
                    {
                        SingleGenerateCompleteEvent(fileName);
                    }
                }
                //发送生成完成事件
                if (GenerateCompleteEvent != null)
                {
                    GenerateCompleteEvent(path);
                    //停止线程，从字典删除
                    StopGenerate(rssMedia.CachePath);
                }
            }
            catch (ThreadAbortException aborted)
            {
                Trace.WriteLine("rss 图片生成线程终止 : " + aborted.Message);
            }
            catch (Exception ex)
            {
                Trace.WriteLine("rss 图片生成遇到bug : " + ex.Message);
            }
            finally
            {
                wicFactory.Dispose();
                d2dFactory.Dispose();
                dwFactory.Dispose();
                wicBitmap.Dispose();
                d2dRenderTarget.Dispose();
                solidColorBrush.Dispose();
                textBodyBrush.Dispose();
                titleColorBrush.Dispose();
                publishDateColorBrush.Dispose();
                rssInfo.Dispose();
                rssMedia.Dispose();
                wicFactory = null;
                d2dFactory = null;
                dwFactory = null;
                wicBitmap = null;
                d2dRenderTarget = null;
                solidColorBrush = null;
                textBodyBrush = null;
                titleColorBrush = null;
                publishDateColorBrush = null;
                rssInfo = null;
                rssMedia = null;
                pages.Clear();
                pages = null;
            }
        }

        Dictionary<string, Thread> threadDic = new Dictionary<string, Thread>();
        public void StartGenerate(RssInfo rssInfo, RssMedia rssMedia, out List<System.Drawing.Image> imgList)
        {
            imgList = new List<System.Drawing.Image>();

            ImageObj obj = new ImageObj(rssInfo, rssMedia);

            Thread thread = new Thread(GetImageFromRss);
            try
            {
                threadDic.Add(rssMedia.CachePath, thread);
                thread.Start(obj);
            }
            catch (ArgumentException aex)
            {
                Trace.WriteLine("rss生成图片线程遇到bug：" + aex.Message);
            }
        }

        object lockobj = new object();
        public void StopGenerate(string path)
        {
            lock (lockobj)
            {
                if (threadDic.ContainsKey(path))
                {
                    threadDic[path].Abort();

                    threadDic.Remove(path);
                }

            }
        }

        public void Dispose()
        {
            foreach (Thread thread in threadDic.Values)
            {
                thread.Abort();
            }

            threadDic.Clear();
        }
    }

    public class ImageObj
    {
        public RssInfo rssInfo;

        public RssMedia rssMedia;

        public ImageObj(RssInfo rssInfo, RssMedia rssMedia)
        {
            this.rssInfo = rssInfo;
            this.rssMedia = rssMedia;
        }
    }
}
