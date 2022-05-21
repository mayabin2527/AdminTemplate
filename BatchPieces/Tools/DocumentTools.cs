using System;
using Spire.Doc;
using Spire.Doc.Fields;
using System.Drawing;
using Spire.Doc.Interface;
using Spire.Doc.Documents;

namespace BatchPieces.Tools
{
    public class DocumentTools
    {
        /*public Bookmark()
        {
        }
        */
        private Document doc = null;
        public DocumentTools(Document document)
        {
            doc = document;
        }

        /// <summary>
        /// 用文本替换指定书签的内容
        /// </summary>
        /// <param name="bookmarkName">书签名</param>
        /// <param name="text">文本</param>
        /// <param name="saveFormatting">删除原始书签内容时，是否保留格式</param>
        /// <returns>TextRange</returns>
        public TextRange ReplaceContent(string bookmarkName, string text, bool saveFormatting)
        {
            BookmarksNavigator navigator = new BookmarksNavigator(doc);
            navigator.MoveToBookmark(bookmarkName);//指向特定书签
            navigator.DeleteBookmarkContent(saveFormatting);//删除原有书签内容     
            Spire.Doc.Interface.ITextRange textRange = navigator.InsertText(text);//写入文本
            return textRange as TextRange;
        }

        /// <summary>
        /// 用图片替换指定书签的内容
        /// </summary>
        /// <param name="bookmarkName">书签名</param>
        /// <param name="picPath">图片路径</param>
        /// <param name="widthScale">宽度缩放比例，0以上正整数</param>
        /// <param name="heightScale">高度缩放比例，0以上正整数</param>
        /// <param name="wrapStyle">环绕方式</param>
        /// <param name="horizontalAlignment"></param>
        public void ReplaceContent(string bookmarkName, string picPath, float widthScale, float heightScale, TextWrappingStyle wrapStyle, ShapeHorizontalAlignment horizontalAlignment)
        {
            BookmarksNavigator navigator = new BookmarksNavigator(doc);
            navigator.MoveToBookmark(bookmarkName);
            navigator.DeleteBookmarkContent(false);
            IParagraphBase paragraphBase = navigator.InsertParagraphItem(ParagraphItemType.Picture);//插入类型为图片
            Image image = Image.FromFile(picPath);//加载图片
            DocPicture picture = paragraphBase.OwnerParagraph.AppendPicture(image);//插入图片
            picture.WidthScale = widthScale;
            picture.HeightScale = heightScale;
            picture.TextWrappingStyle = wrapStyle;
            picture.HorizontalAlignment = horizontalAlignment;
        }

        /// <summary>
        /// 用表格替换指定书签的内容
        /// </summary>
        /// <param name="bookmarkName">书签名</param>
        /// <param name="table">Table实例</param>
        public void ReplaceContent(string bookmarkName, Table table)
        {
            BookmarksNavigator navigator = new BookmarksNavigator(doc);
            navigator.MoveToBookmark(bookmarkName);
            TextBodyPart body = new TextBodyPart(doc);
            body.BodyItems.Add(table);
            navigator.ReplaceBookmarkContent(body);
        }

        /// <summary>
        /// 创建表格并写入数据，返回Table对象
        /// </summary>
        /// <param name="rowsNum">行数</param>
        /// <param name="columnsNum">列数</param>
        /// <param name="columnWidth">列宽</param>
        /// <param name="horizontalAlignment">水平对齐方式</param>
        /// <param name="datatable">DataTable实例</param>
        /// <returns></returns>
        public Table CreateTable(int rowsNum, int columnsNum, float columnWidth, RowAlignment horizontalAlignment, System.Data.DataTable datatable)
        {
            Table table = new Table(doc, true, 1f);//初始化Table对象
            table.ResetCells(rowsNum, columnsNum);//设置行数和列数
            //填充数据
            for (int i = 0; i < datatable.Rows.Count; i++)
            {
                for (int j = 0; j < datatable.Columns.Count; j++)
                {
                    table.Rows[i].Cells[j].AddParagraph().AppendText(datatable.Rows[i][j].ToString());
                }
            }
            //设置列宽
            for (int i = 0; i < rowsNum; i++)
            {
                for (int j = 0; j < columnsNum; j++)
                {
                    table.Rows[i].Cells[j].Width = columnWidth;
                }
            }
            table.TableFormat.HorizontalAlignment = horizontalAlignment;//表格水平对齐方式
            return table;
        }
    }
}
