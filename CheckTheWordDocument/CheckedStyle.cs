using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using Word = Microsoft.Office.Interop.Word;

namespace CheckTheWordDocument
{
    public class CheckedStyle //класс включает в себя вордовский стиль и ...
    {
        //private bool IsChecked;
        private string FontName; //название шрифта
        private float FontSize; //размер шрифта
        private int FontColor; //цвет шрифта
        private int FontBold; //жирное начертание
        private int FontItalic; //курсивное начертание
        private string Alignment; //выравнивание
        private float LineSpacing; //интервал
        private float SpaceBefore; //отступ до абзаца
        private float SpaceAfter; //отступ после абзаца
        private float FirstLineIndent; //красная строка
        private string NameLocal; //название стиля

        public CheckedStyle(Word.Style style)
        {
            //IsChecked = false;
            FontName = style.Font.NameAscii;
            FontSize = style.Font.Size;
            FontColor = style.Font.TextColor.RGB;
            FontBold = style.Font.Bold;
            FontItalic = style.Font.Italic;
            Alignment = style.ParagraphFormat.Alignment.ToString();
            LineSpacing = style.ParagraphFormat.LineSpacing;
            SpaceBefore = style.ParagraphFormat.SpaceBefore;
            SpaceAfter = style.ParagraphFormat.SpaceAfter;
            FirstLineIndent = style.ParagraphFormat.FirstLineIndent;
            NameLocal = style.NameLocal;
        }

        public string GetNameLocal()
        {
            return NameLocal;
        }

        public string GetFontName()
        {
            return FontName;
        }

        public float GetFontSize()
        {
            return FontSize;
        }

        public int GetFontColor()
        {
            return FontColor;
        }

        public int GetFontBold()
        {
            return FontBold;
        }

        public int GetFontItalic()
        {
            return FontItalic;
        }

        public string GetAlignment()
        {
            return Alignment;
        }

        public float GetLineSpacing()
        {
            return LineSpacing;
        }

        public float GetSpaceBefore()
        {
            return SpaceBefore;
        }

        public float GetSpaceAfter()
        {
            return SpaceAfter;
        }

        public float GetFirstLineIndent()
        {
            return FirstLineIndent;
        }
    }
}
