using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace CheckTheWordDocument
{
    public static class Checker
    {
        private static Word.Application wordapp;
        private static Word.Document worddocument;
        //private static Word.Documents worddocuments;
        //private static Word.Paragraphs wordparagraphs;
        //private static Word.Paragraph wordparagraph;
        private static string _path = ""; //путь к выбранному файлу
        private static string _pathtotemplate = ""; //путь к шаблону
        private static List<string> style_names = new List<string>(); //названия эталонных стилей
        private static List<string> style_names_for_removing; //= new List<string>();
      //  private static List<Style> style_from_template = new List<Style>(); //OpenXML-стили из шаблона
        private static List<CheckedStyle> word_style_from_document = new List<CheckedStyle>(); //Word-стили из документа
        private static List<CheckedStyle> word_style_from_template = new List<CheckedStyle>(); //Word-стили из шаблона
      //  private static List<Style> style_from_document = new List<Style>(); //OpenXML-стили из документа

        public static void AddStyleNames(List<string> s_names) //adding standart style names 
        {
            foreach (string s in s_names)
                style_names.Add(s);
        }

        public static void Start(string path)
        {
            _path = path;
            _pathtotemplate = Environment.CurrentDirectory + "\\шаблон.docx";           

            /*style_names.Add("Название статьи"); //запоминаем названия эталонных стилей
            style_names.Add("Авторы");
            style_names.Add("ТЕКСТ");
            style_names.Add("Формула");
            style_names.Add("Рисунок");
            style_names.Add("Подпись к рисунку");
            style_names.Add("Таблица");
            style_names.Add("Заголовок таблицы");
            style_names.Add("Подпись к таблице");
            style_names.Add("Разделы статьи");
            style_names.Add("Список_литературы");*/

            style_names_for_removing = style_names.ToList();

           /* List<Style> styles = GetStylesPart(_pathtotemplate, false, false); //получаем эталонные стили из шаблона
            foreach (Style cs in styles)
            {
                style_from_template.Add(cs);
            }

            List<Style> stylesdoc = GetStylesPart(_path, false, false); //получаем стили из документа, названия которых совпадают с эталонными
            foreach (Style cs in stylesdoc)
            {
                style_from_document.Add(cs);
            }*/
        }

        public static void OpenWordDocument(string path_to)
        {
            wordapp = new Word.Application();
            wordapp.Visible = false; //true;
            Object filename = path_to;
            Object confirmConversions = true;
            Object readOnly = false;
            Object addToRecentFiles = true;
            Object passwordDocument = Type.Missing;
            Object passwordTemplate = Type.Missing;
            Object revert = false;
            Object writePasswordDocument = Type.Missing;
            Object writePasswordTemplate = Type.Missing;
            Object format = Type.Missing;
            Object encoding = Type.Missing; //;
            Object oVisible = Type.Missing;
            Object openConflictDocument = Type.Missing;
            Object openAndRepair = Type.Missing;
            Object documentDirection = Type.Missing;
            Object noEncodingDialog = false;
            Object xmlTransform = Type.Missing;
            //#if OFFICEXP
            //  worddocument=wordapp.Documents.Open2000(ref filename, .....
            //#else
            worddocument = wordapp.Documents.Open(ref filename,
                //#endif 
           ref confirmConversions, ref readOnly, ref addToRecentFiles,
           ref passwordDocument, ref passwordTemplate, ref revert,
           ref writePasswordDocument, ref writePasswordTemplate,
           ref format, ref encoding, ref oVisible,
           ref openAndRepair, ref documentDirection, ref noEncodingDialog, ref xmlTransform);
        }

        public static List<String> CheckAndWrite() //проверка всего документа, исправление, запись примечаний
        {
            List<String> result = new List<string>(); //GetSetStyleFromDoc();
            //var styles = CheckStyles(_path, false); //получаем стили из документа
            //var defaultstyles = CheckStyles(_pathtotemplate, false); //получаем стили из шаблона
            //foreach (Style s in styles.Nodes().)

            //int CountParagraphs;

            OpenWordDocument(_path);
            //Получаем ссылки на параграфы документа
            //wordparagraphs = worddocument.Paragraphs;

            //List<string> SelectedParagraphs = new List<string>();
            for (int i = 1; i < worddocument.Paragraphs.Count; i++) //numeration of paragraphs starts from 1
            {
                string WordP = worddocument.Paragraphs[i].Range.Text; // get paragraph text
                Word.Style style = ((Word.Style)worddocument.Paragraphs[i].get_Style()); //get style of paragraph
                string WordS = ((Word.Style)worddocument.Paragraphs[i].get_Style()).NameLocal; //get paragraph style name

                if (style_names.LastIndexOf(WordS) != -1) //если стиль входит в список эталонных
                {
                    int index = style_names_for_removing.LastIndexOf(WordS);
                    if (index != -1) //если стиль ещё ни разу не добавлялся в список
                    {
                        word_style_from_document.Add(new CheckedStyle(style)); //добавляем стиль
                        style_names_for_removing.RemoveAt(style_names_for_removing.IndexOf(WordS)); //удаляем его из списка недобавленных стилей
                    }
                    //((Word.Style)worddocument.Paragraphs[i].get_Style()).Borders.
                    //CompareStyles(_pathtotemplate, _path, WordS);  // check some fields... 
                    //result.Add(WordS + " шрифт " + style.Font.NameAscii + " размер " + style.Font.Size + " жирный " + style.Font.Bold);
                }
                else
                {
                    if (WordP.Length <= 100)  //выводим первые 100 символов абзаца, если он длиннее, то обрезаем до ста символов
                        result.Add("Стиль " + WordS + " не входит в список эталонов: " + WordP + "...");
                    else
                        result.Add("Стиль " + WordS + " не входит в список эталонов: " + WordP.Substring(0, 100) + "...");
                }
            }
            if (!CheckPageMargins()) //проверяем отступы
            {
                result.Add("Неверные поля");
                /*wordparagraph.Range.Text = "Поля исправлены";
                worddocument.Paragraphs.Add(ref oMissing); //next step
                wordparagraph = (Word.Paragraph)wordparagraphs[wordparagraphs.Count]; */
                //next step
            }
            CloseWordDocument();

            OpenWordDocument(_pathtotemplate); //открываем шаблонный документ

            for (int i = 1; i < worddocument.Paragraphs.Count; i++) //numeration of paragraphs starts from 1
            {
                string WordP = worddocument.Paragraphs[i].Range.Text; // get paragraph text
                Word.Style style = ((Word.Style)worddocument.Paragraphs[i].get_Style()); //get style of paragraph
                string WordS = ((Word.Style)worddocument.Paragraphs[i].get_Style()).NameLocal; //get paragraph style name

                if (style_names.LastIndexOf(WordS) != -1) //если стиль входит в список эталонных
                {
                    word_style_from_template.Add(new CheckedStyle(style));
                }
            }
            CloseWordDocument();

            List<String> styles_comparing_errors = CompareStyles(); //соответствуют ли параметры форматирования стилей документа параметрам стилей из шаблона
            foreach (String s in styles_comparing_errors) //добавляем результаты проверки в список результатов
                result.Add(s);

            return result;
        }

        //worddocument.ListTemplates[0].
        //Выводим текст в первый параграф
        /*wordparagraph.Range.Text = "Текст который мы выводим в 1 абзац";
        //Меняем характеристики текста и параграфа
        wordparagraph.Range.Font.Color = Word.WdColor.wdColorBlue;
        wordparagraph.Range.Font.Size = 20;
        wordparagraph.Range.Font.Name = "Arial";
        wordparagraph.Range.Font.Italic = 1;
        wordparagraph.Range.Font.Bold = 0;
        wordparagraph.Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
        wordparagraph.Range.Font.UnderlineColor = Word.WdColor.wdColorDarkRed;*/

        public static List<String> CompareStyles() //сравниваем эталонные стили из шаблона и стили из текущего документа 
        {
            List<String> result = new List<string>();

            foreach (CheckedStyle style in word_style_from_document) //берём стиль документа
                {
                foreach (CheckedStyle s in word_style_from_template) //просматриваем стили шаблона в поисках совпадающего по имени
                {
                    if (s.GetNameLocal().Equals(style.GetNameLocal()))//(s.GetName().Equals(style.GetName())) //если названия стилей совпадают
                    {
                        if ((s.GetFontName() != null) && (style.GetFontName() != null) &&
                            (!s.GetFontName().Equals(style.GetFontName()))) //название шрифта
                            result.Add("Стиль " + style.GetNameLocal() + " имеет шрифт, отличающийся от эталонного: " +
                                style.GetFontName() + " вместо " + s.GetFontName());
                        //try
                        //{
                        if //((s.GetFontSize() != null) && (style.GetFontSize() != null) &&
                                (!s.GetFontSize().Equals(style.GetFontSize())) //размер шрифта
                            result.Add("Стиль " + style.GetNameLocal() + " имеет размер шрифта, отличающийся от эталонного: " +
                            style.GetFontSize() + " вместо " + s.GetFontSize());
                        //}
                        //catch (NullReferenceException) { }

                        if (!s.GetFontColor().Equals(style.GetFontColor())) //цвет шрифта
                            result.Add("Стиль " + style.GetNameLocal() + " имеет цвет шрифта, отличающийся от эталонного: "); //+
                            //style.GetFontColor() + " вместо " + s.GetFontColor());

                        if (s.GetFontBold() != (style.GetFontBold())) //жирное начертание
                            result.Add("Стиль " + style.GetNameLocal() + " имеет жирное начертание, отличающееся от эталонного: ");

                        if (s.GetFontItalic() != (style.GetFontItalic())) //курсивное начертание
                            result.Add("Стиль " + style.GetNameLocal() + " имеет курсивное начертание, отличающееся от эталонного: ");

                        if (s.GetAlignment() != (style.GetAlignment())) //выравнивание
                            result.Add("Стиль " + style.GetNameLocal() + " имеет выравнивание, отличающееся от эталонного: " +
                                style.GetAlignment() + " вместо " + s.GetAlignment());

                        if (s.GetLineSpacing() != (style.GetLineSpacing())) //интервал
                            result.Add("Стиль " + style.GetNameLocal() + " имеет интервал, отличающийся от эталонного: " +
                                style.GetLineSpacing() + " вместо " + s.GetLineSpacing());

                        if (s.GetSpaceBefore() != (style.GetSpaceBefore())) //интервал до абзаца
                            result.Add("Стиль " + style.GetNameLocal() + " имеет интервал до абзаца, отличающийся от эталонного: " +
                                style.GetSpaceBefore() + " вместо " + s.GetSpaceBefore());

                        if (s.GetSpaceAfter() != (style.GetSpaceAfter())) //интервал после абзаца
                            result.Add("Стиль " + style.GetNameLocal() + " имеет интервал после абзаца, отличающийся от эталонного: " +
                                style.GetSpaceAfter() + " вместо " + s.GetSpaceAfter());

                        if (s.GetFirstLineIndent() != (style.GetFirstLineIndent())) //красная строка
                            result.Add("Стиль " + style.GetNameLocal() + " имеет красную строку, отличающуюся от эталонной: " +
                                style.GetFirstLineIndent() + " вместо " + s.GetFirstLineIndent());

                        //try
                        //{
                           // if //((s.GetStyle().Font.Bold != null) && (style.GetStyle().Font.Bold != null) &&
                              /*  (!s.GetStyle().Font.Bold.Equals(style.GetStyle().Font.Bold)) //цвет текста
                                result.Add("Стиль " + style.GetName() + " имеет жирное выделение, отличное от эталонного: "); //+*/
                                    //style.GetStyle().Font.Bold + " вместо " + s.GetStyle().Font.Bold);
                        //}
                        //catch (NullReferenceException) { }

                       /* try
                        {
                            if ((s.StyleRunProperties.RunFonts != null) && (style.StyleRunProperties.RunFonts != null) &&
                                (!s.StyleRunProperties.RunFonts.LocalName.ToString().Equals(style.StyleRunProperties.RunFonts.LocalName.ToString()))) //название шрифта
                                result.Add("Стиль " + style.StyleName.Val.Value + " имеет название шрифта, отличающееся от эталонного: " +
                                    style.StyleRunProperties.RunFonts.LocalName.ToString() + " вместо " + s.StyleRunProperties.RunFonts.LocalName.ToString());
                        }
                        catch (NullReferenceException) { }*/
                        //else result.Add("Стиль " + style.StyleName.Val.Value + " совпадает с эталонным, ура!");                           
                        //если какого-то из свойств нет, то игнорим его
                        break;
                    }
                }
            }

            return result;
        }

        /*public static List<String> CompareStyles() //сравниваем эталонные стили из шаблона и стили из текущего документа 
        {
            List<String> result = new List<string>();

            foreach (Style style in style_from_document) //берём стиль документа
            {
                foreach (Style s in style_from_template) //просматриваем стили шаблона в поисках совпадающего по имени
                {
                    if (s.StyleName.Val.Value.Equals(style.StyleName.Val.Value)) //если названия стилей совпадают
                    {

                        if ((s.StyleParagraphProperties.SpacingBetweenLines != null) || (style.StyleParagraphProperties.SpacingBetweenLines != null) ||
                            (!s.StyleParagraphProperties.SpacingBetweenLines.ToString().Equals(style.StyleParagraphProperties.SpacingBetweenLines.ToString()))) //расстояние между строками
                            result.Add("Стиль " + style.StyleName.Val.Value + " имеет расстояние между строками, отличающееся от эталонного: " +
                                style.StyleParagraphProperties.SpacingBetweenLines.ToString() + " вместо " + s.StyleParagraphProperties.SpacingBetweenLines.ToString());
                        try
                        {
                            if ((s.StyleRunProperties.Bold != null) && (style.StyleRunProperties.Bold != null) &&
                                (!s.StyleRunProperties.Bold.Val.Value.Equals(style.StyleRunProperties.Bold.Val.Value))) //жирное начертание
                                result.Add("Стиль " + style.StyleName.Val.Value + " имеет жирное начертание, отличающееся от эталонного");
                        }
                        catch (NullReferenceException) { }

                        try
                        {
                            if ((s.StyleRunProperties.Color != null) && (style.StyleRunProperties.Color != null) &&
                                (!s.StyleRunProperties.Color.Val.Value.Equals(style.StyleRunProperties.Color.Val.Value))) //цвет текста
                                result.Add("Стиль " + style.StyleName.Val.Value + " имеет цвет текста, отличный от эталонного: " +
                                    style.StyleRunProperties.Color.Val.Value + " вместо " + s.StyleRunProperties.Color.Val.Value);
                        }
                        catch (NullReferenceException) { }

                        try
                        {
                            if ((s.StyleRunProperties.RunFonts != null) && (style.StyleRunProperties.RunFonts != null) &&
                                (!s.StyleRunProperties.RunFonts.LocalName.ToString().Equals(style.StyleRunProperties.RunFonts.LocalName.ToString()))) //название шрифта
                                result.Add("Стиль " + style.StyleName.Val.Value + " имеет название шрифта, отличающееся от эталонного: " +
                                    style.StyleRunProperties.RunFonts.LocalName.ToString() + " вместо " + s.StyleRunProperties.RunFonts.LocalName.ToString());
                        }
                        catch (NullReferenceException) { }
                        //else result.Add("Стиль " + style.StyleName.Val.Value + " совпадает с эталонным, ура!");                           
                        //если какого-то из свойств нет, то игнорим его
                        break;
                    }
                }
            }

            return result;
        }*/

        private static bool CheckPageMargins()
        {
            Object start = Type.Missing;
            Object end = Type.Missing;
            Word.Range wordrange = worddocument.Range(ref start, ref end);
            if ((wordrange.PageSetup.LeftMargin != wordapp.CentimetersToPoints((float)2.22)) &&
                (wordrange.PageSetup.RightMargin != wordapp.CentimetersToPoints((float)1.15)) &&
                wordrange.PageSetup.TopMargin != wordapp.CentimetersToPoints((float)2.0) &&
                wordrange.PageSetup.BottomMargin != wordapp.CentimetersToPoints((float)2.0))
            {
                //
                /*wordrange.PageSetup.LeftMargin = wordapp.CentimetersToPoints((float)2.22);
                wordrange.PageSetup.RightMargin = wordapp.CentimetersToPoints((float)1.15);
                wordrange.PageSetup.TopMargin = wordapp.CentimetersToPoints((float)2.0);
                wordrange.PageSetup.BottomMargin = wordapp.CentimetersToPoints((float)2.0);*/
                return false;
            }
            else return true;
        }

        /*private static void CompareStyles(string pathtotemplate, string pathtofile, string stylename) //проверяем, являются ли идентичными стили с одинаковым именем
        {
            XDocument templatestyles = GetStylesPart(_pathtotemplate);
            XDocument documentstyles = GetStylesPart(_path);
            //templatestyles.
        }*/

        private static void InsertPageBreak()
        {
            //Сдвигаемся вниз в конец документа
            object unit;
            object extend;
            unit = Word.WdUnits.wdStory;
            extend = Word.WdMovementType.wdMove;
            wordapp.Selection.EndKey(ref unit, ref extend);
            object oType;
            oType = Word.WdBreakType.wdSectionBreakNextPage;
            //И на новый лист
            wordapp.Selection.InsertBreak(ref oType);
        }


        private static XDocument GetStylesDoc(string path, bool getStylesWithEffectsPart = false) //выделяем из документа часть, в которой хранятся стили
        {
            XDocument styles = null;
            if (path != null)
            {
                using (var document = WordprocessingDocument.Open(path, false))
                {
                    var docPart = document.MainDocumentPart;

                    StylesPart stylesPart = null;
                    if (getStylesWithEffectsPart)
                        stylesPart = docPart.StylesWithEffectsPart;
                    else
                        stylesPart = docPart.StyleDefinitionsPart;

                    if (stylesPart != null)
                    {
                        using (var reader = XmlNodeReader.Create(
                          stylesPart.GetStream(FileMode.Open, FileAccess.Read)))
                        {
                            styles = XDocument.Load(reader);
                        }
                    }
                }
            }
            return styles;
        }

        public static List<Style> GetStylesPart(string path, bool selectAll, bool getStylesWithEffectsPart = false) //выделяем из документа часть, в которой хранятся стили
        {
            StylesPart stylesPart = null;
            List<Style> cs = new List<Style>();
            if (path != null)
            {
                using (var document = WordprocessingDocument.Open(path, false))
                {
                    var docPart = document.MainDocumentPart;

                    if (getStylesWithEffectsPart)
                        stylesPart = docPart.StylesWithEffectsPart;
                    else
                        stylesPart = docPart.StyleDefinitionsPart;
                    foreach (Style style in stylesPart.RootElement.Elements<Style>())
                    {
                        if (selectAll) //if we need all styles
                            cs.Add(style); 
                        else if (style_names.LastIndexOf(style.StyleName.Val.Value) != -1) //если данный стиль есть в списке эталонных стилей
                            cs.Add(style);
                    }
                }
            }
            return cs;
        }


        /*   public static List<String> GetSetStyleFromDoc()
           {
               List<String> messages=new List<String>();
               try {
                   using (var document = WordprocessingDocument.Open(_path, true))
                   {
                       // Get the Styles part for this document
                       StyleDefinitionsPart part = document.MainDocumentPart.StyleDefinitionsPart;
                       foreach (Style style in part.RootElement.Elements<Style>())
                       {
                           if (style_names.LastIndexOf(style.StyleName.Val.Value) != -1) //если данный стиль есть в списке эталонных стилей
                           {
                               //messages.Add(style.StyleParagraphProperties.SpacingBetweenLines.Before); //интервал..?
                               //style.StyleParagraphProperties.Indentation.FirstLine //отступ
                               //style.StyleRunProperties.RunFonts //шрифт
                               //style.StyleRunProperties.FontSize //размер шрифта


                               //messages.Add("Стиль " + style.StyleName.Val.Value + " отсутствует в списке эталонов.");//+style.StyleParagraphProperties.FirstChild.ToString());
                               //part.RootElement.Elements().RemoveChild(style); //то удаляем его ...на!!!
                           }
                        
                           //if (style.StyleId.Value.Equals("ТЕКСТ", StringComparison.InvariantCultureIgnoreCase))
                           if (style.StyleName.Val.Value.Equals("ТЕКСТ", StringComparison.InvariantCulture)) //InvariantCultureIgnoreCase))
                           {
                               //style.StyleParagraphProperties.SpacingBetweenLines.Line = "276";
                               //style.StyleRunProperties.FontSize.Val = "24";
                               //style.StyleRunProperties.FontSize.Val=  //.Color.Val = "4F81BD"; // font color
                               //style.StyleParagraphProperties.

                           }
                           else part.RootElement.RemoveChild<Style>(style); //удаляем неправильный стиль 

                           else if (style.StyleRunProperties.Color.Val != null)
                           {
                               //style.StyleRunProperties.Color.Val = "AABBBB"; // font color
                           }
                       }

                   }
               }
               catch(Exception){} //....
               return messages;
           } */

        public static void SaveWordDocument(string path)
        {
            /*worddocuments = wordapp.Documents;
            Object name = "Документ2";
            //Для Visual Studio 2003
            //worddocument=(Word.Document)worddocuments.Item(ref name);
            worddocument = (Word.Document)worddocuments.get_Item(ref name);
            worddocument.Activate();*/
            //Подготавливаем параметры для сохранения документа
            Object fileName = path; //@"C:\Users\LenovoG560\Downloads\anosova+.doc";
            Object fileFormat = Word.WdSaveFormat.wdFormatDocument;
            Object lockComments = false;
            Object password = "";
            Object addToRecentFiles = false;
            Object writePassword = "";
            Object readOnlyRecommended = false;
            Object embedTrueTypeFonts = false;
            Object saveNativePictureFormat = false;
            Object saveFormsData = false;
            Object saveAsAOCELetter = Type.Missing;
            Object encoding = Type.Missing;
            Object insertLineBreaks = Type.Missing;
            Object allowSubstitutions = Type.Missing;
            Object lineEnding = Type.Missing;
            Object addBiDiMarks = Type.Missing;
            //#if OFFICEXP
            //worddocument.SaveAs2000(ref fileName,
            //#else
            worddocument.SaveAs(ref fileName,
                //#endif
        ref fileFormat, ref lockComments,
       ref password, ref addToRecentFiles, ref writePassword,
       ref readOnlyRecommended, ref embedTrueTypeFonts,
       ref saveNativePictureFormat, ref saveFormsData,
       ref saveAsAOCELetter, ref encoding, ref insertLineBreaks,
       ref allowSubstitutions, ref lineEnding, ref addBiDiMarks);
        }

        public static void CloseWordDocument()
        {
            Object saveChanges = Word.WdSaveOptions.wdPromptToSaveChanges;
            Object originalFormat = Word.WdOriginalFormat.wdWordDocument;
            Object routeDocument = Type.Missing;
            if (wordapp != null)
            {
                wordapp.Quit(ref saveChanges,
                    ref originalFormat, ref routeDocument);
                wordapp = null;
            }
        }
    }
}
