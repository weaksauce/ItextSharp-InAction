using System;
using System.Collections.Generic;
using System.Text;

using System.Threading;
using System.Globalization;

using iTextSharp.text;
using iTextSharp.text.pdf;

using System.IO;
using System.Data;
using System.Collections;

namespace ItextInAction
{
    class Program
    {
        //public const string APP_PATH = @"C:\Projetos\ItextInAction\bin\Debug\";
        public const string APP_PATH = @"D:\_projetos\ItextSharpInAction\bin\Debug";


        public static void Main(string[] args)
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("pt-BR");
            DataSet bd = SomeData.setDataSet();
            /*
            MiscPdfUse.CriarPDF();
            MiscPdfUse.PDFWithImage();
            MiscPdfUse.PDFWithTable();
            MiscPdfUse.GetPDFInfo(Program.APP_PATH + "pdf-with-image.pdf");
            MiscPdfUse.WorkWithStamper();
            MiscPdfUse.WorkWithForm(Program.APP_PATH + "pdf-form.pdf");
            MiscPdfUse.FillDataForm(Program.APP_PATH + "pdf-form.pdf");
           
            ArrayList ar_guias = new ArrayList();

            for (int i = 1; i <= 150; i++)
            {
                if (i % 5 == 0)
                {
                    Console.Clear();
                    Console.Write(".");
                }
                else
                {
                    Console.Write(".");
                }

                ar_guias.Add(MiscPdfUse.FillDataForm(Program.APP_PATH + "pdf-form.pdf", MiscPdfUse.GetFormData(), i.ToString()));
            }

            Console.Clear();
            MiscPdfUse.CloseFileGuias(ar_guias); 
            */
            BasicPdfUse.PdfCreateInFiveSteps();
            BasicPdfUse.PdfWhitMirroredMargins();
            BasicPdfUse.PdfWriteInMemoryFirst();
            BasicPdfUse.JustChunks();
            BasicPdfUse.SomePhrases(true);
            BasicPdfUse.SomeParagraph();
            BasicPdfUse.SomeOrderedList();
            BasicPdfUse.SomeImage();


            SomeUsePdfContentByte.helloWord();
        }

        /// <summary>
        /// Converte medidas de mm para pt
        /// </summary>
        /// <param name="measure">mediade em mm</param>
        /// <returns>medida convertida para pt</returns>
        public static float millimetersToPoits(float measure)
        {
            return (72f * measure) / 25.4f;
        }

        /// <summary>
        /// Converte medidas de pt para mm
        /// </summary>
        /// <param name="measure">medida em pt</param>
        /// <returns>medida em mm</returns>
        public static float poitsToMillimeters(float measure)
        {
            return (25.4f * measure) / 72f;
        }

        
    }

    class BasicPdfUse
    {
        public static void PdfCreateInFiveSteps()
        {
            PdfCreateInFiveSteps("Hello Word!!");
        }

        public static void PdfCreateInFiveSteps(string content)
        {
            // definindo e passando o retangulo para o documento
            //Rectangle pagesize = new Rectangle(Program.millimetersToPoits(210f), Program.millimetersToPoits(297f)); //define o tamanho da pagina do documento
            //Document doc = new Document(pagesize, Program.millimetersToPoits(10f), Program.millimetersToPoits(10f), Program.millimetersToPoits(10f), Program.millimetersToPoits(10f));

            // utilizando a classe PageSize que já possui um serie de retangulos estáficos pre definidos.
            Document doc = new Document(PageSize.A5.Rotate(), Program.millimetersToPoits(10f), Program.millimetersToPoits(10f), Program.millimetersToPoits(10f), Program.millimetersToPoits(10f));
            PdfWriter pdfw = PdfWriter.GetInstance(doc, new FileStream(Program.APP_PATH + "hello-word.pdf", FileMode.Create));

            doc.Open();
            doc.Add(new Chunk(content));
            doc.Close();
        }

        public static void PdfWhitMirroredMargins()
        {
            Document doc = new Document(PageSize.A6);
            doc.SetMargins(Program.millimetersToPoits(25), Program.millimetersToPoits(10), Program.millimetersToPoits(10), Program.millimetersToPoits(10));
            PdfWriter pdfw = PdfWriter.GetInstance(doc, new FileStream(Program.APP_PATH + "hello-word-mirroerdmargins.pdf", FileMode.Create));
            doc.SetMarginMirroring(true);

            string txt = @"Cras nec ante eu nunc pulvinar iaculis vitae quis velit. Suspendisse potenti. Maecenas id massa purus, accumsan tincidunt libero? Curabitur ut lacus ac turpis faucibus pharetra. Praesent sapien diam, tincidunt eu ultricies vitae, malesuada id turpis. Integer lectus velit, varius quis semper et, feugiat at orci. Aliquam tempor sapien eget tellus cursus quis auctor nisl congue. Suspendisse pellentesque ullamcorper magna, eget scelerisque purus rutrum eu. Praesent in pharetra metus. Proin ullamcorper egestas suscipit. Suspendisse lectus turpis, porttitor nec porta et, interdum sed dui. Nullam sed turpis at dui aliquam sodales vehicula blandit sem. Nunc a lorem diam, id sollicitudin sem. Nam placerat, velit id euismod convallis, mi libero molestie enim, nec ornare lectus nisi quis quam. Donec magna orci, adipiscing at dapibus nec, viverra eu ipsum! Etiam non egestas ante.
Sed tincidunt mi eget lectus consectetur et eleifend orci pretium. Suspendisse vitae odio vel dui scelerisque consequat tristique non risus. Vestibulum sollicitudin placerat arcu, a iaculis ligula molestie a. Vivamus adipiscing tellus in nisi luctus iaculis sollicitudin metus porttitor? Maecenas bibendum nunc placerat nisi pellentesque sit amet commodo urna pretium. Praesent congue eleifend ornare. Aenean eu elit porta sapien pulvinar tincidunt. In semper neque id dui luctus sed sagittis elit cursus. Nullam condimentum, sem et pretium gravida, nisi sem sollicitudin felis, ac porttitor nisl mi a ligula. Mauris interdum, magna eu pellentesque dapibus, ligula est facilisis dui, sed luctus nibh nunc id nibh. Morbi et est et felis pretium rutrum. Vivamus condimentum lectus et odio scelerisque nec mollis ligula molestie. Sed vitae tellus erat. Suspendisse imperdiet iaculis dolor vel varius. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos.
Praesent pellentesque nibh id odio feugiat placerat. Fusce pharetra; tellus et mattis dapibus, lectus nulla cursus dui, sit amet rutrum lorem tortor sed orci. Sed malesuada velit et augue semper hendrerit. Duis tempus magna nulla, sed mattis magna. Quisque quam purus, tincidunt in condimentum eu, ultrices eget orci. Suspendisse sagittis, lorem in volutpat condimentum, est enim volutpat sapien, non rhoncus turpis ipsum in ante. Maecenas vitae nisl dolor. Pellentesque congue risus vitae ante pharetra hendrerit? Pellentesque sit amet nisi et libero fringilla condimentum at nec lectus. Sed elit dui, semper vel cursus sit amet, molestie ut sem. Etiam semper orci ac elit faucibus sit amet congue nibh ornare!
Sed tincidunt mi eget lectus consectetur et eleifend orci pretium. Suspendisse vitae odio vel dui scelerisque consequat tristique non risus. Vestibulum sollicitudin placerat arcu, a iaculis ligula molestie a. Vivamus adipiscing tellus in nisi luctus iaculis sollicitudin metus porttitor? Maecenas bibendum nunc placerat nisi pellentesque sit amet commodo urna pretium. Praesent congue eleifend ornare. Aenean eu elit porta sapien pulvinar tincidunt. In semper neque id dui luctus sed sagittis elit cursus. Nullam condimentum, sem et pretium gravida, nisi sem sollicitudin felis, ac porttitor nisl mi a ligula. Mauris interdum, magna eu pellentesque dapibus, ligula est facilisis dui, sed luctus nibh nunc id nibh. Morbi et est et felis pretium rutrum. Vivamus condimentum lectus et odio scelerisque nec mollis ligula molestie. Sed vitae tellus erat. Suspendisse imperdiet iaculis dolor vel varius. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos.
";
            doc.Open();
            Paragraph p = new Paragraph(txt);
            p.Alignment = Element.ALIGN_JUSTIFIED_ALL;
            doc.Add(p);
            doc.Close();

        }

        public static void PdfWriteInMemoryFirst()
        {
            Document doc = new Document(PageSize.A5);
            MemoryStream ms = new MemoryStream();
            PdfWriter pdfw = PdfWriter.GetInstance(doc, ms);
            pdfw.CloseStream = false; // impede que o stream seja fechado
            string txt = @"Cras nec ante eu nunc pulvinar iaculis vitae quis velit. Suspendisse potenti. Maecenas id massa purus, accumsan tincidunt libero? Curabitur ut lacus ac turpis faucibus pharetra. Praesent sapien diam, tincidunt eu ultricies vitae, malesuada id turpis. Integer lectus velit, varius quis semper et, feugiat at orci. Aliquam tempor sapien eget tellus cursus quis auctor nisl congue. Suspendisse pellentesque ullamcorper magna, eget scelerisque purus rutrum eu. Praesent in pharetra metus. Proin ullamcorper egestas suscipit. Suspendisse lectus turpis, porttitor nec porta et, interdum sed dui. Nullam sed turpis at dui aliquam sodales vehicula blandit sem. Nunc a lorem diam, id sollicitudin sem. Nam placerat, velit id euismod convallis, mi libero molestie enim, nec ornare lectus nisi quis quam. Donec magna orci, adipiscing at dapibus nec, viverra eu ipsum! Etiam non egestas ante.
Sed tincidunt mi eget lectus consectetur et eleifend orci pretium. Suspendisse vitae odio vel dui scelerisque consequat tristique non risus. Vestibulum sollicitudin placerat arcu, a iaculis ligula molestie a. Vivamus adipiscing tellus in nisi luctus iaculis sollicitudin metus porttitor? Maecenas bibendum nunc placerat nisi pellentesque sit amet commodo urna pretium. Praesent congue eleifend ornare. Aenean eu elit porta sapien pulvinar tincidunt. In semper neque id dui luctus sed sagittis elit cursus. Nullam condimentum, sem et pretium gravida, nisi sem sollicitudin felis, ac porttitor nisl mi a ligula. Mauris interdum, magna eu pellentesque dapibus, ligula est facilisis dui, sed luctus nibh nunc id nibh. Morbi et est et felis pretium rutrum. Vivamus condimentum lectus et odio scelerisque nec mollis ligula molestie. Sed vitae tellus erat. Suspendisse imperdiet iaculis dolor vel varius. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos.
Praesent pellentesque nibh id odio feugiat placerat. Fusce pharetra; tellus et mattis dapibus, lectus nulla cursus dui, sit amet rutrum lorem tortor sed orci. Sed malesuada velit et augue semper hendrerit. Duis tempus magna nulla, sed mattis magna. Quisque quam purus, tincidunt in condimentum eu, ultrices eget orci. Suspendisse sagittis, lorem in volutpat condimentum, est enim volutpat sapien, non rhoncus turpis ipsum in ante. Maecenas vitae nisl dolor. Pellentesque congue risus vitae ante pharetra hendrerit? Pellentesque sit amet nisi et libero fringilla condimentum at nec lectus. Sed elit dui, semper vel cursus sit amet, molestie ut sem. Etiam semper orci ac elit faucibus sit amet congue nibh ornare!
Sed tincidunt mi eget lectus consectetur et eleifend orci pretium. Suspendisse vitae odio vel dui scelerisque consequat tristique non risus. Vestibulum sollicitudin placerat arcu, a iaculis ligula molestie a. Vivamus adipiscing tellus in nisi luctus iaculis sollicitudin metus porttitor? Maecenas bibendum nunc placerat nisi pellentesque sit amet commodo urna pretium. Praesent congue eleifend ornare. Aenean eu elit porta sapien pulvinar tincidunt. In semper neque id dui luctus sed sagittis elit cursus. Nullam condimentum, sem et pretium gravida, nisi sem sollicitudin felis, ac porttitor nisl mi a ligula. Mauris interdum, magna eu pellentesque dapibus, ligula est facilisis dui, sed luctus nibh nunc id nibh. Morbi et est et felis pretium rutrum. Vivamus condimentum lectus et odio scelerisque nec mollis ligula molestie. Sed vitae tellus erat. Suspendisse imperdiet iaculis dolor vel varius. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos.";
            try
            {
                doc.Open();
                doc.Add(new Paragraph("Documento criado primeiramente na memoria e depois escrito em arquivo"));
                doc.Add(new Paragraph(txt));
                doc.Close();

                FileStream fs = new FileStream(Program.APP_PATH + "pdf-memory-first.pdf", FileMode.Create);
                ms.WriteTo(fs);
                fs.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        public static void JustChunks()
        {
            string[] chunks = { "Pellentesque vel ipsum lorem.", "Maecenas ut nisi quis.", "In ante nunc.", "Suspendisse dolor lacus.", "Cras sit amet convallis elit" };
            Document doc = new Document();
            PdfWriter wri = PdfWriter.GetInstance(doc, new FileStream(Program.APP_PATH + "just-chunks.pdf", FileMode.Create));
            wri.InitialLeading = 20;

            doc.Open();
            int i = 1;
            foreach (string ch in chunks)
            {
                doc.Add(new Chunk(ch));
                doc.Add(new Chunk(" "));

                Font font = new Font(Font.FontFamily.HELVETICA, 6, Font.BOLD, BaseColor.WHITE);
                Chunk id = new Chunk(i.ToString(), font);
                id.SetBackground(BaseColor.BLACK, 1f, 0.5f, 1f, 1.5f);
                id.SetTextRise(6);
                doc.Add(id);
                doc.Add(Chunk.NEWLINE);
                i++;
            }
            doc.Close();
        }

        public static void SomePhrases()
        {
            string[] director = { "David", "Martin", "Ethan", "Steven", "Terrence" };
            string[] director_named = { "Lynch", "Scorsese", "Coen", "Soderbergh", "Malick" };            Phrase directors = new Phrase();

            for (int i = 0; i < director.Length; i++ )
            {
                directors.Add(new Chunk(director[i], SomeFonts.BOLD_UNDERLINE));
                directors.Add(new Chunk(",", SomeFonts.BOLD_UNDERLINE));
                directors.Add(new Chunk(" ", SomeFonts.NORMAL));
                directors.Add(new Chunk(director_named[i], SomeFonts.NORMAL));
                directors.Add(Chunk.NEWLINE);
            }

            Document doc = new Document();
            MemoryStream ms = new MemoryStream();
            PdfWriter pdw = PdfWriter.GetInstance(doc, ms);
            pdw.CloseStream = false;

            doc.Open();
            doc.Add(directors);
            doc.Close();

            GeneralUse.PdfToFile(ms, "just-first-phrase.pdf");
        }

        public static void SomePhrases(bool embededFont)
        {
            string[] director = { "David", "Martin", "Ethan", "Steven", "Terrence" };
            string[] director_named = { "Lynch", "Scorsese", "Coen", "Soderbergh", "Malick" }; Phrase directors = new Phrase();

            directors.Leading = 24;

            for (int i = 0; i < director.Length; i++)
            {
                Chunk name = new Chunk(director[i], SomeFonts.EMBED_BOLD_UNDERLINE);
                name.SetUnderline(0.2f, -2f);

                directors.Add(name);
                directors.Add(new Chunk(",", SomeFonts.EMBED_BOLD_UNDERLINE));
                directors.Add(new Chunk(" ", SomeFonts.EMBED_NORMAL));
                directors.Add(new Chunk(director_named[i], SomeFonts.EMBED_NORMAL));
                directors.Add(new Chunk("\n", SomeFonts.EMBED_NORMAL));
            }

            Document doc = new Document();
            MemoryStream ms = new MemoryStream();
            PdfWriter pdw = PdfWriter.GetInstance(doc, ms);
            pdw.CloseStream = false;

            doc.Open();
            doc.Add(directors);
            doc.Close();

            GeneralUse.PdfToFile(ms, "just-first-phrase_embeded-font.pdf");
        }

        public static void SomeParagraph()
        {
            DataTable dt = SomeData.GetMovieInfos();
            Document doc = new Document();
            MemoryStream ms = new MemoryStream();
            PdfWriter pdw = PdfWriter.GetInstance(doc, ms);
            pdw.CloseStream = false;

            doc.Open();
            foreach (DataRow dw in dt.Rows)
            {
                Paragraph movie = new Paragraph();
                movie = CreateMovieInformation(dw);
                movie.Alignment = Element.ALIGN_JUSTIFIED;
                movie.FirstLineIndent = 0;
                movie.IndentationLeft = 0;
                movie.SpacingAfter = 20;
                doc.Add(movie);
            }
            doc.Close();

            GeneralUse.PdfToFile(ms, "some-movies-info_phrase.pdf");
        }

        private static Paragraph CreateYearAndDuration(DataRow rw)
        {
            Paragraph info = new Paragraph();
            info.Font = SomeFonts.FILM_NORMAL;
            info.Add(new Chunk("Ano: ", SomeFonts.FILM_BOLD_ITALIC));
            info.Add(new Chunk(rw["year"].ToString(), SomeFonts.FILM_NORMAL));
            info.Add(" ");
            info.Add(new Chunk("Duração: ", SomeFonts.FILM_BOLD_ITALIC));
            info.Add(new Chunk(rw["duration"].ToString(), SomeFonts.FILM_NORMAL));
            info.Add(new Chunk(" minutos", SomeFonts.FILM_NORMAL));
            info.Add(" ");

            return info;
        }

        private static Paragraph CreateMovieInformation(DataRow rw)
        {
            Paragraph p = new Paragraph();
            p.Font = SomeFonts.FILM_NORMAL;
            p.Add(new Phrase("Titulo: ", SomeFonts.FILM_BOLD_ITALIC));
            p.Add(new Phrase(rw["title"].ToString()));
            p.Add(" ");
            p.Add(new Phrase("Titulo Original: ", SomeFonts.FILM_BOLD_ITALIC));
            p.Add(new Phrase(rw["title_original"].ToString()));
            p.Add(" ");
            p.Add(new Phrase("País de Origem: ", SomeFonts.FILM_BOLD_ITALIC));
            p.Add(new Phrase(rw["country"].ToString()));
            p.Add(" ");
            p.Add(new Phrase("Diretor: ", SomeFonts.FILM_BOLD_ITALIC));
            p.Add(new Phrase(rw["director"].ToString()));
            p.Add(" ");

            p.Add(CreateYearAndDuration(rw));

            return p;
        }

        public static void SomeOrderedList()
        {
            DataTable dt = SomeData.GetMovieInfos();
            Hashtable rootCoutry = new Hashtable();

            foreach (DataRow rw in dt.Rows)
                if (!rootCoutry.ContainsKey(rw["country"]))
                    rootCoutry.Add(rw["country"], rw["country"]);

            List listCountry = new List(List.ORDERED);
            foreach (DictionaryEntry country in rootCoutry)
            {
                ListItem itemCountry = new ListItem(country.Value.ToString(), SomeFonts.FILM_BOLD);
                List listMovie = new List(List.ORDERED, List.ALPHABETICAL);
                listMovie.Lowercase = true;
                foreach (DataRow rw in dt.Rows)
                {
                    if (rw["country"].ToString() == country.Value.ToString())
                    {
                        ListItem itemMovie = new ListItem(String.Format("Titulo: {0} [diretor: {1}]", rw["title"], rw["director"]));
                        listMovie.Add(itemMovie);
                    }
                }
                itemCountry.Add(listMovie);
                listCountry.Add(itemCountry);
            }

            Document doc = new Document();
            MemoryStream ms = new MemoryStream();
            PdfWriter pdfw = PdfWriter.GetInstance(doc, ms);
            pdfw.CloseStream = false;

            doc.Open();
            doc.Add(listCountry);
            doc.Close();

            GeneralUse.PdfToFile(ms, "some-list.pdf");
        }

        public static void SomeImage()
        {
            DataTable dt = SomeData.GetMovieInfos();
            Document doc = new Document(PageSize.A4);
            MemoryStream ms = new MemoryStream();
            PdfWriter pdfw = PdfWriter.GetInstance(doc, ms);
            pdfw.CloseStream = false;

            doc.Open();
            foreach (DataRow dw in dt.Rows){
                if (!String.IsNullOrEmpty(dw["img"].ToString()))
                {
                    doc.NewPage();
                    doc.Add(new Paragraph(dw["title"].ToString(), SomeFonts.FILM_BOLD));
                    doc.Add(Image.GetInstance(Program.APP_PATH + dw["img"].ToString()));
                }
                else
                {
                    doc.Add(new Paragraph(dw["title"].ToString(), SomeFonts.FILM_BOLD));
                }
            }
            doc.Close();

            GeneralUse.PdfToFile(ms, "some-image.pdf");

        }
    }

    public class SomeUseForColumnText
    {
        
    }

    public class SomeUsePdfContentByte
    {
        public static void helloWord()
        {
            Document doc = new Document();
            MemoryStream ms = new MemoryStream();
            PdfWriter writer = PdfWriter.GetInstance(doc, ms);
            writer.CloseStream = false;

            doc.Open();
            PdfContentByte cb = writer.DirectContentUnder;
        }
    }

    public class AdvancedPdfUse
    {
        public static void SomeColumnText()
        {

        }
    }

    abstract class SomeFonts
    {
        public static Font BOLD_UNDERLINE = new Font(Font.FontFamily.TIMES_ROMAN, 12,  Font.BOLD | Font.UNDERLINE);
        public static Font NORMAL = new Font(Font.FontFamily.TIMES_ROMAN, 12);

        public static Font EMBED_BOLD_UNDERLINE = new Font(BaseFont.CreateFont(@"C:\Windows\Fonts\timr65w.ttf", BaseFont.WINANSI, BaseFont.EMBEDDED));
        public static Font EMBED_NORMAL = new Font(BaseFont.CreateFont(@"C:\Windows\Fonts\timr45w.ttf", BaseFont.WINANSI, BaseFont.EMBEDDED));

        public static Font FILM_NORMAL = new Font(BaseFont.CreateFont(@"C:\Windows\Fonts\ncsr55w.ttf", BaseFont.WINANSI, BaseFont.EMBEDDED));
        public static Font FILM_BOLD = new Font(BaseFont.CreateFont(@"C:\Windows\Fonts\ncsr75w.ttf", BaseFont.WINANSI, BaseFont.EMBEDDED));
        public static Font FILM_ITALIC = new Font(BaseFont.CreateFont(@"C:\Windows\Fonts\ncsr56w.ttf", BaseFont.WINANSI, BaseFont.EMBEDDED));
        public static Font FILM_BOLD_ITALIC = new Font(BaseFont.CreateFont(@"C:\Windows\Fonts\ncsr76w.ttf", BaseFont.WINANSI, BaseFont.EMBEDDED));
    }

    abstract class SomeData
    {
        public static DataSet setDataSet()
        {
            DataSet bd = new DataSet("movies");
            bd.Tables.Add(new DataTable("film_director"));
            bd.Tables.Add(new DataTable("film_movietitle"));
            bd.Tables.Add(new DataTable("film_country"));
            bd.Tables.Add(new DataTable("film_movie_director"));
            bd.Tables.Add(new DataTable("film_movie_coutry"));

            bd.Tables["film_director"].Columns.Add(new DataColumn("id", typeof(int)));
            bd.Tables["film_director"].Columns.Add(new DataColumn("name", typeof(string)));
            bd.Tables["film_director"].Columns.Add(new DataColumn("give_name", typeof(string)));

            bd.Tables["film_director"].Columns["id"].AutoIncrement = true;
            bd.Tables["film_director"].Columns["id"].AllowDBNull = false;
            bd.Tables["film_director"].Constraints.Add("pk_filme_director", bd.Tables["film_director"].Columns["id"], true);

            return bd;
        }

        public static DataTable GetFormData()
        {
            string dummy = @"Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos.";
            DataTable dt = new DataTable("formData");
            dt.Columns.Add(new DataColumn("txtNome", typeof(string)));
            dt.Columns.Add(new DataColumn("txtCidade", typeof(string)));
            dt.Columns.Add(new DataColumn("txtDescricao", typeof(string)));

            dt.Rows.Add("Jackson", "Curitiba", dummy);
            dt.Rows.Add("Maria Antonieta", "Boa Vista", dummy);
            dt.Rows.Add("Carlos Drummond", "Santo Antônio da Platina", dummy);
            dt.Rows.Add("Carlos Magno", "Siqueira Campos", dummy);
            dt.Rows.Add("Maria Bethânia", "Natal", dummy);
            dt.Rows.Add("Julio Ceasar", "Sergipe", dummy);
            dt.Rows.Add("Paulinho da Viola", "Terezina", dummy);
            dt.Rows.Add("Pedro Amaral", "Curralinho", dummy);

            return dt;
        }

        public static DataTable GetMovieInfos()
        {
            DataTable dt = new DataTable("moviesinfo");
            dt.Columns.Add("title", typeof(string));
            dt.Columns.Add("title_original", typeof(string));
            dt.Columns.Add("director", typeof(string));
            dt.Columns.Add("year", typeof(int));
            dt.Columns.Add("duration", typeof(int));
            dt.Columns.Add("country", typeof(string));
            dt.Columns.Add("img", typeof(string));

            dt.Rows.Add("Um Sonho de Liberdade", "The Shawshank Redemption", "Frank Darabont", 1994, 142, "USA", "sonho-de-liberdade.jpg");
            dt.Rows.Add("O Poderoso Chefão", "The Godfather", "Francis Ford Coppola", 1972, 175, "USA");
            dt.Rows.Add("Três Homens em Conflito", "Il buono, il brutto, il cattivo", "Sergio Leone", 1966, 161, "Italy");
            dt.Rows.Add("Pulp Fiction - Tempo de Violência", "Pulp Fiction", "Quentin Tarantino", 1994, 154, "USA");
            dt.Rows.Add("A Origem", "Inception", "Christopher Nolan", 2010, 148, "USA", "a-origem.jpg");
            dt.Rows.Add("Um Estranho no Ninho", "One Flew Over the Cuckoo's Nest", "Milos Forman", 1975, 133, "USA", "um-estranho-no-ninho.jpg");
            dt.Rows.Add("12 Homens e uma Sentença", "12 Angry Men", "Sidney Lumet", 1957, 96, "USA");
            dt.Rows.Add("A Lista de Schindler", "Schindler's List", "Steven Spielberg", 1993, 195, "USA");
            dt.Rows.Add("Batman - O Cavaleiro das Trevas", "The Dark Knight", "Christopher Nolan", 2008, 152, "USA");
            dt.Rows.Add("Ensaio Sobre a Cegueira", "Blindness", "Fernando Meirelles", 2008, 121, "Brazil");
            dt.Rows.Add("O Império Contra-Ataca", "Star Wars: Episode V - The Empire Strikes Back", "Irvin Kershner", 1980, 124, "USA");
            dt.Rows.Add("Cidade de Deus", "Cidade de Deus", "Fernando Meirelles", 2002, 130, "Brazil");
            dt.Rows.Add("Tropa de Elite 2 - O Inimigo Agora é Outro", "Tropa de Elite 2 - O Inimigo Agora é Outro", "José Padilha", 2010, 115, "Brazil");
            dt.Rows.Add("Io sono l'amore", "Io sono l'amore", "Luca Guadagnino", 2009, 120, "Italy");
            dt.Rows.Add("Nine", "Nine", "Rob Marshall", 2009, 118, "Italy");
            dt.Rows.Add("A Queda! As Últimas Horas de Hitler", "Der Untergang", "Oliver Hirschbiegel", 2004, 156, "Germany");
            dt.Rows.Add("A Vida É Bela", "La vita è bella", "Roberto Benigni", 1997, 116, "Italy");

            //dt.DefaultView.RowFilter = "duration > 150";
            //dt.DefaultView.Sort = "country ASC";
              
            return dt;
        }

        public static DataTable GetDataTable()
        {
            DataTable dt = new DataTable("pdfdt");
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Nome", typeof(string));

            dt.Rows.Add(1, "Tadeu");
            dt.Rows.Add(2, "Mario");
            dt.Rows.Add(3, "Aberlardo");
            dt.Rows.Add(4, "Juca");
            dt.Rows.Add(5, "Eufrazino");
            dt.Rows.Add(6, "Maria");
            dt.Rows.Add(7, "Janaina");

            return dt;
        }
    }

    abstract class GeneralUse
    {
        public static void PdfToFile(MemoryStream ms, string fileName)
        {
            try
            {
                FileStream fs = new FileStream(Program.APP_PATH + fileName, FileMode.Create);
                ms.WriteTo(fs);
                fs.Close();
                ms.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
    }
    class MiscPdfUse
    {
        public static void CriarPDF()
        {
            Document doc = new Document(iTextSharp.text.PageSize.A4, 10, 10, 30, 30);

            try
            {
                PdfWriter pdfw = PdfWriter.GetInstance(doc, new FileStream(Program.APP_PATH + "pdf-with-text.pdf", FileMode.Create));
                doc.Open();

                Paragraph paragraph = new Paragraph("This is my first line using Paragraph.");
                Phrase pharse = new Phrase("This is my second line using Pharse.");
                Chunk chunk = new Chunk(" This is my third line using Chunk.");

                doc.Add(paragraph);
                doc.Add(pharse);
                doc.Add(chunk);
            }
            catch (DocumentException dex)
            {
                Console.WriteLine(dex.Message);
            }
            finally
            {
                doc.Close();
            }
        }

        public static void PDFWithImage()
        {
            Document doc = new Document(iTextSharp.text.PageSize.A4, 10, 10, 30, 30);

            try
            {
                string pdfPath = Program.APP_PATH + "pdf-with-image.pdf";
                PdfWriter pdfw = PdfWriter.GetInstance(doc, new FileStream(pdfPath, FileMode.Create));
                doc.Open();

                Paragraph paragraph = new Paragraph("This is my first line using Paragraph.");

                string imgPath = Program.APP_PATH + "sample-4.jpg";
                iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(imgPath);
                jpg.ScaleToFit(280f, 260f);
                jpg.SpacingBefore = 30f;
                jpg.SpacingAfter = 10f;
                jpg.Alignment = Element.ALIGN_CENTER;
                doc.Add(paragraph);
                doc.Add(jpg);
            }
            catch (DocumentException dex)
            {
                Console.WriteLine(dex.Message);
            }
            finally
            {
                doc.Close();
            }
        }

        public static void PDFWithTable()
        {
            Document doc = new Document(iTextSharp.text.PageSize.A4, 30, 30, 30, 30);

            try
            {
                string pdfPath = Program.APP_PATH + "pdf-with-table.pdf";
                PdfWriter pdfw = PdfWriter.GetInstance(doc, new FileStream(pdfPath, FileMode.Create));
                doc.Open();

                Font f8 = FontFactory.GetFont("ARIAL", 7);

                Paragraph paragraph = new Paragraph("Using ITextsharp I am going to show how to create simple table in PDF document ");

                DataTable dt = SomeData.GetDataTable();
                
                if (dt != null)
                {
                    PdfPTable table = new PdfPTable(dt.Columns.Count);
                    PdfPCell cell = null;

                    cell = new PdfPCell(new Phrase(new Chunk("ID", f8)));
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(new Chunk("Name", f8)));
                    table.AddCell(cell);

                    for (int rows = 0; rows < dt.Rows.Count; rows++)
                    {
                        for (int column = 0; column < dt.Columns.Count; column++)
                        {
                            cell = new PdfPCell(new Phrase(new Chunk(dt.Rows[rows][column].ToString(), f8)));
                            table.AddCell(cell);
                        }
                    }

                    table.SpacingBefore = 15f; // Give some space after the text or it may overlap the table

                    doc.Add(paragraph);// add paragraph to the document
                    doc.Add(table); // add pdf table to the document
                }
            }
            catch (DocumentException dex)
            {
                Console.WriteLine(dex.Message);
            }
            finally
            {
                doc.Close();
            }
        }

        public static void GetPDFInfo(string pdfPath)
        {
            PdfReader pdfr = new PdfReader(pdfPath);
            Rectangle pgsize = pdfr.GetPageSize(1);

            Console.WriteLine(pdfPath);
            Console.WriteLine("Número de Paginas: {0}", pdfr.NumberOfPages);
            Console.WriteLine("Tamanho da Pagina: {0} x {1}", pgsize.Height, pgsize.Width);
            Console.WriteLine("Rotação da Pagina: {0}", pdfr.GetPageRotation(1));
            Console.WriteLine("Tamanho com rotação: {0}", pdfr.GetPageSizeWithRotation(1));
            Console.WriteLine("Tamanho do Arquivo: {0}", pdfr.FileLength);
            Console.WriteLine("Foi refeito: {0}", pdfr.IsRebuilt());
            Console.WriteLine("Está encriptado: {0}", pdfr.IsEncrypted());
        }

        public static void CopyPdf()
        {
        }

        public static ArrayList GetFieldsKeys(AcroFields af)
        {
            ArrayList aux = new ArrayList();


            PdfReader pdfr = new PdfReader(Program.APP_PATH + "pdf-form.pdf");

   //1. foreach (KeyValuePair<string, AcroFields.Item> de in pdfReader.AcroFields.Fields)  
   //2. {  
   //3.     sb.Append(de.Key.ToString() + Environment.NewLine);  
   //4. }  

            foreach (KeyValuePair<string, AcroFields.Item> de in af.Fields)
            {
                string nome = de.Key.ToString();
                string[] temp = nome.Split(".".ToCharArray());
                nome = temp[temp.Length - 1].TrimEnd("[0]".ToCharArray());

                aux.Add(nome);
            }

            return aux;
        }

        public static void WorkWithStamper()
        {
            PdfReader pdfr = new PdfReader(Program.APP_PATH + "pdf-form.pdf");
            PdfStamper pdfs = new PdfStamper(pdfr, new FileStream(Program.APP_PATH + "pdf-stamper.pdf", FileMode.Create));
            PdfContentByte pdfcb = pdfs.GetOverContent(1);
            ColumnText.ShowTextAligned(pdfcb, Element.ALIGN_CENTER, new Phrase("Hello Word!!"), 36, 540, 90);
            pdfs.Close();
        }

        public static void CloseFileGuias(ArrayList guias)
        {
            Document doc = new Document();
            PdfSmartCopy copy = new PdfSmartCopy(doc, new FileStream(Program.APP_PATH + DateTime.Now.ToFileTime() + ".pdf", FileMode.Create));
            PdfReader reader;
            doc.Open();

            for (int i = 0; i < guias.Count; i++)
            {
                reader = new PdfReader(((MemoryStream)guias[i]).ToArray());
                for (int j = 1; j <= reader.NumberOfPages; j++)
                {
                    copy.AddPage(copy.GetImportedPage(reader, j));
                }
                //
            }

            doc.Close();
        }

        public static MemoryStream FillDataForm(string tplPdfPath, DataTable dt, string numGuia)
        {
            ArrayList ar_ps = new ArrayList();
            ArrayList ar_ms = new ArrayList();
            ArrayList nameFields = new ArrayList();

            foreach (DataRow drw in dt.Rows)
            {
                PdfReader prra = new PdfReader(new RandomAccessFileOrArray(tplPdfPath), null);
                MemoryStream ms = new MemoryStream();
                PdfStamper psra = new PdfStamper(prra, ms);
                AcroFields form = psra.AcroFields;

                GetFieldsKeys(psra.AcroFields);

                form.SetField("txtNumGuia", numGuia);
                form.SetField("txtNome", "[" + numGuia + "] " + drw[dt.Columns.IndexOf("txtNome")].ToString());
                form.SetField("txtCidade", "[" + numGuia + "] " + drw[dt.Columns.IndexOf("txtCidade")].ToString());
                form.SetField("txtDescricao", "[" + numGuia + "] " + drw[dt.Columns.IndexOf("txtDescricao")].ToString());

                ar_ps.Add(psra);
                ar_ms.Add(ms);
            }

            MemoryStream msdoc = new MemoryStream();
            Document doc = new Document();

            //PdfSmartCopy copy = new PdfSmartCopy(doc, new FileStream(Program.APP_PATH + DateTime.Now.ToFileTime() + ".pdf", FileMode.Create));
            PdfSmartCopy copy = new PdfSmartCopy(doc, msdoc);
            PdfReader reader;
            int i = 0;
            doc.Open();

            foreach (PdfStamper ps in ar_ps)
            {
                ps.FormFlattening = true;
                ps.Close();
                reader = new PdfReader(((MemoryStream)ar_ms[i]).ToArray());
                copy.AddPage(copy.GetImportedPage(reader, 1));
                i++;
            }

            doc.Close();

            return msdoc;
            /*
            form.SetField("txtNome", "Maria Fernanda Barroso");
            form.SetField("txtCidade", "Belém do Pará");
            form.SetField("txtDescricao", dummy);
            ps.FormFlattening = true;
            ps.Close();
             */
        }

        public static void FillDataForm(string tplPdfPath)
        {
            PdfReader pr = new PdfReader(tplPdfPath);
            PdfStamper ps = new PdfStamper(pr, new FileStream(Program.APP_PATH + DateTime.Now.ToFileTime() + ".pdf", FileMode.Create));
            AcroFields form = ps.AcroFields;

            string dummy = @"Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos. Morbi accumsan tincidunt venenatis. Nullam non lorem eget sem volutpat sollicitudin! Etiam vehicula, mi eu pellentesque dignissim, quam arcu euismod leo, sed laoreet lacus odio tincidunt arcu. Vivamus nec consectetur diam. Duis congue urna non tellus tempus aliquet. Aenean ac risus gravida justo vestibulum rutrum? Aliquam dignissim, diam nec gravida semper, urna ligula hendrerit nisl, a gravida turpis arcu sit amet libero. Vivamus eget sem et erat consequat placerat eu vitae nisi? Donec tempor, nulla quis facilisis semper, dolor mi porta sem; sed euismod nibh elit at augue.
Ut et felis risus. Nunc suscipit malesuada quam, ut fermentum urna dictum sit amet. Proin tristique, augue non vestibulum volutpat, mi turpis molestie magna, auctor scelerisque nibh urna a sem. Proin eget sapien a tortor accumsan dignissim! Suspendisse a vehicula magna. Morbi nec sem orci. Praesent faucibus enim sed enim sollicitudin blandit sit amet eget augue. Nunc adipiscing diam vitae lectus luctus eget adipiscing metus volutpat. Pellentesque felis mi, laoreet quis semper a, scelerisque vitae elit! Ut eget dui eu nisi placerat eleifend non id sem! Quisque pretium, velit egestas convallis ultrices, nisi justo mattis felis, non condimentum ligula nunc sit amet diam! Cras molestie libero ut justo egestas in consectetur dui porta. Quisque non est a tortor euismod vehicula quis eget diam. Vivamus ultricies pulvinar nulla, vel molestie tortor consectetur a? Sed id nunc urna. Praesent tempus nulla sed est pretium non eleifend ipsum sodales.
";
            form.SetField("txtNome", "Maria Fernanda Barroso");
            form.SetField("txtCidade", "Belém do Pará");
            form.SetField("txtDescricao", dummy);
            ps.FormFlattening = true;
            ps.Close();
        }

        public static void Fill()
        {
        }

        public static void WorkWithForm(string formPath)
        {
            PdfReader pdfr = new PdfReader(formPath);
            AcroFields form = pdfr.AcroFields;

            foreach (KeyValuePair<string, AcroFields.Item> key in form.Fields)
            {
                Console.Write("Key: ");

                switch (form.GetFieldType(key.Key.ToString()))
                {
                    case AcroFields.FIELD_TYPE_CHECKBOX:
                        Console.WriteLine("Checkbox");
                        break;
                    case AcroFields.FIELD_TYPE_COMBO:
                        Console.WriteLine("Combobox");
                        break;
                    case AcroFields.FIELD_TYPE_LIST:
                        Console.WriteLine("List");
                        break;
                    case AcroFields.FIELD_TYPE_NONE:
                        Console.WriteLine("None");
                        break;
                    case AcroFields.FIELD_TYPE_PUSHBUTTON:
                        Console.WriteLine("Pushbutton");
                        break;
                    case AcroFields.FIELD_TYPE_RADIOBUTTON:
                        Console.WriteLine("Radiobutton");
                        break;
                    case AcroFields.FIELD_TYPE_SIGNATURE:
                        Console.WriteLine("Signature");
                        break;
                    case AcroFields.FIELD_TYPE_TEXT:
                        Console.WriteLine("Text");
                        break;
                    default:
                        Console.WriteLine("?");
                        break;
                }
            }
        }
    }
}
