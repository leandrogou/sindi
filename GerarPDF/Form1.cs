using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

using iTextSharp.text;
using iTextSharp.text.pdf;


namespace prj_gerarPDF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            String dia = DateTime.Now.Day.ToString();
            String mes = DateTime.Now.ToString("MMMM");
            String ano = DateTime.Now.Year.ToString();
            
            int totalfonts = FontFactory.RegisterDirectory("C:\\WINDOWS\\Fonts");

            StringBuilder sb = new StringBuilder();

            String hh = "";

            foreach (string fontname in FontFactory.RegisteredFonts)

            {

                sb.Append(fontname + "\n");
                hh += fontname + "\n";

            }
            MySqlConnection conn = new MySqlConnection("Persist Security Info=False;server=localhost;port=3306;database=bd;uid=user;pwd=password");
            MySqlCommand cmd = null;
            String SQL = "SELECT * FROM tabela_usada";
            
            MySqlDataReader dr = null;
            try
            {
                conn.Open();
                cmd = new MySqlCommand(SQL, conn);
            }
            catch (MySqlException ex)
            {
                System.Windows.Forms.MessageBox.Show("Não foi possível acessar o banco de dados.");
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
            String v0 = "";
            String v1 = "";
            String v2 = "";
            String v3 = "";
            try
            {
                dr = cmd.ExecuteReader();
            }
            catch (MySqlException ex)
            {
                System.Windows.Forms.MessageBox.Show(e.ToString());
            }
            if (dr.HasRows)
            {
                dr.Read();
                v0 = dr.GetString(0);
                v1 = dr.GetString(1);
                v2 = dr.GetString(2);
                v3 = dr.GetString(3);
            }
            conn.Close();
            

            using (var fileStream = new System.IO.FileStream("output.pdf", System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.None))
            {
                var document = new iTextSharp.text.Document();
                var pdfWriter = iTextSharp.text.pdf.PdfWriter.GetInstance(document, fileStream);
                document.SetPageSize(iTextSharp.text.PageSize.A4.Rotate());
                document.Open();
                

                iTextSharp.text.FontFactory.RegisterDirectory("C:\\");

                var contentByte = pdfWriter.DirectContent;
                
                var imagem = iTextSharp.text.Image.GetInstance("..\\..\\Fundo Certificado.jpg");
                imagem.SetAbsolutePosition(0, 0);
                imagem.ScaleToFit(900f, 600f);
                imagem.Alignment = iTextSharp.text.Image.ALIGN_JUSTIFIED;
                imagem.SetAbsolutePosition(0, 0);
                document.Add(imagem);
                
                
                
                var font = iTextSharp.text.FontFactory.GetFont("Century Gothic",BaseFont.CP1252,BaseFont.EMBEDDED,20,0,new iTextSharp.text.BaseColor(0,153,153));
                
                var paragrafo = new iTextSharp.text.Paragraph("", font);
                paragrafo.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                document.Add(paragrafo);
                Chunk c1 = new Chunk("\n\n\n\n\n\nCertificamos que ", font);
                font = iTextSharp.text.FontFactory.GetFont("Century Gothic", BaseFont.CP1252, BaseFont.EMBEDDED, 20, 1, new iTextSharp.text.BaseColor(51, 102, 153));
                Chunk c2 = new Chunk(v0, font);
                Phrase p1 = new Phrase();
                p1.Add(c1);
                p1.Add(c2);
                paragrafo.Add(p1);
                document.Add(paragrafo);
                font = iTextSharp.text.FontFactory.GetFont("Century Gothic", BaseFont.CP1252, BaseFont.EMBEDDED, 20, 0, new iTextSharp.text.BaseColor(0, 153, 153));
                //iTextSharp.text.
                paragrafo = new iTextSharp.text.Paragraph("participou do "+v1, font);
                paragrafo.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
               
                document.Add(paragrafo);

                paragrafo = new iTextSharp.text.Paragraph("realizado em "+v2, font);
                paragrafo.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                document.Add(paragrafo);

                paragrafo = new iTextSharp.text.Paragraph("com carga horária de "+v3, font);
                paragrafo.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                document.Add(paragrafo);

                font = iTextSharp.text.FontFactory.GetFont("Century Gothic", BaseFont.CP1252, BaseFont.EMBEDDED, 20, 0, new iTextSharp.text.BaseColor(51, 102, 153));
                paragrafo = new iTextSharp.text.Paragraph("\nSão Paulo, "+dia +" de "+mes+" de "+ano, font);
                paragrafo.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                document.Add(paragrafo);
                
                imagem = iTextSharp.text.Image.GetInstance("..\\..\\Ass.png");
                
                imagem.ScaleToFit(200f, 200f);
                imagem.Alignment = iTextSharp.text.Image.ALIGN_RIGHT;
                document.Add(imagem);

                font = iTextSharp.text.FontFactory.GetFont("Century Gothic (Corpo)", 12, 0, new iTextSharp.text.BaseColor(51, 102, 153));
                paragrafo = new iTextSharp.text.Paragraph("eeeeeeeeeeee                    \naaaaaaaaaaaaaaa               ", font);
                paragrafo.Alignment = iTextSharp.text.Element.ALIGN_RIGHT;
                document.Add(paragrafo);

                document.Close();
                System.Diagnostics.Process.Start("output.pdf");
            }
        }
    }
}
