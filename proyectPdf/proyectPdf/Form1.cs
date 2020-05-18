using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.IO;
using iTextSharp;

//using iTextSharp.text.*;
using iTextSharp.text.pdf;
using Font=  iTextSharp.text.Font;
using FontFamily =  iTextSharp.text.Font.FontFamily;
using Image = iTextSharp.text.Image;
using iTextSharp.text;
using System.Security;
using ExcelDataReader;



namespace proyectPdf
{
    public partial class Form1 : Form
    {


        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            label2.Text = "Generando archivos pdf. Por favor espere!";
            String filepath= label1.Text;
            FileStream stream = null;
            try
            {
                stream = File.Open(filepath, FileMode.Open, FileAccess.Read);
            }
            catch
            {
                MessageBox.Show("Por favor cierre el Archivo Excel ingresado y vuelva a intentar", "Mensaje de error");
                Application.Exit();
            }
            IExcelDataReader excelReader;
            if (Path.GetExtension(filepath).ToUpper() == ".xls" )
            {
                //excel xls
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else{
                //excel xlsx
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }
            //DataSet result = excelReader;
            var result = excelReader.AsDataSet();
            DataTable table = result.Tables[0];
            DataRow row = table.Rows[0];
            //String cell = row[0].ToString();
            String cell = result.Tables[0].Rows.Count.ToString();
            String valor1 = table.Rows[1][1].ToString();
            //label2.Text = cell;

            //fuentes
            iTextSharp.text.pdf.BaseFont bf = iTextSharp.text.pdf.BaseFont.CreateFont(iTextSharp.text.pdf.BaseFont.COURIER, iTextSharp.text.pdf.BaseFont.CP1252, iTextSharp.text.pdf.BaseFont.EMBEDDED);
            iTextSharp.text.Font font = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 11, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            iTextSharp.text.Font fontBlack = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 11, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
            //end fuentes

            String pathOrigin = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            String folder = "pdfs "+ DateTime.Now.ToString("dd-MM-yyyy_HH-mm-ss");
            String pathString = System.IO.Path.Combine(pathOrigin, folder);
            System.IO.Directory.CreateDirectory(pathString);
            for (int i = 1; i < table.Rows.Count; i++)
            {
                //System.Console.WriteLine(table.Rows[i][4].ToString());
                //String valorRow0 = result.Tables[0].Rows[i][1].ToString();// fila 1 columna1
                //Paragraph c = new Paragraph(valorRow0, font);
                //doc.Add(c);
                String dni = table.Rows[i][4].ToString(); // cuil

                //verifico si existe el archivo
                //String pathFile1 = pathString + "/" + dni + ".pdf";
                //String pathFile = File.Exists(pathFile1) ? pathString + "/" + dni + "_" + i + ".pdf" : pathFile1  ; 
                Document doc = new Document(PageSize.A4, 90f, 50f, 140f, 0f);
                int dniRepetidos = 0;
                bool repetido = true;
                while (repetido)
                {
                    if (File.Exists(pathString + "/" + dni + ".pdf"))
                    {
                        dniRepetidos = dniRepetidos + 1;
                        if (File.Exists(pathString + "/" + dni + "_" + dniRepetidos + ".pdf"))
                        {

                        }
                        else
                        {
                            repetido = false;
                            PdfWriter wri = PdfWriter.GetInstance(doc, new FileStream(pathString + "/" + dni + "_" + dniRepetidos + ".pdf", FileMode.Create));
                        }
                    }
                    else
                    {
                        repetido = false;
                        PdfWriter wri = PdfWriter.GetInstance(doc, new FileStream(pathString + "/" + dni + ".pdf", FileMode.Create));
                    }
                }

                
                doc.Open();

                //Image image = Image.GetInstance("c:/users/franco/desktop/01.jpg");
                Image image = Image.GetInstance(@"01.jpg");
                Image image02 = Image.GetInstance(@"02.jpg");
                //image.ScalePercent(18f);
                image.ScaleToFit(150f, 110f);
                image02.ScaleToFit(50f, 30f);
                image.SetAbsolutePosition(90, 770);
                image02.SetAbsolutePosition(480, 770);
                //image.ScaleAbsoluteHeight(50);
                //image.ScaleAbsoluteWidth(100);
                doc.Add(image);
                doc.Add(image02);

                //header 
                String anio = DateTime.Today.ToString("yyyy"); //2020
                String nombreDia = DateTime.Today.ToString("dddd"); //martes 
                String numeroDia = DateTime.Today.ToString("dd"); //22
                String mes = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(DateTime.Now.Month); //abril
                String fechaActual = nombreDia + " " + numeroDia + " de " + mes + " del " + anio;
                Paragraph p = new Paragraph("SAN MIGUEL DE TUCUMÁN, " + fechaActual, fontBlack);
                p.Alignment = Element.ALIGN_CENTER;
                doc.Add(p);
                doc.Add(Chunk.NEWLINE);
                Paragraph p2 = new Paragraph("CONSTANCIA RG (DGR) N° 44/20", fontBlack);
                p2.Alignment = Element.ALIGN_CENTER;
                doc.Add(p2);
                doc.Add(Chunk.NEWLINE);
                doc.Add(Chunk.NEWLINE);
                //end header

                String f = "                Conforme lo dispuesto por el 2° párrafo del artículo 282 del Código Tributario Provincial, se deja constancia que el contribuyente identificado con la CUIT/CUIL:  ";
                Paragraph texto = new Paragraph(f, font);
                doc.Add(texto);

                //Paragraph razonsocial = new Paragraph(" CECILIA", fontBlack);
                dni = dni + ", ";
                String razon_social = table.Rows[i][5].ToString();

                Paragraph dni1 = new Paragraph(dni + razon_social, fontBlack);
                doc.Add(dni1);

                //clase para convertir numeros a letras
                convertirNumerosALetras convertir = new convertirNumerosALetras();
                //
                String obligacion = table.Rows[i][3].ToString();
                String fechaHoy = DateTime.Today.ToString("dd-MM-yyyy") + ", ";
                String texto3 = obligacion + " por el instrumento otorgado en fecha " + fechaHoy;
                //String nombrePrecio = " (pesos ________________________).-";
                String precioNro = table.Rows[i][10].ToString(); // TOTAL
                String nombrePrecio = " (pesos "+convertir.convertir(float.Parse(precioNro)) + ").-";

                String precio = precioNro + nombrePrecio;
                String texto4 = texto3 + "que fue presentado en copia ante la DIRECCIÓN GENERAL DE RENTAS, emitiéndose a los fines del pago del Impuesto de Sellos el formulario 600 (F.600), por un importe total de $ " + precio;
                String texto1 = "presentó ante este Organismo Declaración Jurada del Impuesto de Sellos – F.950, Obligación N° " + texto4;
                Paragraph texto2 = new Paragraph(texto1, font);
                doc.Add(texto2);
                doc.Close();

            }//end for

            MessageBox.Show("PDFs creados en\n"+ pathString);
            Close();
            Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //DialogResult resul = new DialogResult();
            OpenFileDialog fil= new OpenFileDialog();
            //fil.Filter = "Excel  *.xls";
            fil.Filter = "Excel | *.xlsx; *.xls";
            if (fil.ShowDialog() ==System.Windows.Forms.DialogResult.OK)
            {
                button1.Visible = true ;
                try
                {
                    label1.Text = fil.FileName;

                }
                catch (SecurityException ex)
                {
                    MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                    $"Details:\n\n{ex.StackTrace}");
                }
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
