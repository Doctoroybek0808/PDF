using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;


namespace PDF
{
    public class PDF
    {
        public int row_index, i;
        public string tmpFile;
        public PdfPTable mainTable;
        public int count = 0;
        public PDF(string title)
        {
         
            tmpFile = System.IO.Path.GetTempFileName().Replace(".tmp", ".pdf");
            mainTable = new PdfPTable(1);
            mainTable.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;
        }
       
        public void AddData(params object[] args)
        {



            var table1 = new PdfPTable(3);
            int j = i + 4 * i;
            row_index = i * 5;
            int color_num = (int)args[6];
            string sylfaenpath = Environment.GetEnvironmentVariable("SystemRoot") + "\\fonts\\sylfaen.ttf";
            BaseFont sylfaen = BaseFont.CreateFont(sylfaenpath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            iTextSharp.text.Font fontHeader = new iTextSharp.text.Font(sylfaen, 12f, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
            iTextSharp.text.Font fontCell = new iTextSharp.text.Font(sylfaen, 10f, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            iTextSharp.text.Font underline = new iTextSharp.text.Font(sylfaen, 10f, iTextSharp.text.Font.UNDERLINE, BaseColor.BLACK);


            byte[] bits = null;
            var table_img = new PdfPTable(1);
            Paragraph paragraph = new Paragraph("#157117");
            paragraph.Alignment = 0;
            PdfPCell cell = new PdfPCell();
            cell.AddElement(paragraph);
            if (args[0] != null)
            {
                System.Drawing.Image image = args[0] as System.Drawing.Image;

                Bitmap objBitmap = new Bitmap(image, new Size(100, 75));
                bits = ImageToByte(objBitmap);
                var tmpImg = Path.GetTempFileName();
                objBitmap.Save(tmpImg);
                iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(tmpImg);

                cell.AddElement(jpg);
            }

            cell.HorizontalAlignment = 1;
            cell.VerticalAlignment = 1;
            cell.PaddingTop = 10;
            cell.FixedHeight = 180;
            cell.Rowspan = 4;
            Color baseColor = Color.White;
            if (color_num == 1)
            {
                baseColor = Color.FromArgb(197, 224, 179);
            }
            else if (color_num == 2)
            {
                baseColor = Color.FromArgb(244, 176, 131);
            }
            else if (color_num == 0)
            {
                baseColor = Color.White;
            }
            PdfPCell key = new PdfPCell(new Phrase("Номер:", fontHeader));
            PdfPCell val = new PdfPCell(new Phrase(args[2].ToString(), fontCell));
            key.BackgroundColor = new BaseColor(baseColor);
            val.BackgroundColor = new BaseColor(baseColor);
            key.FixedHeight = 45;
            val.FixedHeight = 45;
            PdfPRow row = new PdfPRow(new PdfPCell[] { cell, key, val });

            table1.Rows.Add(row);

            key = new PdfPCell(new Phrase("Дата:", fontHeader));
            val = new PdfPCell(new Phrase(args[3].ToString(), fontCell));
            key.BackgroundColor = new BaseColor(baseColor);
            val.BackgroundColor = new BaseColor(baseColor);
            key.FixedHeight = 45;
            val.FixedHeight = 45;
            row = new PdfPRow(new PdfPCell[] { null, key, val });
            table1.Rows.Add(row);

            key = new PdfPCell(new Phrase("Направления:", fontHeader));
            val = new PdfPCell(new Phrase(args[4].ToString(), fontCell));
            key.BackgroundColor = new BaseColor(baseColor);
            val.BackgroundColor = new BaseColor(baseColor);
            key.FixedHeight = 45;
            val.FixedHeight = 45;
            row = new PdfPRow(new PdfPCell[] { null, key, val });
            table1.Rows.Add(row);

            key = new PdfPCell(new Phrase("В список:", fontHeader));
            val = new PdfPCell(new Phrase(args[5].ToString(), fontCell));
            key.BackgroundColor = new BaseColor(baseColor);
            val.BackgroundColor = new BaseColor(baseColor);
            key.FixedHeight = 45;
            val.FixedHeight = 45;
            row = new PdfPRow(new PdfPCell[] { null, key, val });
            table1.Rows.Add(row);
            
           



            PdfPCell cellMain = new PdfPCell(table1);
            cellMain.Border = iTextSharp.text.Rectangle.NO_BORDER;
            mainTable.Rows.Add(new PdfPRow(new PdfPCell[] { cellMain }));
            PdfPCell blank = new PdfPCell(new Phrase("\n"));
            blank.Border = iTextSharp.text.Rectangle.NO_BORDER;
            mainTable.Rows.Add(new PdfPRow(new PdfPCell[] { blank }));
        }

        public bool Save(string path, out string Error)
        {
            try
            {
                using (var document = new Document(PageSize.A4, 5f, 5f, 20f, 10f))
                {
                    PdfWriter.GetInstance(document, new System.IO.FileStream(tmpFile, System.IO.FileMode.Create));
                    document.Open();

                    document.Add(mainTable);

                    document.Close();
                }
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
                File.Copy(tmpFile, path);
                string strFileName = tmpFile;
                string ChangeExtension = path;

                File.Delete(tmpFile);
                Error = string.Empty;
                return true;
            }
            catch (Exception ex)
            {
                Error = ex.Message;
                return false;
            }
        }

        public byte[] ImageToByte(System.Drawing.Image img)
        {
            ImageConverter converter = new ImageConverter();
            return (byte[])converter.ConvertTo(img, typeof(byte[]));
        }

        public void Dispose()
        {

        }
    }
}
