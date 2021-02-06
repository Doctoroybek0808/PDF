namespace PDF
{
    public class PDFInit
    {
        public static void Init()
        {
            string _namespace = "PDF";
            string[] recources = new string[] { "itextsharp.dll", "Spire.XLS.dll", "Spire.Doc.dll", "Spire.License.dll", "Spire.Pdf.dll" };
 
            foreach (string recource in recources)
                EmbeddedAssembly.Load(_namespace + "." + recource, recource);
        }
    }
}
