namespace ConsultaPantone.Models
{
    public class PantoneItem
    {
        public string? PantoneTpx { get; set; }
        public string? Pagina2 { get; set; }
        public string? Coluna2 { get; set; }
        public string? Linha2 { get; set; }
        public string? NomePantone { get; set; }
        public double CorCo { get; set; }
        public double CorCoAlt { get; set; }
        public double CorCv { get; set; }
        public double CorPa { get; set; }
        public double CorPes { get; set; }
        public double CorPoli { get; set; }
        public int Red { get; set; }
        public int Green { get; set; }
        public int Blue { get; set; }
    }
}