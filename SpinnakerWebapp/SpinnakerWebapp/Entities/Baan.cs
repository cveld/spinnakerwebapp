namespace SpinnakerWebapp.Entities
{
    public class Baan
    {
        public string Windrichting { get; set; }
        public List<Boei> Boeien { get; set; } = new List<Boei>();
    }
}
