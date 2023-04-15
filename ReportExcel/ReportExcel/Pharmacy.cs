
namespace ReportExcel
{
    class Pharmacy
    {
        public string Name;
        public string Price;

        public Pharmacy()
        {
            Name = null;
            Price = null;
        }

        public Pharmacy(string name, string price)
        {
            Name = name;
            Price = price;
        }
    }
}
