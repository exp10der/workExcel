using System;

namespace lol
{
  public  class Import
    {
        public string ObjectBilder { get; set; }
        public string K { get; set; }
       // public Status Status { get; set; }
        public string Status { get; set; }
        public double Area { get; set; }
        public decimal PriceMeter { get; set; }
        public decimal PriceApartment { get; set; }
        public int? CountDayArmor { get; set; }
        public DateTime? DayArmor { get; set; }
        public int Access { get; set; }
        public int Floor { get; set; }
        public int LevelRoom { get; set; }
        public string Room { get; set; }
    }
}
public enum Status : byte
{
    Free, Reservations
}
