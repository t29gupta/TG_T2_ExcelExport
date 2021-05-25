using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace TG_Reporting.Model
{
    public class TG_Hotel
    {
        [JsonPropertyName("hotel")]
        public Hotel Hotel { get; set; }

        [JsonPropertyName("hotelRates")]
        public List<HotelRate> HotelRates { get; set; }
    }
    // Root myDeserializedClass = JsonSerializer.Deserialize<Root>(myJsonResponse);
    public class Hotel
    {
        [JsonPropertyName("hotelID")]
        public int HotelID { get; set; }

        [JsonPropertyName("classification")]
        public int Classification { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("reviewscore")]
        public double Reviewscore { get; set; }
    }

    public class Price
    {
        [JsonPropertyName("currency")]
        public string Currency { get; set; }

        [JsonPropertyName("numericFloat")]
        public decimal NumericFloat { get; set; }

        [JsonPropertyName("numericInteger")]
        public int NumericInteger { get; set; }
    }

    public class RateTag
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("shape")]
        public bool Shape { get; set; }
    }

    public class HotelRate
    {
        [JsonPropertyName("adults")]
        public int Adults { get; set; }

        [JsonPropertyName("los")]
        public int Los { get; set; }

        [JsonPropertyName("price")]
        public Price Price { get; set; }

        [JsonPropertyName("rateDescription")]
        public string RateDescription { get; set; }

        [JsonPropertyName("rateID")]
        public string RateID { get; set; }

        [JsonPropertyName("rateName")]
        public string RateName { get; set; }

        [JsonPropertyName("rateTags")]
        public List<RateTag> RateTags { get; set; }

        [JsonPropertyName("targetDay")]
        public DateTime TargetDay { get; set; }
    }

}
