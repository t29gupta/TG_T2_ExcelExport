using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TG_Reporting.Model
{
    public class ReportModel
    {
        public DateTime Arrival_Date { get; set; }

        public DateTime Departure_Date { get; set; }

        public decimal Price { get; set; }

        public string Currency { get; set; }

        public string RateName { get; set; }

        public int Adults { get; set; }

        public int Breakfast_Included { get; set; }
    }
}
