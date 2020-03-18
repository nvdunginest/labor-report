using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace LaborReport.Datas.Models
{
    public class InOut
    {
        public Guid Id { get; set; }
        public DateTime Time { get; set; }
        public string CardNumber { get; set; }
        public string Event { get; set; }
        public string Status { get; set; }
        public string Description { get; set; }
    }
}
