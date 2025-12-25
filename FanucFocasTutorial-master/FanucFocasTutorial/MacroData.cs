using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace FanucFocasTutorial
{
    [Table("MacroData")]
    public class MacroData
    {
        [Key, Column(Order = 0)]
        [StringLength(50)]
        public string IpAddress { get; set; }

        [Key, Column(Order = 1)]
        public short MacroNumber { get; set; }

        public double Value { get; set; }

        public DateTime LastUpdated { get; set; }

        public MacroData()
        {
            LastUpdated = DateTime.Now;
        }
    }
}
