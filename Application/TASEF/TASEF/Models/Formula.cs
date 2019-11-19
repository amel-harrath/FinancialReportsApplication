using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace TASEF.Models
{
    public class Formula
    {
        [Key]
        public int Id { get; set; }

        [Required]
        public string code { get; set; }

        [Required]
        public string type { get; set; }

        [Required]
        public string parameter { get; set; }

    }
}