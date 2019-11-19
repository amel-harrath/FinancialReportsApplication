using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace TASEF.Models
{
    public class FluxTresorerieMRParamModel
    {
        [Key]
        [Column(Order = 1)]
        [Display(Name = "Code")]
        public string code { get; set; }

        [Display(Name = "Libelle")]
        public string libelle { get; set; }

        [Display(Name = "Net N ")]
        public float netN { get; set; }

        [Display(Name = "Net N-1 ")]
        public float netN1 { get; set; }

        public string type { get; set; }

        [Key]
        [Column(Order = 2)]
        public int exercice { get; set; }

        [Key]
        [Column(Order = 3)]
        public string matricule { get; set; }

        [Key]
        [Column(Order = 0)]
        public string ownerId { get; set; }

        public string state { get; set; }

        public int priority { get; set; }
    }
}