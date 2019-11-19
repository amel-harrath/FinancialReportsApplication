using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace TASEF.Models
{
    public class ExcelInfo
    {
        public int periode { get; set; }

        [Display(Name = "Journale")]
        public string journal { get; set; }

        [Display(Name = "Compte")]
        public string compte { get; set; }

        [Display(Name = "Date Ecriture")]
        public DateTime dateEcriture { get; set; }

        [Display(Name = "Débit")]
        public float debit { get; set; }

        [Display(Name = "Crédit")]
        public float credit { get; set; }

    }
}