using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace TASEF.Models
{
    public class ExcelModelView
    {
        [Key]
        public int Id { get; set; }

        public string matricule { get; set; }

        public int exercice { get; set; }

        public string ownerId { get; set; }

        [Required]
        [Display ( Name = "Début de la première période")]
        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
        public DateTime firstPeriodStart { get; set; }

        [Required]
        [Display(Name = "Fin de la première période")]
        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
        public DateTime firstPeriodEnd { get; set; }

        [Required]
        [Display(Name = "Fichier de la premier période")]
        public string firstfile { get; set; }

        [Required]
        [Display(Name = "Début de la deuxième période")]
        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
        public DateTime secondPeriodStart { get; set; }

        [Required]
        [Display(Name = "Fin de la deuxième période")]
        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
        public DateTime secondPeriodEnd { get; set; }

        [Required]
        [Display(Name = "Fichier de la deuxième période")]
        public string secondfile { get; set; }

        public string journaleRAN { get; set; }

    }
}