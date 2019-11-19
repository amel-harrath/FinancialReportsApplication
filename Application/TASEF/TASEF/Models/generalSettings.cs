using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace TASEF.Models
{
    public class generalSettings
    {
        [Key]
        [Column(Order =0)]
        [Required]
        public String ownerId { get; set; }

        [Key]
        [Column(Order = 1)]
        [Required]
        [Display(Name = "Matricule")]
        public String matricule { get; set; }

        [Required]
        [Display(Name = "Nom et prenom raison sociale")]
        public String nomEtPrenomRaisonSociale { get; set; }

        [Required]
        [Display(Name = "Activité")]
        public String activite { get; set; }

        [Required]
        [Display(Name = "Adresse")]
        public String adresse { get; set; }

        [Key]
        [Column(Order = 2)]
        [Required]
        [Display(Name = "Exercice")]
        public int exercice { get; set; }

        [Required]
        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
        [Display(Name = "Date debut exercice")]
        public DateTime dateDebutExercice { get; set; }

        [Required]
        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
        [Display(Name = "Date cloture exercice")]
        public DateTime dateClotureExercice { get; set; }


        [Required]
        [Display(Name = "Acte de depot")]
        public String actededepot { get; set; }

        [Required]
        [Display(Name = "Nature depot")]
        public String natureDepot { get; set; }

        


    }
}