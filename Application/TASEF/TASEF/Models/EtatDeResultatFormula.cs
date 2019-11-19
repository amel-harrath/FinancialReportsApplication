using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace TASEF.Models
{
    public class EtatDeResultatFormula
    {
        [Key]
        public int Id { get; set; }

        public int exercice { get; set; }

        public string matricule { get; set; }

        public string ownerId { get; set; }

        [Display(Name = "code Parameters")]
        public string codeParam { get; set; }

        [Required]
        [Display(Name = "Compte")]
        public string codeDonnee { get; set; }

        [Required]
        [Display(Name = "Nom du Compte")]
        public string nomCompte { get; set; }

        //type means if it is Solde, Solde débiteur, Solde créditeur, ... 
        [Required]
        [Display(Name = "type Formule")]
        public string typeFormule { get; set; }
    }
}