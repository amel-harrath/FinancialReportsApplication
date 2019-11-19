using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace TASEF.Models
{
    public class ParametersSetting
    {
        [Key]
        public int Id { get; set; }

        public string ownerId { get; set; }
        public string matricule { get; set; }
        public int exercice { get; set; }

        public bool hasParamActif { get; set; } = false;
        public bool hasParamPassif { get; set; } = false;
        public bool hasParamFluxMA { get; set; } = false;
        public bool hasParamFluxMR { get; set; } = false;
        public bool hasParamRes { get; set; } = false;
        public bool hasParamEtatDeRes { get; set; } = false;
    }
}