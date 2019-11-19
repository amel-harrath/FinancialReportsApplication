using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;
using TASEF.Models;

namespace TASEF.Infrastructure
{
    public class ExcelProjectContext : DbContext
    {
        public ExcelProjectContext() : base("name=ExcelProjectContext")
        {
        }

        public virtual DbSet<ActifParamModel> ActifModel { get; set; }
        public virtual DbSet<ActifFormula> ActifFormula { get; set; }

        public virtual DbSet<PassifParamModel> PassifModel { get; set; }
        public virtual DbSet<PassifFormula> PassifFormula { get; set; }

        public virtual DbSet<EtatDeResultatParamModel> EtatDeResultatModel { get; set; }
        public virtual DbSet<EtatDeResultatFormula> EtatDeResultatFormula { get; set; }

        public virtual DbSet<FluxTresorerieMAParamModel> FluxTresorerieMAModel { get; set; }
        public virtual DbSet<FluxTresorerieMAFormula> FluxTresorerieMAFormula { get; set; }

        public virtual DbSet<FluxTresorerieMRParamModel> FluxTresorerieMRModel { get; set; }
        public virtual DbSet<FluxTresorerieMRFormula> FluxTresorerieMRFormula { get; set; }

        public virtual DbSet<ResultatFiscalParamModel> ResultatFiscalModel { get; set; }
        public virtual DbSet<ResultatFiscalFormula> ResultatFiscalFormula { get; set; }

        public virtual DbSet<generalSettings> GeneralSettings { get; set; }

        public virtual DbSet<ExcelModelView> ExcelModelViews { get; set; }

        public virtual DbSet<ParametersSetting> ParametersSetting { get; set; }

        public virtual DbSet<Formula> DefinedFormulas { get; set; }

    }
}