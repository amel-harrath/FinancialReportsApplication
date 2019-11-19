namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class OwnerIdMigrations : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.ActifFormulas",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        codeParam = c.String(),
                        codeDonnee = c.String(nullable: false),
                        nomCompte = c.String(nullable: false),
                        typeFormule = c.String(nullable: false),
                        genre = c.String(nullable: false),
                    })
                .PrimaryKey(t => t.Id);
            
            CreateTable(
                "dbo.ActifParamModels",
                c => new
                    {
                        code = c.String(nullable: false, maxLength: 128),
                        libelle = c.String(),
                        brutN = c.Int(nullable: false),
                        amortProvN = c.Int(nullable: false),
                        netN = c.Int(nullable: false),
                        netN1 = c.Int(nullable: false),
                        type = c.String(),
                    })
                .PrimaryKey(t => t.code);
            
            CreateTable(
                "dbo.EtatDeResultatFormulas",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        codeParam = c.String(),
                        codeDonnee = c.String(nullable: false),
                        nomCompte = c.String(nullable: false),
                        typeFormule = c.String(nullable: false),
                    })
                .PrimaryKey(t => t.Id);
            
            CreateTable(
                "dbo.EtatDeResultatParamModels",
                c => new
                    {
                        code = c.String(nullable: false, maxLength: 128),
                        libelle = c.String(),
                        netN = c.Int(nullable: false),
                        netN1 = c.Int(nullable: false),
                        type = c.String(),
                    })
                .PrimaryKey(t => t.code);
            
            CreateTable(
                "dbo.FluxTresorerieMAFormulas",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        codeParam = c.String(),
                        codeDonnee = c.String(nullable: false),
                        nomCompte = c.String(nullable: false),
                        typeFormule = c.String(nullable: false),
                    })
                .PrimaryKey(t => t.Id);
            
            CreateTable(
                "dbo.FluxTresorerieMAParamModels",
                c => new
                    {
                        code = c.String(nullable: false, maxLength: 128),
                        libelle = c.String(),
                        netN = c.Int(nullable: false),
                        netN1 = c.Int(nullable: false),
                        type = c.String(),
                    })
                .PrimaryKey(t => t.code);
            
            CreateTable(
                "dbo.FluxTresorerieMRFormulas",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        codeParam = c.String(),
                        codeDonnee = c.String(nullable: false),
                        nomCompte = c.String(nullable: false),
                        typeFormule = c.String(nullable: false),
                    })
                .PrimaryKey(t => t.Id);
            
            CreateTable(
                "dbo.FluxTresorerieMRParamModels",
                c => new
                    {
                        code = c.String(nullable: false, maxLength: 128),
                        libelle = c.String(),
                        netN = c.Int(nullable: false),
                        netN1 = c.Int(nullable: false),
                        type = c.String(),
                    })
                .PrimaryKey(t => t.code);
            
            CreateTable(
                "dbo.generalSettings",
                c => new
                    {
                        matricule = c.String(nullable: false, maxLength: 128),
                        exercice = c.Int(nullable: false),
                        ownerId = c.String(nullable: false),
                        nomEtPrenomRaisonSociale = c.String(nullable: false),
                        activite = c.String(nullable: false),
                        adresse = c.String(nullable: false),
                        dateDebutExercice = c.DateTime(nullable: false),
                        dateClotureExercice = c.DateTime(nullable: false),
                        actededepot = c.String(nullable: false),
                        natureDepot = c.String(nullable: false),
                    })
                .PrimaryKey(t => new { t.matricule, t.exercice });
            
            CreateTable(
                "dbo.PassifFormulas",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        codeParam = c.String(),
                        codeDonnee = c.String(nullable: false),
                        nomCompte = c.String(nullable: false),
                        typeFormule = c.String(nullable: false),
                    })
                .PrimaryKey(t => t.Id);
            
            CreateTable(
                "dbo.PassifParamModels",
                c => new
                    {
                        code = c.String(nullable: false, maxLength: 128),
                        libelle = c.String(),
                        netN = c.Int(nullable: false),
                        netN1 = c.Int(nullable: false),
                        type = c.String(),
                    })
                .PrimaryKey(t => t.code);
            
            CreateTable(
                "dbo.ResultatFiscalFormulas",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        codeParam = c.String(),
                        codeDonnee = c.String(nullable: false),
                        nomCompte = c.String(nullable: false),
                        typeFormule = c.String(nullable: false),
                    })
                .PrimaryKey(t => t.Id);
            
            CreateTable(
                "dbo.ResultatFiscalParamModels",
                c => new
                    {
                        code = c.String(nullable: false, maxLength: 128),
                        libelle = c.String(),
                        netN = c.Int(nullable: false),
                        netN1 = c.Int(nullable: false),
                        type = c.String(),
                    })
                .PrimaryKey(t => t.code);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.ResultatFiscalParamModels");
            DropTable("dbo.ResultatFiscalFormulas");
            DropTable("dbo.PassifParamModels");
            DropTable("dbo.PassifFormulas");
            DropTable("dbo.generalSettings");
            DropTable("dbo.FluxTresorerieMRParamModels");
            DropTable("dbo.FluxTresorerieMRFormulas");
            DropTable("dbo.FluxTresorerieMAParamModels");
            DropTable("dbo.FluxTresorerieMAFormulas");
            DropTable("dbo.EtatDeResultatParamModels");
            DropTable("dbo.EtatDeResultatFormulas");
            DropTable("dbo.ActifParamModels");
            DropTable("dbo.ActifFormulas");
        }
    }
}
