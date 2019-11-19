namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class AddedOwnerIdEverywhere : DbMigration
    {
        public override void Up()
        {
            DropPrimaryKey("dbo.ActifParamModels");
            DropPrimaryKey("dbo.EtatDeResultatParamModels");
            DropPrimaryKey("dbo.FluxTresorerieMAParamModels");
            DropPrimaryKey("dbo.FluxTresorerieMRParamModels");
            DropPrimaryKey("dbo.PassifParamModels");
            DropPrimaryKey("dbo.ResultatFiscalParamModels");
            AddColumn("dbo.ActifFormulas", "ownerId", c => c.String());
            AddColumn("dbo.ActifParamModels", "ownerId", c => c.String(nullable: false, maxLength: 128));
            AddColumn("dbo.EtatDeResultatFormulas", "ownerId", c => c.String());
            AddColumn("dbo.EtatDeResultatParamModels", "ownerId", c => c.String(nullable: false, maxLength: 128));
            AddColumn("dbo.ExcelModelViews", "ownerId", c => c.String());
            AddColumn("dbo.FluxTresorerieMAFormulas", "ownerId", c => c.String());
            AddColumn("dbo.FluxTresorerieMAParamModels", "ownerId", c => c.String(nullable: false, maxLength: 128));
            AddColumn("dbo.FluxTresorerieMRFormulas", "ownerId", c => c.String());
            AddColumn("dbo.FluxTresorerieMRParamModels", "ownerId", c => c.String(nullable: false, maxLength: 128));
            AddColumn("dbo.PassifFormulas", "ownerId", c => c.String());
            AddColumn("dbo.PassifParamModels", "ownerId", c => c.String(nullable: false, maxLength: 128));
            AddColumn("dbo.ResultatFiscalFormulas", "ownerId", c => c.String());
            AddColumn("dbo.ResultatFiscalParamModels", "ownerId", c => c.String(nullable: false, maxLength: 128));
            AddPrimaryKey("dbo.ActifParamModels", new[] { "ownerId", "code", "exercice", "matricule" });
            AddPrimaryKey("dbo.EtatDeResultatParamModels", new[] { "ownerId", "code", "exercice", "matricule" });
            AddPrimaryKey("dbo.FluxTresorerieMAParamModels", new[] { "ownerId", "code", "exercice", "matricule" });
            AddPrimaryKey("dbo.FluxTresorerieMRParamModels", new[] { "ownerId", "code", "exercice", "matricule" });
            AddPrimaryKey("dbo.PassifParamModels", new[] { "ownerId", "code", "exercice", "matricule" });
            AddPrimaryKey("dbo.ResultatFiscalParamModels", new[] { "ownerId", "code", "exercice", "matricule" });
        }
        
        public override void Down()
        {
            DropPrimaryKey("dbo.ResultatFiscalParamModels");
            DropPrimaryKey("dbo.PassifParamModels");
            DropPrimaryKey("dbo.FluxTresorerieMRParamModels");
            DropPrimaryKey("dbo.FluxTresorerieMAParamModels");
            DropPrimaryKey("dbo.EtatDeResultatParamModels");
            DropPrimaryKey("dbo.ActifParamModels");
            DropColumn("dbo.ResultatFiscalParamModels", "ownerId");
            DropColumn("dbo.ResultatFiscalFormulas", "ownerId");
            DropColumn("dbo.PassifParamModels", "ownerId");
            DropColumn("dbo.PassifFormulas", "ownerId");
            DropColumn("dbo.FluxTresorerieMRParamModels", "ownerId");
            DropColumn("dbo.FluxTresorerieMRFormulas", "ownerId");
            DropColumn("dbo.FluxTresorerieMAParamModels", "ownerId");
            DropColumn("dbo.FluxTresorerieMAFormulas", "ownerId");
            DropColumn("dbo.ExcelModelViews", "ownerId");
            DropColumn("dbo.EtatDeResultatParamModels", "ownerId");
            DropColumn("dbo.EtatDeResultatFormulas", "ownerId");
            DropColumn("dbo.ActifParamModels", "ownerId");
            DropColumn("dbo.ActifFormulas", "ownerId");
            AddPrimaryKey("dbo.ResultatFiscalParamModels", new[] { "code", "exercice", "matricule" });
            AddPrimaryKey("dbo.PassifParamModels", new[] { "code", "exercice", "matricule" });
            AddPrimaryKey("dbo.FluxTresorerieMRParamModels", new[] { "code", "exercice", "matricule" });
            AddPrimaryKey("dbo.FluxTresorerieMAParamModels", new[] { "code", "exercice", "matricule" });
            AddPrimaryKey("dbo.EtatDeResultatParamModels", new[] { "code", "exercice", "matricule" });
            AddPrimaryKey("dbo.ActifParamModels", new[] { "code", "exercice", "matricule" });
        }
    }
}
