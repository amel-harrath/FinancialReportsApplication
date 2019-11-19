namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class etatdeFlux : DbMigration
    {
        public override void Up()
        {
            DropPrimaryKey("dbo.FluxTresorerieMAParamModels");
            DropPrimaryKey("dbo.FluxTresorerieMRParamModels");
            AddColumn("dbo.FluxTresorerieMAFormulas", "exercice", c => c.Int(nullable: false));
            AddColumn("dbo.FluxTresorerieMAFormulas", "matricule", c => c.String());
            AddColumn("dbo.FluxTresorerieMAParamModels", "exercice", c => c.Int(nullable: false));
            AddColumn("dbo.FluxTresorerieMAParamModels", "matricule", c => c.String(nullable: false, maxLength: 128));
            AddColumn("dbo.FluxTresorerieMAParamModels", "state", c => c.String());
            AddColumn("dbo.FluxTresorerieMRFormulas", "exercice", c => c.Int(nullable: false));
            AddColumn("dbo.FluxTresorerieMRFormulas", "matricule", c => c.String());
            AddColumn("dbo.FluxTresorerieMRParamModels", "exercice", c => c.Int(nullable: false));
            AddColumn("dbo.FluxTresorerieMRParamModels", "matricule", c => c.String(nullable: false, maxLength: 128));
            AddColumn("dbo.FluxTresorerieMRParamModels", "state", c => c.String());
            AddPrimaryKey("dbo.FluxTresorerieMAParamModels", new[] { "code", "exercice", "matricule" });
            AddPrimaryKey("dbo.FluxTresorerieMRParamModels", new[] { "code", "exercice", "matricule" });
        }
        
        public override void Down()
        {
            DropPrimaryKey("dbo.FluxTresorerieMRParamModels");
            DropPrimaryKey("dbo.FluxTresorerieMAParamModels");
            DropColumn("dbo.FluxTresorerieMRParamModels", "state");
            DropColumn("dbo.FluxTresorerieMRParamModels", "matricule");
            DropColumn("dbo.FluxTresorerieMRParamModels", "exercice");
            DropColumn("dbo.FluxTresorerieMRFormulas", "matricule");
            DropColumn("dbo.FluxTresorerieMRFormulas", "exercice");
            DropColumn("dbo.FluxTresorerieMAParamModels", "state");
            DropColumn("dbo.FluxTresorerieMAParamModels", "matricule");
            DropColumn("dbo.FluxTresorerieMAParamModels", "exercice");
            DropColumn("dbo.FluxTresorerieMAFormulas", "matricule");
            DropColumn("dbo.FluxTresorerieMAFormulas", "exercice");
            AddPrimaryKey("dbo.FluxTresorerieMRParamModels", "code");
            AddPrimaryKey("dbo.FluxTresorerieMAParamModels", "code");
        }
    }
}
