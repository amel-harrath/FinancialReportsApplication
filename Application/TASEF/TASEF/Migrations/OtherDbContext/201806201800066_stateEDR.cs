namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class stateEDR : DbMigration
    {
        public override void Up()
        {
            DropPrimaryKey("dbo.EtatDeResultatParamModels");
            AddColumn("dbo.EtatDeResultatFormulas", "exercice", c => c.Int(nullable: false));
            AddColumn("dbo.EtatDeResultatFormulas", "matricule", c => c.String());
            AddColumn("dbo.EtatDeResultatParamModels", "exercice", c => c.Int(nullable: false));
            AddColumn("dbo.EtatDeResultatParamModels", "matricule", c => c.String(nullable: false, maxLength: 128));
            AddColumn("dbo.EtatDeResultatParamModels", "state", c => c.String());
            AddPrimaryKey("dbo.EtatDeResultatParamModels", new[] { "code", "exercice", "matricule" });
        }
        
        public override void Down()
        {
            DropPrimaryKey("dbo.EtatDeResultatParamModels");
            DropColumn("dbo.EtatDeResultatParamModels", "state");
            DropColumn("dbo.EtatDeResultatParamModels", "matricule");
            DropColumn("dbo.EtatDeResultatParamModels", "exercice");
            DropColumn("dbo.EtatDeResultatFormulas", "matricule");
            DropColumn("dbo.EtatDeResultatFormulas", "exercice");
            AddPrimaryKey("dbo.EtatDeResultatParamModels", "code");
        }
    }
}
