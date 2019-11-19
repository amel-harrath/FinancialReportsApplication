namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class RF : DbMigration
    {
        public override void Up()
        {
            DropPrimaryKey("dbo.ResultatFiscalParamModels");
            AddColumn("dbo.ResultatFiscalFormulas", "exercice", c => c.Int(nullable: false));
            AddColumn("dbo.ResultatFiscalFormulas", "matricule", c => c.String());
            AddColumn("dbo.ResultatFiscalParamModels", "exercice", c => c.Int(nullable: false));
            AddColumn("dbo.ResultatFiscalParamModels", "matricule", c => c.String(nullable: false, maxLength: 128));
            AddColumn("dbo.ResultatFiscalParamModels", "state", c => c.String());
            AddPrimaryKey("dbo.ResultatFiscalParamModels", new[] { "code", "exercice", "matricule" });
        }
        
        public override void Down()
        {
            DropPrimaryKey("dbo.ResultatFiscalParamModels");
            DropColumn("dbo.ResultatFiscalParamModels", "state");
            DropColumn("dbo.ResultatFiscalParamModels", "matricule");
            DropColumn("dbo.ResultatFiscalParamModels", "exercice");
            DropColumn("dbo.ResultatFiscalFormulas", "matricule");
            DropColumn("dbo.ResultatFiscalFormulas", "exercice");
            AddPrimaryKey("dbo.ResultatFiscalParamModels", "code");
        }
    }
}
