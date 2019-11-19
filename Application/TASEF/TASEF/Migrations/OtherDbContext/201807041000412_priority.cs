namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class priority : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.ActifParamModels", "priority", c => c.Int(nullable: false));
            AddColumn("dbo.EtatDeResultatParamModels", "priority", c => c.Int(nullable: false));
            AddColumn("dbo.FluxTresorerieMAParamModels", "priority", c => c.Int(nullable: false));
            AddColumn("dbo.FluxTresorerieMRParamModels", "priority", c => c.Int(nullable: false));
            AddColumn("dbo.PassifParamModels", "priority", c => c.Int(nullable: false));
            AddColumn("dbo.ResultatFiscalParamModels", "priority", c => c.Int(nullable: false));
        }
        
        public override void Down()
        {
            DropColumn("dbo.ResultatFiscalParamModels", "priority");
            DropColumn("dbo.PassifParamModels", "priority");
            DropColumn("dbo.FluxTresorerieMRParamModels", "priority");
            DropColumn("dbo.FluxTresorerieMAParamModels", "priority");
            DropColumn("dbo.EtatDeResultatParamModels", "priority");
            DropColumn("dbo.ActifParamModels", "priority");
        }
    }
}
