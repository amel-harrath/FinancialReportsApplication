namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class fromIntToFloat : DbMigration
    {
        public override void Up()
        {
            AlterColumn("dbo.ActifParamModels", "brutN", c => c.Single(nullable: false));
            AlterColumn("dbo.ActifParamModels", "amortProvN", c => c.Single(nullable: false));
            AlterColumn("dbo.ActifParamModels", "netN", c => c.Single(nullable: false));
            AlterColumn("dbo.ActifParamModels", "netN1", c => c.Single(nullable: false));
            AlterColumn("dbo.EtatDeResultatParamModels", "netN", c => c.Single(nullable: false));
            AlterColumn("dbo.EtatDeResultatParamModels", "netN1", c => c.Single(nullable: false));
            AlterColumn("dbo.FluxTresorerieMAParamModels", "netN", c => c.Single(nullable: false));
            AlterColumn("dbo.FluxTresorerieMAParamModels", "netN1", c => c.Single(nullable: false));
            AlterColumn("dbo.FluxTresorerieMRParamModels", "netN", c => c.Single(nullable: false));
            AlterColumn("dbo.FluxTresorerieMRParamModels", "netN1", c => c.Single(nullable: false));
            AlterColumn("dbo.PassifParamModels", "netN", c => c.Single(nullable: false));
            AlterColumn("dbo.PassifParamModels", "netN1", c => c.Single(nullable: false));
            AlterColumn("dbo.ResultatFiscalParamModels", "netN", c => c.Single(nullable: false));
            AlterColumn("dbo.ResultatFiscalParamModels", "netN1", c => c.Single(nullable: false));
        }
        
        public override void Down()
        {
            AlterColumn("dbo.ResultatFiscalParamModels", "netN1", c => c.Int(nullable: false));
            AlterColumn("dbo.ResultatFiscalParamModels", "netN", c => c.Int(nullable: false));
            AlterColumn("dbo.PassifParamModels", "netN1", c => c.Int(nullable: false));
            AlterColumn("dbo.PassifParamModels", "netN", c => c.Int(nullable: false));
            AlterColumn("dbo.FluxTresorerieMRParamModels", "netN1", c => c.Int(nullable: false));
            AlterColumn("dbo.FluxTresorerieMRParamModels", "netN", c => c.Int(nullable: false));
            AlterColumn("dbo.FluxTresorerieMAParamModels", "netN1", c => c.Int(nullable: false));
            AlterColumn("dbo.FluxTresorerieMAParamModels", "netN", c => c.Int(nullable: false));
            AlterColumn("dbo.EtatDeResultatParamModels", "netN1", c => c.Int(nullable: false));
            AlterColumn("dbo.EtatDeResultatParamModels", "netN", c => c.Int(nullable: false));
            AlterColumn("dbo.ActifParamModels", "netN1", c => c.Int(nullable: false));
            AlterColumn("dbo.ActifParamModels", "netN", c => c.Int(nullable: false));
            AlterColumn("dbo.ActifParamModels", "amortProvN", c => c.Int(nullable: false));
            AlterColumn("dbo.ActifParamModels", "brutN", c => c.Int(nullable: false));
        }
    }
}
