namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class ActifParam1 : DbMigration
    {
        public override void Up()
        {
            DropPrimaryKey("dbo.ActifParamModels");
            AlterColumn("dbo.ActifParamModels", "matricule", c => c.String(nullable: false, maxLength: 128));
            AddPrimaryKey("dbo.ActifParamModels", new[] { "code", "exercice", "matricule" });
        }
        
        public override void Down()
        {
            DropPrimaryKey("dbo.ActifParamModels");
            AlterColumn("dbo.ActifParamModels", "matricule", c => c.String());
            AddPrimaryKey("dbo.ActifParamModels", "code");
        }
    }
}
