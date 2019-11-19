namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class ActifParam : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.ActifParamModels", "exercice", c => c.Int(nullable: false));
            AddColumn("dbo.ActifParamModels", "matricule", c => c.String());
        }
        
        public override void Down()
        {
            DropColumn("dbo.ActifParamModels", "matricule");
            DropColumn("dbo.ActifParamModels", "exercice");
        }
    }
}
