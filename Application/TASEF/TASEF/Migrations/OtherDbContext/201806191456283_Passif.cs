namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class Passif : DbMigration
    {
        public override void Up()
        {
            DropPrimaryKey("dbo.PassifParamModels");
            AddColumn("dbo.PassifFormulas", "exercice", c => c.Int(nullable: false));
            AddColumn("dbo.PassifFormulas", "matricule", c => c.String());
            AddColumn("dbo.PassifParamModels", "exercice", c => c.Int(nullable: false));
            AddColumn("dbo.PassifParamModels", "matricule", c => c.String(nullable: false, maxLength: 128));
            AddPrimaryKey("dbo.PassifParamModels", new[] { "code", "exercice", "matricule" });
        }
        
        public override void Down()
        {
            DropPrimaryKey("dbo.PassifParamModels");
            DropColumn("dbo.PassifParamModels", "matricule");
            DropColumn("dbo.PassifParamModels", "exercice");
            DropColumn("dbo.PassifFormulas", "matricule");
            DropColumn("dbo.PassifFormulas", "exercice");
            AddPrimaryKey("dbo.PassifParamModels", "code");
        }
    }
}
