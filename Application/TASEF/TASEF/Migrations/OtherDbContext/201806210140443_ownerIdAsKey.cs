namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class ownerIdAsKey : DbMigration
    {
        public override void Up()
        {
            DropPrimaryKey("dbo.generalSettings");
            AlterColumn("dbo.generalSettings", "ownerId", c => c.String(nullable: false, maxLength: 128));
            AddPrimaryKey("dbo.generalSettings", new[] { "ownerId", "matricule", "exercice" });
        }
        
        public override void Down()
        {
            DropPrimaryKey("dbo.generalSettings");
            AlterColumn("dbo.generalSettings", "ownerId", c => c.String(nullable: false));
            AddPrimaryKey("dbo.generalSettings", new[] { "matricule", "exercice" });
        }
    }
}
