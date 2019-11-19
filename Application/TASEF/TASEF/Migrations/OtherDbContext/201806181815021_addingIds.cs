namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class addingIds : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.ActifFormulas", "exercice", c => c.Int(nullable: false));
            AddColumn("dbo.ActifFormulas", "matricule", c => c.String());
        }
        
        public override void Down()
        {
            DropColumn("dbo.ActifFormulas", "matricule");
            DropColumn("dbo.ActifFormulas", "exercice");
        }
    }
}
