namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class journalRAN : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.FluxTresorerieMAFormulas", "RANjournal", c => c.String(nullable: false));
            AddColumn("dbo.FluxTresorerieMRFormulas", "RANjournal", c => c.String(nullable: false));
        }
        
        public override void Down()
        {
            DropColumn("dbo.FluxTresorerieMRFormulas", "RANjournal");
            DropColumn("dbo.FluxTresorerieMAFormulas", "RANjournal");
        }
    }
}
