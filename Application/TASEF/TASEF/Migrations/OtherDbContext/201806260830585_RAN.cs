namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class RAN : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.ExcelModelViews", "journaleRAN", c => c.String());
        }
        
        public override void Down()
        {
            DropColumn("dbo.ExcelModelViews", "journaleRAN");
        }
    }
}
