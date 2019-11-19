namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class hasParameters : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.generalSettings", "hasParameters", c => c.Boolean(nullable: false));
        }
        
        public override void Down()
        {
            DropColumn("dbo.generalSettings", "hasParameters");
        }
    }
}
