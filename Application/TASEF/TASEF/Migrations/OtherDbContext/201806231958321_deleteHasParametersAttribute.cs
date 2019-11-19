namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class deleteHasParametersAttribute : DbMigration
    {
        public override void Up()
        {
            DropColumn("dbo.generalSettings", "hasParameters");
        }
        
        public override void Down()
        {
            AddColumn("dbo.generalSettings", "hasParameters", c => c.Boolean(nullable: false));
        }
    }
}
