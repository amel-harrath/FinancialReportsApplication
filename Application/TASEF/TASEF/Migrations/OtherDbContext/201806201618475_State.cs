namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class State : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.ActifParamModels", "state", c => c.String());
        }
        
        public override void Down()
        {
            DropColumn("dbo.ActifParamModels", "state");
        }
    }
}
