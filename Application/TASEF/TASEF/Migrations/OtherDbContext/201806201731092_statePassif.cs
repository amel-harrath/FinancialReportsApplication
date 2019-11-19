namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class statePassif : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.PassifParamModels", "state", c => c.String());
        }
        
        public override void Down()
        {
            DropColumn("dbo.PassifParamModels", "state");
        }
    }
}
