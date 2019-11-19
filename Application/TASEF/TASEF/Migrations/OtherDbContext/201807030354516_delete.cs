namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class delete : DbMigration
    {
        public override void Up()
        {
            DropTable("dbo.Formulae");
        }
        
        public override void Down()
        {
        }
    }
}
