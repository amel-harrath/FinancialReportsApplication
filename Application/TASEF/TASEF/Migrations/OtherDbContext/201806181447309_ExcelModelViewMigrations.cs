namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class ExcelModelViewMigrations : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.ExcelModelViews",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        matricule = c.String(),
                        exercice = c.Int(nullable: false),
                        firstPeriodStart = c.DateTime(nullable: false),
                        firstPeriodEnd = c.DateTime(nullable: false),
                        firstfile = c.String(nullable: false),
                        secondPeriodStart = c.DateTime(nullable: false),
                        secondPeriodEnd = c.DateTime(nullable: false),
                        secondfile = c.String(nullable: false),
                    })
                .PrimaryKey(t => t.Id);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.ExcelModelViews");
        }
    }
}
