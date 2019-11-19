namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class testing : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.Formulae",
                c => new
                {
                    Id = c.String(nullable: false, maxLength: 128),
                    code = c.String(),
                    type = c.String(),
                })
                .PrimaryKey(t => t.Id);
            DropTable("dbo.Formulae");
        }
        
        public override void Down()
        {
            CreateTable(
                "dbo.Formulae",
                c => new
                    {
                        Id = c.String(nullable: false, maxLength: 128),
                        code = c.String(),
                        type = c.String(),
                    })
                .PrimaryKey(t => t.Id);
            
        }
    }
}
