namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class AddingDefinedFormulaTable : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.Formulae",
                c => new
                {
                    Id = c.Int(nullable: false, identity: true),
                   
                    code = c.String(nullable: false),
                    type = c.String(nullable: false),
                    parameter = c.String(nullable: false),
                })
                .PrimaryKey(t => t.Id);
        }
        
        public override void Down()
        {
            DropTable("dbo.Formulae");

        }
    }
}
