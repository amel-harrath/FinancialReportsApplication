namespace TASEF.Migrations.OtherDbContext
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class AddParamSettingModel : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.ParametersSettings",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        ownerId = c.String(),
                        matricule = c.String(),
                        exercice = c.Int(nullable: false),
                        hasParamActif = c.Boolean(nullable: false),
                        hasParamPassif = c.Boolean(nullable: false),
                        hasParamFluxMA = c.Boolean(nullable: false),
                        hasParamFluxMR = c.Boolean(nullable: false),
                        hasParamRes = c.Boolean(nullable: false),
                        hasParamEtatDeRes = c.Boolean(nullable: false),
                    })
                .PrimaryKey(t => t.Id);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.ParametersSettings");
        }
    }
}
