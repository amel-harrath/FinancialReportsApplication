namespace TASEF.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class newTable : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.UserProfileInfoes",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        FirstName = c.String(),
                    })
                .PrimaryKey(t => t.Id);
            
            AddColumn("dbo.AspNetUsers", "userProfileInfo_Id", c => c.Int());
            CreateIndex("dbo.AspNetUsers", "userProfileInfo_Id");
            AddForeignKey("dbo.AspNetUsers", "userProfileInfo_Id", "dbo.UserProfileInfoes", "Id");
            DropColumn("dbo.AspNetUsers", "FirstName");
        }
        
        public override void Down()
        {
            AddColumn("dbo.AspNetUsers", "FirstName", c => c.String());
            DropForeignKey("dbo.AspNetUsers", "userProfileInfo_Id", "dbo.UserProfileInfoes");
            DropIndex("dbo.AspNetUsers", new[] { "userProfileInfo_Id" });
            DropColumn("dbo.AspNetUsers", "userProfileInfo_Id");
            DropTable("dbo.UserProfileInfoes");
        }
    }
}
