using Microsoft.AspNet.Identity;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web;
using System.Web.Mvc;
using TASEF.Infrastructure;
using TASEF.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace TASEF.Controllers
{
    public class ExcelController : Controller
    {
        private ExcelProjectContext db = new ExcelProjectContext();

        // GET: Excel
        [Authorize]
        public ActionResult Index()
        {
            Session.Clear();
            return View();
        }

        [Authorize]
        public ActionResult NewFiles(string Id)
        {
            generalSettings gs = (generalSettings)Session["SteInformation"];
            ViewBag.fps = String.Format("{0:yyyy-MM-dd}", gs.dateDebutExercice.AddYears(-1));
            ViewBag.fpe = String.Format("{0:yyyy-MM-dd}", gs.dateClotureExercice.AddYears(-1));
            ViewBag.sps = String.Format("{0:yyyy-MM-dd}", gs.dateDebutExercice);
            ViewBag.spe = String.Format("{0:yyyy-MM-dd}", gs.dateClotureExercice);
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult ImportNewFiles([Bind(Include = "firstPeriodStart,firstPeriodEnd,firstfile,secondPeriodStart,secondPeriodEnd,secondfile")] ExcelModelView model)
        {
            var files = HttpContext.Request.Files;
            if (files.Get(0).ContentLength == 0)
            {
                ViewBag.firstfileError = "Le fichier que vous avez impoter est vide!";
                return View("Input");
            }
            else if (files.Get(1).ContentLength == 0)
            {
                ViewBag.secondfileError = "Le fichier que vous avez impoter est vide!";
                return View("Input");
            }
            else if (model.firstPeriodStart > model.firstPeriodEnd)
            {
                @ViewBag.firstPeriodError = " La date de début et la date de clotûre de la première période doivent être en ordre!";
                return View("Input");
            }
            else if (model.secondPeriodStart > model.secondPeriodEnd)
            {
                @ViewBag.firstPeriodError = "La date de début et la date de clotûre de la deuxième période doivent être en ordre!";
                return View("Input");
            }
            else if (model.secondPeriodStart < model.firstPeriodEnd)
            {
                @ViewBag.firstPeriodError = "Les deux périodes ne doivent pas s'intercepter!";
                return View("Input");
            }
            else
            {
                if ((files.Get(0).FileName.EndsWith("xls") || files.Get(0).FileName.EndsWith("xlsx")) && (files.Get(1).FileName.EndsWith("xls") || files.Get(1).FileName.EndsWith("xlsx")))
                {
                    //saving the file
                    #region Unique File Name Generation

                    string finaldate = DateTime.Now.ToString("_MMddyyyy_HHmmss");

                    string firstname = Path.GetFileNameWithoutExtension(files.Get(0).FileName);
                    string firstext = Path.GetExtension(files.Get(0).FileName);
                    string secondname = Path.GetFileNameWithoutExtension(files.Get(1).FileName);
                    string secondext = Path.GetExtension(files.Get(1).FileName);
                    string FFN = firstname + finaldate + firstext;
                    string SFN = secondname + finaldate + secondext;

                    #endregion

                    string firstPath = Server.MapPath("~/Content/" + FFN);
                    string secondPath = Server.MapPath("~/Content/" + SFN);

                    files.Get(0).SaveAs(firstPath);
                    files.Get(1).SaveAs(secondPath);

                    #region Read the data from the first Excel File
                    Excel.Application Fapplication = new Excel.Application();
                    Excel.Workbooks Fworkbooks = Fapplication.Workbooks;
                    Excel.Workbook Fworkbook = Fworkbooks.Open(firstPath);
                    Excel.Worksheet Fworksheet = (Excel.Worksheet)Fworkbook.ActiveSheet;
                    Excel.Range Frange = Fworksheet.UsedRange;

                    List<ExcelInfo> firstFileContent = new List<ExcelInfo>();
                    for (int row = 1; row <= Frange.Rows.Count; row++)
                    {
                        ExcelInfo ei = new ExcelInfo();
                        ei.periode = 1;
                        ei.journal = ((string)((Excel.Range)Frange.Cells[row, 1]).Text);
                        ei.compte = ((string)((Excel.Range)Frange.Cells[row, 2]).Text);

                        string date = (string)(((Excel.Range)Frange.Cells[row, 3]).Value2.ToString());
                        //DateTime myDate = DateTime.Parse(date, new CultureInfo("en-US", true));

                        //string sDate = (xlRange.Cells[4, 3] as Excel.Range).Value2.ToString();
                        double ddate = double.Parse(date);
                        var myDate = DateTime.FromOADate(ddate);
                        if (!(myDate > model.firstPeriodStart && myDate < model.firstPeriodEnd))
                        {
                            @ViewBag.firstPeriodError = "Cette date n'appartient pas à l'intervalle de la première période." + myDate.ToString("D");
                            return View("Input");
                        }
                        ei.dateEcriture = myDate;
                        ei.debit = Int32.Parse((string)((Excel.Range)Frange.Cells[row, 4]).Text);
                        ei.credit = Int32.Parse((string)((Excel.Range)Frange.Cells[row, 5]).Text);
                        firstFileContent.Add(ei);
                    }
                    Fworkbook.Close();
                    Marshal.ReleaseComObject(Fworkbook);
                    Marshal.ReleaseComObject(Fworksheet);
                    Marshal.ReleaseComObject(Frange);
                    Fapplication.Quit();
                    Fworkbook = null;
                    Fworksheet = null;
                    Fapplication = null;
                    #endregion

                    #region Read the data from the second Excel File
                    Excel.Application Sapplication = new Excel.Application();
                    Excel.Workbooks Sworkbooks = Sapplication.Workbooks;
                    Excel.Workbook Sworkbook = Sworkbooks.Open(secondPath);
                    Excel.Worksheet Sworksheet = (Excel.Worksheet)Sworkbook.ActiveSheet;
                    Excel.Range Srange = Sworksheet.UsedRange;

                    List<ExcelInfo> secondFileContent = new List<ExcelInfo>();
                    for (int row = 1; row <= Srange.Rows.Count; row++)
                    {
                        ExcelInfo ei = new ExcelInfo();
                        ei.periode = 2;
                        ei.journal = (string)((Excel.Range)Srange.Cells[row, 1]).Text;
                        ei.compte = ((string)((Excel.Range)Srange.Cells[row, 2]).Text);

                        string date = (string)(((Excel.Range)Srange.Cells[row, 3]).Value2.ToString());
                        //DateTime myDate = DateTime.Parse(date, new CultureInfo("en-US", true));

                        //string sDate = (xlRange.Cells[4, 3] as Excel.Range).Value2.ToString();
                        double ddate = double.Parse(date);
                        var myDate = DateTime.FromOADate(ddate);
                        if (!(myDate > model.secondPeriodStart && myDate < model.secondPeriodEnd))
                        {
                            @ViewBag.secondPeriodError = "Cette date n'appartient pas à l'intervalle de la deuxième période." + myDate.ToString("D");
                            return View("Input");
                        }
                        ei.dateEcriture = myDate;
                        ei.debit = Int32.Parse((string)((Excel.Range)Srange.Cells[row, 4]).Text);
                        ei.credit = Int32.Parse((string)((Excel.Range)Srange.Cells[row, 5]).Text);
                        secondFileContent.Add(ei);
                    }
                    Sworkbook.Close();
                    Marshal.ReleaseComObject(Sworkbook);
                    Marshal.ReleaseComObject(Sworksheet);
                    Marshal.ReleaseComObject(Srange);
                    Sapplication.Quit();
                    Sworkbook = null;
                    Sworksheet = null;
                    Sapplication = null;
                    #endregion

                    var firstFileContentSession = from a in firstFileContent
                                                  select a;
                    var secondFileContentSession = from a in secondFileContent
                                                   select a;

                    generalSettings ste = (generalSettings)Session["SteInformation"];
                    model.exercice = ste.exercice;
                    model.ownerId = ste.ownerId;
                    model.matricule = ste.matricule;
                    model.firstfile = firstPath;
                    model.secondfile = secondPath;
                    db.ExcelModelViews.Add(model);
                    db.SaveChanges();

                    Session["firstInputFile"] = firstFileContentSession;
                    Session["secondInputFile"] = secondFileContentSession;
                    Session["ExcelModelView"] = model;

                    #region delete old files

                    ExcelModelView emv = (from a in db.ExcelModelViews
                                          where a.exercice == ste.exercice &&
                                                a.matricule.Equals(ste.matricule) &&
                                                a.ownerId.Equals(ste.ownerId)
                                          select a).FirstOrDefault();
                    if (emv != null)
                    {
                        db.ExcelModelViews.Remove(emv);
                        if (System.IO.File.Exists(emv.firstfile))
                        {
                            System.IO.File.Delete(emv.firstfile);
                        }
                        if (System.IO.File.Exists(emv.secondfile))
                        {
                            System.IO.File.Delete(emv.secondfile);
                        }
                    }

                    db.ExcelModelViews.Remove(emv);
                    db.SaveChanges();
                    #endregion
                    return RedirectToAction("Success");
                }
                else
                {
                    ViewBag.Error = "Ce n'est pas un fichier Excel";
                    return View("Index");
                }
            }
        }

        [Authorize]
        public ActionResult ClearSession()
        {
            Session.Clear();
            Session.Abandon();
            return RedirectToAction("Input");
        }

        [Authorize]
        public ActionResult Input()
        {
            if (Session["SteInformation"] == null)
            {
                ViewBag.Msg = " Vous devez d'abord saisir les informations de l'entreprise";
                return RedirectToAction("Index");
            }
            else if (Session["firstInputFile"] != null || Session["secondInputFile"] != null)
            {
                ViewBag.Error = "Verifier votre fichiers ! <a href=\"/Excel/Success\"> Fichiers importés </a> <br> ";
                return View("FileError");
            }
            generalSettings gs = (generalSettings)Session["SteInformation"];
            ViewBag.fps = String.Format("{0:yyyy-MM-dd}", gs.dateDebutExercice.AddYears(-1));
            ViewBag.fpe = String.Format("{0:yyyy-MM-dd}", gs.dateClotureExercice.AddYears(-1));
            ViewBag.sps = String.Format("{0:yyyy-MM-dd}", gs.dateDebutExercice);
            ViewBag.spe = String.Format("{0:yyyy-MM-dd}", gs.dateClotureExercice);
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult RegisterSte([Bind(Include = "matricule,nomEtPrenomRaisonSociale,activite,adresse,exercice,dateDebutExercice,dateClotureExercice,actededepot,natureDepot")] generalSettings gs)
        {
            if (gs.dateDebutExercice.Year != gs.exercice)
            {
                @ViewBag.dateDebutYearError = " Cette date n'est pas dans le même année que l'exercice !";
                return View("Index", gs);
            }
            if (gs.dateClotureExercice.Year != gs.exercice)
            {
                @ViewBag.dateClotureYearError = " Cette date n'est pas dans le même année que l'exercice !";
                return View("Index", gs);
            }
            if (gs.dateDebutExercice > gs.dateClotureExercice)
            {
                @ViewBag.dateError = " Date début exercice doit etre inéerieur à la date de clotûre !";
                return View("Index", gs);
            }
            else
            {
                try
                {
                    //Saving the company into the database
                    gs.ownerId = User.Identity.GetUserId();
                    db.GeneralSettings.Add(gs);
                    db.SaveChanges();
                    Session["SteInformation"] = gs;

                    //creating the associate setting 
                    ParametersSetting ps = new ParametersSetting();
                    ps.ownerId = gs.ownerId;
                    ps.matricule = gs.matricule;
                    ps.exercice = gs.exercice;


                    db.ParametersSetting.Add(ps);
                    db.SaveChanges();

                    return RedirectToAction("Input", gs);
                }
                catch (Exception e)
                {
                  
                    ViewBag.sqlError = "Une entreprise avec la même matricule et le même exercice existe déja ! Verifier la liste de vos entreprises <a href=\"/Home/Index\"> Tous les entreprises </a> ou modifier votre entrée";
                    return View("Index", gs);
                }

            }



        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Import([Bind(Include = "firstPeriodStart,firstPeriodEnd,firstfile,secondPeriodStart,secondPeriodEnd,secondfile")] ExcelModelView model)
        {
            var files = HttpContext.Request.Files;
            if (files.Get(0).ContentLength == 0)
            {
                ViewBag.firstfileError = "Le fichier que vous avez impoter est vide!";
                return View("Input");
            }
            else if (files.Get(1).ContentLength == 0)
            {
                ViewBag.secondfileError = "Le fichier que vous avez impoter est vide!";
                return View("Input");
            }
            else if (model.firstPeriodStart > model.firstPeriodEnd)
            {
                @ViewBag.firstPeriodError = " La date de début et la date de clotûre de la première période doivent être en ordre!";
                return View("Input");
            }
            else if (model.secondPeriodStart > model.secondPeriodEnd)
            {
                @ViewBag.firstPeriodError = "La date de début et la date de clotûre de la deuxième période doivent être en ordre!";
                return View("Input");
            }
            else if (model.secondPeriodStart < model.firstPeriodEnd)
            {
                @ViewBag.firstPeriodError = "Les deux périodes ne doivent pas s'intercepter!";
                return View("Input");
            }
            else
            {
                if ((files.Get(0).FileName.EndsWith("xls") || files.Get(0).FileName.EndsWith("xlsx")) && (files.Get(1).FileName.EndsWith("xls") || files.Get(1).FileName.EndsWith("xlsx")))
                {
                    //saving the file
                    #region Unique File Name Generation

                    string finaldate = DateTime.Now.ToString("_MMddyyyy_HHmmss");

                    string firstname = Path.GetFileNameWithoutExtension(files.Get(0).FileName);
                    string firstext = Path.GetExtension(files.Get(0).FileName);
                    string secondname = Path.GetFileNameWithoutExtension(files.Get(1).FileName);
                    string secondext = Path.GetExtension(files.Get(1).FileName);
                    string FFN = firstname + finaldate + firstext;
                    string SFN = secondname + finaldate + secondext;

                    #endregion

                    string firstPath = Server.MapPath("~/Content/" + FFN);
                    string secondPath = Server.MapPath("~/Content/" + SFN);

                    files.Get(0).SaveAs(firstPath);
                    files.Get(1).SaveAs(secondPath);

                    #region Read the data from the first Excel File
                    Excel.Application Fapplication = new Excel.Application();
                    Excel.Workbooks Fworkbooks = Fapplication.Workbooks;
                    Excel.Workbook Fworkbook = Fworkbooks.Open(firstPath);
                    Excel.Worksheet Fworksheet = (Excel.Worksheet)Fworkbook.ActiveSheet;
                    Excel.Range Frange = Fworksheet.UsedRange;

                    List<ExcelInfo> firstFileContent = new List<ExcelInfo>();
                    for (int row = 1; row <= Frange.Rows.Count; row++)
                    {
                        ExcelInfo ei = new ExcelInfo();
                        ei.periode = 1;
                        ei.journal = (string)((Excel.Range)Frange.Cells[row, 1]).Text;
                        ei.compte = ((string)((Excel.Range)Frange.Cells[row, 2]).Text);

                        string date = (string)(((Excel.Range)Frange.Cells[row, 3]).Value2.ToString());
                        //DateTime myDate = DateTime.Parse(date, new CultureInfo("en-US", true));

                        //string sDate = (xlRange.Cells[4, 3] as Excel.Range).Value2.ToString();
                        double ddate = double.Parse(date);
                        var myDate = DateTime.FromOADate(ddate);
                        if (!(myDate > model.firstPeriodStart && myDate < model.firstPeriodEnd))
                        {
                            @ViewBag.firstPeriodError = "Cette date n'appartient pas à l'intervalle de la première période." + myDate.ToString("D");
                            return View("Input");
                        }
                        ei.dateEcriture = myDate;
                        ei.debit = float.Parse((string)((Excel.Range)Frange.Cells[row, 4]).Text, System.Globalization.CultureInfo.InvariantCulture);
                        ei.credit = float.Parse((string)((Excel.Range)Frange.Cells[row, 5]).Text, System.Globalization.CultureInfo.InvariantCulture);                      
                        firstFileContent.Add(ei);
                    }
                    Fworkbook.Close();
                    Marshal.ReleaseComObject(Fworkbook);
                    Marshal.ReleaseComObject(Fworksheet);
                    Marshal.ReleaseComObject(Frange);
                    Fapplication.Quit();
                    Fworkbook = null;
                    Fworksheet = null;
                    Fapplication = null;
                    #endregion

                    #region Read the data from the second Excel File
                    Excel.Application Sapplication = new Excel.Application();
                    Excel.Workbooks Sworkbooks = Sapplication.Workbooks;
                    Excel.Workbook Sworkbook = Sworkbooks.Open(secondPath);
                    Excel.Worksheet Sworksheet = (Excel.Worksheet)Sworkbook.ActiveSheet;
                    Excel.Range Srange = Sworksheet.UsedRange;

                    List<ExcelInfo> secondFileContent = new List<ExcelInfo>();
                    for (int row = 1; row <= Srange.Rows.Count; row++)
                    {
                        ExcelInfo ei = new ExcelInfo();
                        ei.periode = 2;
                        ei.journal = (string)((Excel.Range)Srange.Cells[row, 1]).Text;
                        ei.compte = ((string)((Excel.Range)Srange.Cells[row, 2]).Text);

                        string date = (string)(((Excel.Range)Srange.Cells[row, 3]).Value2.ToString());
                        //DateTime myDate = DateTime.Parse(date, new CultureInfo("en-US", true));

                        //string sDate = (xlRange.Cells[4, 3] as Excel.Range).Value2.ToString();
                        double ddate = double.Parse(date);
                        var myDate = DateTime.FromOADate(ddate);
                        if (!(myDate > model.secondPeriodStart && myDate < model.secondPeriodEnd))
                        {
                            @ViewBag.secondPeriodError = "Cette date n'appartient pas à l'intervalle de la deuxième période." + myDate.ToString("D");
                            return View("Input");
                        }
                        ei.dateEcriture = myDate;
                        ei.debit = float.Parse((string)((Excel.Range)Srange.Cells[row, 4]).Text, System.Globalization.CultureInfo.InvariantCulture);
                        ei.credit = float.Parse((string)((Excel.Range)Srange.Cells[row, 5]).Text, System.Globalization.CultureInfo.InvariantCulture);
                        secondFileContent.Add(ei);
                    }
                    Sworkbook.Close();
                    Marshal.ReleaseComObject(Sworkbook);
                    Marshal.ReleaseComObject(Sworksheet);
                    Marshal.ReleaseComObject(Srange);
                    Sapplication.Quit();
                    Sworkbook = null;
                    Sworksheet = null;
                    Sapplication = null;
                    #endregion

                    var firstFileContentSession = from a in firstFileContent
                                                  select a;
                    var secondFileContentSession = from a in secondFileContent
                                                   select a;

                    generalSettings ste = (generalSettings)Session["SteInformation"];
                    model.exercice = ste.exercice;
                    model.ownerId = ste.ownerId;
                    model.matricule = ste.matricule;
                    model.firstfile = firstPath;
                    model.secondfile = secondPath;
                    model.journaleRAN = "RAN";
                    db.ExcelModelViews.Add(model);
                    db.SaveChanges();

                    Session["firstInputFile"] = firstFileContentSession;
                    Session["secondInputFile"] = secondFileContentSession;
                    Session["ExcelModelView"] = model;
                    return RedirectToAction("Success");
                }
                else
                {
                    ViewBag.Error = "Ce n'est pas un fichier Excel!";
                    return View("Index");
                }
            }
        }

        [Authorize]
        public ActionResult Success()
        {
            GC.Collect();
            if (Session["SteInformation"] == null)
            {
                ViewBag.Error = "Vous devez d'abord choisir une entreprise ! <a href=\"/Home/Index\"> Tous les entreprises </a>";
                return View("FileError");
            }
            else if (Session["firstInputFile"] == null || Session["secondInputFile"] == null || Session["ExcelModelView"] == null)
            {
                ViewBag.Error = "Vous devez d'abord importer les fichiers  ! Rediriger vers <a href=\"/Excel/Input\"> Importer </a>";
                return View("FileError");
            }
            IEnumerable<ExcelInfo> firstInputFile = (IEnumerable<ExcelInfo>)Session["firstInputFile"];
            IEnumerable<ExcelInfo> secondInputFile = (IEnumerable<ExcelInfo>)Session["secondInputFile"];
            ViewBag.secondInputFile = secondInputFile;
            ViewBag.gs = Session["SteInformation"];
            return View(firstInputFile);
        }

        [Authorize]
        public ActionResult EditPeriodes()
        {
            return View(Session["Periodes"]);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize]
        public ActionResult EditPeriodes([Bind(Include = "Id,firstFile,secondFile,matricule,exercice,firstPeriodStart,firstPeriodEnd,secondPeriodStart,secondPeriodEnd,ownerId")] ExcelModelView periode)
        {

            if (periode.firstPeriodStart > periode.firstPeriodEnd)
            {
                @ViewBag.firstPeriodError = " La date de début et la date de clotûre de la première période doivent être en ordre!";
                return View(periode);
            }
            else if (periode.secondPeriodStart > periode.secondPeriodEnd)
            {
                @ViewBag.secondPeriodError = " La date de début et la date de clotûre de la deuxième période doivent être en ordre!";
                return View(periode);
            }
            else if (periode.secondPeriodStart < periode.firstPeriodEnd)
            {
                @ViewBag.bothPeriodError = "Les deux périodes ne doivent pas s'intercepter!";
                return View(periode);
            }
            IEnumerable<ExcelInfo> first =(IEnumerable<ExcelInfo>) Session["firstInputFile"];
            IEnumerable<ExcelInfo> second = (IEnumerable<ExcelInfo>)Session["secondInputFile"];

            foreach (ExcelInfo ei in first)
            {
                if ((ei.dateEcriture > periode.firstPeriodEnd) ||(ei.dateEcriture < periode.firstPeriodStart))
                {
                    @ViewBag.firstPeriodError = "La première période n'inclus pas tous les dates du fichier.";
                    return View(periode);
                }
            }
            foreach (ExcelInfo ei in second)
            {
                if ((ei.dateEcriture > periode.secondPeriodEnd) || (ei.dateEcriture < periode.secondPeriodStart))
                {
                    @ViewBag.secondPeriodError = "La deuxième période n'inclus pas tous les dates du fichier.";
                    return View(periode);
                }
            }

            db.Entry(periode).State = EntityState.Modified;
            db.SaveChanges();
            Session["Periode"] = periode;

            return RedirectToAction("DetailsPeriodes");

        }

        [Authorize]
        public ActionResult DetailsPeriodes()
        {
            if (Session["SteInformation"] == null)
            {
                ViewBag.Error = "Vous devez d'abord choisir une entreprise ! <a href=\"/Home/Index\"> Tous les entreprises </a>";
                return View("FileError");
            }
            else if (Session["firstInputFile"] == null || Session["secondInputFile"] == null || Session["ExcelModelView"] == null)
            {
                ViewBag.Error = "Vous devez d'abord importer les fichiers  ! Rediriger vers <a href=\"/Excel/Input\"> Importer </a>";
                return View("FileError");
            }
            generalSettings gs = (generalSettings)Session["SteInformation"];
            var periodes = from a in db.ExcelModelViews
                           where a.exercice == gs.exercice && a.matricule == gs.matricule && a.ownerId.Equals(gs.ownerId)
                           select a;
            ExcelModelView emv = periodes.FirstOrDefault();
            Session["Periodes"] = emv;

            return View(emv);
        }


        public ActionResult Documentation()
        {
            return View();
        }

        [Authorize]
        public ActionResult RAN()
        {
            if (Session["SteInformation"] == null)
            {
                ViewBag.Error = "Vous devez d'abord choisir une entreprise ! <a href=\"/Home/Index\"> Tous les entreprises </a>";
                return View("FileError");
            }
            else if (Session["firstInputFile"] == null || Session["secondInputFile"] == null || Session["ExcelModelView"] == null)
            {
                ViewBag.Error = "Vous devez d'abord importer les fichiers  ! Rediriger vers <a href=\"/Excel/Input\"> Importer </a>";
                return View("FileError");
            }
            generalSettings gs = (generalSettings)Session["SteInformation"];

            ExcelModelView emv = (from e in db.ExcelModelViews
                                  where e.exercice == gs.exercice && e.ownerId.Equals(gs.ownerId) && e.matricule.Equals(gs.matricule)
                                  select e).FirstOrDefault();
            ViewBag.ran = emv.journaleRAN;
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        [Authorize]
        public ActionResult RAN([Bind(Include ="journale")] RANViewModel ranViewModel)
        {
            generalSettings gs = (generalSettings)Session["SteInformation"];

            ExcelModelView emv = (from e in db.ExcelModelViews
                                  where e.exercice == gs.exercice && e.ownerId.Equals(gs.ownerId) && e.matricule.Equals(gs.matricule)
                                  select e).FirstOrDefault();
            emv.journaleRAN = ranViewModel.journale;

            db.Entry(emv).State = EntityState.Modified;
            db.SaveChanges();
            ViewBag.ran = emv.journaleRAN;
            return View();
        }


    }
}