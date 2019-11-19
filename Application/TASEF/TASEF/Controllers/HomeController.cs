using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.EntityFramework;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Web;
using System.Web.Mvc;
using TASEF.Infrastructure;
using TASEF.Models;
using Excel = Microsoft.Office.Interop.Excel;


namespace TASEF.Controllers
{
    public class HomeController : Controller
    {
        private ExcelProjectContext db = new ExcelProjectContext();

        [Authorize]
        public ActionResult Index()
        {
            var manager = new UserManager<ApplicationUser>(new UserStore<ApplicationUser>(new ApplicationDbContext()));
            var currentUser = manager.FindById(User.Identity.GetUserId());
            string Id = currentUser.Id;
            var companies = from a in db.GeneralSettings
                            where a.ownerId == Id
                            select a;
            Session["Companies"] = companies;
            return View(companies);

        }

        [Authorize]
        public ActionResult Select(string ownerId,string matricule, int exercice)
        {
            Session.Clear();

            generalSettings gs = db.GeneralSettings.Find(ownerId,matricule, exercice);
            Session["SteInformation"] = gs;
            var em = from e in db.ExcelModelViews
                     where e.exercice == gs.exercice && e.matricule == gs.matricule && e.ownerId==gs.ownerId
                     select e;
            ExcelModelView emv = em.FirstOrDefault();
            if (emv == null)
                return RedirectToAction("Input", "Excel");
            else
            {
                string firstPath = emv.firstfile;
                string secondPath = emv.secondfile;

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

                Session["firstInputFile"] = firstFileContentSession;
                Session["secondInputFile"] = secondFileContentSession;
                Session["ExcelModelView"] = emv;


                return RedirectToAction("Success", "Excel");
            }

        }

        [Authorize]
        public ActionResult Edit(string ownerId,string matricule, int exercice)
        {
            generalSettings gs = db.GeneralSettings.Find(ownerId,matricule, exercice);
            return View(gs);
        }

        [Authorize]
        public ActionResult Details(string ownerId,string matricule, int exercice, int test)
        {
            generalSettings gs = db.GeneralSettings.Find(ownerId, matricule, exercice);
            if (test == 2)
            {
                ViewBag.action = "Success";
                ViewBag.controller = "Excel";
            }
            else
            {
                ViewBag.action = "Index";
                ViewBag.controller = "Home";
            }
            return View(gs);
        }

        [HttpPost]
        [Authorize]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "matricule, nomEtPrenomRaisonSociale, activite, adresse, exercice, dateDebutExercice, dateClotureExercice, actededepot, natureDepot,ownerId")] generalSettings generalsettings)
        {
            if (generalsettings.dateDebutExercice.Year != generalsettings.exercice)
            {
                @ViewBag.dateDebutYearError = " Cette date n'est pas dans le même année que l'exercice !";
                return View(generalsettings);
            }
            if (generalsettings.dateClotureExercice.Year != generalsettings.exercice)
            {
                @ViewBag.dateClotureYearError = " Cette date n'est pas dans le même année que l'exercice !";
                return View(generalsettings);
            }
            if (generalsettings.dateDebutExercice > generalsettings.dateClotureExercice)
            {
                @ViewBag.PeriodError = "Date début doit être inférieur à la date de clotûre";
                return View(generalsettings);
            }

            db.Entry(generalsettings).State = EntityState.Modified;
            db.SaveChanges();
            Session["SteInformation"] = generalsettings;

            //TODO : change only one company 
            var manager = new UserManager<ApplicationUser>(new UserStore<ApplicationUser>(new ApplicationDbContext()));
            var currentUser = manager.FindById(User.Identity.GetUserId());
            string Id = currentUser.Id;
            var companies = from a in db.GeneralSettings
                            where a.ownerId == Id
                            select a;
            Session["Companies"] = companies;
            return View("Index", companies);

        }

        [Authorize]
        public ActionResult Delete(string ownerId,string matricule, string exercice)
        {
            if ((matricule == null) || (exercice == null))
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            generalSettings ste = db.GeneralSettings.Find(ownerId,matricule, Int32.Parse(exercice));
            if (ste == null)
            {
                return HttpNotFound();
            }
            return View(ste);
        }

        [Authorize]
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string ownerId,string matricule, string exercice)
        {
            generalSettings ste = db.GeneralSettings.Find(ownerId,matricule, Int32.Parse(exercice));
            int e = Int32.Parse(exercice);
            #region Delete all actif formula and parameters
            var actifFormula = from af in db.ActifFormula
                               where af.exercice == e && af.matricule.Equals(matricule) && af.ownerId.Equals(ownerId)
                               select af;
            var actifParam = from ap in db.ActifModel
                             where ap.exercice == e && ap.matricule.Equals(matricule) && ap.ownerId.Equals(ownerId)
                             select ap;
            foreach (var a in actifFormula)
            {
                db.ActifFormula.Remove(a);
            }
            foreach (var a in actifParam)
            {
                db.ActifModel.Remove(a);
            }
            #endregion

            #region Delete all passif formula and parameters
            var passifFormula = from af in db.PassifFormula
                                where af.exercice == e && af.matricule.Equals(matricule)
                                select af;
            var passifParam = from ap in db.PassifModel
                              where ap.exercice == e && ap.matricule.Equals(matricule)
                              select ap;
            foreach (var a in passifFormula)
            {
                db.PassifFormula.Remove(a);
            }
            foreach (var a in passifParam)
            {
                db.PassifModel.Remove(a);
            }
            #endregion

            #region Delete all etat de résultat formula and parameters
            var EDRFormula = from af in db.EtatDeResultatFormula
                             where af.exercice == e && af.matricule.Equals(matricule)
                             select af;
            var EDRParam = from ap in db.EtatDeResultatModel
                           where ap.exercice == e && ap.matricule.Equals(matricule)
                           select ap;
            foreach (var a in EDRFormula)
            {
                db.EtatDeResultatFormula.Remove(a);
            }
            foreach (var a in EDRParam)
            {
                db.EtatDeResultatModel.Remove(a);
            }
            #endregion

            #region Delete all flux de trésorie modelle autorisé formula and parameters
            var FMAFormula = from af in db.FluxTresorerieMAFormula
                             where af.exercice == e && af.matricule.Equals(matricule)
                             select af;
            var FMAParam = from ap in db.FluxTresorerieMAModel
                           where ap.exercice == e && ap.matricule.Equals(matricule)
                           select ap;
            foreach (var a in FMAFormula)
            {
                db.FluxTresorerieMAFormula.Remove(a);
            }
            foreach (var a in FMAParam)
            {
                db.FluxTresorerieMAModel.Remove(a);
            }
            #endregion 

            #region Delete all flux de trésorie modele de référence formula and parameters
            var FMRFormula = from af in db.FluxTresorerieMRFormula
                             where af.exercice == e && af.matricule.Equals(matricule)
                             select af;
            var FMRParam = from ap in db.FluxTresorerieMRModel
                           where ap.exercice == e && ap.matricule.Equals(matricule)
                           select ap;
            foreach (var a in FMRFormula)
            {
                db.FluxTresorerieMRFormula.Remove(a);
            }
            foreach (var a in FMRParam)
            {
                db.FluxTresorerieMRModel.Remove(a);
            }
            #endregion

            #region Delete all resultat fiscal formula and parameters
            var RFFormula = from af in db.ResultatFiscalFormula
                            where af.exercice == e && af.matricule.Equals(matricule) && af.ownerId.Equals(ownerId)
                            select af;
            var RFParam = from ap in db.ResultatFiscalModel
                          where ap.exercice == e && ap.matricule.Equals(matricule) && ap.ownerId.Equals(ownerId)
                          select ap;
            foreach (var a in RFFormula)
            {
                db.ResultatFiscalFormula.Remove(a);
            }
            foreach (var a in RFParam)
            {
                db.ResultatFiscalModel.Remove(a);
            }
            #endregion

            #region Delete parameters setting
            var pr = from a in db.ParametersSetting
                     where a.exercice == e && a.matricule.Equals(matricule) && a.ownerId.Equals(ownerId)
                     select a;
            foreach (var a in pr)
            {
                db.ParametersSetting.Remove(a);
            }
            #endregion

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
                

            db.GeneralSettings.Remove(ste);
            db.SaveChanges();
            Session.Clear();
            return RedirectToAction("Index");
        }

    }
}