using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.EntityFramework;
using Rotativa;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using System.Xml;
using TASEF.Infrastructure;
using TASEF.Models;

namespace TASEF.Controllers
{
    public class FluxTresorerieMAParametersController : Controller
    {


        private ExcelProjectContext db = new ExcelProjectContext();

        // GET: Parameters
        [Authorize]
        public ActionResult Index()
        {
            if (Session["SteInformation"] == null)
            {
                ViewBag.Error = "Vous devez d'abord choisir une entreprise ! <a href=\"/Home/Index\"> Tous les entreprises </a>";
                return View("FileError");
            }
            else if (Session["firstInputFile"] == null || Session["secondInputFile"] == null)
            {
                ViewBag.Error = "Vous devez d'abord importer les fichiers ! <a href=\"/Excel/Input\"> Importer </a>";
                return View("FileError");
            }

            generalSettings gs = (generalSettings)Session["SteInformation"];

            var ParamSetting = from ps in db.ParametersSetting
                               where ps.ownerId.Equals(gs.ownerId) && ps.matricule.Equals(gs.matricule) && ps.exercice == gs.exercice
                               select ps;

            ParametersSetting paramFluxMA = ParamSetting.FirstOrDefault();
            if (!paramFluxMA.hasParamFluxMA)
                return View("ParamSetting");
            else
            {

                IEnumerable<ExcelInfo> firstInputFile = (IEnumerable<ExcelInfo>)Session["firstInputFile"];
                IEnumerable<ExcelInfo> secondInputFile = (IEnumerable<ExcelInfo>)Session["secondInputFile"];

                IEnumerable<ExcelInfo> InputFile = firstInputFile.Concat(secondInputFile);

                //calculate the values foreach parameter
                int exercice = gs.exercice;
                string matricule = gs.matricule;
                string ownerId = gs.ownerId;

                var af = from a in db.FluxTresorerieMAModel
                         where a.exercice == exercice && a.matricule.Equals(matricule) && a.ownerId.Equals(ownerId)
                         select a;

                var databaseFormulaList = from fl in db.FluxTresorerieMAFormula
                                          where fl.exercice == exercice && fl.matricule.Equals(matricule) && fl.ownerId.Equals(ownerId)
                                          select fl;


                //there are no FluxTresorerieMA parameters for that company so we need to create them
                if (!af.Any())
                {
                    List<FluxTresorerieMAParamModel> FluxTresorerieMAParamList = new List<FluxTresorerieMAParamModel>
            {
                new FluxTresorerieMAParamModel(){ code="0001" , libelle ="Flux de Trésorerie provenant de (affectés à) l'Exploitation" , type="Formula",state="Stable",priority=0},
                new FluxTresorerieMAParamModel(){ code="0002" , libelle ="Résultat net" , type="Calculated",state="Stable"},
                new FluxTresorerieMAParamModel(){ code="0003" , libelle ="Ajustement amortissement et provision" , type="Calculated",state="Stable"},
                new FluxTresorerieMAParamModel(){ code="0004" , libelle ="variation des stocks" , type="Calculated",state="Stable"},
                new FluxTresorerieMAParamModel(){ code="0005" , libelle ="Variation des créances" , type="Calculated",state="Stable"},
                new FluxTresorerieMAParamModel(){ code="0006" , libelle ="Variation des autres actifs" , type="Calculated",state="Stable"},
                new FluxTresorerieMAParamModel(){ code="0007" , libelle ="Variation des fournisseurs et autres dettes" , type="Calculated",state="Stable"},
                new FluxTresorerieMAParamModel(){ code="0008" , libelle ="Ajustement plus au moins values de cession" , type="Calculated",state="Stable"},
                new FluxTresorerieMAParamModel(){ code="0009" , libelle ="Ajustement transfert des charges" , type="Calculated",state="Stable"},
                new FluxTresorerieMAParamModel(){ code="0010" , libelle ="Autres (Flux de Trésorerie provenant de (affectés à) l'Exploitation ) " , type="Calculated",state="Stable"},
                new FluxTresorerieMAParamModel(){ code="0011" , libelle ="Flux de Trésorerie provenant de (affectés aux)  activités d'Investissement(N)" , type="Formula",state="Stable",priority=0},
                new FluxTresorerieMAParamModel(){ code="0012" , libelle ="Décaissements provenant de l'acquisition d'immobilisations corporelles et incorporelles" , type="Calculated",state="Stable"},
                new FluxTresorerieMAParamModel(){ code="0013" , libelle ="Encaissements provenant de l'acquisition d'immobilisations corporelles et incorporelles" , type="Calculated",state="Stable"},
                new FluxTresorerieMAParamModel(){ code="0014" , libelle ="Décaissements provenant de l'acquisition d'immobilisations financières" , type="Calculated",state="Stable"},
                new FluxTresorerieMAParamModel(){ code="0015" , libelle ="Encaissements provenant de l'acquisition d'immobilisations financières" , type="Calculated",state="Stable"},
                new FluxTresorerieMAParamModel(){ code="0016" , libelle ="Autres (Flux de Trésorerie provenant de (affectés aux)  activités d'Investissement) " , type="Calculated",state="Stable"},
                new FluxTresorerieMAParamModel(){ code="0017" , libelle ="Flux de Trésorerie provenant des (affectés aux) activités de Financement " , type="Formula",state="Stable",priority=0},
                new FluxTresorerieMAParamModel(){ code="0018" , libelle ="Encaissement suite à l'émission d'actions " , type="Calculated",state="Stable"},
                new FluxTresorerieMAParamModel(){ code="0019" , libelle ="Dividendes et autres distribution " , type="Calculated",state="Stable"},
                new FluxTresorerieMAParamModel(){ code="0020" , libelle ="Encaissements provenant des emprunts " , type="Calculated",state="Stable"},
                new FluxTresorerieMAParamModel(){ code="0021" , libelle ="Remboursement d'emprunts " , type="Calculated",state="Stable"},
                new FluxTresorerieMAParamModel(){ code="0022" , libelle =" Autres (Flux de Trésorerie provenant des (affectés aux) activités de Financement) " , type="Calculated",state="Stable"},
                new FluxTresorerieMAParamModel(){ code="0023" , libelle ="Incidences des variations des taux de change/les liquidités&équiv°(N et) " , type="Calculated",state="Stable"},
                new FluxTresorerieMAParamModel(){ code="0024" , libelle ="Autres Postes des Flux de Trésorerie" , type="Calculated",state="Stable"},
                new FluxTresorerieMAParamModel(){ code="0025" , libelle ="Variation de Trésorerie" , type="Formula",state="Stable",priority=1},
                new FluxTresorerieMAParamModel(){ code="0026" , libelle ="Trésorerie au début de l'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMAParamModel(){ code="0027" , libelle ="Trésorerie à la clôture de l'exercice" , type="Calculated",state="Stable"},
            };

                    foreach (var param in FluxTresorerieMAParamList)
                    {
                        var manager = new UserManager<ApplicationUser>(new UserStore<ApplicationUser>(new ApplicationDbContext()));
                        var currentUser = manager.FindById(User.Identity.GetUserId());
                        param.ownerId = currentUser.Id;
                        param.exercice = exercice;
                        param.matricule = matricule;
                        param.netN = 0;
                        param.netN1 = 0;
                        db.FluxTresorerieMAModel.Add(param);
                        db.SaveChanges();
                    }
                    return View(FluxTresorerieMAParamList);

                }
                else return View(af);
            }
        }

        //GET /Parameters/Create/0001
        [Authorize]
        public ActionResult Create(string id)
        {
            List<string> definedParam = new List<string>(new string[] { "0001", "0011", "0017", "0025" });
            if (definedParam.Contains(id))
            {
                ViewBag.Error = "Ces paramétres ne peuvent pas être modifier! <a href=\"/FluxTresorerieMAParametersParameters/Index\"> Paramétres </a>";
                return View("FileError");
            }
            generalSettings gs = (generalSettings)Session["SteInformation"];
            string matricule = gs.matricule;
            int exercice = gs.exercice;
            string ownerId = gs.ownerId;


            var listFormula = from lf in db.FluxTresorerieMAFormula
                              where lf.codeParam == id
                              && lf.matricule.Equals(matricule)
                              && lf.exercice == exercice
                              && lf.ownerId.Equals(ownerId)
                              select lf;
            ViewBag.listFormula = listFormula;
            FluxTresorerieMAFormula af = new FluxTresorerieMAFormula() { codeParam = id };
            return View(af);
        }

        // POST: FluxTresorerieMAFormulas/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [Authorize]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,codeParam,codeDonnee,nomCompte,typeFormule,RANjournal")] FluxTresorerieMAFormula FluxTresorerieMAFormula)
        {
            if (ModelState.IsValid)
            {
                generalSettings gs = (generalSettings)Session["SteInformation"];
                string matricule = gs.matricule;
                int exercice = gs.exercice;
                string ownerId = gs.ownerId;

                string code = FluxTresorerieMAFormula.codeParam;
                FluxTresorerieMAParamModel apm = db.FluxTresorerieMAModel.Find(ownerId, code, exercice, matricule);
                apm.state = "Stable";

                db.Entry(apm).State = EntityState.Modified;
                db.SaveChanges();

                FluxTresorerieMAFormula.exercice = exercice;
                FluxTresorerieMAFormula.matricule = matricule;
                FluxTresorerieMAFormula.ownerId = ownerId;

                db.FluxTresorerieMAFormula.Add(FluxTresorerieMAFormula);
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(FluxTresorerieMAFormula);
        }

        [Authorize]
        public ActionResult Recalculate()
        {
            generalSettings gs = (generalSettings)Session["SteInformation"];
            int exercice = gs.exercice;
            string matricule = gs.matricule;
            string ownerId = gs.ownerId;

            var FluxTresorerieMAParam = from ap in db.FluxTresorerieMAModel
                                        where ap.matricule.Equals(matricule) && ap.exercice == exercice && ap.ownerId.Equals(ownerId)
                                        select ap;
            var FluxTresorerieMAFormula = from af in db.FluxTresorerieMAFormula
                                          where af.matricule.Equals(matricule) && af.exercice == exercice && af.ownerId.Equals(ownerId)
                                          select af;
            var calculated = from c in FluxTresorerieMAParam
                             where c.type.Equals("Calculated")
                             select c;
            var formula = from c in FluxTresorerieMAParam
                          where c.type.Equals("Formula")
                          orderby c.priority
                          select c;

            //List<Formula> FluxTresorerieMAFormulaList = new List<Formula>
            //{
            //   new Formula(){ code="0001",type="FluxTresorerieMA",parameters = new List<string>() {"0002","0003","0004","0005","0006","0007","0008","0009","0010"}},
            //   new Formula(){ code="0011",type="FluxTresorerieMA",parameters = new List<string>() {"0012","0013","0014","0015","0016"}},
            //   new Formula(){ code="0017",type="FluxTresorerieMA",parameters = new List<string>() {"0018","0019","0020","0021","0022"}},
            //   new Formula(){ code="0025",type="FluxTresorerieMA",parameters = new List<string>() {"0024","0023","0017","0011","0001"}}
            //};

            var FluxTresorerieMAFormulaList = from f in db.DefinedFormulas
                                              where f.type.Equals("FluxTresorerieMA")
                                              select f;

            IEnumerable<ExcelInfo> firstInputFile = (IEnumerable<ExcelInfo>)Session["firstInputFile"];
            IEnumerable<ExcelInfo> secondInputFile = (IEnumerable<ExcelInfo>)Session["secondInputFile"];

            IEnumerable<ExcelInfo> InputFile = firstInputFile.Concat(secondInputFile);
            IEnumerable<ExcelInfo> specificInputFile = InputFile;

            foreach (var param in calculated.ToList())
            {
                if (param.state.Equals("Stable"))
                {
                    float netN = 0;
                    float netN1 = 0;
                    float valeur = 0;
                    string code = param.code;


                    var one = FluxTresorerieMAFormula.Where(AF => AF.codeParam.Equals(code));
                    var two = one.Where(AF => AF.matricule.Equals(matricule));
                    var three = two.Where(AF => AF.matricule.Equals(ownerId));
                    var specificFluxTresorerieMAFormula = two.Where(AF => AF.exercice == exercice);
                    foreach (var formulas in specificFluxTresorerieMAFormula)
                    {
                        if (formulas.typeFormule.Equals("Solde"))
                        {
                            if (formulas.RANjournal.Equals("Tous"))
                            {
                                specificInputFile = InputFile;
                            }
                            else
                            {
                                ExcelModelView emv = (ExcelModelView)Session["ExcelModelView"];
                                specificInputFile = from f in InputFile
                                                    where f.journal.Equals(emv.journaleRAN)
                                                    select f;
                            }
                            foreach (var input in specificInputFile)
                            {
                                if (input.compte.StartsWith(formulas.codeDonnee))
                                {
                                    valeur = input.debit - input.credit;
                                    if (input.periode == 2)
                                        netN += valeur;
                                    else if (input.periode == 1)
                                        netN1 += valeur;
                                }
                            }
                        }
                        else if (formulas.typeFormule.Equals("MvtDebit"))
                        {
                            foreach (var input in specificInputFile)
                            {
                                if (input.compte.StartsWith(formulas.codeDonnee))
                                {
                                    valeur = input.debit;
                                    if (input.periode == 2)
                                        netN += valeur;
                                    else if (input.periode == 1)
                                        netN1 += valeur;
                                }
                            }
                        }
                        else if (formulas.typeFormule.Equals("MvtCredit"))
                        {
                            foreach (var input in specificInputFile)
                            {
                                if (input.compte.StartsWith(formulas.codeDonnee))
                                {
                                    valeur = input.credit;
                                    if (input.periode == 2)
                                        netN += valeur;
                                    else if (input.periode == 1)
                                        netN1 += valeur;
                                }
                            }
                        }
                        else if (formulas.typeFormule.Equals("SoldeSiD"))
                        {
                            foreach (var input in specificInputFile)
                            {
                                if (input.compte.StartsWith(formulas.codeDonnee))
                                {
                                    valeur = input.debit - input.credit;
                                    if (valeur >= 0)
                                    {
                                        if (input.periode == 2)
                                            netN += valeur;
                                        else if (input.periode == 1)
                                            netN1 += valeur;
                                    }

                                }
                            }
                        }
                        else if (formulas.typeFormule.Equals("SoldeSiC"))
                        {
                            foreach (var input in specificInputFile)
                            {
                                if (input.compte.StartsWith(formulas.codeDonnee))
                                {
                                    valeur = input.debit - input.credit;
                                    if (valeur <= 0)
                                    {
                                        if (input.periode == 2)
                                            netN += valeur;
                                        else if (input.periode == 1)
                                            netN1 += valeur;
                                    }
                                }
                            }
                        }
                    }
                    param.ownerId = ownerId;
                    param.exercice = exercice;
                    param.matricule = matricule;
                    param.netN = netN;
                    param.netN1 = netN1;
                    db.Entry(param).State = EntityState.Modified;
                    db.SaveChanges();


                }

            }
            foreach (var param in formula.ToList())
            {
                param.netN = 0;
                param.netN1 = 0;

                var f = from fo in FluxTresorerieMAFormulaList
                        where fo.code.Equals(param.code)
                        select fo;

                List<string> parameters = new List<string>();

                foreach (Formula form in f.ToList())
                {
                    parameters.Add(form.parameter);
                }
                //Formula formulas = f.FirstOrDefault();
                foreach (var code in/* formulas.*/parameters)
                {
                    string cleanCode;
                    if (code.StartsWith("-"))
                    {
                        cleanCode = code.Trim('-');
                    }
                    else
                    {
                        cleanCode = code;
                    }
                    FluxTresorerieMAParamModel apm = FluxTresorerieMAParam.Where(AF => AF.code.Equals(cleanCode)).FirstOrDefault();

                    if (code.StartsWith("-"))
                    {
                        param.netN -= apm.netN;
                        param.netN1 -= apm.netN1;
                    }
                    else
                    {
                        param.netN += apm.netN;
                        param.netN1 += apm.netN1;
                    }
                }
                param.exercice = exercice;
                param.matricule = matricule;
                param.ownerId = ownerId;
                db.Entry(param).State = EntityState.Modified;
                db.SaveChanges();
            }
            return RedirectToAction("Index");
        }

        [Authorize]
        public ActionResult Show(string id)
        {
            //List<Formula> FluxTresorerieMAFormulaList = new List<Formula>
            //{
            //   new Formula(){ code="0001",type="FluxTresorerieMA",parameters = new List<string>() {"0002","0003","0004","0005","0006","0007","0008","0009","0010"}},
            //   new Formula(){ code="0011",type="FluxTresorerieMA",parameters = new List<string>() {"0012","0013","0014","0015","0016"}},
            //   new Formula(){ code="0017",type="FluxTresorerieMA",parameters = new List<string>() {"0018","0019","0020","0021","0022"}},
            //   new Formula(){ code="0025",type="FluxTresorerieMA",parameters = new List<string>() {"0024","0023","0017","0011","0001"}}
            //};

            //Formula specificFormula = FluxTresorerieMAFormulaList.Where(AF => AF.code.Equals(id)).FirstOrDefault();

            var FluxTresorerieMAFormulaList = from flux in db.DefinedFormulas
                                              where flux.type.Equals("FluxTresorerieMA")
                                              select flux;

            var f = from fo in FluxTresorerieMAFormulaList
                    where fo.code.Equals(id)
                    select fo;

            List<string> parameters = new List<string>();

            foreach (Formula form in f.ToList())
            {
                parameters.Add(form.parameter);
            }

            List<String> CleanParameters = new List<string>();

            List<string> minus = new List<string>();
            foreach (string code in /*specificFormula.*/parameters/*.ToList()*/)
            {
                string cleanCode;
                if (code.StartsWith("-"))
                {
                    cleanCode = code.Trim('-');
                    minus.Add(cleanCode);
                }
                else
                {
                    cleanCode = code;
                }
                CleanParameters.Add(cleanCode);
            }

            generalSettings gs = (generalSettings)Session["SteInformation"];
            var FluxTresorerieMAFormulas = from af in db.FluxTresorerieMAModel
                                           where af.exercice == gs.exercice &&
                                                 af.matricule.Equals(gs.matricule) &&
                                                 af.ownerId.Equals(gs.ownerId) &&
                                                 /*specificFormula.*/CleanParameters.Contains(af.code)
                                           select af;

            ViewBag.minus = minus;
            return View(FluxTresorerieMAFormulas);
        }

        [Authorize]
        public ActionResult EditParam(string ownerId, string code, string exercice, string matricule)
        {
            if ((code == null) || (exercice == null) || (matricule == null))
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            FluxTresorerieMAParamModel FluxTresorerieMAParamModel = db.FluxTresorerieMAModel.Find(ownerId, code, Int32.Parse(exercice), matricule);
            if (FluxTresorerieMAParamModel == null)
            {
                return HttpNotFound();
            }
            return View(FluxTresorerieMAParamModel);
        }

        [HttpPost]
        [Authorize]
        [ValidateAntiForgeryToken]
        public ActionResult EditParam([Bind(Include = "code,ownerId,libelle,netN,netN1,type,exercice,matricule,ownerId,state")] FluxTresorerieMAParamModel param)
        {
            if (ModelState.IsValid)
            {
                param.state = "Changed";
                db.Entry(param).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(param);
        }

        [Authorize]
        public ActionResult PrintFluxTresorerieMAsAsPdf()
        {
            if (Session["SteInformation"] == null)
            {
                ViewBag.Error = "Vous devez d'abord choisir une entreprise ! <a href=\"/Home/Index\"> Tous les entreprises </a>";
                return View("FileError");
            }
            else if (Session["firstInputFile"] == null || Session["secondInputFile"] == null || Session["ExcelModelView"] == null)
            {
                ViewBag.Error = "Vous devez d'abord importer les fichiers  !  <a href=\"/Excel/Input\"> Importer </a>";
                return View("FileError");
            }

            generalSettings gs = (generalSettings)Session["SteInformation"];
            string matricule = gs.matricule;
            int exercice = gs.exercice;
            string ownerId = gs.ownerId;

            var af = from a in db.FluxTresorerieMAModel
                     where a.exercice == exercice && a.matricule.Equals(matricule) && a.ownerId.Equals(ownerId)
                     select a;
            ViewBag.info1 = gs.nomEtPrenomRaisonSociale;
            ViewBag.info2 = gs.adresse;
            ViewBag.info3 = gs.activite;
            ViewBag.info4 = gs.dateDebutExercice;
            ViewBag.info5 = gs.dateClotureExercice;
            var report = new ViewAsPdf("FluxTresorerieMAsAsPdf", af)
            {
                PageOrientation = Rotativa.Options.Orientation.Landscape,
                PageSize = Rotativa.Options.Size.A4,
                CustomSwitches = "--footer-center \"  Créer le : " + DateTime.Now.Date.ToString("dd/MM/yyyy") + "  Page: [page]/[toPage]\"" + " --footer-spacing 1 --footer-font-name \"Segoe UI\""
            };
            return report;
        }

        [Authorize]
        public ActionResult EditFormula(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            FluxTresorerieMAFormula FluxTresorerieMAFormula = db.FluxTresorerieMAFormula.Find(Int32.Parse(id));
            if (FluxTresorerieMAFormula == null)
            {
                return HttpNotFound();
            }
            return View(FluxTresorerieMAFormula);
        }

        [HttpPost]
        [Authorize]
        [ValidateAntiForgeryToken]
        public ActionResult EditFormula([Bind(Include = "")] FluxTresorerieMAFormula FluxTresorerieMAFormula)
        {
            if (ModelState.IsValid)
            {
                generalSettings gs = (generalSettings)Session["SteInformation"];
                int exercice = gs.exercice;
                string matricule = gs.matricule;
                string ownerId = gs.ownerId;

                string code = FluxTresorerieMAFormula.codeParam;
                FluxTresorerieMAParamModel apm = db.FluxTresorerieMAModel.Find(ownerId, code, exercice, matricule);
                apm.state = "Stable";

                db.Entry(apm).State = EntityState.Modified;
                db.SaveChanges();

                db.Entry(FluxTresorerieMAFormula).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(FluxTresorerieMAFormula);
        }

        [Authorize]
        public ActionResult Delete(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            FluxTresorerieMAFormula FluxTresorerieMAFormula = db.FluxTresorerieMAFormula.Find(Int32.Parse(id));
            if (FluxTresorerieMAFormula == null)
            {
                return HttpNotFound();
            }
            return View(FluxTresorerieMAFormula);
        }

        [Authorize]
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string id)
        {
            FluxTresorerieMAFormula FluxTresorerieMAFormula = db.FluxTresorerieMAFormula.Find(Int32.Parse(id));

            generalSettings gs = (generalSettings)Session["SteInformation"];
            int exercice = gs.exercice;
            string matricule = gs.matricule;
            string ownerId = gs.ownerId;

            string code = FluxTresorerieMAFormula.codeParam;
            FluxTresorerieMAParamModel apm = db.FluxTresorerieMAModel.Find(ownerId, code, exercice, matricule);
            apm.state = "Stable";

            db.Entry(apm).State = EntityState.Modified;
            db.SaveChanges();

            db.FluxTresorerieMAFormula.Remove(FluxTresorerieMAFormula);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        [Authorize]
        public ActionResult GenerateFiles()
        {
            if (Session["SteInformation"] == null)
            {
                ViewBag.Error = "Vous devez d'abord choisir une entreprise ! <a href=\"/Home/Index\"> Tous les entreprises </a>";
                return View("FileError");
            }
            else if (Session["firstInputFile"] == null || Session["secondInputFile"] == null || Session["ExcelModelView"] == null)
            {
                ViewBag.Error = "Vous devez d'abord importer les fichiers  ! <a href=\"/Excel/Input\"> Importer </a>";
                return View("FileError");
            }

            generalSettings gs = (generalSettings)Session["SteInformation"];
            int exercice = gs.exercice;
            string matricule = gs.matricule;
            string ownerId = gs.ownerId;

            var FluxParam = from ap in db.FluxTresorerieMAModel
                            where ap.exercice == exercice && ap.matricule.Equals(matricule) && ap.ownerId.Equals(ownerId)
                            select ap;
            if (!FluxParam.Any())
            {
                ViewBag.Error = "Vous devez d'abord configurer les paramétres Flux de Trésorerie - Modèle Autorisé  !  <a href=\"/FluxTresorerieMAParameters/Index\"> Flux de Trésorerie - Modèle Autorisé</a>";
                return View("FileError");
            }
            return View();
        }

        [Authorize]
        public void PrintFluxTresorerieMAsAsXml()
        {

            generalSettings gs = (generalSettings)Session["SteInformation"];
            var FluxTresorerieMAParam = from a in db.FluxTresorerieMAModel
                                        where a.exercice == gs.exercice && a.matricule.Equals(gs.matricule) && a.ownerId.Equals(gs.ownerId)
                                        select a;
            string fileName = "FluxTresorerieMA-" + gs.matricule + "-" + gs.exercice + ".xml";

            using (MemoryStream stream = new MemoryStream())
            {
                // Create an XML document. Write our specific values into the document.
                XmlTextWriter xmlWriter = new XmlTextWriter(stream, System.Text.Encoding.UTF8);
                // Write the XML document header.
                xmlWriter.WriteStartDocument();
                xmlWriter.WriteRaw("<?xml-stylesheet type=\"text/xsl\"?>");
                xmlWriter.WriteStartElement("lf:F6004");
                xmlWriter.WriteAttributeString("xmlns:lf", "http://www.impots.finances.gov.tn/liasse");
                xmlWriter.WriteAttributeString("xmlns:vc", "http://www.w3.org/2007/XMLSchema-versioning");
                xmlWriter.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
                xmlWriter.WriteAttributeString("xsi:schemaLocation", "http://www.impots.finances.gov.tn/liasse F6004-MODELE-AUT.xsd");
                xmlWriter.WriteElementString("lf:VersionDocument", "1");
                // Write our first XML header.
                xmlWriter.WriteStartElement("lf:Entete", "");
                //xmlWriter.WriteRaw("<lf:Entete>");

                xmlWriter.WriteElementString("lf:MatriculeFiscalDeclarant", gs.matricule);
                xmlWriter.WriteElementString("lf:NometPrenomouRaisonSociale", gs.nomEtPrenomRaisonSociale);
                xmlWriter.WriteElementString("lf:Activite", gs.activite);
                xmlWriter.WriteElementString("lf:Adresse", gs.adresse);
                xmlWriter.WriteElementString("lf:Exercice", gs.exercice.ToString());
                xmlWriter.WriteElementString("lf:DateDebutExercice", gs.dateDebutExercice.ToString("dd/MM/yyyy"));
                xmlWriter.WriteElementString("lf:DateClotureExercice", gs.dateClotureExercice.ToString("dd/MM/yyyy"));
                xmlWriter.WriteElementString("lf:ActeDeDepot", gs.actededepot);
                xmlWriter.WriteElementString("lf:NatureDepot", gs.natureDepot);

                xmlWriter.WriteEndElement();
                //xmlWriter.WriteRaw("</lf:Entete>");

                //xmlWriter.WriteStartElement("lf:Details");
                xmlWriter.WriteRaw("<lf:Details>");

                for (int i = 0; i < 2; i++)
                {
                    int add = i * 1000;
                    string code;
                    foreach (var param in FluxTresorerieMAParam)
                    {
                        code = (Int32.Parse(param.code) + add).ToString();
                        if (i == 0)
                        {
                            xmlWriter.WriteElementString($"lf:{param.code}", param.netN.ToString());
                        }
                        else if (i == 1)
                        {
                            xmlWriter.WriteElementString($"lf:{code}", param.netN1.ToString());
                        }

                    }
                }

                //xmlWriter.WriteEndElement();
                xmlWriter.WriteRaw("</lf:Details>");

                xmlWriter.WriteEndElement();
                xmlWriter.WriteEndDocument();
                // To be safe, flush the document to the memory stream.
                xmlWriter.Flush();
                // Convert the memory stream to an array of bytes.
                byte[] byteArray = stream.ToArray();
                // Send the XML file to the web browser for download.
                Response.Clear();
                Response.AppendHeader("Content-Disposition", $"filename={fileName}");
                Response.AppendHeader("Content-Length", byteArray.Length.ToString());
                Response.ContentType = "application/octet-stream";
                Response.BinaryWrite(byteArray);
                xmlWriter.Close();
            }
        }

        [Authorize]
        public ActionResult SelectParameters(int id)
        {
            generalSettings gs = (generalSettings)Session["SteInformation"];

            var ParamSetting = from ps in db.ParametersSetting
                               where ps.ownerId.Equals(gs.ownerId) && ps.matricule.Equals(gs.matricule) && ps.exercice == gs.exercice
                               select ps;

            ParametersSetting paramFluxMA = ParamSetting.FirstOrDefault();

            //3==> Specific parameters are chosen
            if (id == 3)
            {
                //Setting the hasParamActif to true
                paramFluxMA.hasParamFluxMA = true;

                db.Entry(paramFluxMA).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");

            }
            //2==> Duplicate an othe company's parameters
            else
                if (id == 2)
            {
                List<generalSettings> companies = new List<generalSettings>();

                var Allcompanies = from a in db.GeneralSettings
                                   where a.ownerId == gs.ownerId
                                   select a;

                var current = gs.matricule.Insert(gs.matricule.Length, gs.exercice.ToString());

                foreach (generalSettings company in Allcompanies)
                {
                    if (company.matricule.Insert(company.matricule.Length, company.exercice.ToString()) != current)
                    {
                        companies.Add(company);
                    }
                }

                return View("CompaniesList", companies);


            }

            //1==> Default Paramaters are chosen
            else
            {
                //Setting the hasParamActif to true
                paramFluxMA.hasParamFluxMA = true;

                db.Entry(paramFluxMA).State = EntityState.Modified;
                db.SaveChanges();


                return SelectCompany("0", "Default", 0);

            }

        }

        [Authorize]
        public ActionResult SelectCompany(string ownerId, string matricule, int exercice)
        {
            generalSettings gs = (generalSettings)Session["SteInformation"];

            var ParamSetting = from ps in db.ParametersSetting
                               where ps.ownerId.Equals(gs.ownerId) && ps.matricule.Equals(gs.matricule) && ps.exercice == gs.exercice
                               select ps;

            ParametersSetting paramFluxMA = ParamSetting.FirstOrDefault();
            paramFluxMA.hasParamFluxMA = true;

            db.Entry(paramFluxMA).State = EntityState.Modified;
            db.SaveChanges();

            var databaseFormulaList = from fl in db.FluxTresorerieMAFormula
                                      where fl.exercice == exercice && fl.matricule.Equals(matricule)
                                      select fl;

            foreach (var formula in databaseFormulaList)
            {
                FluxTresorerieMAFormula copy = new FluxTresorerieMAFormula();
                copy.ownerId = gs.ownerId;
                copy.matricule = gs.matricule;
                copy.exercice = gs.exercice;
                copy.codeParam = formula.codeParam;
                copy.codeDonnee = formula.codeDonnee;
                copy.nomCompte = formula.nomCompte;
                copy.typeFormule = formula.typeFormule;
                db.FluxTresorerieMAFormula.Add(copy);

            }
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        [Authorize]
        public ActionResult PrintFluxTresorerieMANotesAsPdf()
        {
            if (Session["SteInformation"] == null)
            {
                ViewBag.Error = "Vous devez d'abord choisir une entreprise ! <a href=\"/Home/Index\"> Tous les entreprises </a>";
                return View("FileError");
            }
            else if (Session["firstInputFile"] == null || Session["secondInputFile"] == null)
            {
                ViewBag.Error = "Vous devez d'abord importer les fichiers ! <a href=\"/Excel/Input\"> Importer </a>";
                return View("FileError");
            }

            generalSettings gs = (generalSettings)Session["SteInformation"];
            string matricule = gs.matricule;
            int exercice = gs.exercice;
            string ownerId = gs.ownerId;

            IEnumerable<ExcelInfo> firstFile = (IEnumerable<ExcelInfo>)Session["firstInputFile"];
            IEnumerable<ExcelInfo> secondFile = (IEnumerable<ExcelInfo>)Session["secondInputFile"];

            ViewBag.firstFile = firstFile;
            ViewBag.secondFile = secondFile;


            var af = from a in db.FluxTresorerieMAModel
                     where a.exercice == exercice && a.matricule.Equals(matricule) && a.ownerId.Equals(ownerId)
                     select a;
            ViewBag.info1 = gs.nomEtPrenomRaisonSociale;
            ViewBag.info2 = gs.adresse;
            ViewBag.info3 = gs.activite;
            ViewBag.info4 = gs.dateDebutExercice;
            ViewBag.info5 = gs.dateClotureExercice;


            var listFormula = from lf in db.FluxTresorerieMAFormula
                              where lf.matricule.Equals(matricule)
                              && lf.exercice == exercice
                              && lf.ownerId.Equals(ownerId)
                              select lf;

            ViewBag.listFormula = listFormula;

            var report = new ViewAsPdf("FluxTresorerieMANotesAsPdf", af)
            {
                PageOrientation = Rotativa.Options.Orientation.Landscape,
                PageSize = Rotativa.Options.Size.A4,
                CustomSwitches = "--footer-center \"  Créer le : " + DateTime.Now.Date.ToString("dd/MM/yyyy") + "  Page: [page]/[toPage]\"" + " --footer-spacing 1 --footer-font-name \"Segoe UI\""
            };
            return report;
        }

        [Authorize]
        public ActionResult GenerateNoteFiles()
        {
            if (Session["SteInformation"] == null)
            {
                ViewBag.Error = "Vous devez d'abord choisir une entreprise ! <a href=\"/Home/Index\"> Tous les entreprises </a>";
                return View("FileError");
            }
            else if (Session["firstInputFile"] == null || Session["secondInputFile"] == null)
            {
                ViewBag.Error = "Vous devez d'abord importer les fichiers ! <a href=\"/Excel/Input\"> Importer </a>";
                return View("FileError");
            }

            generalSettings gs = (generalSettings)Session["SteInformation"];
            int exercice = gs.exercice;
            string matricule = gs.matricule;
            string ownerId = gs.ownerId;

            var ActifParam = from ap in db.FluxTresorerieMAModel
                             where ap.exercice == exercice && ap.matricule.Equals(matricule) && ap.ownerId.Equals(ownerId)
                             select ap;
            if (!ActifParam.Any())
            {
                ViewBag.Error = " Vous devez d'abord configurer les paramétres Flux de Trésorerie - Modèle Autorisé  !  <a href=\"/FluxTresorerieMAParameters/Index\"> Flux de Trésorerie - Modèle Autorisé </a>";
                return View("FileError");
            }
            return View();
        }
    }
}
