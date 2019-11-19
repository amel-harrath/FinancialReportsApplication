using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.EntityFramework;
using Rotativa;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using System.Xml;
using TASEF.Infrastructure;
using TASEF.Migrations.OtherDbContext;
using TASEF.Models;

namespace TASEF.Controllers
{
    public class PassifParametersController : Controller
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
            else if (Session["firstInputFile"] == null || Session["secondInputFile"] == null || Session["ExcelModelView"] == null)
            {
                ViewBag.Error = "Vous devez d'abord importer les fichiers  !  <a href=\"/Excel/Input\"> Importer </a>";
                return View("FileError");
            }

            generalSettings gs = (generalSettings)Session["SteInformation"];

            var ParamSetting = from ps in db.ParametersSetting
                               where ps.ownerId.Equals(gs.ownerId) && ps.matricule.Equals(gs.matricule) && ps.exercice == gs.exercice
                               select ps;
            ParametersSetting paramPassif = ParamSetting.FirstOrDefault();

            if (!paramPassif.hasParamPassif)
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

                var af = from a in db.PassifModel
                         where a.exercice == exercice && a.matricule.Equals(matricule) && a.ownerId.Equals(ownerId)
                         select a;

                var databaseFormulaList = from fl in db.PassifFormula
                                          where fl.exercice == exercice && fl.matricule.Equals(matricule) && fl.ownerId.Equals(ownerId)
                                          select fl;


                //there are no Passif parameters for that company so we need to create them
                if (!af.Any())
                {
                    List<PassifParamModel> PassifParamList = new List<PassifParamModel>
            {
                new PassifParamModel(){ code="0001" , libelle ="Capitaux propres" , type="Formula",state="Stable",priority=1},
                new PassifParamModel(){ code="0002" , libelle ="Capital social" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0003" , libelle ="Réserves" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0004" , libelle ="Autres capitaux propres" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0005" , libelle ="Résultats reportés" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0006" , libelle ="Capitaux propres avant résultat de l'exercice" , type="Formula",state="Stable",priority=0},
                new PassifParamModel(){ code="0007" , libelle ="Résultat de l'exercice" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0008" , libelle ="Total Passifs" , type="Formula",state="Stable",priority=2},
                new PassifParamModel(){ code="0009" , libelle ="Passifs non courants" , type="Formula",state="Stable",priority=1},
                new PassifParamModel(){ code="0010" , libelle ="Emprunts" , type="Formula",state="Stable",priority=0},
                new PassifParamModel(){ code="0011" , libelle ="Emprunts obligataires (assortis de sûretés)" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0012" , libelle ="Empts auprès d'étab.Fin. (assortis de sûretés)" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0013" , libelle ="Empts auprès d'étab.Fin. (assorti de sûretés)" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0014" , libelle ="Empts et dettes assorties de condit. particulières" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0015" , libelle ="Emprunts non assortis de sûretés" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0016" , libelle ="Dettes rattachées à des participations" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0017" , libelle ="Dépôts  et  cautionnements reçus" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0018" , libelle ="Autres emprunts et dettes" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0019" , libelle ="Autres Passifs Financiers" , type="Formula",state="Stable",priority=0},
                new PassifParamModel(){ code="0020" , libelle ="Écarts de conversion" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0021" , libelle ="Autres passifs financiers" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0022" , libelle ="Provisions" , type="Formula",state="Stable",priority=0},
                new PassifParamModel(){ code="0023" , libelle ="Provisions pour risques" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0024" , libelle ="Prov.pour charges à répartir/plusieurs exercices" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0025" , libelle ="Prov.pour retraites et obligations similaires" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0026" , libelle ="Provisions d'origine réglementaire" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0027" , libelle ="Provisions pour impôts" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0028" , libelle ="Prov.pour renouvellement des immobilisations" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0029" , libelle ="Provisions pour amortissement" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0030" , libelle ="Autres provisions pour charges" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0031" , libelle ="Passifs courants" , type="Formula",state="Stable",priority=1},
                new PassifParamModel(){ code="0032" , libelle ="Fournisseurs et Comptes Rattachés" , type="Formula",state="Stable",priority=0},
                new PassifParamModel(){ code="0033" , libelle ="Fournisseurs d'exploitation" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0034" , libelle ="Fournisseurs d'exploitation - effets à payer" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0035" , libelle ="Fournisseurs d'immobilisations" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0036" , libelle ="Fournisseurs d'immobilisations - effets à payer" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0037" , libelle ="Fournisseurs - factures non parvenues" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0038" , libelle ="Autres passifs courants" , type="Formula",state="Stable",priority=0},
                new PassifParamModel(){ code="0039" , libelle ="Clients créditeurs" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0040" , libelle ="Sociétés du groupe  et  associés" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0041" , libelle ="État et collectivités publiques" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0042" , libelle ="Sociétés du groupe  et  associés" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0043" , libelle ="Débiteurs divers et Créditeurs divers" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0044" , libelle ="Comptes transitoires ou d'attente" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0045" , libelle ="Comptes de régularisation" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0046" , libelle ="Provisions courantes pour risques et charges" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0047" , libelle ="Concours Bancaires et Autres Passifs Financiers" , type="Formula",state="Stable",priority=0},
                new PassifParamModel(){ code="0048" , libelle ="Emprunts et autres dettes financières courants" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0049" , libelle ="Emprunts échus et impayés" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0050" , libelle ="Intérêts courus" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0051" , libelle ="Banques, établissements financiers et assimilés" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0052" , libelle ="Autres Postes des Capitaux Propres et Passifs du Bilan" , type="Calculated",state="Stable"},
                new PassifParamModel(){ code="0053" , libelle ="Total des capitaux propres et passifs" , type="Formula",state="Stable",priority=3}

            };
                    foreach (var param in PassifParamList)
                    {
                        var manager = new UserManager<ApplicationUser>(new UserStore<ApplicationUser>(new ApplicationDbContext()));
                        var currentUser = manager.FindById(User.Identity.GetUserId());
                        param.ownerId = currentUser.Id;
                        param.exercice = exercice;
                        param.matricule = matricule;
                        param.netN = 0;
                        param.netN1 = 0;
                        db.PassifModel.Add(param);
                        db.SaveChanges();
                    }
                    return View(PassifParamList);

                }
                else return View(af);
            }
       
        }

        //GET /Parameters/Create/0001
        [Authorize]
        public ActionResult Create(string id)
        {
            List<string> definedParam = new List<string>(new string[] { "0001","0006","0009","0010","0019","0022","0031","0032","0038","0047","0053" });
            if (definedParam.Contains(id))
            {
                ViewBag.Error = "Ces paramétres ne peuvent pas être modifier! <a href=\"/PassifParameters/Index\"> Paramétres </a>";
                return View("FileError");
            }
            generalSettings gs = (generalSettings)Session["SteInformation"];
            string matricule = gs.matricule;
            int exercice = gs.exercice;
            string ownerId = gs.ownerId;

            var listFormula = from lf in db.PassifFormula
                              where lf.codeParam == id
                              && lf.matricule.Equals(matricule)
                              && lf.exercice == exercice
                              && lf.ownerId.Equals(ownerId)
                              select lf;
            ViewBag.listFormula = listFormula;
            PassifFormula af = new PassifFormula() { codeParam = id };
            return View(af);
        }

        // POST: PassifFormulas/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [Authorize]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,codeParam,codeDonnee,nomCompte,typeFormule")] PassifFormula PassifFormula)
        {
            if (ModelState.IsValid)
            {
                generalSettings gs = (generalSettings)Session["SteInformation"];
                string matricule = gs.matricule;
                int exercice = gs.exercice;
                string ownerId = gs.ownerId;

                string code = PassifFormula.codeParam;
                PassifParamModel apm = db.PassifModel.Find(ownerId,code, exercice, matricule);
                apm.state = "Stable";

                db.Entry(apm).State = EntityState.Modified;
                db.SaveChanges();

                PassifFormula.exercice = exercice;
                PassifFormula.matricule = matricule;
                PassifFormula.ownerId = ownerId;

                db.PassifFormula.Add(PassifFormula);
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(PassifFormula);
        }
    

        [Authorize]
        public ActionResult Recalculate()
        {
            generalSettings gs = (generalSettings)Session["SteInformation"];
            int exercice = gs.exercice;
            string matricule = gs.matricule;
            string ownerId = gs.ownerId;

            var PassifParam = from ap in db.PassifModel
                              where ap.matricule.Equals(matricule) && ap.exercice == exercice && ap.ownerId.Equals(ownerId)
                              select ap;
            var PassifFormula = from af in db.PassifFormula
                                where af.matricule.Equals(matricule) && af.exercice == exercice && af.ownerId.Equals(ownerId)
                                select af;
            var calculated = from c in PassifParam
                             where c.type.Equals("Calculated")
                             select c;
            var formula = from c in PassifParam
                          where c.type.Equals("Formula")
                          orderby c.priority
                          select c;


            //List<Formula> PassifFormulaList = new List<Formula>
            //{
            //    new Formula(){ code="0001",type="Passif",parameters = new List<string>() {"0006","0007"}},
            //    new Formula(){ code="0006",type="Passif",parameters = new List<string>() {"0002","0003","0004","0005"}},
            //    new Formula(){ code="0008",type="Passif",parameters = new List<string>() {"0009","0031"}},
            //    new Formula(){ code="0009",type="Passif",parameters = new List<string>() {"0010","0019","0022"}},
            //    new Formula(){ code="0010",type="Passif",parameters = new List<string>() {"0011","0012","0013","0014","0015","0016","0017","0018"}},
            //    new Formula(){ code="0019",type="Passif",parameters = new List<string>() {"0020","0021"}},
            //    new Formula(){ code="0022",type="Passif",parameters = new List<string>() {"0023","0024","0025","0026","0027","0028","0029","0030"}},
            //    new Formula(){ code="0031",type="Passif",parameters = new List<string>() {"0032","0038","0047"}},
            //    new Formula(){ code="0032",type="Passif",parameters = new List<string>() {"0033","0034","0035","0036","0037"}},
            //    new Formula(){ code="0038",type="Passif",parameters = new List<string>() {"0039","0040","0041","0042","0043","0044","0045","0046"}},
            //    new Formula(){ code="0047",type="Passif",parameters = new List<string>() {"0048","0049","0050","0051"}},
            //    new Formula(){ code="0053",type="Passif",parameters = new List<string>() {"0001","0008","0052"}}
            //};

           

            var PassifFormulaList = from pf in db.DefinedFormulas
                                    where pf.type.Equals("Passif")
                                    select pf;

            IEnumerable < ExcelInfo > firstInputFile = (IEnumerable<ExcelInfo>)Session["firstInputFile"];
            IEnumerable<ExcelInfo> secondInputFile = (IEnumerable<ExcelInfo>)Session["secondInputFile"];

            IEnumerable<ExcelInfo> InputFile = firstInputFile.Concat(secondInputFile);

            foreach (var param in calculated.ToList())
            {
                if (param.state.Equals("Stable"))
                {
                    float netN = 0;
                    float netN1 = 0;
                    float valeur = 0;
                    string code = param.code;
                    

                    var one = PassifFormula.Where(AF => AF.codeParam.Equals(code));
                    var two = one.Where(AF => AF.matricule.Equals(matricule));
                    var three = two.Where(AF => AF.matricule.Equals(ownerId));
                    var specificPassifFormula = two.Where(AF => AF.exercice == exercice);
                    foreach (var formulas in specificPassifFormula)
                    {
                        if (formulas.typeFormule.Equals("Solde"))
                        {

                            foreach (var input in InputFile)
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
                            foreach (var input in InputFile)
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
                            foreach (var input in InputFile)
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
                            foreach (var input in InputFile)
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
                            foreach (var input in InputFile)
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

                var f = from fo in PassifFormulaList
                        where fo.code.Equals(param.code)
                        select fo;

                List<string> parameters = new List<string>();

                foreach (Formula form in f.ToList())
                {
                    parameters.Add(form.parameter);
                }

                //Formula formulas = f.FirstOrDefault();
                foreach (string code in parameters/*formulas.parameters.ToList()*/)
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
                    PassifParamModel apm = PassifParam.Where(AF => AF.code.Equals(cleanCode)).FirstOrDefault();

                    param.netN += apm.netN;
                    param.netN1 += apm.netN1;
                }
                param.ownerId = ownerId;
                param.exercice = exercice;
                param.matricule = matricule;
                db.Entry(param).State = EntityState.Modified;
                db.SaveChanges();
            }
            return RedirectToAction("Index");
        }

        [Authorize]
        public ActionResult Show(string id)
        {


            //List<Formula> PassifFormulaList = new List<Formula>
            //{
            //    new Formula(){ code="0001",type="Passif",parameters = new List<string>() {"0006","0007"}},
            //    new Formula(){ code="0006",type="Passif",parameters = new List<string>() {"0002","0003","0004","0005"}},
            //    new Formula(){ code="0009",type="Passif",parameters = new List<string>() {"0010","0019","0022"}},
            //    new Formula(){ code="0010",type="Passif",parameters = new List<string>() {"0011","0012","0013","0014","0015","0016","0017","0018"}},
            //    new Formula(){ code="0019",type="Passif",parameters = new List<string>() {"0020","0021"}},
            //    new Formula(){ code="0022",type="Passif",parameters = new List<string>() {"0023","0024","0025","0026","0027","0028","0029","0030"}},
            //    new Formula(){ code="0031",type="Passif",parameters = new List<string>() {"0032","0038","0047"}},
            //    new Formula(){ code="0032",type="Passif",parameters = new List<string>() {"0033","0034","0035","0036","0037"}},
            //    new Formula(){ code="0038",type="Passif",parameters = new List<string>() {"0039","0040","0041","0042","0043","0044","0045","0046"}},
            //    new Formula(){ code="0047",type="Passif",parameters = new List<string>() {"0048","0049","0050","0051"}},
            //    new Formula(){ code="0053",type="Passif",parameters = new List<string>() {"0001","0008","0052"}}
            //};



            //Formula specificFormula = PassifFormulaList.Where(AF => AF.code.Equals(id)).FirstOrDefault();
            var PassifFormulaList = from pf in db.DefinedFormulas
                                    where pf.type.Equals("Passif")
                                    select pf;

            var f = from fo in PassifFormulaList
                    where fo.code.Equals(id)
                    select fo;

            List<string> parameters = new List<string>();
            foreach (Formula form in f.ToList())
            {
                parameters.Add(form.parameter);
            }

            List<String> CleanParameters = new List<string>();

            List<string> minus = new List<string>();
            foreach (string code in parameters /*specificFormula.parameters.ToList()*/)
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
            var PassifFormulas = from af in db.PassifModel
                                           where af.exercice == gs.exercice &&
                                                 af.matricule.Equals(gs.matricule) &&
                                                 af.ownerId.Equals(gs.ownerId) &&
                                                 /*specificFormula.parameters*/CleanParameters.Contains(af.code)
                                           select af;

            ViewBag.minus = minus;
            return View(PassifFormulas);
        }

        [Authorize]
        public ActionResult EditParam(string ownerId,string code, string exercice, string matricule)
        {
            if ((code == null) || (exercice == null) || (matricule == null))
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            PassifParamModel PassifParamModel = db.PassifModel.Find(ownerId,code, Int32.Parse(exercice), matricule);
            if (PassifParamModel == null)
            {
                return HttpNotFound();
            }
            return View(PassifParamModel);
        }

        [HttpPost]
        [Authorize]
        [ValidateAntiForgeryToken]
        public ActionResult EditParam([Bind(Include = "code,ownerId,libelle,netN,netN1,type,exercice,matricule,state")] PassifParamModel param)
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
        public ActionResult PrintPassifsAsPdf()
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

            var af = from a in db.PassifModel
                     where a.exercice == exercice && a.matricule.Equals(matricule) && a.ownerId.Equals(ownerId)
                     select a;
            ViewBag.info1 = gs.nomEtPrenomRaisonSociale;
            ViewBag.info2 = gs.adresse;
            ViewBag.info3 = gs.activite;
            ViewBag.info4 = gs.dateDebutExercice;
            ViewBag.info5 = gs.dateClotureExercice;
            var report = new ViewAsPdf("PassifsAsPdf", af)
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
            PassifFormula PassifFormula = db.PassifFormula.Find(Int32.Parse(id));
            if (PassifFormula == null)
            {
                return HttpNotFound();
            }
            return View(PassifFormula);
        }

        [HttpPost]
        [Authorize]
        [ValidateAntiForgeryToken]
        public ActionResult EditFormula([Bind(Include = "")] PassifFormula PassifFormula)
        {
            if (ModelState.IsValid)
            {
                generalSettings gs = (generalSettings)Session["SteInformation"];
                int exercice = gs.exercice;
                string matricule = gs.matricule;
                string ownerId = gs.ownerId;

                string code = PassifFormula.codeParam;
                PassifParamModel apm = db.PassifModel.Find(ownerId,code, exercice, matricule);
                apm.state = "Stable";

                db.Entry(apm).State = EntityState.Modified;
                db.SaveChanges();

                db.Entry(PassifFormula).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(PassifFormula);
        }

        [Authorize]
        public ActionResult Delete(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            PassifFormula PassifFormula = db.PassifFormula.Find(Int32.Parse(id));
            if (PassifFormula == null)
            {
                return HttpNotFound();
            }
            return View(PassifFormula);
        }

        [Authorize]
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string id)
        {
            PassifFormula PassifFormula = db.PassifFormula.Find(Int32.Parse(id));

            generalSettings gs = (generalSettings)Session["SteInformation"];
            int exercice = gs.exercice;
            string matricule = gs.matricule;
            string ownerId = gs.ownerId;

            string code = PassifFormula.codeParam;
            PassifParamModel apm = db.PassifModel.Find(ownerId,code, exercice, matricule);
            apm.state = "Stable";

            db.Entry(apm).State = EntityState.Modified;
            db.SaveChanges();

            db.PassifFormula.Remove(PassifFormula);
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
            else if (Session["firstInputFile"] == null || Session["secondInputFile"] == null)
            {
                ViewBag.Error = "Vous devez d'abord importer les fichiers ! <a href=\"/Excel/Input\"> Importer </a>";
                return View("FileError");
            }

            generalSettings gs = (generalSettings)Session["SteInformation"];
            int exercice = gs.exercice;
            string matricule = gs.matricule;
            string ownerId = gs.ownerId;

            var PassifParam = from ap in db.PassifModel
                             where ap.exercice == exercice && ap.matricule.Equals(matricule) && ap.ownerId.Equals(ownerId)
                             select ap;
            if (!PassifParam.Any())
            {
                ViewBag.Error = "Vous devez d'abord configurer les paramétres passifs! <a href=\"/PassifParameters/Index\"> Paramétrage Passifs  </a>";
                return View("FileError");
            }
            return View();
        }

        [Authorize]
        public void PrintPassifsAsXml()
        {

            generalSettings gs = (generalSettings)Session["SteInformation"];
            var PassifParam = from a in db.PassifModel
                              where a.exercice == gs.exercice && a.matricule.Equals(gs.matricule) && a.ownerId.Equals(gs.ownerId)
                             select a;
            string fileName = "Passif-" + gs.matricule + "-" + gs.exercice + ".xml";

            using (MemoryStream stream = new MemoryStream())
            {
                // Create an XML document. Write our specific values into the document.
                XmlTextWriter xmlWriter = new XmlTextWriter(stream, System.Text.Encoding.UTF8);
                // Write the XML document header.
                xmlWriter.WriteStartDocument();
                xmlWriter.WriteRaw("<?xml-stylesheet type=\"text/xsl\" href=\"F6002.xsl\"?>");
                xmlWriter.WriteStartElement("lf:F6002");
                xmlWriter.WriteAttributeString("xmlns:lf", "http://www.impots.finances.gov.tn/liasse");
                xmlWriter.WriteAttributeString("xmlns:vc", "http://www.w3.org/2007/XMLSchema-versioning");
                xmlWriter.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
                xmlWriter.WriteAttributeString("xsi:schemaLocation", "http://www.impots.finances.gov.tn/liasse F6002.xsd");
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
                    int add = i * 53;
                    string code;
                    foreach (var param in PassifParam)
                    {
                        code = (Int32.Parse(param.code) + add).ToString();
                        code = code.PadLeft(4, '0');
                         if (i == 0)
                        {
                            xmlWriter.WriteElementString($"lf:{code}", param.netN.ToString());
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

            ParametersSetting paramPassif = ParamSetting.FirstOrDefault();

            //3==> Specific parameters are chosen
            if (id == 3)
            {
                //Setting the hasParamActif to true
                paramPassif.hasParamPassif = true;

                db.Entry(paramPassif).State = EntityState.Modified;
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
                paramPassif.hasParamPassif = true;

                db.Entry(paramPassif).State = EntityState.Modified;
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

            ParametersSetting paramPassif = ParamSetting.FirstOrDefault();
            paramPassif.hasParamPassif = true;

            db.Entry(paramPassif).State = EntityState.Modified;
            db.SaveChanges();


            

            var databaseFormulaList = from fl in db.PassifFormula
                                      where fl.exercice == exercice && fl.matricule.Equals(matricule)
                                      select fl;
            
            foreach (var formula in databaseFormulaList)
            {
               PassifFormula copy = new PassifFormula();
                copy.ownerId = gs.ownerId;
                copy.matricule = gs.matricule;
                copy.exercice = gs.exercice;
                copy.codeParam = formula.codeParam;
                copy.codeDonnee = formula.codeDonnee;
                copy.nomCompte = formula.nomCompte;
                copy.typeFormule = formula.typeFormule;
                db.PassifFormula.Add(copy);

            }
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        [Authorize]
        public ActionResult PrintPassifNotesAsPdf()
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

            IEnumerable<ExcelInfo> firstFile = (IEnumerable<ExcelInfo>)Session["firstInputFile"];
            IEnumerable<ExcelInfo> secondFile = (IEnumerable<ExcelInfo>)Session["secondInputFile"];

            ViewBag.firstFile = firstFile;
            ViewBag.secondFile = secondFile;


            var af = from a in db.ActifModel
                     where a.exercice == exercice && a.matricule.Equals(matricule) && a.ownerId.Equals(ownerId)
                     select a;
            ViewBag.info1 = gs.nomEtPrenomRaisonSociale;
            ViewBag.info2 = gs.adresse;
            ViewBag.info3 = gs.activite;
            ViewBag.info4 = gs.dateDebutExercice;
            ViewBag.info5 = gs.dateClotureExercice;



            var listFormula = from lf in db.PassifFormula
                              where lf.matricule.Equals(matricule)
                              && lf.exercice == exercice
                              && lf.ownerId.Equals(ownerId)
                              select lf;

            ViewBag.listFormula = listFormula;

            var report = new ViewAsPdf("PassifNotesAsPdf", af)
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
            else if (Session["firstInputFile"] == null || Session["secondInputFile"] == null || Session["ExcelModelView"] == null)
            {
                ViewBag.Error = "Vous devez d'abord importer les fichiers  ! <a href=\"/Excel/Input\"> Importer </a>";
                return View("FileError");
            }


            generalSettings gs = (generalSettings)Session["SteInformation"];
            int exercice = gs.exercice;
            string matricule = gs.matricule;
            string ownerId = gs.ownerId;

            var PassifParam = from ap in db.PassifModel
                             where ap.exercice == exercice && ap.matricule.Equals(matricule) && ap.ownerId.Equals(ownerId)
                             select ap;
            if (!PassifParam.Any())
            {
                ViewBag.Error = "Vous devez d'abord configurer les paramétres passifs! <a href=\"/PassifParameters/Index\"> Paramétrage Passifs  </a>";
                return View("FileError");
            }
            return View();
        }
    }
}
