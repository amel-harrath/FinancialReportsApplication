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
using TASEF.Models;

namespace TASEF.Controllers
{
    public class ParametersController : Controller
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
            ParametersSetting paramActif = ParamSetting.FirstOrDefault();

            if (!paramActif.hasParamActif)
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

                var af = from a in db.ActifModel
                         where a.exercice == exercice && a.matricule.Equals(matricule) && a.ownerId.Equals(ownerId)
                         select a;

                var databaseFormulaList = from fl in db.ActifFormula
                                          where fl.exercice == exercice && fl.matricule.Equals(matricule) && fl.ownerId.Equals(ownerId)
                                          select fl;


                //there are no actif parameters for that company so we need to create them
                if (!af.Any())
                {
                    List<ActifParamModel> actifParamList = new List<ActifParamModel>
            {
                new ActifParamModel(){ code="0001" , libelle ="Actifs non courants" , type="Formula",state="Stable",priority=3},
                new ActifParamModel(){ code="0002" , libelle ="Actifs immobilises" , type="Formula",state="Stable",priority=2},
                new ActifParamModel(){ code="0003" , libelle ="Immobilisations Incorporelles" , type="Formula",state="Stable",priority=1},
                new ActifParamModel(){ code="0004" , libelle ="Investissement recherche et développement" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0005" , libelle ="Concess. marque,brevet,licence,marque" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0006" , libelle ="Logiciels" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0007" , libelle ="Fonds commercial" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0008" , libelle ="Droit au bail" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0009" , libelle ="Autres Immobilisations Incorporelles" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0010" , libelle ="Immobilisations Incorporelles en cours" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0011" , libelle ="Av. et Ac.  Verses/Cmde.Immob.Incorp." , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0012" , libelle ="Immobilisations corporelles" , type="Formula",state="Stable",priority=1},
                new ActifParamModel(){ code="0013" , libelle ="Terrains" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0014" , libelle ="Constructions" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0015" , libelle ="Inst. Tech., materiel et outillages Industriels" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0016" , libelle ="Materiel de transport" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0017" , libelle ="Autres Immobilisations Corporelles" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0018" , libelle ="Immob. Corporelles en cours" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0019" , libelle ="Av. et Ac. Verses/Commande Immob.Corp." , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0020" , libelle ="Immob. a statut juridique particulier" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0021" , libelle ="Immobilisations Financieres" , type="Formula",state="Stable",priority=1},
                new ActifParamModel(){ code="0022" , libelle ="Actions" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0023" , libelle ="Autres creances rattach. a des participat." , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0024" , libelle ="Creances rattach. a des stes en participat." , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0025" , libelle ="Vers.a eff./titre de participation non liberes" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0026" , libelle ="Titres immobilises (droit de propriete)" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0027" , libelle ="Titres immobilises (droit de creance)" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0028" , libelle ="Depots et cautionnements verses" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0029" , libelle ="Autres creances immobilisees" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0030" , libelle ="Vers.a eff./Titres immobilises non liberes" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0031" , libelle ="Autres Actifs Non Courants" , type="Formula",state="Stable",priority=1},
                new ActifParamModel(){ code="0032" , libelle ="Frais preliminaires" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0033" , libelle ="Charges a repartir" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0034" , libelle ="Frais d'emission et primes de Remb. Empts" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0035" , libelle ="ecarts de conversion" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0036" , libelle ="Actifs courants" , type="Formula",state="Stable",priority=2},
                new ActifParamModel(){ code="0037" , libelle ="Stocks" , type="Formula",state="Stable",priority=1},
                new ActifParamModel(){ code="0038" , libelle ="Stocks Matieres Premieres et Fournit. Liees" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0039" , libelle ="Stocks Autres Approvisionnements" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0040" , libelle ="Stocks En-cours de production de biens" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0041" , libelle ="Stocks En-cours de production services" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0042" , libelle ="Stocks de produits" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0043" , libelle ="Stocks de marchandises" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0044" , libelle ="Clients et Comptes Rattaches" , type="Formula",state="Stable",priority=1},
                new ActifParamModel(){ code="0045" , libelle ="Clients  et  comptes rattaches" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0046" , libelle ="Clients - effets a recevoir" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0047" , libelle ="Clients douteux ou litigieux" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0048" , libelle ="Creances/travaux non encore facturables" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0049" , libelle ="Clt-pdts non encore factures (pdt a recev.)" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0050" , libelle ="Autres Actifs Courants" , type="Formula",state="Stable",priority=1},
                new ActifParamModel(){ code="0051" , libelle ="Fournisseurs debiteurs" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0052" , libelle ="Personnel et comptes rattaches" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0053" , libelle ="etat et collectivites publiques" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0054" , libelle ="Societes du groupe  et  associes" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0055" , libelle ="Debiteurs divers et Crediteurs divers" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0056" , libelle ="Comptes transitoires ou d'attente" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0057" , libelle ="Comptes de regularisation" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0058" , libelle ="Autres" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0059" , libelle ="Placements et Autres Actifs Financiers" , type="Formula",state="Stable",priority=1},
                new ActifParamModel(){ code="0060" , libelle ="Prets et autres creances Fin. courants" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0061" , libelle ="Placements courants" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0062" , libelle ="Regies d'avances et accreditifs" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0063" , libelle ="Autres" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0064" , libelle ="Liquidites et equivalents de liquidites" , type="Formula",state="Stable",priority=1},
                new ActifParamModel(){ code="0065" , libelle ="Banques, etabl. Financiers et assimiles" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0066" , libelle ="Caisse" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0067" , libelle ="Autres Postes des Actifs du Bilan" , type="Calculated",state="Stable"},
                new ActifParamModel(){ code="0068" , libelle ="Total des actifs" , type="Formula",state="Stable",priority=4}

            };
                    foreach (var param in actifParamList)
                    {
                        var manager = new UserManager<ApplicationUser>(new UserStore<ApplicationUser>(new ApplicationDbContext()));
                        var currentUser = manager.FindById(User.Identity.GetUserId());
                        param.ownerId = currentUser.Id;
                        param.exercice = exercice;
                        param.matricule = matricule;
                        param.brutN = 0;
                        param.amortProvN = 0;
                        param.netN = 0;
                        param.netN1 = 0;
                        db.ActifModel.Add(param);
                        db.SaveChanges();
                    }
                    return View(actifParamList);

                }
                else return View(af);

            }

        }

        //GET /Parameters/Create/0001
        [Authorize]
        public ActionResult Create(string id)
        {
            List<string> definedParam = new List<string>(new string[] { "0001", "0002", "0003", "0012", "0021", "0031", "0036", "0037", "0044", "0050", "0059", "0064", "0068" });
            if (definedParam.Contains(id))
            {
                ViewBag.Error = "Ces paramétres ne peuvent pas être modifier! <a href=\"/Parameters/Index\"> Paramétres </a>";
                return View("FileError");
            }
            generalSettings gs = (generalSettings)Session["SteInformation"];
            string matricule = gs.matricule;
            int exercice = gs.exercice;
            string ownerId = gs.ownerId;


            var listFormula = from lf in db.ActifFormula
                              where lf.codeParam == id
                              && lf.matricule.Equals(matricule)
                              && lf.exercice == exercice
                              && lf.ownerId.Equals(ownerId)
                              select lf;
            ViewBag.listFormula = listFormula;
            ActifFormula af = new ActifFormula() { codeParam = id };
            return View(af);
        }

        // POST: ActifFormulas/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [Authorize]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,codeParam,codeDonnee,nomCompte,typeFormule,genre")] ActifFormula actifFormula)
        {
            if (ModelState.IsValid)
            {
                generalSettings gs = (generalSettings)Session["SteInformation"];
                string matricule = gs.matricule;
                int exercice = gs.exercice;
                string ownerId = gs.ownerId;

                string code = actifFormula.codeParam;
                ActifParamModel apm = db.ActifModel.Find(ownerId, code, exercice, matricule);
                apm.state = "Stable";

                db.Entry(apm).State = EntityState.Modified;
                db.SaveChanges();

                actifFormula.exercice = exercice;
                actifFormula.matricule = matricule;
                actifFormula.ownerId = ownerId;

                db.ActifFormula.Add(actifFormula);
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(actifFormula);
        }

        [Authorize]
        public ActionResult Recalculate()
        {
            generalSettings gs = (generalSettings)Session["SteInformation"];
            int exercice = gs.exercice;
            string matricule = gs.matricule;
            string ownerId = gs.ownerId;

            var ActifParam = from ap in db.ActifModel
                             where ap.matricule.Equals(matricule) && ap.exercice == exercice && ap.ownerId.Equals(ownerId)
                             select ap;
            var ActifFormula = from af in db.ActifFormula
                               where af.matricule.Equals(matricule) && af.exercice == exercice && af.ownerId.Equals(ownerId)
                               select af;
            var calculated = from c in ActifParam
                             where c.type.Equals("Calculated")
                             select c;
            var formula = from c in ActifParam
                          where c.type.Equals("Formula")
                          orderby c.priority
                          select c;

            //List<Formula> actifFormulaList = new List<Formula>
            //{
            //    new Formula(){ code="0001",type="Actif",parameters = new List<string>() {"0002","0031"}},
            //    new Formula(){ code="0002",type="Actif",parameters = new List<string>() {"0003","0012","0021"}},
            //    new Formula(){ code="0003",type="Actif",parameters = new List<string>() {"0004","0005","0006","0007","0008","0009","0010","0011"}},
            //    new Formula(){ code="0012",type="Actif",parameters = new List<string>() {"0013","0014","0015","0016","0017","0018","0019","0020"}},
            //    new Formula(){ code="0021",type="Actif",parameters = new List<string>() {"0022","0023","0024","0025","0026","0027","0028","0029","0030"}},
            //    new Formula(){ code="0031",type="Actif",parameters = new List<string>() {"0032","0033","0034","0035"}},
            //    new Formula(){ code="0036",type="Actif",parameters = new List<string>() {"0037","0044","0050","0059","0064"}},
            //    new Formula(){ code="0037",type="Actif",parameters = new List<string>() {"0038","0039","0040","0041","0042","0043"}},
            //    new Formula(){ code="0044",type="Actif",parameters = new List<string>() {"0045","0046","0047","0048","0049"}},
            //    new Formula(){ code="0050",type="Actif",parameters = new List<string>() {"0051","0052","0053","0054","0055","0056","0057","0058"}},
            //    new Formula(){ code="0059",type="Actif",parameters = new List<string>() {"0060","0061","0062","0063"}},
            //    new Formula(){ code="0064",type="Actif",parameters = new List<string>() {"0065","0066"}},
            //    new Formula(){ code="0068",type="Actif",parameters = new List<string>() {"0001","0031"}}
            //};

            
            var actifFormulaList = from f in db.DefinedFormulas
                                   where f.type.Equals("Actif")
                                   select f;

            IEnumerable<ExcelInfo> firstInputFile = (IEnumerable<ExcelInfo>)Session["firstInputFile"];
            IEnumerable<ExcelInfo> secondInputFile = (IEnumerable<ExcelInfo>)Session["secondInputFile"];

            IEnumerable<ExcelInfo> InputFile = firstInputFile.Concat(secondInputFile);

            foreach (var param in calculated.ToList())
            {
                if (param.state.Equals("Stable"))
                {
                    float sumBrutN = 0;
                    float sumAmorN = 0;
                    float sumBrutN1 = 0;
                    float sumAmorN1 = 0;
                    float valeur = 0;
                    string code = param.code;


                    var one = ActifFormula.Where(AF => AF.codeParam.Equals(code));
                    var two = one.Where(AF => AF.matricule.Equals(matricule));
                    var three = two.Where(AF => AF.matricule.Equals(ownerId));
                    var specificActifFormula = two.Where(AF => AF.exercice == exercice);
                    foreach (var formulas in specificActifFormula)
                    {
                        if (formulas.typeFormule.Equals("Solde"))
                        {

                            foreach (var input in InputFile)
                            {
                                if (input.compte.StartsWith(formulas.codeDonnee))
                                {
                                    valeur = input.debit - input.credit;
                                    if ((formulas.genre.Equals("Amor")) && (input.periode == 2))
                                        sumAmorN += valeur;
                                    else if ((formulas.genre.Equals("Amor")) && (input.periode == 1))
                                        sumAmorN1 += valeur;
                                    else if ((formulas.genre.Equals("Brut")) && (input.periode == 2))
                                        sumBrutN += valeur;
                                    else if ((formulas.genre.Equals("Brut")) && (input.periode == 1))
                                        sumBrutN1 += valeur;
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
                                    if ((formulas.genre.Equals("Amor")) && (input.periode == 2))
                                        sumAmorN += valeur;
                                    else if ((formulas.genre.Equals("Amor")) && (input.periode == 1))
                                        sumAmorN1 += valeur;
                                    else if ((formulas.genre.Equals("Brut")) && (input.periode == 2))
                                        sumBrutN += valeur;
                                    else if ((formulas.genre.Equals("Brut")) && (input.periode == 1))
                                        sumBrutN1 += valeur;
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
                                    if ((formulas.genre.Equals("Amor")) && (input.periode == 2))
                                        sumAmorN += valeur;
                                    else if ((formulas.genre.Equals("Amor")) && (input.periode == 1))
                                        sumAmorN1 += valeur;
                                    else if ((formulas.genre.Equals("Brut")) && (input.periode == 2))
                                        sumBrutN += valeur;
                                    else if ((formulas.genre.Equals("Brut")) && (input.periode == 1))
                                        sumBrutN1 += valeur;
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
                                        if ((formulas.genre.Equals("Amor")) && (input.periode == 2))
                                            sumAmorN += valeur;
                                        else if ((formulas.genre.Equals("Amor")) && (input.periode == 1))
                                            sumAmorN1 += valeur;
                                        else if ((formulas.genre.Equals("Brut")) && (input.periode == 2))
                                            sumBrutN += valeur;
                                        else if ((formulas.genre.Equals("Brut")) && (input.periode == 1))
                                            sumBrutN1 += valeur;
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
                                        if ((formulas.genre.Equals("Amor")) && (input.periode == 2))
                                            sumAmorN += valeur;
                                        else if ((formulas.genre.Equals("Amor")) && (input.periode == 1))
                                            sumAmorN1 += valeur;
                                        else if ((formulas.genre.Equals("Brut")) && (input.periode == 2))
                                            sumBrutN += valeur;
                                        else if ((formulas.genre.Equals("Brut")) && (input.periode == 1))
                                            sumBrutN1 += valeur;
                                    }
                                }
                            }
                        }
                    }
                    param.ownerId = ownerId;
                    param.exercice = exercice;
                    param.matricule = matricule;
                    param.brutN = sumBrutN;
                    param.amortProvN = sumAmorN;
                    param.netN = sumBrutN - sumAmorN;
                    param.netN1 = sumBrutN1 - sumAmorN1;
                    db.Entry(param).State = EntityState.Modified;
                    db.SaveChanges();


                }

            }
            foreach (var param in formula.ToList())
            {

                param.brutN = 0;
                param.amortProvN = 0;
                param.netN = 0;
                param.netN1 = 0;

                var f = from fo in actifFormulaList
                        where fo.code.Equals(param.code)
                        select fo;
                List<string> parameters = new List<string>();

                foreach (Formula form in f.ToList())
                {
                    parameters.Add(form.parameter);
                }

                //Formula formulas = f.FirstOrDefault();
                foreach (var code in parameters/*formulas.parameters*/)
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
                    ActifParamModel apm = ActifParam.Where(AF => AF.code.Equals(cleanCode)).FirstOrDefault();
                    param.brutN += apm.brutN;
                    param.amortProvN += apm.amortProvN;
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
            //List<Formula> actifFormulaList = new List<Formula>
            //{
            //    new Formula(){ code="0001",type="Actif",parameters = new List<string>() {"0002","0031"}},
            //    new Formula(){ code="0002",type="Actif",parameters = new List<string>() {"0003","0012","0021"}},
            //    new Formula(){ code="0003",type="Actif",parameters = new List<string>() {"0004","0005","0006","0007","0008","0009","0010","0011"}},
            //    new Formula(){ code="0012",type="Actif",parameters = new List<string>() {"0013","0014","0015","0016","0017","0018","0019","0020"}},
            //    new Formula(){ code="0021",type="Actif",parameters = new List<string>() {"0022","0023","0024","0025","0026","0027","0028","0029","0030"}},
            //    new Formula(){ code="0031",type="Actif",parameters = new List<string>() {"0032","0033","0034","0035"}},
            //    new Formula(){ code="0036",type="Actif",parameters = new List<string>() {"0037","0044","0050","0059","0064"}},
            //    new Formula(){ code="0037",type="Actif",parameters = new List<string>() {"0038","0039","0040","0041","0042","0043"}},
            //    new Formula(){ code="0044",type="Actif",parameters = new List<string>() {"0045","0046","0047","0048","0049"}},
            //    new Formula(){ code="0050",type="Actif",parameters = new List<string>() {"0051","0052","0053","0054","0055","0056","0057","0058"}},
            //    new Formula(){ code="0059",type="Actif",parameters = new List<string>() {"0060","0061","0062","0063"}},
            //    new Formula(){ code="0064",type="Actif",parameters = new List<string>() {"0065","0066"}},
            //    new Formula(){ code="0068",type="Actif",parameters = new List<string>() {"0001","0031"}}
            //};

            var ActifFormulaList = from pf in db.DefinedFormulas
                                    where pf.type.Equals("Actif")
                                    select pf;

            var f = from fo in ActifFormulaList
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
            var actifFormulas = from af in db.ActifModel
                                where af.exercice == gs.exercice &&
                                      af.matricule.Equals(gs.matricule) &&
                                      af.ownerId.Equals(gs.ownerId) &&
                                      /*specificFormula.parameters*/CleanParameters.Contains(af.code)
                                select af;
            ViewBag.minus = minus;
            return View(actifFormulas);
        }

        [Authorize]
        public ActionResult EditParam(string ownerId, string code, string exercice, string matricule)
        {
            if ((code == null) || (exercice == null) || (matricule == null))
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ActifParamModel actifParamModel = db.ActifModel.Find(ownerId, code, Int32.Parse(exercice), matricule);
            if (actifParamModel == null)
            {
                return HttpNotFound();
            }
            return View(actifParamModel);
        }

        [HttpPost]
        [Authorize]
        [ValidateAntiForgeryToken]
        public ActionResult EditParam([Bind(Include = "code,ownerId,libelle,brutN,amortProvN,netN,netN1,type,exercice,matricule,ownerId,state")] ActifParamModel param)
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
        public ActionResult PrintActifsAsPdf()
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

            var af = from a in db.ActifModel
                     where a.exercice == exercice && a.matricule.Equals(matricule) && a.ownerId.Equals(ownerId)
                     select a;
            ViewBag.info1 = gs.nomEtPrenomRaisonSociale;
            ViewBag.info2 = gs.adresse;
            ViewBag.info3 = gs.activite;
            ViewBag.info4 = gs.dateDebutExercice;
            ViewBag.info5 = gs.dateClotureExercice;
            var report = new ViewAsPdf("ActifsAsPdf", af)
            {
                PageOrientation = Rotativa.Options.Orientation.Landscape,
                PageSize = Rotativa.Options.Size.A4,
                CustomSwitches = "--footer-center \"  Créer le : " + DateTime.Now.Date.ToString("dd/MM/yyyy") + "  Page: [page]/[toPage]\"" + " --footer-spacing 1 --footer-font-name \"Segoe UI\""
            };
            return report;
        }

        [Authorize]
        public void PrintActifsAsXml()
        {

            generalSettings gs = (generalSettings)Session["SteInformation"];
            var ActifParam = from a in db.ActifModel
                             where a.exercice == gs.exercice && a.matricule.Equals(gs.matricule) && a.ownerId.Equals(gs.ownerId)
                             select a;
            string fileName = "Actif-" + gs.matricule + "-" + gs.exercice + ".xml";

            using (MemoryStream stream = new MemoryStream())
            {
                // Create an XML document. Write our specific values into the document.
                XmlTextWriter xmlWriter = new XmlTextWriter(stream, System.Text.Encoding.UTF8);
                // Write the XML document header.
                xmlWriter.WriteStartDocument();
                xmlWriter.WriteRaw("<?xml-stylesheet type=\"text/xsl\"?>");
                xmlWriter.WriteStartElement("lf:F6001");
                xmlWriter.WriteAttributeString("xmlns:lf", "http://www.impots.finances.gov.tn/liasse");
                xmlWriter.WriteAttributeString("xmlns:vc", "http://www.w3.org/2007/XMLSchema-versioning");
                xmlWriter.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
                xmlWriter.WriteAttributeString("xsi:schemaLocation", "http://www.impots.finances.gov.tn/liasse F6001.xsd");
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

                for (int i = 0; i < 4; i++)
                {
                    int add = i * 1000;
                    string code;
                    foreach (var param in ActifParam)
                    {
                        code = (Int32.Parse(param.code) + add).ToString();
                        if (i == 0)
                        {
                            xmlWriter.WriteElementString($"lf:{param.code}", param.brutN.ToString());
                        }
                        else if (i == 1)
                        {
                            xmlWriter.WriteElementString($"lf:{code}", param.amortProvN.ToString());
                        }
                        else if (i == 2)
                        {
                            xmlWriter.WriteElementString($"lf:{code}", param.netN.ToString());
                        }
                        else if (i == 3)
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

            var ActifParam = from ap in db.ActifModel
                             where ap.exercice == exercice && ap.matricule.Equals(matricule) && ap.ownerId.Equals(ownerId)
                             select ap;
            if (!ActifParam.Any())
            {
                ViewBag.Error = "Vous devez d'abord configurer les paramétres Actifs! <a href=\"/Parameters/Index\"> Paramétrage Actifs  </a>";
                return View("FileError");
            }
            return View();
        }

        [Authorize]
        public ActionResult EditFormula(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ActifFormula actifFormula = db.ActifFormula.Find(Int32.Parse(id));
            if (actifFormula == null)
            {
                return HttpNotFound();
            }
            return View(actifFormula);
        }

        [HttpPost]
        [Authorize]
        [ValidateAntiForgeryToken]
        public ActionResult EditFormula([Bind(Include = "")] ActifFormula actifFormula)
        {
            if (ModelState.IsValid)
            {
                generalSettings gs = (generalSettings)Session["SteInformation"];
                int exercice = gs.exercice;
                string matricule = gs.matricule;
                string ownerId = gs.ownerId;

                string code = actifFormula.codeParam;
                ActifParamModel apm = db.ActifModel.Find(ownerId, code, exercice, matricule);
                apm.state = "Stable";

                db.Entry(apm).State = EntityState.Modified;
                db.SaveChanges();

                db.Entry(actifFormula).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(actifFormula);
        }

        [Authorize]
        public ActionResult Delete(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ActifFormula actifFormula = db.ActifFormula.Find(Int32.Parse(id));
            if (actifFormula == null)
            {
                return HttpNotFound();
            }
            return View(actifFormula);
        }

        [Authorize]
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string id)
        {
            ActifFormula actifFormula = db.ActifFormula.Find(Int32.Parse(id));

            generalSettings gs = (generalSettings)Session["SteInformation"];
            int exercice = gs.exercice;
            string matricule = gs.matricule;
            string ownerId = gs.ownerId;

            string code = actifFormula.codeParam;
            ActifParamModel apm = db.ActifModel.Find(ownerId, code, exercice, matricule);
            apm.state = "Stable";

            db.Entry(apm).State = EntityState.Modified;
            db.SaveChanges();

            db.ActifFormula.Remove(actifFormula);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        [Authorize]
        public ActionResult SelectParameters(int id)
        {
            generalSettings gs = (generalSettings)Session["SteInformation"];

            var ParamSetting = from ps in db.ParametersSetting
                               where ps.ownerId.Equals(gs.ownerId) && ps.matricule.Equals(gs.matricule) && ps.exercice == gs.exercice
                               select ps;

            ParametersSetting paramActif = ParamSetting.FirstOrDefault();

            //3==> Specific parameters are chosen
            if (id == 3)
            {
                //Setting the hasParamActif to true
                paramActif.hasParamActif = true;

                db.Entry(paramActif).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");

            }
            //2==> Duplicate an othe company's parameters
            else if (id == 2)
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
                        ParametersSetting ps = (from p in db.ParametersSetting
                                                where p.exercice == company.exercice && p.matricule.Equals(company.matricule) && p.ownerId.Equals(company.ownerId)
                                                select p).FirstOrDefault();
                        if (ps.hasParamActif)
                        {
                            companies.Add(company);
                        }
                    }
                }

                return View("CompaniesList", companies);


            }

            //1==> Default Paramaters are chosen
            else
            {
                //Setting the hasParamActif to true
                paramActif.hasParamActif = true;

                db.Entry(paramActif).State = EntityState.Modified;
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

            ParametersSetting paramActif = ParamSetting.FirstOrDefault();
            paramActif.hasParamActif = true;

            db.Entry(paramActif).State = EntityState.Modified;
            db.SaveChanges();



            var databaseFormulaList = from fl in db.ActifFormula
                                      where fl.exercice == exercice && fl.matricule.Equals(matricule)
                                      select fl;

            foreach (var formula in databaseFormulaList)
            {
                ActifFormula copy = new ActifFormula();
                copy.ownerId = gs.ownerId;
                copy.matricule = gs.matricule;
                copy.exercice = gs.exercice;
                copy.codeParam = formula.codeParam;
                copy.codeDonnee = formula.codeDonnee;
                copy.nomCompte = formula.nomCompte;
                copy.typeFormule = formula.typeFormule;
                copy.genre = formula.genre;
                db.ActifFormula.Add(copy);

            }
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        [Authorize]
        public ActionResult PrintActifNotesAsPdf()
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


            var af = from a in db.ActifModel
                     where a.exercice == exercice && a.matricule.Equals(matricule) && a.ownerId.Equals(ownerId)
                     select a;
            ViewBag.info1 = gs.nomEtPrenomRaisonSociale;
            ViewBag.info2 = gs.adresse;
            ViewBag.info3 = gs.activite;
            ViewBag.info4 = gs.dateDebutExercice;
            ViewBag.info5 = gs.dateClotureExercice;

            var listFormula = from lf in db.ActifFormula
                              where lf.matricule.Equals(matricule)
                              && lf.exercice == exercice
                              && lf.ownerId.Equals(ownerId)
                              select lf;

            ViewBag.listFormula = listFormula;

            var report = new ViewAsPdf("ActifNotesAsPdf", af)
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

            var ActifParam = from ap in db.ActifModel
                             where ap.exercice == exercice && ap.matricule.Equals(matricule) && ap.ownerId.Equals(ownerId)
                             select ap;
            if (!ActifParam.Any())
            {
                ViewBag.Error = "Vous devez d'abord configurer les paramétres actifs! <a href=\"/Parameters/Index\"> Paramétrage Actifs  </a>";
                return View("FileError");
            }
            return View();
        }
    }
}