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
    public class EtatDeResultatParametersController : Controller
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
                ViewBag.Error = "Vous devez d'abord importer les fichiers  ! <a href=\"/Excel/Input\"> Importer </a>";
                return View("FileError");
            }


            generalSettings gs = (generalSettings)Session["SteInformation"];

            var ParamSetting = from ps in db.ParametersSetting
                               where ps.ownerId.Equals(gs.ownerId) && ps.matricule.Equals(gs.matricule) && ps.exercice == gs.exercice
                               select ps;

            ParametersSetting paramEtatDeRes = ParamSetting.FirstOrDefault();
            if (!paramEtatDeRes.hasParamEtatDeRes)
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

                var af = from a in db.EtatDeResultatModel
                         where a.exercice == exercice && a.matricule == matricule && a.ownerId.Equals(ownerId)
                         select a;

                var databaseFormulaList = from fl in db.EtatDeResultatFormula
                                          where fl.exercice == exercice && fl.matricule.Equals(matricule) && fl.ownerId.Equals(ownerId)
                                          select fl;


                //there are no EtatDeResultat parameters for that company so we need to create them
                if (!af.Any())
                {
                    List<EtatDeResultatParamModel> EtatDeResultatParamList = new List<EtatDeResultatParamModel>
            {
                new EtatDeResultatParamModel(){ code="0001" , libelle ="Produits d'exploitation" , type="Formula",state="Stable",priority=3},
                new EtatDeResultatParamModel(){ code="0002" , libelle ="Revenus" , type="Formula",state="Stable",priority=2},
                new EtatDeResultatParamModel(){ code="0003" , libelle ="Ventes nettes des marchandises" , type="Formula",state="Stable",priority=1},
                new EtatDeResultatParamModel(){ code="0004" , libelle ="Ventes de Marchandises" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0005" , libelle ="Rabais, Remises et Ristournes (3R) accordés/ventes de Marchandises" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0006" , libelle ="Ventes nettes de la production" , type="Formula",state="Stable",priority=1},
                new EtatDeResultatParamModel(){ code="0007" , libelle ="Ventes de Produits Finis" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0008" , libelle ="Ventes de Produits Intermédiaires" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0009" , libelle ="Ventes de Produits Résiduels" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0010" , libelle ="Ventes des Travaux" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0011" , libelle ="Ventes des Études et Prestations de Services" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0012" , libelle ="Produits des Activités Annexes" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0013" , libelle ="Rabais, Remises et Ristournes (3R) accordés/ventes de la Production" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0014" , libelle ="Production immobilisée" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0015" , libelle ="Autres produits d'exploitation" , type="Formula",state="Stable",priority=2},
                new EtatDeResultatParamModel(){ code="0016" , libelle ="Produits divers ordin.(sans gains/cession immo.)" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0017" , libelle ="Subventions d'exploitation et d'équilibre" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0018" , libelle ="Reprises sur amortissements et provisions" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0019" , libelle ="Transferts de charges" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0020" , libelle ="Charges d'exploitation" , type="Formula",state="Stable",priority=3},
                new EtatDeResultatParamModel(){ code="0021" , libelle ="Variation stocks produits finis et encours" , type="Formula",state="Stable",priority=2},
                new EtatDeResultatParamModel(){ code="0022" , libelle ="Variations des en-cours de production biens" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0023" , libelle ="Variation des en-cours de production services" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0024" , libelle ="Variation des stocks de produits" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0025" , libelle ="Achats de marchandises consommées" , type="Formula",state="Stable",priority=2},
                new EtatDeResultatParamModel(){ code="0026" , libelle ="Achats de marchandises" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0027" , libelle ="Rabais, Remises et Ristournes (3R) obtenus sur achats marchandises" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0028" , libelle ="Variation des stocks de marchandises" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0029" , libelle ="Achats d'approvisionnements consommés" , type="Formula",state="Stable",priority=2},
                new EtatDeResultatParamModel(){ code="0030" , libelle ="Achats stockés-Mat.Premières et Fournit. liées" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0031" , libelle ="Achats stockés - Autres approvisionnements" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0032" , libelle ="Rabais, Remises et Ristournes (3R) obtenus/achats Mat.Premières et Fournit. liées" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0033" , libelle ="Rabais, Remises et Ristournes (3R) obtenus/achats autres approvisionnements" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0034" , libelle ="Var.de stocks Mat.Premières et Fournitures" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0035" , libelle ="Var.de stocks des autres approvisionnements" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0036" , libelle ="Charges de personnel" , type="Formula",state="Stable",priority=2},
                new EtatDeResultatParamModel(){ code="0037" , libelle ="Salaires et compléments de salaires" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0038" , libelle ="Appointements et compléments d'appoint." , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0039" , libelle ="Indemnités représentatives de frais" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0040" , libelle ="Commissions au personnel" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0041" , libelle ="Rémun.des administrateurs, gérants et associés" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0042" , libelle ="Ch.connexes sal., appoint., comm. et rémun." , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0043" , libelle ="Charges sociales légales" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0044" , libelle ="Ch.PL/Modif.Compt.à imputer au Réslt de l'exerc.ou Activ.abandonnée" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0045" , libelle ="Autres charges de PL et autres charges sociales" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0046" , libelle ="Dotations aux amortissements et aux provisions" , type="Formula",state="Stable",priority=2},
                new EtatDeResultatParamModel(){ code="0047" , libelle ="Dot.amort. et prov.-Ch.ord.(autres que Fin.)" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0048" , libelle ="Dot. aux résorptions des charges reportées" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0049" , libelle ="Dot. Prov. Risques et Charges d'exploitation" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0050" , libelle ="Dot.Prov.dépréc.immob. Incorp. et Corporelles" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0051" , libelle ="Dot.Prov.dépréc.actifs courants (autres que Val.Mobil.de Placem. et équiv. de liquidités)" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0052" , libelle ="Dot.aux amort. et prov./Modif.Compt. à imputer au Réslt de l'exerc. ou Activ. abandonnée" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0053" , libelle ="Autres charges d'exploitation" , type="Formula",state="Stable",priority=2},
                new EtatDeResultatParamModel(){ code="0054" , libelle ="Achats d’études et prestations services (y compris achat de soustraitance production)" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0055" , libelle ="Achats de matériel, équipements et travaux" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0056" , libelle ="Achats non stockés non rattachés" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0057" , libelle ="Services extérieurs" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0058" , libelle ="Autres services extérieurs" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0059" , libelle ="Charges diverses ordinaires" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0060" , libelle ="Impôts, taxes et versements assimilés" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0061" , libelle ="Resultat d'exploitation" , type="Formula",state="Stable",priority=4},
                new EtatDeResultatParamModel(){ code="0062" , libelle ="Charges financières nettes" , type="Formula",state="Stable",priority=4},
                new EtatDeResultatParamModel(){ code="0063" , libelle ="Charges financières" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0064" , libelle ="Dot.amort. et provisions - charges financières" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0065" , libelle ="Produits des placements" , type="Formula",state="Stable",priority=4},
                new EtatDeResultatParamModel(){ code="0066" , libelle ="Produits financiers" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0067" , libelle ="Reprise/prov.(à inscrire dans les pdts financ.)" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0068" , libelle ="Transferts de charges financières" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0069" , libelle ="Autres gains ordinaires" , type="Formula",state="Stable",priority=4},
                new EtatDeResultatParamModel(){ code="0070" , libelle ="Produits nets sur cessions d'immobilisations" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0071" , libelle ="Autres gains/élém.non récurrents ou except." , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0072" , libelle ="Autres pertes ordinanires" , type="Formula",state="Stable",priority=4},
                new EtatDeResultatParamModel(){ code="0073" , libelle ="Charges Nettes/cession immobilisations" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0074" , libelle ="Autres pertes/élém.non récurrents ou except." , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0075" , libelle ="Réduction de valeur" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0076" , libelle ="Résultat des Activités Ordinaires avant Impôt" , type="Formula",state="Stable",priority=5},
                new EtatDeResultatParamModel(){ code="0077" , libelle ="Impôt sur les bénéfices" , type="Formula",state="Stable",priority=5},
                new EtatDeResultatParamModel(){ code="0078" , libelle ="Impôts/Bénéfices calculés/Résultat/activ./ ord." , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0079" , libelle ="Autres impôts/Bénéfice (régimes particuliers)" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0080" , libelle ="Résultat des Activités Ordinaires après Impôt" , type="Formula",state="Stable",priority=6},
                new EtatDeResultatParamModel(){ code="0081" , libelle ="Elements extraordinanires (Gains/pertes)" , type="Formula",state="Stable",priority=6},
                new EtatDeResultatParamModel(){ code="0082" , libelle ="Gains extraordinaires" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0083" , libelle ="Pertes extraordinaires" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0084" , libelle ="Résultat net de l'exercice" , type="Formula",state="Stable",priority=7},
                new EtatDeResultatParamModel(){ code="0085" , libelle ="Effets des modif. Comptables (net d'impôt)" , type="Formula",state="Stable",priority=0},
                new EtatDeResultatParamModel(){ code="0086" , libelle ="Effet positif/Modif.C.affectant Réslts Reportés" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0087" , libelle ="Effet négatif/Modif.C.affectant Réslts Reportés" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0088" , libelle ="Autres Postes des Comptes de Résultat" , type="Calculated",state="Stable"},
                new EtatDeResultatParamModel(){ code="0089" , libelle ="Resultat apres modifications comptables" , type="Formula",state="Stable",priority=8},

            };

                    foreach (var param in EtatDeResultatParamList)
                    {
                        var manager = new UserManager<ApplicationUser>(new UserStore<ApplicationUser>(new ApplicationDbContext()));
                        var currentUser = manager.FindById(User.Identity.GetUserId());
                        param.ownerId = currentUser.Id;
                        param.exercice = exercice;
                        param.matricule = matricule;
                        param.netN = 0;
                        param.netN1 = 0;
                        db.EtatDeResultatModel.Add(param);
                        db.SaveChanges();
                    }
                    return View(EtatDeResultatParamList);

                }
                else
                    return View(af);
            }

        }

        //GET /Parameters/Create/0001
        [Authorize]
        public ActionResult Create(string id)
        {
            List<string> definedParam = new List<string>(new string[] { "0001", "0002", "0003" });
            if (definedParam.Contains(id))
            {
                ViewBag.Error = "Ces paramétres ne peuvent pas être modifier! <a href=\"/EtatDeResultatParameters/Index\"> Paramétres </a>";
                return View("FileError");
            }
            generalSettings gs = (generalSettings)Session["SteInformation"];
            string matricule = gs.matricule;
            int exercice = gs.exercice;
            string ownerId = gs.ownerId;


            var listFormula = from lf in db.EtatDeResultatFormula
                              where lf.codeParam == id
                              && lf.matricule == matricule
                              && lf.exercice == exercice
                              && lf.ownerId.Equals(ownerId)
                              select lf;
            ViewBag.listFormula = listFormula;
            EtatDeResultatFormula af = new EtatDeResultatFormula() { codeParam = id };
            return View(af);
        }

        // POST: EtatDeResultatFormulas/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [Authorize]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,codeParam,codeDonnee,nomCompte,typeFormule")] EtatDeResultatFormula EtatDeResultatFormula)
        {
            if (ModelState.IsValid)
            {
                generalSettings gs = (generalSettings)Session["SteInformation"];
                string matricule = gs.matricule;
                int exercice = gs.exercice;
                string ownerId = gs.ownerId;

                string code = EtatDeResultatFormula.codeParam;
                EtatDeResultatParamModel apm = db.EtatDeResultatModel.Find(ownerId, code, exercice, matricule);
                apm.state = "Stable";

                db.Entry(apm).State = EntityState.Modified;
                db.SaveChanges();

                EtatDeResultatFormula.exercice = exercice;
                EtatDeResultatFormula.matricule = matricule;
                EtatDeResultatFormula.ownerId = ownerId;

                db.EtatDeResultatFormula.Add(EtatDeResultatFormula);
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(EtatDeResultatFormula);
        }

        [Authorize]
        public ActionResult Recalculate()
        {
            generalSettings gs = (generalSettings)Session["SteInformation"];
            int exercice = gs.exercice;
            string matricule = gs.matricule;
            string ownerId = gs.ownerId;

            var EtatDeResultatParam = from ap in db.EtatDeResultatModel
                                      where ap.matricule == matricule && ap.exercice == exercice && ap.ownerId.Equals(ownerId)
                                      select ap;
            var EtatDeResultatFormula = from af in db.EtatDeResultatFormula
                                        where af.matricule == matricule && af.exercice == exercice && af.ownerId.Equals(ownerId)
                                        select af;
            var calculated = from c in EtatDeResultatParam
                             where c.type.Equals("Calculated")
                             select c;
            var formula = from c in EtatDeResultatParam
                          where c.type.Equals("Formula")
                          orderby c.priority
                          select c;
            //List<Formula> EtatDeResultatFormulaList = new List<Formula>
            //{
            //    new Formula(){ code="0001",type="EtatDeResultat",parameters = new List<string>() {"0002","0014","0015"}},
            //    new Formula(){ code="0002",type="EtatDeResultat",parameters = new List<string>() {"0003","0006"}},
            //    new Formula(){ code="0003",type="EtatDeResultat",parameters = new List<string>() {"0004","-0005"}},
            //    new Formula(){ code="0006",type="EtatDeResultat",parameters = new List<string>() {"0007","0008","0009","0010","0011","0012","-0013"}},
            //    new Formula(){ code="0015",type="EtatDeResultat",parameters = new List<string>() {"0016","0017","0018","0019"}},
            //    new Formula(){ code="0020",type="EtatDeResultat",parameters = new List<string>() {"0021","0025","0029","0036","0046","0053"}},
            //    new Formula(){ code="0021",type="EtatDeResultat",parameters = new List<string>() {"0022","0023","0024"}},
            //    new Formula(){ code="0025",type="EtatDeResultat",parameters = new List<string>() {"0026","-0027","0028"}},
            //    new Formula(){ code="0029",type="EtatDeResultat",parameters = new List<string>() {"0030","0031","-0032","-0033","0034","0035"}},
            //    new Formula(){ code="0036",type="EtatDeResultat",parameters = new List<string>() {"0037","0038","0038","0039","0040","0041","0042","0043","0044","0045"}},
            //    new Formula(){ code="0046",type="EtatDeResultat",parameters = new List<string>() {"0047","0048","0049","0050","0051","0052"}},
            //    new Formula(){ code="0053",type="EtatDeResultat",parameters = new List<string>() {"0054","0055","0056","0057","0058","0059","0060"}},
            //    new Formula(){ code="0061",type="EtatDeResultat",parameters = new List<string>() {"0001","-0020"}},
            //    new Formula(){ code="0062",type="EtatDeResultat",parameters = new List<string>() {"0063","0064"}},
            //    new Formula(){ code="0065",type="EtatDeResultat",parameters = new List<string>() {"0066","0067","0068"}},
            //    new Formula(){ code="0069",type="EtatDeResultat",parameters = new List<string>() {"0070","0071"}},
            //    new Formula(){ code="0072",type="EtatDeResultat",parameters = new List<string>() {"0073","0074","0075"}},
            //    new Formula(){ code="0076",type="EtatDeResultat",parameters = new List<string>() {"0061","0062","0065","0069","-0072"}},
            //    new Formula(){ code="0077",type="EtatDeResultat",parameters = new List<string>() {"0078","0079"}},
            //    new Formula(){ code="0080",type="EtatDeResultat",parameters = new List<string>() {"0076","-0077"}},
            //    new Formula(){ code="0081",type="EtatDeResultat",parameters = new List<string>() {"0082","-0083"}},
            //    new Formula(){ code="0084",type="EtatDeResultat",parameters = new List<string>() {"0080","0081"}},
            //    new Formula(){ code="0085",type="EtatDeResultat",parameters = new List<string>() {"0086","-0087"}},
            //    new Formula(){ code="0089",type="EtatDeResultat",parameters = new List<string>() {"0084","0085","0088"}}
            //};

            var EtatDeResultatFormulaList = from f in db.DefinedFormulas
                                            where f.type.Equals("EtatDeResultat")
                                            select f;


            IEnumerable<ExcelInfo> firstInputFile = (IEnumerable<ExcelInfo>)Session["firstInputFile"];
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

                    var one = EtatDeResultatFormula.Where(AF => AF.codeParam.Equals(code));
                    var two = one.Where(AF => AF.matricule.Equals(matricule));
                    var three = two.Where(AF => AF.matricule.Equals(ownerId));
                    var specificEtatDeResultatFormula = two.Where(AF => AF.exercice == exercice);
                    foreach (var formulas in specificEtatDeResultatFormula)
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

                var f = from fo in EtatDeResultatFormulaList
                        where fo.code.Equals(param.code)
                        select fo;

                List<string> parameters = new List<string>();

                foreach (Formula form in f.ToList())
                {
                    parameters.Add(form.parameter);
                }

                Formula formulas = f.FirstOrDefault();
                foreach (var code in /*formulas.*/parameters)
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

                    EtatDeResultatParamModel apm = EtatDeResultatParam.Where(AF => AF.code.Equals(cleanCode)).FirstOrDefault();

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
            //List<Formula> EtatDeResultatFormulaList = new List<Formula>
            //{
            //    new Formula(){ code="0001",type="EtatDeResultat",parameters = new List<string>() {"0002","0014","0015"}},
            //    new Formula(){ code="0002",type="EtatDeResultat",parameters = new List<string>() {"0003","0006"}},
            //    new Formula(){ code="0003",type="EtatDeResultat",parameters = new List<string>() {"0004","-0005"}},
            //    new Formula(){ code="0006",type="EtatDeResultat",parameters = new List<string>() {"0007","0008","0009","0010","0011","0012","-0013"}},
            //    new Formula(){ code="0015",type="EtatDeResultat",parameters = new List<string>() {"0016","0017","0018","0019"}},
            //    new Formula(){ code="0020",type="EtatDeResultat",parameters = new List<string>() {"0021","0025","0029","0036","0046","0053"}},
            //    new Formula(){ code="0021",type="EtatDeResultat",parameters = new List<string>() {"0022","0023","0024"}},
            //    new Formula(){ code="0025",type="EtatDeResultat",parameters = new List<string>() {"0026","-0027","0028"}},
            //    new Formula(){ code="0029",type="EtatDeResultat",parameters = new List<string>() {"0030","0031","-0032","-0033","0034","0035"}},
            //    new Formula(){ code="0036",type="EtatDeResultat",parameters = new List<string>() {"0037","0038","0038","0039","0040","0041","0042","0043","0044","0045"}},
            //    new Formula(){ code="0046",type="EtatDeResultat",parameters = new List<string>() {"0047","0048","0049","0050","0051","0052"}},
            //    new Formula(){ code="0053",type="EtatDeResultat",parameters = new List<string>() {"0054","0055","0056","0057","0058","0059","0060"}},
            //    new Formula(){ code="0061",type="EtatDeResultat",parameters = new List<string>() {"0001","-0020"}},
            //    new Formula(){ code="0062",type="EtatDeResultat",parameters = new List<string>() {"0063","0064"}},
            //    new Formula(){ code="0065",type="EtatDeResultat",parameters = new List<string>() {"0066","0067","0068"}},
            //    new Formula(){ code="0069",type="EtatDeResultat",parameters = new List<string>() {"0070","0071"}},
            //    new Formula(){ code="0072",type="EtatDeResultat",parameters = new List<string>() {"0073","0074","0075"}},
            //    new Formula(){ code="0076",type="EtatDeResultat",parameters = new List<string>() {"0061","0062","0065","0069","-0072"}},
            //    new Formula(){ code="0077",type="EtatDeResultat",parameters = new List<string>() {"0078","0079"}},
            //    new Formula(){ code="0080",type="EtatDeResultat",parameters = new List<string>() {"0076","-0077"}},
            //    new Formula(){ code="0081",type="EtatDeResultat",parameters = new List<string>() {"0082","-0083"}},
            //    new Formula(){ code="0084",type="EtatDeResultat",parameters = new List<string>() {"0080","0081"}},
            //    new Formula(){ code="0085",type="EtatDeResultat",parameters = new List<string>() {"0086","-0087"}},
            //    new Formula(){ code="0089",type="EtatDeResultat",parameters = new List<string>() {"0084","0085","0088"}}
            //};

            var EtatDeResultatFormulaList = from e in db.DefinedFormulas
                                            where e.type.Equals("EtatDeResultat")
                                            select e;

            //Formula specificFormula = EtatDeResultatFormulaList.Where(AF => AF.code.Equals(id)).FirstOrDefault();
            var f = from fo in EtatDeResultatFormulaList
                    where fo.code.Equals(id)
                    select fo;

            List<string> parameters = new List<string>();

            foreach (Formula form in f.ToList())
            {
                parameters.Add(form.parameter);
            }
            List<String> CleanParameters = new List<string>();

            List<string> minus = new List<string>();
            foreach (string code in/* specificFormula.*/parameters/*.ToList()*/)
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
            var EtatDeResultatFormulas = from af in db.EtatDeResultatModel
                                         where af.exercice == gs.exercice &&
                                               af.matricule.Equals(gs.matricule) &&
                                               af.ownerId.Equals(gs.ownerId) &&
                                               /*specificFormula.*/CleanParameters.Contains(af.code)
                                         select af;

            ViewBag.minus = minus;
            return View(EtatDeResultatFormulas);
        }

        [Authorize]
        public ActionResult EditParam(string ownerId, string code, string exercice, string matricule)
        {
            if ((code == null) || (exercice == null) || (matricule == null))
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            EtatDeResultatParamModel EtatDeResultatParamModel = db.EtatDeResultatModel.Find(ownerId, code, Int32.Parse(exercice), matricule);
            if (EtatDeResultatParamModel == null)
            {
                return HttpNotFound();
            }
            return View(EtatDeResultatParamModel);
        }

        [HttpPost]
        [Authorize]
        [ValidateAntiForgeryToken]
        public ActionResult EditParam([Bind(Include = "code,ownerId,libelle,netN,netN1,type,exercice,matricule,ownerId,state")] EtatDeResultatParamModel param)
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
        public ActionResult PrintEtatDeResultatsAsPdf()
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
            string matricule = gs.matricule;
            int exercice = gs.exercice;
            string ownerId = gs.ownerId;

            var af = from a in db.EtatDeResultatModel
                     where a.exercice == exercice && a.matricule.Equals(matricule) && a.ownerId.Equals(ownerId)
                     select a;
            ViewBag.info1 = gs.nomEtPrenomRaisonSociale;
            ViewBag.info2 = gs.adresse;
            ViewBag.info3 = gs.activite;
            ViewBag.info4 = gs.dateDebutExercice;
            ViewBag.info5 = gs.dateClotureExercice;
            var report = new ViewAsPdf("EtatDeResultatsAsPdf", af)
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
            EtatDeResultatFormula EtatDeResultatFormula = db.EtatDeResultatFormula.Find(Int32.Parse(id));
            if (EtatDeResultatFormula == null)
            {
                return HttpNotFound();
            }
            return View(EtatDeResultatFormula);
        }

        [HttpPost]
        [Authorize]
        [ValidateAntiForgeryToken]
        public ActionResult EditFormula([Bind(Include = "")] EtatDeResultatFormula EtatDeResultatFormula)
        {
            if (ModelState.IsValid)
            {
                generalSettings gs = (generalSettings)Session["SteInformation"];
                int exercice = gs.exercice;
                string matricule = gs.matricule;
                string ownerId = gs.ownerId;

                string code = EtatDeResultatFormula.codeParam;
                EtatDeResultatParamModel apm = db.EtatDeResultatModel.Find(ownerId, code, exercice, matricule);
                apm.state = "Stable";

                db.Entry(apm).State = EntityState.Modified;
                db.SaveChanges();

                db.Entry(EtatDeResultatFormula).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(EtatDeResultatFormula);
        }

        [Authorize]
        public ActionResult Delete(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            EtatDeResultatFormula EtatDeResultatFormula = db.EtatDeResultatFormula.Find(Int32.Parse(id));
            if (EtatDeResultatFormula == null)
            {
                return HttpNotFound();
            }
            return View(EtatDeResultatFormula);
        }

        [Authorize]
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string id)
        {
            EtatDeResultatFormula EtatDeResultatFormula = db.EtatDeResultatFormula.Find(Int32.Parse(id));

            generalSettings gs = (generalSettings)Session["SteInformation"];
            int exercice = gs.exercice;
            string matricule = gs.matricule;
            string ownerId = gs.ownerId;

            string code = EtatDeResultatFormula.codeParam;
            EtatDeResultatParamModel apm = db.EtatDeResultatModel.Find(ownerId, code, exercice, matricule);
            apm.state = "Stable";

            db.Entry(apm).State = EntityState.Modified;
            db.SaveChanges();

            db.EtatDeResultatFormula.Remove(EtatDeResultatFormula);
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
                ViewBag.Error = "Vous devez d'abord importer les fichiers  !  <a href=\"/Excel/Input\"> Importer </a>";
                return View("FileError");
            }

            generalSettings gs = (generalSettings)Session["SteInformation"];
            int exercice = gs.exercice;
            string matricule = gs.matricule;
            string ownerId = gs.ownerId;

            var EtatParam = from ap in db.EtatDeResultatModel
                            where ap.exercice == exercice && ap.matricule.Equals(matricule) && ap.ownerId.Equals(ownerId)
                            select ap;
            if (!EtatParam.Any())
            {
                ViewBag.Error = "Vous devez d'abord configurer les paramétres Etat de résultat!  <a href=\"/EtatDeResultatParameters/Index\"> Paramétrage Etat de résultat  </a>";
                return View("FileError");
            }
            return View();
        }

        [Authorize]
        public void PrintEtatDeResultatsAsXml()
        {

            generalSettings gs = (generalSettings)Session["SteInformation"];
            var EtatDeResultatParam = from a in db.EtatDeResultatModel
                                      where a.exercice == gs.exercice && a.matricule.Equals(gs.matricule) && a.ownerId.Equals(gs.ownerId)
                                      select a;
            string fileName = "EtatDeResultat-" + gs.matricule + "-" + gs.exercice + ".xml";

            using (MemoryStream stream = new MemoryStream())
            {
                // Create an XML document. Write our specific values into the document.
                XmlTextWriter xmlWriter = new XmlTextWriter(stream, System.Text.Encoding.UTF8);
                // Write the XML document header.
                xmlWriter.WriteStartDocument();
                xmlWriter.WriteRaw("<?xml-stylesheet type=\"text/xsl\" href=\"F6003.xsl\"?>");
                xmlWriter.WriteStartElement("lf:F6003");
                xmlWriter.WriteAttributeString("xmlns:lf", "http://www.impots.finances.gov.tn/liasse");
                xmlWriter.WriteAttributeString("xmlns:vc", "http://www.w3.org/2007/XMLSchema-versioning");
                xmlWriter.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
                xmlWriter.WriteAttributeString("xsi:schemaLocation", "http://www.impots.finances.gov.tn/liasse F6003.xsd");
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
                    int add = i * 89;
                    string code;
                    foreach (var param in EtatDeResultatParam)
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

            ParametersSetting paramEtatDeRes = ParamSetting.FirstOrDefault();

            //3==> Specific parameters are chosen
            if (id == 3)
            {
                //Setting the hasParamActif to true
                paramEtatDeRes.hasParamEtatDeRes = true;

                db.Entry(paramEtatDeRes).State = EntityState.Modified;
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
                paramEtatDeRes.hasParamEtatDeRes = true;

                db.Entry(paramEtatDeRes).State = EntityState.Modified;
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

            ParametersSetting paramEtatDeRes = ParamSetting.FirstOrDefault();
            paramEtatDeRes.hasParamEtatDeRes = true;

            db.Entry(paramEtatDeRes).State = EntityState.Modified;
            db.SaveChanges();



            var databaseFormulaList = from fl in db.EtatDeResultatFormula
                                      where fl.exercice == exercice && fl.matricule.Equals(matricule)
                                      select fl;

            foreach (var formula in databaseFormulaList)
            {
                EtatDeResultatFormula copy = new EtatDeResultatFormula();
                copy.ownerId = gs.ownerId;
                copy.matricule = gs.matricule;
                copy.exercice = gs.exercice;
                copy.codeParam = formula.codeParam;
                copy.codeDonnee = formula.codeDonnee;
                copy.nomCompte = formula.nomCompte;
                copy.typeFormule = formula.typeFormule;
                db.EtatDeResultatFormula.Add(copy);

            }
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        [Authorize]
        public ActionResult PrintEtatDeResultatNotesAsPdf()
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


            var af = from a in db.EtatDeResultatModel
                     where a.exercice == exercice && a.matricule.Equals(matricule) && a.ownerId.Equals(ownerId)
                     select a;
            ViewBag.info1 = gs.nomEtPrenomRaisonSociale;
            ViewBag.info2 = gs.adresse;
            ViewBag.info3 = gs.activite;
            ViewBag.info4 = gs.dateDebutExercice;
            ViewBag.info5 = gs.dateClotureExercice;

            var listFormula = from lf in db.EtatDeResultatFormula
                              where lf.matricule.Equals(matricule)
                              && lf.exercice == exercice
                              && lf.ownerId.Equals(ownerId)
                              select lf;

            ViewBag.listFormula = listFormula;

            var report = new ViewAsPdf("EtatDeResultatNotesAsPdf", af)
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

            var ActifParam = from ap in db.EtatDeResultatModel
                             where ap.exercice == exercice && ap.matricule.Equals(matricule) && ap.ownerId.Equals(ownerId)
                             select ap;
            if (!ActifParam.Any())
            {
                ViewBag.Error = "Vous devez d'abord configurer les paramétres Etat de résultat! <a href=\"/EtatDeResultatParameters/Index\"> Paramétrage Etat de résultat  </a>";
                return View("FileError");
            }
            return View();
        }
    }
}
