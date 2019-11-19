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
    public class FluxTresorerieMRParametersController : Controller
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

            ParametersSetting paramFluxMR = ParamSetting.FirstOrDefault();
            if (!paramFluxMR.hasParamFluxMR)
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

                var af = from a in db.FluxTresorerieMRModel
                         where a.exercice == exercice && a.matricule.Equals(matricule) && a.ownerId.Equals(ownerId)
                         select a;

                var databaseFormulaList = from fl in db.FluxTresorerieMRFormula
                                          where fl.exercice == exercice && fl.matricule.Equals(matricule) && fl.ownerId.Equals(ownerId)
                                          select fl;


                //there are no FluxTresorerieMR parameters for that company so we need to create them
                if (!af.Any())
                {
                    List<FluxTresorerieMRParamModel> FluxTresorerieMRParamList = new List<FluxTresorerieMRParamModel>
            {
                new FluxTresorerieMRParamModel(){ code="0001" , libelle ="Flux de trésorerie lies a l'exploitation " , type="Formula",state="Stable",priority=2},
                new FluxTresorerieMRParamModel(){ code="0002" , libelle ="Encaissements reçus des clients " , type="Formula",state="Stable",priority=1},
                new FluxTresorerieMRParamModel(){ code="0003" , libelle ="S.Debiteurs Clts et Rattaches et Regul.bruts en Debut d’exercice " , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0004" , libelle ="S.Crediteurs Clts et Rattaches et Regul.bruts en Debut d’exercice " , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0005" , libelle ="Ventes TTC " , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0006" , libelle ="Ajustements des ventes des exercices anterieurs portes en modifications comptables (compte 128) majores de la TVA " , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0007" , libelle ="Creances clients passees en pertes (comptes 634)" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0008" , libelle ="Gains de change sur creances clients en devises" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0009" , libelle ="Pertes de change sur creances client en devises" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0010" , libelle ="S.Debiteurs Clts et Rattaches et Regul.bruts en Fin d’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0011" , libelle ="S.Crediteurs Clts et Rattaches et Regul.bruts en Fin d’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0012" , libelle ="Sommes versees aux fournisseurs (d'exploitation)" , type="Formula",state="Stable" ,priority=1},
                new FluxTresorerieMRParamModel(){ code="0013" , libelle ="S.Crediteurs Frs Expl. et Rattaches et Regul. en Debut d’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0014" , libelle ="S.Debiteurs Frs Expl. et Rattaches et Regul. en Debut d’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0015" , libelle ="S.C. Etat, RaS  et autres I et T/Ch.Exploitation en Debut d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0016" , libelle ="Achats TTC (des comptes 60, 61, 62 et 63 en partie)" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0017" , libelle ="Ajustements des achats des exercices anterieurs portes en modifications comptables (compte 128) majores de la TVA" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0018" , libelle ="Gains de change sur dettes Frs Expl. en devises" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0019" , libelle ="Pertes de change sur dettes Frs Expl. en devises" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0020" , libelle ="S.Crediteurs Frs Expl. et Rattaches et Regul. en Fin d’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0021" , libelle ="S.Debiteurs Frs Expl. et Rattaches et Regul. en Fin d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0022" , libelle ="S.C. Etat, RaS  et autres I et T/Ch.Exploitation en Fin d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0023" , libelle ="Sommes versees au personnel" , type="Formula",state="Stable",priority=1},
                new FluxTresorerieMRParamModel(){ code="0024" , libelle ="S.Crediteurs PL(Org.Sociaux) et Lies et Regul.en Debut d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0025" , libelle ="S.Debiteurs PL(Org.Sociaux) et Lies et Regul.en Debut d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0026" , libelle ="S.C. Etat, RaS  et autres I et T/Ch.du personnel en Debut d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0027" , libelle ="Charges de personnel de l’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0028" , libelle ="Ajustements des charges de personnel des exercices anterieurs portes en modifications comptables (compte 128)" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0029" , libelle ="S.Crediteurs PL(Org.Sociaux) et Lies et Regul.en Fin d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0030" , libelle ="S.Debiteurs PL(Org.Sociaux) et Lies et Regul.en Fin d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0031" , libelle ="S.C. Etat, RaS  et autres I et T/Ch.du personnel en Fin d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0032" , libelle ="Interets payes" , type="Formula",state="Stable",priority=1},
                new FluxTresorerieMRParamModel(){ code="0033" , libelle ="S.C. Interets dus et Rattaches et Regul.a Payer en Debut d’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0034" , libelle ="S.D. Interets comptes de Regul.d'Avance en Debut d’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0035" , libelle ="S.C. Etat, RaS/Revenus de Capitaux en Debut d’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0036" , libelle ="Charges Financieres de l’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0037" , libelle ="Ajustements des charges d’interet des exercices anterieurs portes en modifications comptables (compte 128)" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0038" , libelle ="Frais d’emission d’emprunt" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0039" , libelle ="Dot.Resorp. Frais d’emission et Primes de Rembours.Empts" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0040" , libelle ="Dot.Prov. Risques et charges financiers  et  Risques de change  et  deprec.immo.financieres  et  deprec.placements et prets courants" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0041" , libelle ="Reprise/Prov. Risques  et charges financiers  et  Risques de change  et  deprec.immo.financieres  et  deprec.placements et prets courants" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0042" , libelle ="S.C. Interets dus et Rattaches et Regul.a Payer en Fin d’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0043" , libelle ="S.D. Interets Comptes de Regul.d'Avance en Fin d’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0044" , libelle ="S.C. Etat, RaS/Revenus de Capitaux en Fin d’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0045" , libelle ="Impots et taxes payes" , type="Formula",state="Stable",priority=1},
                new FluxTresorerieMRParamModel(){ code="0046" , libelle ="S.Crediteurs ( Etat, impots et taxes ) en Debut d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0047" , libelle ="S.Debiteurs ( Etat, impots et taxes ) en Debut d’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0048" , libelle ="S.C.Impot/Resultat differe (+)actif;(-)passif en Debut d’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0049" , libelle ="Impot sur les Resultats (69)" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0050" , libelle ="Impots et Taxes de l’exercice  (66)" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0051" , libelle ="TVA et autres Taxes/B et S Hors exploitation" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0052" , libelle ="Impot/Resultat differe (+)actif;(-)passif constate durant l'exercice sans passer par l’Etat de resultat" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0053" , libelle ="Impots et Taxes portes en Modif.Compt.(cpt128)" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0054" , libelle ="Impot/Resultat a(+)liquider;(-)imputer portes en Modif.Compt." , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0055" , libelle ="S.Crediteurs ( Etat, impot et taxes ) en Fin d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0056" , libelle ="S.Debiteurs ( Etat, impot et taxes ) en Fin d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0057" , libelle ="S.C.Impot/Resultat differe (+)actif;(-)passif en Fin d’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0058" , libelle ="Flux de tresorerie lies aux activites d'investissement" , type="Formula",state="Stable",priority=2},
                new FluxTresorerieMRParamModel(){ code="0059" , libelle ="Decaissements lies aux immo. Corporelles et incorporelles" , type="Formula",state="Stable",priority=1},
                new FluxTresorerieMRParamModel(){ code="0060" , libelle ="S.C.FRS d’Invest. et Rattaches et Regul. en Debut d’exercice " , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0061" , libelle ="S.C. Etat, RaS operee/plus value immobiliere en Fin d’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0062" , libelle ="Valeurs brutes des Invest. d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0063" , libelle ="TVA payee/Investissements de l’exercice " , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0064" , libelle ="S.C. FRS d’Invest. et Rattaches et Regul. en Fin d’exercice " , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0065" , libelle ="S.C. Etat, RaS operee/plus value immobiliere en Fin d’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0066" , libelle ="Encaissements lies aux immo. Corporelles et incorporelles " , type="Formula",state="Stable",priority=1},
                new FluxTresorerieMRParamModel(){ code="0067" , libelle ="S.D. Immo.Corporelles et Incorporelles en Debut d'exercice " , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0068" , libelle ="S.D. Debiteurs  et  autres Creances TTC / cession des Immo.en Debut d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0069" , libelle ="S.C.TVA collectee/cession Investissements en Debut d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0070" , libelle ="S.C.TVA a reverser/cession Investissements en Debut d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0071" , libelle ="S.D.Etat, RaS supportee/plus value immobiliere en Debut d’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0072" , libelle ="TVA a reverser/cession d'Invest. durant l'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0073" , libelle ="Produits nets/Cession des Invest.(-TVA a reverser comprise) durant l'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0074" , libelle ="Charges Nettes/Cession des Invest. durant l'exercice " , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0075" , libelle ="S.D. Immo.Corporelles et Incorporelles en Fin d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0076" , libelle ="S.D.Debiteurs  et  autres Creances TTC / cession des Invest. en Fin d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0077" , libelle ="S.C. TVA collectee/cession investissements en Fin d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0078" , libelle ="S.C. TVA a reverser/cession investissements en Fin d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0079" , libelle ="S.D.Etat, RaS supportee/plus value immobiliere en Fin d’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0080" , libelle ="Decaissements lies aux immobilisations financieres" , type="Formula",state="Stable",priority=1},
                new FluxTresorerieMRParamModel(){ code="0081" , libelle ="Dettes/acquisition d'immo. Financieres en fin d'exercice " , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0082" , libelle ="Valeur brute des titres acquis durant l’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0083" , libelle ="Dettes/acquisition d'immo. Finan. en debut d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0084" , libelle ="Encaissements lies aux immobilisations financieres" , type="Formula",state="Stable",priority=1},
                new FluxTresorerieMRParamModel(){ code="0085" , libelle ="Creances sur cessions d’immo.Fin. en debut d’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0086" , libelle ="Cessions /Immo.Financieres durant l'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0087" , libelle ="Remboursements/Immo.Financieres durant l'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0088" , libelle ="Creances sur cessions d’immo.financ. en fin d’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0089" , libelle ="Flux de tresorerie lies aux activites de financement" , type="Formula",state="Stable",priority=2},
                new FluxTresorerieMRParamModel(){ code="0090" , libelle ="Encaissements suite a l'emission d'actions" , type="Formula",state="Stable",priority=1},
                new FluxTresorerieMRParamModel(){ code="0091" , libelle ="Capital en fin d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0092" , libelle =" Primes liees au capital en fin d'exercice " , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0093" , libelle ="Augmentations du capital durant l'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0094" , libelle ="Conversion de dettes en capital durant l'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0095" , libelle ="Capital en debut d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0096" , libelle ="Primes liees au capital en debut d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0097" , libelle ="Dividendes at autres distributions " , type="Formula",state="Stable" ,priority=1},
                new FluxTresorerieMRParamModel(){ code="0098" , libelle ="Dividendes dus aux actionnaires en debut d’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0099" , libelle ="Dividendes attribues en (N)" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0100" , libelle ="Prelevements sur les reserves en (N)" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0101" , libelle ="Rachat d'actions et autres reductions de capital (non motivees par des pertes ou modif.Comptable) en (N)" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0102" , libelle ="Dividendes dus aux actionnaires en fin d’exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0103" , libelle ="Encaissements/remboursements d'emprunts" , type="Formula",state="Stable",priority=1},
                new FluxTresorerieMRParamModel(){ code="0104" , libelle ="S.C. (Emprunts et dettes assimilees) en fin d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0105" , libelle ="S.C. (Emprunts courants) en fin d'exercice " , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0106" , libelle ="S.C. (Emprunts et dettes assimilees) en debut d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0107" , libelle ="S.C. (Emprunts courants) en debut d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0108" , libelle ="Decaissements/remboursements de prets et des placements" , type="Formula",state="Stable",priority=1},
                new FluxTresorerieMRParamModel(){ code="0109" , libelle ="S.D.(Prets et Creances Fin. courants) en fin d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0110" , libelle ="S.D. (Placements Courants) en fin d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0111" , libelle ="S.D.(Prets et Creances Fin. courants) en debut d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0112" , libelle ="S.D. (Placements Courants) en debut d'exercice" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0113" , libelle ="Incidences des variations des taux de change/les liquidites et equiv." , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0114" , libelle ="Autres Postes des Flux de Tresorerie" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0115" , libelle ="Variation de Tresorerie" , type="Formula",state="Stable",priority=3},
                new FluxTresorerieMRParamModel(){ code="0116" , libelle ="Tresorerie au debut de la periode" , type="Calculated",state="Stable"},
                new FluxTresorerieMRParamModel(){ code="0117" , libelle ="Tresorerie a la cloture de la periode" , type="Calculated",state="Stable"}

            };
                    foreach (var param in FluxTresorerieMRParamList)
                    {
                        var manager = new UserManager<ApplicationUser>(new UserStore<ApplicationUser>(new ApplicationDbContext()));
                        var currentUser = manager.FindById(User.Identity.GetUserId());
                        param.ownerId = currentUser.Id;
                        param.exercice = exercice;
                        param.matricule = matricule;
                        param.netN = 0;
                        param.netN1 = 0;
                        db.FluxTresorerieMRModel.Add(param);
                        db.SaveChanges();
                    }
                    return View(FluxTresorerieMRParamList);

                }
                else return View(af);
            }


        }

        //GET /Parameters/Create/0001
        [Authorize]
        public ActionResult Create(string id)
        {
            List<string> definedParam = new List<string>(new string[] { "0001", "0002", "0012", "0023", "0032", "0045", "0058", "0059", "0066", "0080", "0084", "0089", "0090", "0097", "0103", "0108", "0115" });
            if (definedParam.Contains(id))
            {
                ViewBag.Error = " Ces paramétres ne peuvent pas être modifier! <a href=\"/FluxTresorerieMRParameters/Index\"> Paramétres </a>";
                return View("FileError");
            }
            generalSettings gs = (generalSettings)Session["SteInformation"];
            string matricule = gs.matricule;
            int exercice = gs.exercice;
            string ownerId = gs.ownerId;


            var listFormula = from lf in db.FluxTresorerieMRFormula
                              where lf.codeParam == id
                              && lf.matricule.Equals(matricule)
                              && lf.exercice == exercice
                              && lf.ownerId.Equals(ownerId)
                              select lf;
            ViewBag.listFormula = listFormula;
            FluxTresorerieMRFormula af = new FluxTresorerieMRFormula() { codeParam = id };
            return View(af);
        }

        // POST: FluxTresorerieMRFormulas/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [Authorize]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,codeParam,codeDonnee,nomCompte,typeFormule,RANjournal")] FluxTresorerieMRFormula FluxTresorerieMRFormula)
        {
            if (ModelState.IsValid)
            {
                generalSettings gs = (generalSettings)Session["SteInformation"];
                string matricule = gs.matricule;
                int exercice = gs.exercice;
                string ownerId = gs.ownerId;

                string code = FluxTresorerieMRFormula.codeParam;
                FluxTresorerieMRParamModel apm = db.FluxTresorerieMRModel.Find(ownerId, code, exercice, matricule);
                apm.state = "Stable";

                db.Entry(apm).State = EntityState.Modified;
                db.SaveChanges();

                FluxTresorerieMRFormula.exercice = exercice;
                FluxTresorerieMRFormula.matricule = matricule;
                FluxTresorerieMRFormula.ownerId = ownerId;

                db.FluxTresorerieMRFormula.Add(FluxTresorerieMRFormula);
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(FluxTresorerieMRFormula);
        }

        [Authorize]
        public ActionResult Recalculate()
        {
            generalSettings gs = (generalSettings)Session["SteInformation"];
            int exercice = gs.exercice;
            string matricule = gs.matricule;
            string ownerId = gs.ownerId;

            var FluxTresorerieMRParam = from ap in db.FluxTresorerieMRModel
                                        where ap.matricule.Equals(matricule) && ap.exercice == exercice && ap.ownerId.Equals(ownerId)
                                        select ap;
            var FluxTresorerieMRFormula = from af in db.FluxTresorerieMRFormula
                                          where af.matricule.Equals(matricule) && af.exercice == exercice && af.ownerId.Equals(ownerId)
                                          select af;
            var calculated = from c in FluxTresorerieMRParam
                             where c.type.Equals("Calculated")
                             select c;
            var Paramformula = from c in FluxTresorerieMRParam
                               where c.type.Equals("Formula")
                               orderby c.priority
                               select c;

            //List<Formula> FluxTresorerieMRFormulaList = new List<Formula>
            //{
            //   new Formula(){ code="0001",type="FluxTresorerieMR",parameters = new List<string>() {"0002","-0012","-0023","-0032","-0045"}},
            //    new Formula(){ code="0002",type="FluxTresorerieMR",parameters = new List<string>() {"0003","-0004","0005","0006","-0007","0008","-0009","-0010","0011"}},
            //    new Formula(){ code="0012",type="FluxTresorerieMR",parameters = new List<string>() {"0013","-0014","0015","0016","0017","-0018","0019","-0020","0021","-0022"}},
            //    new Formula(){ code="0023",type="FluxTresorerieMR",parameters = new List<string>() {"0024","-0025","0026","0027","0028","-0029","0030","0031"}},
            //    new Formula(){ code="0032",type="FluxTresorerieMR",parameters = new List<string>() {"0033","-0034","0035","0036","0037","-0038","-0039","-0040","0041","-0042","0043","-0044"}},
            //    new Formula(){ code="0045",type="FluxTresorerieMR",parameters = new List<string>() {"0046","-0047","-0048","0049","0050","0051","-0052","0053","0054","-0055","0056","0057"}},
            //    new Formula(){ code="0058",type="FluxTresorerieMR",parameters = new List<string>() {"0059","0066","-0080","0084"}},
            //    new Formula(){ code="0059",type="FluxTresorerieMR",parameters = new List<string>() {"0060","0061","0062","0063","-0064","-0065"}},
            //    new Formula(){ code="0066",type="FluxTresorerieMR",parameters = new List<string>() {"0067","0068","-0069","-0070","0071","-0072","-0073","0074","-0075","-0076","0077","0078","-0079"}},
            //    new Formula(){ code="0080",type="FluxTresorerieMR",parameters = new List<string>() {"0081","0082","-0083"}},
            //    new Formula(){ code="0084",type="FluxTresorerieMR",parameters = new List<string>() {"0085","0086","0087","-0088"}},
            //    new Formula(){ code="0089",type="FluxTresorerieMR",parameters = new List<string>() {"0090","-0097","0103","-0108","-0045"}},
            //    new Formula(){ code="0090",type="FluxTresorerieMR",parameters = new List<string>() {"0091","0092","-0093","-0094","-0095","-0096"}},
            //    new Formula(){ code="0097",type="FluxTresorerieMR",parameters = new List<string>() {"0098","0099","0100","0101","-0102"}},
            //    new Formula(){ code="0103",type="FluxTresorerieMR",parameters = new List<string>() {"0104","0105","-0106","-0107"}},
            //    new Formula(){ code="0108",type="FluxTresorerieMR",parameters = new List<string>() {"0109","0110","-0111","-0112"}},
            //    new Formula(){ code="0115",type="FluxTresorerieMR",parameters = new List<string>() {"0001","0058","0089","0113","0114"}}
            //};

            var FluxTresorerieMRFormulaList = from f in db.DefinedFormulas
                                              where f.type.Equals("FluxTresorerieMR")
                                              select f;



            IEnumerable<ExcelInfo> firstInputFile = (IEnumerable<ExcelInfo>)Session["firstInputFile"];
            IEnumerable<ExcelInfo> secondInputFile = (IEnumerable<ExcelInfo>)Session["secondInputFile"];

            IEnumerable<ExcelInfo> InputFile = firstInputFile.Concat(secondInputFile);
            IEnumerable<ExcelInfo> specificInputFile;

            foreach (var param in calculated.ToList())
            {
                if (param.state.Equals("Stable"))
                {
                    float netN = 0;
                    float netN1 = 0;
                    float valeur = 0;
                    string code = param.code;


                    var one = FluxTresorerieMRFormula.Where(AF => AF.codeParam.Equals(code));
                    var two = one.Where(AF => AF.matricule.Equals(matricule));
                    var three = two.Where(AF => AF.matricule.Equals(ownerId));
                    var specificFluxTresorerieMRFormula = two.Where(AF => AF.exercice == exercice);
                    foreach (var formulas in specificFluxTresorerieMRFormula)
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
                        if (formulas.typeFormule.Equals("Solde"))
                        {

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
            foreach (var param in Paramformula.ToList())
            {
                param.netN = 0;
                param.netN1 = 0;

                var f = from fo in FluxTresorerieMRFormulaList
                        where fo.code.Equals(param.code)
                        select fo;

                List<string> parameters = new List<string>();

                foreach (Formula form in f.ToList())
                {
                    parameters.Add(form.parameter);
                }
                //Formula formulas = f.FirstOrDefault();
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
                    FluxTresorerieMRParamModel apm = FluxTresorerieMRParam.Where(AF => AF.code.Equals(cleanCode)).FirstOrDefault();
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
            //List<Formula> FluxTresorerieMRFormulaList = new List<Formula>
            //{
            //   new Formula(){ code="0001",type="FluxTresorerieMR",parameters = new List<string>() {"0002","-0012","-0023","-0032","-0045"}},
            //    new Formula(){ code="0002",type="FluxTresorerieMR",parameters = new List<string>() {"0003","-0004","0005","0006","-0007","0008","-0009","-0010","0011"}},
            //    new Formula(){ code="0012",type="FluxTresorerieMR",parameters = new List<string>() {"0013","-0014","0015","0016","0017","-0018","0019","-0020","0021","-0022"}},
            //    new Formula(){ code="0023",type="FluxTresorerieMR",parameters = new List<string>() {"0024","-0025","0026","0027","0028","-0029","0030","0031"}},
            //    new Formula(){ code="0032",type="FluxTresorerieMR",parameters = new List<string>() {"0033","-0034","0035","0036","0037","-0038","-0039","-0040","0041","-0042","0043","-0044"}},
            //    new Formula(){ code="0045",type="FluxTresorerieMR",parameters = new List<string>() {"0046","-0047","-0048","0049","0050","0051","-0052","0053","0054","-0055","0056","0057"}},
            //    new Formula(){ code="0058",type="FluxTresorerieMR",parameters = new List<string>() {"0059","0066","-0080","0084"}},
            //    new Formula(){ code="0059",type="FluxTresorerieMR",parameters = new List<string>() {"0060","0061","0062","0063","-0064","-0065"}},
            //    new Formula(){ code="0066",type="FluxTresorerieMR",parameters = new List<string>() {"0067","0068","-0069","-0070","0071","-0072","-0073","0074","-0075","-0076","0077","0078","-0079"}},
            //    new Formula(){ code="0080",type="FluxTresorerieMR",parameters = new List<string>() {"0081","0082","-0083"}},
            //    new Formula(){ code="0084",type="FluxTresorerieMR",parameters = new List<string>() {"0085","0086","0087","-0088"}},
            //    new Formula(){ code="0089",type="FluxTresorerieMR",parameters = new List<string>() {"0090","-0097","0103","-0108","-0045"}},
            //    new Formula(){ code="0090",type="FluxTresorerieMR",parameters = new List<string>() {"0091","0092","-0093","-0094","-0095","-0096"}},
            //    new Formula(){ code="0097",type="FluxTresorerieMR",parameters = new List<string>() {"0098","0099","0100","0101","-0102"}},
            //    new Formula(){ code="0103",type="FluxTresorerieMR",parameters = new List<string>() {"0104","0105","-0106","-0107"}},
            //    new Formula(){ code="0108",type="FluxTresorerieMR",parameters = new List<string>() {"0109","0110","-0111","-0112"}},
            //    new Formula(){ code="0115",type="FluxTresorerieMR",parameters = new List<string>() {"0001","0058","0089","0113","0114"}}
            //};

            var FluxTresorerieMRFormulaList = from flux in db.DefinedFormulas
                                              where flux.type.Equals("FluxTresorerieMR")
                                              select flux;

            //Formula specificFormula = FluxTresorerieMRFormulaList.Where(AF => AF.code.Equals(id)).FirstOrDefault();
            var f = from fo in FluxTresorerieMRFormulaList
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
            var FluxTresorerieMRFormulas = from af in db.FluxTresorerieMRModel
                                           where af.exercice == gs.exercice &&
                                                 af.matricule.Equals(gs.matricule) &&
                                                 af.ownerId.Equals(gs.ownerId) &&
                                                 /*specificFormula.*/CleanParameters.Contains(af.code)
                                           select af;

            ViewBag.minus = minus;
            return View(FluxTresorerieMRFormulas);
        }

        [Authorize]
        public ActionResult EditParam(string ownerId, string code, string exercice, string matricule)
        {
            if ((code == null) || (exercice == null) || (matricule == null))
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            FluxTresorerieMRParamModel FluxTresorerieMRParamModel = db.FluxTresorerieMRModel.Find(ownerId, code, Int32.Parse(exercice), matricule);
            if (FluxTresorerieMRParamModel == null)
            {
                return HttpNotFound();
            }
            return View(FluxTresorerieMRParamModel);
        }

        [HttpPost]
        [Authorize]
        [ValidateAntiForgeryToken]
        public ActionResult EditParam([Bind(Include = "code,ownerId,libelle,netN,netN1,type,exercice,matricule,ownerId,state")] FluxTresorerieMRParamModel param)
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
        public ActionResult PrintFluxTresorerieMRsAsPdf()
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

            var af = from a in db.FluxTresorerieMRModel
                     where a.exercice == exercice && a.matricule.Equals(matricule) && a.ownerId.Equals(ownerId)
                     select a;
            ViewBag.info1 = gs.nomEtPrenomRaisonSociale;
            ViewBag.info2 = gs.adresse;
            ViewBag.info3 = gs.activite;
            ViewBag.info4 = gs.dateDebutExercice;
            ViewBag.info5 = gs.dateClotureExercice;
            var report = new ViewAsPdf("FluxTresorerieMRsAsPdf", af)
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
            FluxTresorerieMRFormula FluxTresorerieMRFormula = db.FluxTresorerieMRFormula.Find(Int32.Parse(id));
            if (FluxTresorerieMRFormula == null)
            {
                return HttpNotFound();
            }
            return View(FluxTresorerieMRFormula);
        }

        [HttpPost]
        [Authorize]
        [ValidateAntiForgeryToken]
        public ActionResult EditFormula([Bind(Include = "")] FluxTresorerieMRFormula FluxTresorerieMRFormula)
        {
            if (ModelState.IsValid)
            {
                generalSettings gs = (generalSettings)Session["SteInformation"];
                int exercice = gs.exercice;
                string matricule = gs.matricule;
                string ownerId = gs.ownerId;

                string code = FluxTresorerieMRFormula.codeParam;
                FluxTresorerieMRParamModel apm = db.FluxTresorerieMRModel.Find(ownerId, code, exercice, matricule);
                apm.state = "Stable";

                db.Entry(apm).State = EntityState.Modified;
                db.SaveChanges();

                db.Entry(FluxTresorerieMRFormula).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(FluxTresorerieMRFormula);
        }

        [Authorize]
        public ActionResult Delete(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            FluxTresorerieMRFormula FluxTresorerieMRFormula = db.FluxTresorerieMRFormula.Find(Int32.Parse(id));
            if (FluxTresorerieMRFormula == null)
            {
                return HttpNotFound();
            }
            return View(FluxTresorerieMRFormula);
        }

        [Authorize]
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string id)
        {
            FluxTresorerieMRFormula FluxTresorerieMRFormula = db.FluxTresorerieMRFormula.Find(Int32.Parse(id));

            generalSettings gs = (generalSettings)Session["SteInformation"];
            int exercice = gs.exercice;
            string matricule = gs.matricule;
            string ownerId = gs.ownerId;

            string code = FluxTresorerieMRFormula.codeParam;
            FluxTresorerieMRParamModel apm = db.FluxTresorerieMRModel.Find(ownerId, code, exercice, matricule);
            apm.state = "Stable";

            db.Entry(apm).State = EntityState.Modified;
            db.SaveChanges();

            db.FluxTresorerieMRFormula.Remove(FluxTresorerieMRFormula);
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

            var FluxParam = from ap in db.FluxTresorerieMRModel
                            where ap.exercice == exercice && ap.matricule.Equals(matricule) && ap.ownerId.Equals(ownerId)
                            select ap;
            if (!FluxParam.Any())
            {
                ViewBag.Error = " Vous devez d'abord configurer les paramétres Flux de Trésorerie - Modèle de référence !  <a href=\"/FluxTresorerieMRParameters/Index\"> Flux de Trésorerie - Modèle de Référence </a>";
                return View("FileError");
            }
            return View();
        }

        [Authorize]
        public void PrintFluxTresorerieMRsAsXml()
        {

            generalSettings gs = (generalSettings)Session["SteInformation"];
            var FluxTresorerieMRParam = from a in db.FluxTresorerieMRModel
                                        where a.exercice == gs.exercice && a.matricule.Equals(gs.matricule) && a.ownerId.Equals(gs.ownerId)
                                        select a;
            string fileName = "FluxTresorerieMR-" + gs.matricule + "-" + gs.exercice + ".xml";

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
                xmlWriter.WriteAttributeString("xsi:schemaLocation", "http://www.impots.finances.gov.tn/liasse F6004.xsd");
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
                    foreach (var param in FluxTresorerieMRParam)
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

            ParametersSetting paramFluxMR = ParamSetting.FirstOrDefault();

            //3==> Specific parameters are chosen
            if (id == 3)
            {
                //Setting the hasParamActif to true
                paramFluxMR.hasParamFluxMR = true;

                db.Entry(paramFluxMR).State = EntityState.Modified;
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
                paramFluxMR.hasParamFluxMR = true;

                db.Entry(paramFluxMR).State = EntityState.Modified;
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

            ParametersSetting paramFluxMR = ParamSetting.FirstOrDefault();
            paramFluxMR.hasParamFluxMR = true;

            db.Entry(paramFluxMR).State = EntityState.Modified;
            db.SaveChanges();



            var databaseFormulaList = from fl in db.FluxTresorerieMRFormula
                                      where fl.exercice == exercice && fl.matricule.Equals(matricule)
                                      select fl;

            foreach (var formula in databaseFormulaList)
            {
                FluxTresorerieMRFormula copy = new FluxTresorerieMRFormula();
                copy.ownerId = gs.ownerId;
                copy.matricule = gs.matricule;
                copy.exercice = gs.exercice;
                copy.codeParam = formula.codeParam;
                copy.codeDonnee = formula.codeDonnee;
                copy.nomCompte = formula.nomCompte;
                copy.typeFormule = formula.typeFormule;
                db.FluxTresorerieMRFormula.Add(copy);

            }
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        [Authorize]
        public ActionResult PrintFluxTresorerieMRNotesAsPdf()
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


            var af = from a in db.FluxTresorerieMRModel
                     where a.exercice == exercice && a.matricule.Equals(matricule) && a.ownerId.Equals(ownerId)
                     select a;
            ViewBag.info1 = gs.nomEtPrenomRaisonSociale;
            ViewBag.info2 = gs.adresse;
            ViewBag.info3 = gs.activite;
            ViewBag.info4 = gs.dateDebutExercice;
            ViewBag.info5 = gs.dateClotureExercice;

            var listFormula = from lf in db.FluxTresorerieMRFormula
                              where lf.matricule.Equals(matricule)
                              && lf.exercice == exercice
                              && lf.ownerId.Equals(ownerId)
                              select lf;

            ViewBag.listFormula = listFormula;

            var report = new ViewAsPdf("FluxTresorerieMRNotesAsPdf", af)
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

            var ActifParam = from ap in db.FluxTresorerieMRModel
                             where ap.exercice == exercice && ap.matricule.Equals(matricule) && ap.ownerId.Equals(ownerId)
                             select ap;
            if (!ActifParam.Any())
            {
                ViewBag.Error = " Vous devez d'abord configurer les paramétres Flux de Trésorerie - Modèle de référence !  <a href=\"/FluxTresorerieMRParameters/Index\"> Flux de Trésorerie - Modèle de référence </a>";
                return View("FileError");
            }
            return View();
        }
    }
}
