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
    public class ResultatFiscalParametersController : Controller
    {
        private ExcelProjectContext db = new ExcelProjectContext();

        // GET: ResultatFiscalParameters
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
            ParametersSetting paramRes = ParamSetting.FirstOrDefault();
            if (!paramRes.hasParamRes)
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

                var af = from a in db.ResultatFiscalModel
                         where a.exercice == exercice && a.matricule.Equals(matricule) && a.ownerId.Equals(ownerId)
                         select a;

                var EDRparameters = from a in db.EtatDeResultatModel
                                    where a.exercice == exercice && a.matricule == matricule && a.ownerId.Equals(ownerId)
                                    select a;

                var databaseFormulaList = from fl in db.ResultatFiscalFormula
                                          where fl.exercice == exercice && fl.matricule.Equals(matricule) && fl.ownerId.Equals(ownerId)
                                          select fl;
                //if the parameters of "etat de resultat are not set 
                if (!EDRparameters.Any())
                {
                    ViewBag.Error = "Vous devez d'abord configurer les paramétres Etat de résultat! <a href=\"/EtatDeResultatParameters/Index\"> Paramétrage Etat de résultat  </a>";
                    return View("FileError");
                }
                //there are no Passif parameters for that company so we need to create them
                if (!af.Any())
                {
                    List<ResultatFiscalParamModel> ResultatFiscalParamList = new List<ResultatFiscalParamModel>
            {
                new ResultatFiscalParamModel(){ code="0002" , libelle ="Resultat net comptable apres modifications comptables (avant impôt) " , type="Special",state="Stable",priority=0},
                new ResultatFiscalParamModel(){ code="0003" , libelle ="Charges non déductibles ∑ 4 à 25 " , type="Formula",state="Stable",priority=0},
                new ResultatFiscalParamModel(){ code="0004" , libelle ="Rénumérations de l'exploitant individuel, ou des associés en nom des sociétés de personnes et assimilés " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0005" , libelle ="harges relatifs aux établissements situés à l'étranger SC " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0006" , libelle ="Quote‐part des frais de siège imputable aux établissements situés à l'étranger ( frais du siège x (chiffre d'affaires de l'établissement stable /chiffre d'affaires total))  SC " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0007" , libelle ="Charges relatives aux résidences secondaires, avions et bateaux de plaisance ne faisant pas l'objet de l'exploitation " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0008" , libelle ="Charges relatives aux véhicules de tourismes d'une puissance supérieur à 9 CV ne faisant pas l'objet de l'exploitation" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0009" , libelle ="Cadeaux et frais de réception non déductibles " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0010" , libelle ="Cadeaux et frais de réception excédentaires " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0011" , libelle ="Commissions, courtages, ristournes commerciales ou autres, vacations, honoraires er rémunérations de performance non déclarés dans la déclaration de l'employeur" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0012" , libelle ="Dons et subventions non déductibles" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0013" , libelle ="Dons et subventions excédentaires " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0014" , libelle ="Abandon de créances non déductibles " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0015" , libelle ="Pertes de change non réalisées" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0016" , libelle ="Gains de change non réalisés antérieurement non imposés" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0017" , libelle ="Intérêts servis à l'exploitant ou aux associés des sociétés de personnes et assimilés" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0018" , libelle ="Rémunération excédentaires des titres participatifs et des comptes courants associés" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0019" , libelle ="Charges d'une valeur supérieure ou égale à 5.000 dinars payée en espèces" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0020" , libelle ="Moins values de cession des titres des organismes de placement collectif en valeurs mobilières dans la limite des dividendes réalisés" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0021" , libelle ="Impôts directs supportés aux lieu et place d'autrui PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0022" , libelle ="Taxe de voyages (PP et SP+SC)" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0023" , libelle ="Transactions, amendes, confiscations et pénalités non déductibles     PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0024" , libelle ="Dépenses excédentaires engagées pour la réalisation des  opérations d'essaimage  SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0025" , libelle ="Amortissements non déductibles relatifs aux établissements situés à l'étranger  SC" , type="Formula",state="Stable",priority=0},
                new ResultatFiscalParamModel(){ code="0026" , libelle ="Amortissements non déductibles relatifs aux établissements situés à l'étranger  SC " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0027" , libelle ="Amortissements non déductibles relatifs aux résidences secondaires, avions et bateaux de plaisance ne faisant pas l'objet de l'exploitation    PP et SP+SC " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0028" , libelle ="Amortissements non déductibles relatifs aux véhicules de tourisme d'une puissance fiscale supérieure à 9 CV ne faisant pas l'objet de l'exploitation    PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0029" , libelle ="Amortissements non déductibles relatifs aux terrains et fonds de commerce    PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0030" , libelle ="Amortissements non déductibles relatifs aux actifs d'une valeur supérieure ou égale à 5.000 dinars payée en espèces   PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0031" , libelle ="Partie des amortissements ayant dépassé la limite autorisée par la législation en vigueur  PP et SP+SC " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0032" , libelle ="Partie des amortissements correspondant à une période inférieure à la période autorisée par la législation en vigueur, pour les immobilisations acquises dans le cadre d'un contrat de leasing ou d'ijara   PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0033" , libelle ="Provisions" , type="Formula",state="Stable",priority=0},
                new ResultatFiscalParamModel(){ code="0034" , libelle ="Provisions non déductibles    PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0035" , libelle ="Provisions déductibles pour créances douteuses (autres que celles constituées par les établissements de crédit)    PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0036" , libelle ="Provisions déductibles pour dépréciation des actions cotées en bourse autres que celles constituées par les sociétés d'investissement à capital risque)    PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0037" , libelle ="Provisions déductibles pour dépréciation des stocks destinés à la vente    PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0038" , libelle ="Provisions déductibles pour risque d'exigibilité des engagements techniques (compagnies d'assurance)   SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0039" , libelle ="Produits non comptabilisés ou insuffisamment comptabilisés ∑ 40 à 43" , type="Formula",state="Stable",priority=0},
                new ResultatFiscalParamModel(){ code="0040" , libelle ="Intérêts non décomptés relatifs aux comptes courants associés et aux créances non commerciales    PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0041" , libelle ="Intérêts insuffisamment décomptés relatifs aux comptes courants associés et aux créances non commerciales    PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0042" , libelle ="Plus value de cession des actifs non comptabilisée   PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0043" , libelle ="Plus value de cession des actifs insuffisamment comptabilisée     PP et SP+SC " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0044" , libelle ="Autres réintégrations " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0045" , libelle ="TOTAL DES RÉINTÉGRATIONS" , type="Formula",state="Stable",priority=1},
                new ResultatFiscalParamModel(){ code="0046" , libelle ="Produits réalisés par les établissements situés à l'étranger  SC " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0047" , libelle ="Reprise sur provisions réintégrées au résultat fiscal de l'année de leur constitution    PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0048" , libelle ="Amortissements excédentaires réintégrés aux résultats des années antérieures PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0049" , libelle ="Gains de change réintégrés aux résultats des années antérieures     PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0050" , libelle ="Gains de change non réalisés PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0051" , libelle ="Pertes de change antérieurement constatées    PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0052" , libelle ="50% des salaires servis aux demandeurs d'emploi recrutés pour la première fois  PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0053" , libelle ="Autres déductions PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0054" , libelle ="TOTAL DES DÉDUCTIONS ∑ 0046 à 0053" , type="Formula",state="Stable",priority=1},
                new ResultatFiscalParamModel(){ code="0055" , libelle ="Nature resultat fiscal avant déduction des provisions (Benefice, Perte)  " , type="Special",state="Stable",priority=2},
                new ResultatFiscalParamModel(){ code="0056" , libelle ="Cas bénéficiaire : RÉSULTAT FISCAL 56=sup(55;0) " , type="Special",state="Stable",priority=3},
                new ResultatFiscalParamModel(){ code="0057" , libelle ="Provisions pour créances douteuses    PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0058" , libelle ="Provisions pour dépréciation des stocks destinés à la vente PP et SP+SC " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0059" , libelle ="Prov° pour dépréciation de la valeur des actions cotées en bourse   PP et SP+SC " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0060" , libelle ="Prov° pour risuqes d'exégibilité des engagements techniques(companies d'assurances)   SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0061" , libelle ="RÉSULTAT FISCAL après déduction des provisions" , type="Special",state="Stable",priority=4},
                new ResultatFiscalParamModel(){ code="0062" , libelle ="Déduction de la moins value provenant de la levée de l'option par les salariés de souscription au capital des sociétés ou d'acquisition de leurs actions ou parts sociales dans la limite de 5% du résultat fiscal après déduction des provisions" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0063" , libelle ="Résultat fiscal après déduction des provisions et avant déduction des déficits et des amortissements" , type="Special",state="Stable",priority=5},
                new ResultatFiscalParamModel(){ code="0064" , libelle ="Réintégration des amortissements de l'exercice PP et  SP+SC" , type="Manuel",state="Stable",priority=0},
                new ResultatFiscalParamModel(){ code="0065" , libelle ="Déduction des déficits reportés  PP et SP+SC " , type="Special",state="Stable",priority=6},
                new ResultatFiscalParamModel(){ code="0066" , libelle ="Déduction des amortissements de l'exercice   PP et SP+SC " , type="Special",state="Stable",priority=7},
                new ResultatFiscalParamModel(){ code="0067" , libelle ="Déduction des amortissements diféfrés en périodes défici taires     PP et SP+SC" , type="Special",state="Stable",priority=8},
                new ResultatFiscalParamModel(){ code="0068" , libelle ="RÉSULTAT FISCAL après déduction des provisions, des déficits et amortissements différés " , type="Special",state="Stable",priority=9},
                new ResultatFiscalParamModel(){ code="0069" , libelle ="Dividendes et assimilés distribués par des sociétés établis en Tunisie    SC " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0070" , libelle ="Plus‐value de cession des actions dans le cadre d'une opération d'introduction en bourse PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0071" , libelle ="Plus‐value de cession des actions cotées à la bourse des valeurs mobilières de Tunis  cédées après l'expiration de l'année suivant celle de leur acquisition ou de leur souscription   PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0072" , libelle ="Plus‐value de cession des actions et des parts sociales réalisée par l'intermédiaire des sociétés d'investissement à capital rique (totalement ou dans la limite de 50%)   PP et SP+SC " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0073" , libelle ="Plus‐value de cession des parts des fonds communs de placement à risque (totalement ou dans la limite de 50%)  PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0074" , libelle ="Plus‐value de cession des parts des fonds d'amorçage    PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0075" , libelle ="Plus‐value d'apport dans le cadre d'une opération de fusion ou de scission totale de sociétés, ou d'une opération d'apport des entreprises individuelles dans le capital d'une société (au niveau de la société absorbée, scindée ou de la société ayant fait l'apport SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0076" , libelle ="Plus‐value provenant de l'apport d'actions ou de parts sociales au capital de la société mère ou de la société holding dans le cadre des opérations de restructuration des entreprises ayant pour objet l'introduction de la société mère ou de la société holding à la bourse	PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0077" , libelle ="Plus‐value provenant de la cession totale ou partielle des éléments de l'actif constituant une unité indépendante et autonome suite au départ à la retraite du propriétaire de l'entreprise ou à cause de l'incapacité de poursuivre la gestion de l'entreprise PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0078" , libelle ="Plus‐value provenant de la cession des entreprises en difficultés économiques dans le cadre du règlement judiciare prévu par la loi relative au redressement des entreprises    PP et SP+SC " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0079" , libelle ="intérêts des dépôts et de titres en devises ou en dinars convertibles    PP et SP+SC " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0080" , libelle ="Résultat fiscal avant déduction des bénéfices provenant de l'exploitation  " , type="Special",state="Stable",priority=10},
                new ResultatFiscalParamModel(){ code="0081" , libelle ="Revenus accessoires    PP et SP+SC " , type="Special",state="Stable",priority=0},
                new ResultatFiscalParamModel(){ code="0082" , libelle ="Loyers  PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0083" , libelle ="Revenus de capitaux mobiliers  PP et SP+SC " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0084" , libelle ="Dividendes de source étrangère PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0085" , libelle ="Autres revenus accessoires   PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0086" , libelle ="Gains Exceptionnels    PP et SP+SC" , type="Special",state="Stable",priority=0},
                new ResultatFiscalParamModel(){ code="0087" , libelle ="Plus value de cession des immeubles bâtis et non bâtis et des fonds de commerce    PP et SP+SC " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0088" , libelle ="Gains de change non rattachés à l'activité principale  PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0089" , libelle ="Plus value provenant de la cession des titres    PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0090" , libelle ="Autres gains exceptionnels    PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0091" , libelle ="Total" , type="Special",state="Stable",priority=1},
                new ResultatFiscalParamModel(){ code="0092" , libelle ="Bénéfice servant de base pour la détermination de la quote‐part des bénéfices provenant de l'exploitation déductible" , type="Special",state="Stable",priority=11},
                new ResultatFiscalParamModel(){ code="0093" , libelle ="Au titre de l'exportation    PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0094" , libelle ="Au titre du développement régional  PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0095" , libelle ="Au titre du développement agricole PP et SP+SC " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0096" , libelle ="Autres déductions   PP et SP+SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0097" , libelle ="Total" , type="Special",state="Stable",priority=0},
                new ResultatFiscalParamModel(){ code="0098" , libelle ="Bénéfice fiscal après déduction des bénéfices au titre de l'exploitation" , type="Special",state="Stable",priority=12},
                new ResultatFiscalParamModel(){ code="0099" , libelle ="Déductions des revenus réinvestis PP et SP+SC " , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0100" , libelle ="Réintégration du cinquième de la plus‐value provenant d’opérations de fusion, scission ou d’opérations d’apport (dans la limite de 50%) pour la société absorbante ou la société bénéficiaire de la scission ou de l'apport d'entreprise individuelle. SC" , type="Calculated",state="Stable"},
                new ResultatFiscalParamModel(){ code="0101" , libelle ="Résultat Imposable" , type="Special",state="Stable",priority=13},
                new ResultatFiscalParamModel(){ code="0102" , libelle ="Cas déficitaire RÉSULTAT FISCAL" , type="Special",state="Stable",priority=3},
                new ResultatFiscalParamModel(){ code="0103" , libelle ="Réintégration des amortissements de l'exercice " , type="Manuel",state="Stable",priority=13},
                new ResultatFiscalParamModel(){ code="0104" , libelle ="Déduction des déficits reportés  " , type="Special",state="Stable",priority=13},
                new ResultatFiscalParamModel(){ code="0105" , libelle ="Déduction des amortissements de l'exercice" , type="Special",state="Stable",priority=13},
                new ResultatFiscalParamModel(){ code="0106" , libelle ="Déduction des amortissements différés en périodes déficitaires " , type="Special",state="Stable",priority=13},
                new ResultatFiscalParamModel(){ code="0107" , libelle ="RÉSULTAT FISCAL (déficit reportable) " , type="Formula",state="Stable",priority=14},
                new ResultatFiscalParamModel(){ code="0108" , libelle ="Autre résultat imposable (bénéfice exportation période supérieure à 10 ans)" , type="Formula",state="Stable",priority=0},

            };
                    foreach (var param in ResultatFiscalParamList)
                    {
                        var manager = new UserManager<ApplicationUser>(new UserStore<ApplicationUser>(new ApplicationDbContext()));
                        var currentUser = manager.FindById(User.Identity.GetUserId());
                        param.ownerId = currentUser.Id;
                        param.exercice = exercice;
                        param.matricule = matricule;
                        param.netN = 0;
                        param.netN1 = 0;
                        db.ResultatFiscalModel.Add(param);
                        db.SaveChanges();
                    }
                    return View(ResultatFiscalParamList);

                }
                else return View(af);
            }

        }

        [Authorize]
        public ActionResult Create(string id)
        {
            List<string> definedParam = new List<string>(new string[] { "0003", "0025", "0033", "0039", "0045", "0054", "0068", "0081", "0086", "0091", "0092", "0097", "0101", "0107", "0108", "0002", "0055", "0061", "0064", "0063", "0080", "0098", "0102", "0103", "0104", "0105", "0106", "0037" });
            if (definedParam.Contains(id))
            {
                ViewBag.Error = "Ces paramétres ne peuvent pas être modifier! <a href=\"/ResultatFiscalParameters/Index\"> Paramétres </a>";
                return View("FileError");
            }
            generalSettings gs = (generalSettings)Session["SteInformation"];
            string matricule = gs.matricule;
            int exercice = gs.exercice;
            string ownerId = gs.ownerId;

            var listFormula = from lf in db.ResultatFiscalFormula
                              where lf.codeParam == id
                              && lf.matricule.Equals(matricule)
                              && lf.exercice == exercice
                              && lf.ownerId.Equals(ownerId)
                              select lf;
            ViewBag.listFormula = listFormula;
            ResultatFiscalFormula af = new ResultatFiscalFormula() { codeParam = id };
            return View(af);
        }

        [HttpPost]
        [Authorize]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,codeParam,codeDonnee,nomCompte,typeFormule")] ResultatFiscalFormula RFFormula)
        {
            if (ModelState.IsValid)
            {
                generalSettings gs = (generalSettings)Session["SteInformation"];
                string matricule = gs.matricule;
                int exercice = gs.exercice;
                string ownerId = gs.ownerId;

                string code = RFFormula.codeParam;
                ResultatFiscalParamModel apm = db.ResultatFiscalModel.Find(ownerId, code, exercice, matricule);
                apm.state = "Stable";

                db.Entry(apm).State = EntityState.Modified;
                db.SaveChanges();

                RFFormula.exercice = exercice;
                RFFormula.matricule = matricule;
                RFFormula.ownerId = ownerId;

                db.ResultatFiscalFormula.Add(RFFormula);
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(RFFormula);
        }

        [Authorize]
        public ActionResult Recalculate()
        {
            generalSettings gs = (generalSettings)Session["SteInformation"];
            int exercice = gs.exercice;
            string matricule = gs.matricule;
            string ownerId = gs.ownerId;

            var ResultatFiscalParam = from ap in db.ResultatFiscalModel
                                      where ap.matricule.Equals(matricule) && ap.exercice == exercice && ap.ownerId.Equals(ownerId)
                                      select ap;
            var ResultatFiscalFormula = from af in db.ResultatFiscalFormula
                                        where af.matricule.Equals(matricule) && af.exercice == exercice && af.ownerId.Equals(ownerId)
                                        select af;
            var calculated = from c in ResultatFiscalParam
                             where c.type.Equals("Calculated")
                             select c;
            var rest = from c in ResultatFiscalParam
                       where (!c.type.Equals("Calculated"))
                       orderby c.priority
                       select c;

            //List<Formula> RFFormulaList = new List<Formula>
            //{
            //    new Formula(){ code="0003",type="ResultatFiscal",parameters = new List<string>() {"0004","0005","0006","0007","0008","0009","0010","0011","0012","0012","0014","0015","0016","0017","0018","0019","0020","0021","0022","0023","0024"}},
            //    new Formula(){ code="0025",type="ResultatFiscal",parameters = new List<string>() {"0026","0027","0028","0029","0030","0031","0032"}},
            //    new Formula(){ code="0033",type="ResultatFiscal",parameters = new List<string>() {"0034","0035","0036","0037","0038"}},
            //    new Formula(){ code="0039",type="ResultatFiscal",parameters = new List<string>() {"0040","0041","0042","0043"}},
            //    new Formula(){ code="0045",type="ResultatFiscal",parameters = new List<string>() {"0003","0025","0033","0039","0044"}},
            //    new Formula(){ code="0054",type="ResultatFiscal",parameters = new List<string>() {"0046","0047","0048","0049","0050","0051","0052","0053"}},
            //    new Formula(){ code="0107",type="ResultatFiscal",parameters = new List<string>() {"0102","0103","-0104","-0105","-0106"}},
            //    new Formula(){ code="0108",type="ResultatFiscal",parameters = new List<string>() {"0093"}},
            //};

            var RFFormulaList = from f in db.DefinedFormulas
                                              where f.type.Equals("ResultatFiscal")
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


                    var one = ResultatFiscalFormula.Where(AF => AF.codeParam.Equals(code));
                    var two = one.Where(AF => AF.matricule.Equals(matricule));
                    var three = two.Where(AF => AF.matricule.Equals(ownerId));
                    var specificResultatFormula = two.Where(AF => AF.exercice == exercice);
                    foreach (var formulas in specificResultatFormula)
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
            //Pour les tests sur N
            bool case1 = false;
            bool case2 = false;
            bool case3 = false;
            bool case4 = false;
            bool case5 = false;
            //Pour les tests sur N-1
            bool case1N1 = false;
            bool case2N1 = false;
            bool case3N1 = false;
            bool case4N1 = false;
            bool case5N1 = false;

            foreach (var param in rest.ToList())
            {
               
                if (param.type.Equals("Formula"))
                {
                    param.netN = 0;
                    param.netN1 = 0;

                    var f = from fo in RFFormulaList
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
                        ResultatFiscalParamModel apm = ResultatFiscalParam.Where(AF => AF.code.Equals(cleanCode)).FirstOrDefault();

                        param.netN += apm.netN;
                        param.netN1 += apm.netN1;
                    }

                }
                else if (param.type.Equals("Manuel"))
                {
                    if (param.code.Equals("0103"))
                    {
                        ResultatFiscalParamModel apm = ResultatFiscalParam.Where(AF => AF.code.Equals("0064")).FirstOrDefault();
                        if (case2 || case3 || case4)
                        {
                            param.netN = apm.netN;
                        }
                        if (case2N1 || case3N1 || case4N1)
                        {
                            param.netN1 = apm.netN1;
                        }
                    }
                }
                else if (param.type.Equals("Special"))
                {
                    switch (param.code)
                    {
                        case "0002":
                            {
                                EtatDeResultatParamModel etatDeResultatParamModel = db.EtatDeResultatModel.Find(gs.ownerId, "0076", gs.exercice, gs.matricule);
                                param.netN = etatDeResultatParamModel.netN;
                                param.netN1 = etatDeResultatParamModel.netN1;
                                break;
                            }

                        case "0055":
                            {
                                ResultatFiscalParamModel P0002 = ResultatFiscalParam.Where(AF => AF.code.Equals("0002")).FirstOrDefault();
                                ResultatFiscalParamModel P0045 = ResultatFiscalParam.Where(AF => AF.code.Equals("0045")).FirstOrDefault();
                                ResultatFiscalParamModel P0054 = ResultatFiscalParam.Where(AF => AF.code.Equals("0054")).FirstOrDefault();

                                param.netN = P0002.netN + P0045.netN - P0054.netN;
                                param.netN1 = P0002.netN1 + P0045.netN1 - P0054.netN1;

                                if (param.netN < 0)
                                {
                                    case1 = true;
                                }

                                if (param.netN1 < 0)
                                {
                                    case1N1 = true;
                                }


                                break;
                            }
                        case "0056":
                            {
                                ResultatFiscalParamModel RFP = ResultatFiscalParam.Where(AF => AF.code.Equals("0055")).FirstOrDefault();

                                if (RFP.netN > 0)
                                {
                                    param.netN = RFP.netN;
                                }
                                else
                                {
                                    param.netN = 0;
                                }
                                if (RFP.netN1 > 0)
                                {
                                    param.netN1 = RFP.netN1;
                                }
                                else
                                {
                                    param.netN1 = 0;
                                }
                                break;
                            }
                        case "0061":
                            {
                                ResultatFiscalParamModel P0056 = ResultatFiscalParam.Where(AF => AF.code.Equals("0056")).FirstOrDefault();
                                ResultatFiscalParamModel P0057 = ResultatFiscalParam.Where(AF => AF.code.Equals("0057")).FirstOrDefault();
                                ResultatFiscalParamModel P0058 = ResultatFiscalParam.Where(AF => AF.code.Equals("0058")).FirstOrDefault();
                                ResultatFiscalParamModel P0059 = ResultatFiscalParam.Where(AF => AF.code.Equals("0059")).FirstOrDefault();
                                ResultatFiscalParamModel P0060 = ResultatFiscalParam.Where(AF => AF.code.Equals("0060")).FirstOrDefault();

                                float devidedN = P0056.netN / 2;
                                float devidedN1 = P0056.netN1 / 2;

                                float sumN = P0057.netN + P0058.netN + P0059.netN + P0060.netN;
                                float sumN1 = P0057.netN1 + P0058.netN1 + P0059.netN1 + P0060.netN1;

                                if (sumN > devidedN)
                                {
                                    param.netN = P0056.netN - devidedN;
                                }
                                else
                                {
                                    param.netN = P0056.netN - sumN;
                                }
                                if (sumN1 > devidedN1)
                                {
                                    param.netN1 = P0056.netN1 - devidedN1;
                                }
                                else
                                {
                                    param.netN1 = P0056.netN1 - sumN1;
                                }

                                break;
                            }
                        case "0063":
                            {
                                ResultatFiscalParamModel P0061 = ResultatFiscalParam.Where(AF => AF.code.Equals("0061")).FirstOrDefault();
                                ResultatFiscalParamModel P0062 = ResultatFiscalParam.Where(AF => AF.code.Equals("0062")).FirstOrDefault();

                                if ((P0061.netN - P0062.netN) > 0)
                                {
                                    param.netN = P0061.netN - P0062.netN;
                                }
                                else
                                {
                                    param.netN = 0;
                                }
                                if ((P0061.netN1 - P0062.netN1) > 0)
                                {
                                    param.netN1 = P0061.netN1 - P0062.netN1;
                                }
                                else
                                {
                                    param.netN1 = 0;
                                }


                                break;
                            }
                        case "0065":
                            {
                                ResultatFiscalParamModel P0063 = ResultatFiscalParam.Where(AF => AF.code.Equals("0063")).FirstOrDefault();
                                ResultatFiscalParamModel P0064 = ResultatFiscalParam.Where(AF => AF.code.Equals("0064")).FirstOrDefault();
                                ResultatFiscalParamModel P0065 = ResultatFiscalParam.Where(AF => AF.code.Equals("0065")).FirstOrDefault();

                                if ((P0063.netN + P0064.netN - P0065.netN) < 0)
                                {
                                    case2 = true;
                                }


                                if ((P0063.netN1 + P0064.netN1 - P0065.netN1) < 0)
                                {
                                    case2N1 = true;
                                }

                                break;

                            }


                        case "0066":
                            {
                                ResultatFiscalParamModel P0063 = ResultatFiscalParam.Where(AF => AF.code.Equals("0063")).FirstOrDefault();
                                ResultatFiscalParamModel P0064 = ResultatFiscalParam.Where(AF => AF.code.Equals("0064")).FirstOrDefault();
                                ResultatFiscalParamModel P0065 = ResultatFiscalParam.Where(AF => AF.code.Equals("0065")).FirstOrDefault();
                                ResultatFiscalParamModel P0066 = ResultatFiscalParam.Where(AF => AF.code.Equals("0066")).FirstOrDefault();

                                if ((P0063.netN + P0064.netN - P0065.netN - P0066.netN) < 0)
                                {
                                    case3 = true;
                                }


                                if ((P0063.netN1 + P0064.netN1 - P0065.netN1 - P0066.netN1) < 0)
                                {
                                    case3N1 = true;
                                }

                                break;
                            }
                        case "0067":
                            {
                                ResultatFiscalParamModel P0063 = ResultatFiscalParam.Where(AF => AF.code.Equals("0063")).FirstOrDefault();
                                ResultatFiscalParamModel P0064 = ResultatFiscalParam.Where(AF => AF.code.Equals("0064")).FirstOrDefault();
                                ResultatFiscalParamModel P0065 = ResultatFiscalParam.Where(AF => AF.code.Equals("0065")).FirstOrDefault();
                                ResultatFiscalParamModel P0066 = ResultatFiscalParam.Where(AF => AF.code.Equals("0066")).FirstOrDefault();
                                ResultatFiscalParamModel P0067 = ResultatFiscalParam.Where(AF => AF.code.Equals("0067")).FirstOrDefault();

                                if ((P0063.netN + P0064.netN - P0065.netN - P0066.netN - P0067.netN) < 0)
                                {
                                    case4 = true;
                                }
                                if ((P0063.netN1 + P0064.netN1 - P0065.netN1 - P0066.netN1 - P0067.netN1) < 0)
                                {
                                    case4N1 = true;
                                }

                                break;
                            }
                        case "0068":
                            {
                                ResultatFiscalParamModel P0063 = ResultatFiscalParam.Where(AF => AF.code.Equals("0063")).FirstOrDefault();
                                ResultatFiscalParamModel P0064 = ResultatFiscalParam.Where(AF => AF.code.Equals("0064")).FirstOrDefault();
                                ResultatFiscalParamModel P0065 = ResultatFiscalParam.Where(AF => AF.code.Equals("0065")).FirstOrDefault();
                                ResultatFiscalParamModel P0066 = ResultatFiscalParam.Where(AF => AF.code.Equals("0066")).FirstOrDefault();
                                ResultatFiscalParamModel P0067 = ResultatFiscalParam.Where(AF => AF.code.Equals("0067")).FirstOrDefault();

                                param.netN = P0063.netN + P0064.netN - P0065.netN - P0066.netN - P0067.netN;
                                param.netN1 = P0063.netN1 + P0064.netN1 - P0065.netN1 - P0066.netN1 - P0067.netN1;

                                break;

                            }

                        case "0080":
                            {
                                ResultatFiscalParamModel P0068 = ResultatFiscalParam.Where(AF => AF.code.Equals("0068")).FirstOrDefault();
                                ResultatFiscalParamModel P0069 = ResultatFiscalParam.Where(AF => AF.code.Equals("0069")).FirstOrDefault();
                                ResultatFiscalParamModel P0070 = ResultatFiscalParam.Where(AF => AF.code.Equals("0070")).FirstOrDefault();
                                ResultatFiscalParamModel P0071 = ResultatFiscalParam.Where(AF => AF.code.Equals("0071")).FirstOrDefault();
                                ResultatFiscalParamModel P0072 = ResultatFiscalParam.Where(AF => AF.code.Equals("0072")).FirstOrDefault();
                                ResultatFiscalParamModel P0073 = ResultatFiscalParam.Where(AF => AF.code.Equals("0073")).FirstOrDefault();
                                ResultatFiscalParamModel P0074 = ResultatFiscalParam.Where(AF => AF.code.Equals("0074")).FirstOrDefault();
                                ResultatFiscalParamModel P0075 = ResultatFiscalParam.Where(AF => AF.code.Equals("0075")).FirstOrDefault();
                                ResultatFiscalParamModel P0076 = ResultatFiscalParam.Where(AF => AF.code.Equals("0076")).FirstOrDefault();
                                ResultatFiscalParamModel P0077 = ResultatFiscalParam.Where(AF => AF.code.Equals("0077")).FirstOrDefault();
                                ResultatFiscalParamModel P0078 = ResultatFiscalParam.Where(AF => AF.code.Equals("0078")).FirstOrDefault();
                                ResultatFiscalParamModel P0079 = ResultatFiscalParam.Where(AF => AF.code.Equals("0079")).FirstOrDefault();

                                float sumN = P0069.netN + P0070.netN + P0071.netN + P0072.netN + P0073.netN + P0074.netN + P0075.netN + P0076.netN + P0077.netN + P0078.netN + P0079.netN;
                                float sumN1 = P0069.netN1 + P0070.netN1 + P0071.netN1 + P0072.netN1 + P0073.netN1 + P0074.netN1 + P0075.netN1 + P0076.netN1 + P0077.netN1 + P0078.netN1 + P0079.netN1;

                                if ((P0068.netN - sumN) > 0)
                                {
                                    param.netN = P0068.netN - sumN;
                                }
                                else
                                {
                                    param.netN = 0;
                                }
                                if ((P0068.netN1 - sumN1) > 0)
                                {
                                    param.netN1 = P0068.netN1 - sumN1;
                                }
                                else
                                {
                                    param.netN1 = 0;
                                }
                                break;
                            }
                        case "0081":
                            {

                                ResultatFiscalParamModel P0082 = ResultatFiscalParam.Where(AF => AF.code.Equals("0082")).FirstOrDefault();
                                ResultatFiscalParamModel P0083 = ResultatFiscalParam.Where(AF => AF.code.Equals("0083")).FirstOrDefault();
                                ResultatFiscalParamModel P0084 = ResultatFiscalParam.Where(AF => AF.code.Equals("0084")).FirstOrDefault();
                                ResultatFiscalParamModel P0085 = ResultatFiscalParam.Where(AF => AF.code.Equals("0085")).FirstOrDefault();

                                param.netN = P0082.netN + P0083.netN + P0084.netN + P0085.netN;
                                param.netN1 = P0082.netN1 + P0083.netN1 + P0084.netN1 + P0085.netN1;
                                break;
                            }
                        case "0086":
                            {
                                ResultatFiscalParamModel P0087 = ResultatFiscalParam.Where(AF => AF.code.Equals("0087")).FirstOrDefault();
                                ResultatFiscalParamModel P0088 = ResultatFiscalParam.Where(AF => AF.code.Equals("0088")).FirstOrDefault();
                                ResultatFiscalParamModel P0089 = ResultatFiscalParam.Where(AF => AF.code.Equals("0089")).FirstOrDefault();
                                ResultatFiscalParamModel P0090 = ResultatFiscalParam.Where(AF => AF.code.Equals("0090")).FirstOrDefault();

                                param.netN = P0087.netN + P0088.netN + P0089.netN + P0090.netN;
                                param.netN1 = P0087.netN1 + P0089.netN1 + P0089.netN1 + P0090.netN1;
                                break;
                            }
                        case "0091":
                            {
                                ResultatFiscalParamModel P0081 = ResultatFiscalParam.Where(AF => AF.code.Equals("0081")).FirstOrDefault();
                                ResultatFiscalParamModel P0086 = ResultatFiscalParam.Where(AF => AF.code.Equals("0086")).FirstOrDefault();

                                param.netN = P0081.netN + P0086.netN;
                                param.netN1 = P0081.netN1 + P0086.netN1;
                                break;
                            }
                        case "0092":
                            {
                                ResultatFiscalParamModel P0080 = ResultatFiscalParam.Where(AF => AF.code.Equals("0080")).FirstOrDefault();
                                ResultatFiscalParamModel P0091 = ResultatFiscalParam.Where(AF => AF.code.Equals("0091")).FirstOrDefault();

                                param.netN = P0080.netN - P0091.netN;
                                param.netN1 = P0080.netN1 - P0091.netN1;

                                break;
                            }
                        case "0097":
                            {

                                ResultatFiscalParamModel P0093 = ResultatFiscalParam.Where(AF => AF.code.Equals("0093")).FirstOrDefault();
                                ResultatFiscalParamModel P0094 = ResultatFiscalParam.Where(AF => AF.code.Equals("0094")).FirstOrDefault();
                                ResultatFiscalParamModel P0095 = ResultatFiscalParam.Where(AF => AF.code.Equals("0095")).FirstOrDefault();
                                ResultatFiscalParamModel P0096 = ResultatFiscalParam.Where(AF => AF.code.Equals("0096")).FirstOrDefault();

                                param.netN = P0093.netN + P0094.netN + P0095.netN + P0096.netN;
                                param.netN1 = P0093.netN1 + P0094.netN1 + P0095.netN1 + P0096.netN1;
                                break;
                            }
                        case "0098":
                            {
                                ResultatFiscalParamModel P0092 = ResultatFiscalParam.Where(AF => AF.code.Equals("0092")).FirstOrDefault();
                                ResultatFiscalParamModel P0097 = ResultatFiscalParam.Where(AF => AF.code.Equals("0097")).FirstOrDefault();

                                param.netN = P0092.netN - P0097.netN;
                                param.netN1 = P0092.netN1 - P0097.netN1;

                                if (param.netN < 0)
                                {
                                    case5 = true;
                                }
                                if (param.netN1 < 0)
                                {
                                    case5N1 = true;
                                }

                                break;
                            }
                        case "0101":
                            {
                                ResultatFiscalParamModel P0098 = ResultatFiscalParam.Where(AF => AF.code.Equals("0098")).FirstOrDefault();
                                ResultatFiscalParamModel P0099 = ResultatFiscalParam.Where(AF => AF.code.Equals("0099")).FirstOrDefault();
                                ResultatFiscalParamModel P0100 = ResultatFiscalParam.Where(AF => AF.code.Equals("0100")).FirstOrDefault();

                                param.netN = P0098.netN - P0099.netN + P0100.netN;
                                param.netN1 = P0098.netN1 - P0099.netN1 + P0100.netN1;
                                break;
                            }
                        case "0102":
                            {
                                ResultatFiscalParamModel P0055 = ResultatFiscalParam.Where(AF => AF.code.Equals("0055")).FirstOrDefault();
                                if (P0055.netN < 0)
                                {
                                    param.netN = P0055.netN;
                                }
                                else
                                {
                                    param.netN = 0;
                                }
                                if (P0055.netN1 < 0)
                                {
                                    param.netN1 = P0055.netN1;
                                }
                                else
                                {
                                    param.netN1 = 0;
                                }
                                break;
                            }
                        case "0104":
                            {
                                ResultatFiscalParamModel P0065 = ResultatFiscalParam.Where(AF => AF.code.Equals("0065")).FirstOrDefault();

                                if (case1)
                                {

                                }
                                else if (case2 || case3 || case4)
                                {
                                    param.netN = P0065.netN;

                                }

                                if (case2N1 || case3N1 || case4N1)
                                {
                                    param.netN1 = P0065.netN1;
                                }
                                break;
                            }
                        case "0105":
                            {
                                ResultatFiscalParamModel P0066 = ResultatFiscalParam.Where(AF => AF.code.Equals("0066")).FirstOrDefault();

                                if (case1 || case2)
                                {

                                }
                                else if (case2 || case3 || case4)
                                {
                                    param.netN = P0066.netN;
                                }

                                if (case1N1 || case2N1)
                                {

                                }
                                else if (case2N1 || case3N1 || case4N1)
                                {
                                    param.netN1 = P0066.netN1;
                                }
                                break;
                            }
                        case "0106":
                            {
                                ResultatFiscalParamModel P0067 = ResultatFiscalParam.Where(AF => AF.code.Equals("0067")).FirstOrDefault();
                                if (case1 || case2 || case3)
                                {

                                }
                                else if (case2 || case3 || case4)
                                {
                                    param.netN = P0067.netN;
                                }

                                if (case1N1 || case2N1 || case3N1)
                                {

                                }
                                else if (case2N1 || case3N1 || case4N1)
                                {
                                    param.netN1 = P0067.netN1;
                                }
                                break;
                            }
                        default: { break; }
                    }

                }

                param.ownerId = ownerId;
                param.exercice = exercice;
                param.matricule = matricule;
                db.Entry(param).State = EntityState.Modified;
                db.SaveChanges();
            }


            if (case1)
            {
                for (int i = 56; i <= 101; i++)
                {
                    string deleteCode = i.ToString().PadLeft(4, '0');
                    ResultatFiscalParamModel apm = ResultatFiscalParam.Where(AF => AF.code.Equals(deleteCode)).FirstOrDefault();
                    apm.netN = 0;
                    apm.ownerId = ownerId;
                    apm.exercice = exercice;
                    apm.matricule = matricule;
                    db.Entry(apm).State = EntityState.Modified;
                    db.SaveChanges();

                }
            }
            if (case1N1)
            {
                for (int i=56 ; i<= 101 ; i++)
                {
                    string deleteCode = i.ToString().PadLeft(4, '0');
                    ResultatFiscalParamModel apm = ResultatFiscalParam.Where(AF => AF.code.Equals(deleteCode)).FirstOrDefault();
                    apm.netN1 = 0;
                    apm.ownerId = ownerId;
                    apm.exercice = exercice;
                    apm.matricule = matricule;
                    db.Entry(apm).State = EntityState.Modified;
                    db.SaveChanges();

                }
            }

            if (case2)
            {
                for (int i = 66; i <= 101; i++)
                {
                    string deleteCode = i.ToString().PadLeft(4, '0');
                    ResultatFiscalParamModel apm = ResultatFiscalParam.Where(AF => AF.code.Equals(deleteCode)).FirstOrDefault();
                    apm.netN = 0;
                    apm.ownerId = ownerId;
                    apm.exercice = exercice;
                    apm.matricule = matricule;
                    db.Entry(apm).State = EntityState.Modified;
                    db.SaveChanges();
                }
            }
            if (case2N1)
            {
                for (int i = 66; i <= 101; i++)
                {
                    string deleteCode = i.ToString().PadLeft(4, '0');
                    ResultatFiscalParamModel apm = ResultatFiscalParam.Where(AF => AF.code.Equals(deleteCode)).FirstOrDefault();
                    apm.netN1 = 0;
                    apm.ownerId = ownerId;
                    apm.exercice = exercice;
                    apm.matricule = matricule;
                    db.Entry(apm).State = EntityState.Modified;
                    db.SaveChanges();
                }
            }

            if (case3)
            {
                for (int i = 67; i <= 101; i++)
                {
                    string deleteCode = i.ToString().PadLeft(4, '0');
                    ResultatFiscalParamModel apm = ResultatFiscalParam.Where(AF => AF.code.Equals(deleteCode)).FirstOrDefault();
                    apm.netN = 0;
                    apm.ownerId = ownerId;
                    apm.exercice = exercice;
                    apm.matricule = matricule;
                    db.Entry(apm).State = EntityState.Modified;
                    db.SaveChanges();
                }
            }
            if (case3N1)
            {
                for (int i = 67; i <= 101; i++)
                {
                    string deleteCode = i.ToString().PadLeft(4, '0');
                    ResultatFiscalParamModel apm = ResultatFiscalParam.Where(AF => AF.code.Equals(deleteCode)).FirstOrDefault();
                    apm.netN1 = 0;
                    apm.ownerId = ownerId;
                    apm.exercice = exercice;
                    apm.matricule = matricule;
                    db.Entry(apm).State = EntityState.Modified;
                    db.SaveChanges();
                }
            }

            if (case4)
            {
                for (int i = 68; i <= 101; i++)
                {
                    string deleteCode = i.ToString().PadLeft(4, '0');
                    ResultatFiscalParamModel apm = ResultatFiscalParam.Where(AF => AF.code.Equals(deleteCode)).FirstOrDefault();
                    apm.netN = 0;
                    apm.ownerId = ownerId;
                    apm.exercice = exercice;
                    apm.matricule = matricule;
                    db.Entry(apm).State = EntityState.Modified;
                    db.SaveChanges();
                }
            }

            if (case4N1)
            {
                for (int i = 68; i <= 101; i++)
                {
                    string deleteCode = i.ToString().PadLeft(4, '0');
                    ResultatFiscalParamModel apm = ResultatFiscalParam.Where(AF => AF.code.Equals(deleteCode)).FirstOrDefault();
                    apm.netN1 = 0;
                    apm.ownerId = ownerId;
                    apm.exercice = exercice;
                    apm.matricule = matricule;
                    db.Entry(apm).State = EntityState.Modified;
                    db.SaveChanges();
                }
            }


            if (case5)
            {
                for (int i = 99; i <= 101; i++)
                {
                    string deleteCode = i.ToString().PadLeft(4, '0');
                    ResultatFiscalParamModel apm = ResultatFiscalParam.Where(AF => AF.code.Equals(deleteCode)).FirstOrDefault();
                    apm.netN = 0;
                    apm.ownerId = ownerId;
                    apm.exercice = exercice;
                    apm.matricule = matricule;
                    db.Entry(apm).State = EntityState.Modified;
                    db.SaveChanges();
                }
            }
            if (case5N1)
            {
                for (int i = 99; i <= 101; i++)
                {
                    string deleteCode = i.ToString().PadLeft(4, '0');
                    ResultatFiscalParamModel apm = ResultatFiscalParam.Where(AF => AF.code.Equals(deleteCode)).FirstOrDefault();
                    apm.netN1 = 0;
                    apm.ownerId = ownerId;
                    apm.exercice = exercice;
                    apm.matricule = matricule;
                    db.Entry(apm).State = EntityState.Modified;
                    db.SaveChanges();
                }
            }
            return RedirectToAction("Index");
        }

        [Authorize]
        public ActionResult Show(string id)
        {
            List<string> definedParam = new List<string>(new string[] { "0003", "0025", "0033", "0039", "0045", "0054", "0107", "0108", });
            if (definedParam.Contains(id))
            {
                //    List<Formula> ResultatFiscalFormulaList1 = new List<Formula>
                //{
                //    new Formula(){ code="0003",type="ResultatFiscal",parameters = new List<string>() {"0004","0005","0006","0007","0008","0009","0010","0011","0012","0012","0014","0015","0016","0017","0018","0019","0020","0021","0022","0023","0024"}},
                //    new Formula(){ code="0025",type="ResultatFiscal",parameters = new List<string>() {"0026","0027","0028","0029","0030","0031","0032"}},
                //    new Formula(){ code="0033",type="ResultatFiscal",parameters = new List<string>() {"0034","0035","0036","0037"}},
                //    new Formula(){ code="0039",type="ResultatFiscal",parameters = new List<string>() {"0040","0041","0042","0043"}},
                //    new Formula(){ code="0045",type="ResultatFiscal",parameters = new List<string>() {"0003","0025","0033","0039","0044"}},
                //    new Formula(){ code="0054",type="ResultatFiscal",parameters = new List<string>() {"0046","0047","0048","0049","0050","0051","0052","0053"}},
                //    new Formula(){ code="0107",type="ResultatFiscal",parameters = new List<string>() {"0102","0103","-0104","-0105","-0106"}},
                //    new Formula(){ code="0108",type="ResultatFiscal",parameters = new List<string>() {"0093"}}
                //};

                //Formula specificFormula = ResultatFiscalFormulaList1.Where(AF => AF.code.Equals(id)).FirstOrDefault();

                var ResultatFiscalFormulaList1 = from rf in db.DefinedFormulas
                                                  where rf.type.Equals("ResultatFiscal")
                                                  select rf;

                var f = from fo in ResultatFiscalFormulaList1
                        where fo.code.Equals(id)
                        select fo;

                List<string> parameters = new List<string>();

                foreach (Formula form in f.ToList())
                {
                    parameters.Add(form.parameter);
                }

                List<String> CleanParameters = new List<string>();

                List<string> minus = new List<string>();
                foreach (string code in parameters/*specificFormula.parameters.ToList()*/)
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
                var ResultatFiscalFormulas = from af in db.ResultatFiscalModel
                                             where af.exercice == gs.exercice &&
                                                   af.matricule.Equals(gs.matricule) &&
                                                   af.ownerId.Equals(gs.ownerId) &&
                                                   /*specificFormula.*/CleanParameters.Contains(af.code)
                                             select af;

                ViewBag.minus = minus;
                switch (id)
                {
                    case "0003": { ViewBag.formula = "Somme de F60050004 à F60050024"; break; }
                    case "0025": { ViewBag.formula = "Somme de F60050026 à F60050032"; break; }
                    case "0033": { ViewBag.formula = "Somme de F60050034 à F60050037"; break; }
                    case "0039": { ViewBag.formula = "Somme de F60050040 à F60050043"; break; }
                    case "0045": { ViewBag.formula = "F60051003 + F60051025 +F60051033 + F60051039 + F60051044"; break; }
                    case "0054": { ViewBag.formula = "F60051046 + F60051047 + F60051048 + F60051049 + F60051050 + F60051051 + F60051052 + F60051053"; break; }
                    case "0107": { ViewBag.formula = "F60050102 + F60050103 ‐ F60050104 ‐ F60050105 ‐ F60050106"; break; }
                    case "0108": { ViewBag.formula = "F60050108 = F60050093"; break; }
                    default: { ViewBag.formula = ""; break; }
                }
                return View(ResultatFiscalFormulas);
            }
            else
            {
                //List<Formula> ResultatFiscalFormulaList2 = new List<Formula>
                //{
                //    new Formula() { code = "0002", type = "ResultatFiscal", parameters = new List<string>() { } },
                //    new Formula() { code = "0055", type = "ResultatFiscal", parameters = new List<string>() {"0002","0045","-0054" } },
                //    new Formula() { code = "0056", type = "ResultatFiscal", parameters = new List<string>() {"0055" } },
                //    new Formula() { code = "0061", type = "ResultatFiscal", parameters = new List<string>() {"0056","-0057","-0058","-0059","-0060"} },
                //    new Formula() { code = "0063", type = "ResultatFiscal", parameters = new List<string>() {"0061","-0062"} },
                //    new Formula() { code = "0064", type = "ResultatFiscal", parameters = new List<string>() { } },
                //    new Formula() { code = "0065", type = "ResultatFiscal", parameters = new List<string>() {"0063","0064","-0065" } },
                //    new Formula() { code = "0066", type = "ResultatFiscal", parameters = new List<string>() {"0063","0064","-0065", "-0066" } },
                //    new Formula() { code = "0067", type = "ResultatFiscal", parameters = new List<string>() {"0063","0064","-0065", "-0066", "-0067" } },
                //    new Formula() { code = "0068", type = "ResultatFiscal", parameters = new List<string>() {"0063", "0064", "-0065", "-0066", "-0067" } },
                //    new Formula() { code = "0080", type = "ResultatFiscal", parameters = new List<string>() {"0068","-0069","-0070","-0071","-0072","-0073","-0074","-0075","-0076","-0078","-0079" } },
                //    new Formula() { code = "0081", type = "ResultatFiscal", parameters = new List<string>() {"0082","0083","0084","0085" } },
                //    new Formula() { code = "0086", type = "ResultatFiscal", parameters = new List<string>() {"0087","0088","0089","0090" } },
                //    new Formula() { code = "0091", type = "ResultatFiscal", parameters = new List<string>() {"0081","0086" } },
                //    new Formula() { code = "0092", type = "ResultatFiscal", parameters = new List<string>() {"0080","-0091" } },
                //    new Formula() { code = "0097", type = "ResultatFiscal", parameters = new List<string>() {"0093","0094","0095","0096" } },
                //    new Formula() { code = "0098", type = "ResultatFiscal", parameters = new List<string>() {"0092","-0097" } },
                //    new Formula() { code = "0101", type = "ResultatFiscal", parameters = new List<string>() {"0098","-0099","100" } },
                //    new Formula() { code = "0102", type = "ResultatFiscal", parameters = new List<string>() {"0055" } },
                //    new Formula() { code = "0103", type = "ResultatFiscal", parameters = new List<string>() {"0064" } },
                //    new Formula() { code = "0104", type = "ResultatFiscal", parameters = new List<string>() {"0065"} },
                //    new Formula() { code = "0105", type = "ResultatFiscal", parameters = new List<string>() { "0066"} },
                //    new Formula() { code = "0106", type = "ResultatFiscal", parameters = new List<string>() {"0067" } }
                //};

                //Formula specificFormula = ResultatFiscalFormulaList2.Where(AF => AF.code.Equals(id)).FirstOrDefault();
                var ResultatFiscalFormulaList1 = from rf in db.DefinedFormulas
                                                 where rf.type.Equals("ResultatFiscal")
                                                 select rf;

                var f = from fo in ResultatFiscalFormulaList1
                        where fo.code.Equals(id)
                        select fo;

                List<string> parameters = new List<string>();

                foreach (Formula form in f.ToList())
                {
                    parameters.Add(form.parameter);
                }

                List<String> CleanParameters = new List<string>();

                List<string> minus = new List<string>();
                foreach (string code in parameters/*specificFormula.parameters.ToList()*/)
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
                var ResultatFiscalFormulas = from af in db.ResultatFiscalModel
                                             where af.exercice == gs.exercice &&
                                                   af.matricule.Equals(gs.matricule) &&
                                                   af.ownerId.Equals(gs.ownerId) &&
                                                   /*specificFormula.*/CleanParameters.Contains(af.code)
                                             select af;

                ViewBag.minus = minus;
                switch (id)
                {
                    case "0002": { ViewBag.formula = "Cette valeur est récupérée à partir du paramétre F60030076 de l'Etat de resultat."; break; }
                    case "0038": { ViewBag.formula = "Sup(F60050037,0) "; break; }
                    case "0055": { ViewBag.formula = "Cas 1: si ce paramétre est negatif alors les parametres de F60050056 jusqu'à F60050101 doivent avoir des zéros"; break; }
                    case "0056": { ViewBag.formula = "Sup(F60050055,0) "; break; }
                    case "0061": { ViewBag.formula = "F60050056 - inf(F60050056/2,sum(F60050057 à F60050060)"; break; }
                    case "0063": { ViewBag.formula = "Sup(F60050061- F60050062,0)"; break; }
                    case "0064": { ViewBag.formula = "Ce paramétre est manuel"; break; }
                    case "0065": { ViewBag.formula = "Cas 2: si (F60050063+F60050064‐F60050065) est negatif alors les paramétres de F60050066 à F60050101 doivent avoir des zéros"; break; }
                    case "0066": { ViewBag.formula = "Cas 3: si (F60050063+F60050064‐F60050065‐F60050066) est negatif alors les paramétres de F60050067 à F60050101 doivent avoir des zéros"; break; }
                    case "0067": { ViewBag.formula = "Cas 4: si (F60050063+F60050064‐F60050065‐F60050066‐F60050067) est negatif alors les paramétres de F60050068 à F60050101 doivent avoir des zéros"; break; }
                    case "0068": { ViewBag.formula = "F60050063+ F60050064-F60050065 -F60050066 - F60050067"; break; }
                    case "0080": { ViewBag.formula = "Sup(F60050068- sum(F60050069 à F60050079),0)"; break; }
                    case "0081": { ViewBag.formula = "F60050082 + F60050083 + F60050084 + F60050085"; break; }
                    case "0086": { ViewBag.formula = "F60050087 +F60050088 + F60050089 + F60050090"; break; }
                    case "0091": { ViewBag.formula = "F60050081 + F60050086"; break; }
                    case "0092": { ViewBag.formula = "F60050080 - F60050091"; break; }
                    case "0097": { ViewBag.formula = "F60050093+F60050094+F60050095+F60050096"; break; }
                    case "0098": { ViewBag.formula = "F60050092-F60050097 \n\n Cas 5: si ce paramétre est negatif alors les paramétres de 0099 à 0101 doivent avoir des zéros"; break; }
                    case "0101": { ViewBag.formula = "F60050098-F60050099 +F60050100"; break; }
                    case "0102": { ViewBag.formula = "Inf(F60050055,0)"; break; }
                    case "0103": { ViewBag.formula = "Si cas 1 est valide alors F60050103 est Numérique sinon si cas 2 ou cas 3 ou cas 4 est valide alors F60050103 = F60050064"; break; }
                    case "0104": { ViewBag.formula = "Si cas 1 est valide alors F60050104 est Numérique sinon si cas 2 ou cas 3 ou cas 4 est valide alors F60050104 = F60050065"; break; }
                    case "0105": { ViewBag.formula = "Si cas 1 ou cas 2 est valide alors F60050105 est Numérique sinon si cas 2 ou cas 3 ou cas 4 est valide alors F60050105 = F60050066"; break; }
                    case "0106": { ViewBag.formula = "Si cas 1 ou cas 2 ou cas 3 est valide alors F60050106 est Numérique sinon si cas 2 ou cas 3 ou cas 4 est valide alors F60050106 = F60050067"; break; }
                    default: { ViewBag.formula = ""; break; }
                }
                if (id.Equals("0002"))
                {
                    EtatDeResultatParamModel etatDeResultatParamModel = db.EtatDeResultatModel.Find(gs.ownerId, "0076", gs.exercice, gs.matricule);

                    List<ResultatFiscalParamModel> list = new List<ResultatFiscalParamModel>
                    {
                        new ResultatFiscalParamModel() { code = etatDeResultatParamModel.code,libelle =etatDeResultatParamModel.libelle,netN= etatDeResultatParamModel.netN,netN1=etatDeResultatParamModel.netN1}
                    };
                    var ResultatFiscalFormulas2 = from ff in list
                                                  select ff;
                    return View(ResultatFiscalFormulas2);
                }
                return View(ResultatFiscalFormulas);

            }

        }

        [Authorize]
        public ActionResult EditParam(string ownerId, string code, string exercice, string matricule)
        {
            if ((code == null) || (exercice == null) || (matricule == null))
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ResultatFiscalParamModel RFParamModel = db.ResultatFiscalModel.Find(ownerId, code, Int32.Parse(exercice), matricule);
            if (RFParamModel == null)
            {
                return HttpNotFound();
            }
            return View(RFParamModel);
        }

        [HttpPost]
        [Authorize]
        [ValidateAntiForgeryToken]
        public ActionResult EditParam([Bind(Include = "code,ownerId,libelle,netN,netN1,type,exercice,matricule,state")] ResultatFiscalParamModel param)
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
        public ActionResult EditFormula(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ResultatFiscalFormula RFFormula = db.ResultatFiscalFormula.Find(Int32.Parse(id));
            if (RFFormula == null)
            {
                return HttpNotFound();
            }
            return View(RFFormula);
        }

        [HttpPost]
        [Authorize]
        [ValidateAntiForgeryToken]
        public ActionResult EditFormula([Bind(Include = "")] ResultatFiscalFormula RFFormula)
        {
            if (ModelState.IsValid)
            {
                generalSettings gs = (generalSettings)Session["SteInformation"];
                int exercice = gs.exercice;
                string matricule = gs.matricule;
                string ownerId = gs.ownerId;

                string code = RFFormula.codeParam;
                ResultatFiscalParamModel apm = db.ResultatFiscalModel.Find(ownerId, code, exercice, matricule);
                apm.state = "Stable";

                db.Entry(apm).State = EntityState.Modified;
                db.SaveChanges();

                db.Entry(RFFormula).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(RFFormula);
        }

        [Authorize]
        public ActionResult Delete(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ResultatFiscalFormula RFFormula = db.ResultatFiscalFormula.Find(Int32.Parse(id));
            if (RFFormula == null)
            {
                return HttpNotFound();
            }
            return View(RFFormula);
        }

        [Authorize]
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string id)
        {
            ResultatFiscalFormula RFFormula = db.ResultatFiscalFormula.Find(Int32.Parse(id));

            generalSettings gs = (generalSettings)Session["SteInformation"];
            int exercice = gs.exercice;
            string matricule = gs.matricule;
            string ownerId = gs.ownerId;

            string code = RFFormula.codeParam;
            PassifParamModel apm = db.PassifModel.Find(ownerId, code, exercice, matricule);
            apm.state = "Stable";

            db.Entry(apm).State = EntityState.Modified;
            db.SaveChanges();

            db.ResultatFiscalFormula.Remove(RFFormula);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        [Authorize]
        public ActionResult PrintResultatFiscalAsPdf()
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

            var af = from a in db.ResultatFiscalModel
                     where a.exercice == exercice && a.matricule.Equals(matricule) && a.ownerId.Equals(ownerId)
                     select a;
            ViewBag.info1 = gs.nomEtPrenomRaisonSociale;
            ViewBag.info2 = gs.adresse;
            ViewBag.info3 = gs.activite;
            ViewBag.info4 = gs.dateDebutExercice;
            ViewBag.info5 = gs.dateClotureExercice;
            var report = new ViewAsPdf("ResultatFiscalAsPdf", af)
            {
                PageOrientation = Rotativa.Options.Orientation.Landscape,
                PageSize = Rotativa.Options.Size.A4,
                CustomSwitches = "--footer-center \"  Créer le : " + DateTime.Now.Date.ToString("dd/MM/yyyy") + "  Page: [page]/[toPage]\"" + " --footer-spacing 1 --footer-font-name \"Segoe UI\""
            };
            return report;
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

            var ResultatFiscal = from ap in db.ResultatFiscalModel
                                 where ap.exercice == exercice && ap.matricule.Equals(matricule) && ap.ownerId.Equals(ownerId)
                                 select ap;
            if (!ResultatFiscal.Any())
            {
                ViewBag.Error = "Vous devez d'abord configurer les paramétres resultat fiscal!  <a href=\"/ResultatFiscalParameters/Index\"> Paramétrage Resultat fiscal  </a>";
                return View("FileError");
            }
            return View();
        }

        [Authorize]
        public void PrintResultatFiscalsAsXml()
        {

            generalSettings gs = (generalSettings)Session["SteInformation"];
            var ResultatFiscalParam = from a in db.ResultatFiscalModel
                                      where a.exercice == gs.exercice && a.matricule.Equals(gs.matricule) && a.ownerId.Equals(gs.ownerId)
                                      select a;
            string fileName = "ResultatFiscal-" + gs.matricule + "-" + gs.exercice + ".xml";

            using (MemoryStream stream = new MemoryStream())
            {
                // Create an XML document. Write our specific values into the document.
                XmlTextWriter xmlWriter = new XmlTextWriter(stream, System.Text.Encoding.UTF8);
                // Write the XML document header.
                xmlWriter.WriteStartDocument();
                xmlWriter.WriteRaw("<?xml-stylesheet type=\"text/xsl\"?>");
                xmlWriter.WriteStartElement("lf:F6005");
                xmlWriter.WriteAttributeString("xmlns:lf", "http://www.impots.finances.gov.tn/liasse");
                xmlWriter.WriteAttributeString("xmlns:vc", "http://www.w3.org/2007/XMLSchema-versioning");
                xmlWriter.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
                xmlWriter.WriteAttributeString("xsi:schemaLocation", "http://www.impots.finances.gov.tn/liasse F6005.xsd");
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
                    string resultat;
                    foreach (var param in ResultatFiscalParam)
                    {
                        code = (Int32.Parse(param.code) + add).ToString();
                        if (i == 0)
                        {
                            if (param.code.Equals("0002"))
                            {
                                xmlWriter.WriteRaw("<lf:F60050000 categorie=\"M\" codeformejuridique=\"SC\" />");
                                if (param.netN < 0)
                                    resultat = "p";
                                else
                                    resultat = "B";
                                xmlWriter.WriteRaw($"<lf:F60050001 resultat=\"{resultat}\" F60050002=\"{param.netN * 1000}\"/>");

                            }
                            else if (param.code.Equals("0055"))
                            {
                                if (param.netN < 0)
                                    resultat = "p";
                                else
                                    resultat = "B";
                                xmlWriter.WriteRaw($"<lf:F60050055 resultat=\"{resultat}\" F60050955=\"{param.netN * 1000}\"/>");
                            }
                            else
                            {
                                xmlWriter.WriteElementString($"lf:{param.code}", (param.netN * 1000).ToString());
                            }
                        }
                        else if (i == 1)
                        {
                            if (param.code.Equals("0002"))
                            {
                                xmlWriter.WriteRaw("<lf:F60050000 categorie=\"M\" codeformejuridique=\"SC\" />");

                                if (param.netN1 < 0)
                                    resultat = "p";
                                else
                                    resultat = "B";

                                xmlWriter.WriteRaw($"<lf:F60050001 resultat=\"{resultat}\" F60050002=\"{param.netN1 * 1000}\"/>");

                            }
                            else if (param.code.Equals("0055"))
                            {
                                if (param.netN1 < 0)
                                    resultat = "p";
                                else
                                    resultat = "B";
                                xmlWriter.WriteRaw($"<lf:F60050055 resultat=\"{resultat}\" F60050955=\"{param.netN1 * 1000}\"/>");
                            }
                            else
                            {
                                xmlWriter.WriteElementString($"lf:{code}", (param.netN * 1000).ToString());
                            }
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

            ParametersSetting paramRes = ParamSetting.FirstOrDefault();

            //3==> Specific parameters are chosen
            if (id == 3)
            {
                //Setting the hasParamActif to true
                paramRes.hasParamRes = true;

                db.Entry(paramRes).State = EntityState.Modified;
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
                paramRes.hasParamRes = true;

                db.Entry(paramRes).State = EntityState.Modified;
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

            ParametersSetting paramRes = ParamSetting.FirstOrDefault();
            paramRes.hasParamRes = true;

            db.Entry(paramRes).State = EntityState.Modified;
            db.SaveChanges();



            var databaseFormulaList = from fl in db.ResultatFiscalFormula
                                      where fl.exercice == exercice && fl.matricule.Equals(matricule)
                                      select fl;

            foreach (var formula in databaseFormulaList)
            {
                ResultatFiscalFormula copy = new ResultatFiscalFormula();
                copy.ownerId = gs.ownerId;
                copy.matricule = gs.matricule;
                copy.exercice = gs.exercice;
                copy.codeParam = formula.codeParam;
                copy.codeDonnee = formula.codeDonnee;
                copy.nomCompte = formula.nomCompte;
                copy.typeFormule = formula.typeFormule;
                db.ResultatFiscalFormula.Add(copy);

            }
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        [Authorize]
        public ActionResult PrintResultatFiscalNotesAsPdf()
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

            IEnumerable<ExcelInfo> firstFile = (IEnumerable<ExcelInfo>)Session["firstInputFile"];
            IEnumerable<ExcelInfo> secondFile = (IEnumerable<ExcelInfo>)Session["secondInputFile"];

            ViewBag.firstFile = firstFile;
            ViewBag.secondFile = secondFile;


            var af = from a in db.ResultatFiscalModel
                     where a.exercice == exercice && a.matricule.Equals(matricule) && a.ownerId.Equals(ownerId)
                     select a;
            ViewBag.info1 = gs.nomEtPrenomRaisonSociale;
            ViewBag.info2 = gs.adresse;
            ViewBag.info3 = gs.activite;
            ViewBag.info4 = gs.dateDebutExercice;
            ViewBag.info5 = gs.dateClotureExercice;



            var listFormula = from lf in db.ResultatFiscalFormula
                              where lf.matricule.Equals(matricule)
                              && lf.exercice == exercice
                              && lf.ownerId.Equals(ownerId)
                              select lf;

            ViewBag.listFormula = listFormula;

            var report = new ViewAsPdf("ResultatFiscalNotesAsPdf", af)
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

            var ResultatFiscalParam = from ap in db.ResultatFiscalModel
                                      where ap.exercice == exercice && ap.matricule.Equals(matricule) && ap.ownerId.Equals(ownerId)
                                      select ap;
            if (!ResultatFiscalParam.Any())
            {
                ViewBag.Error = "Vous devez d'abord paramétrer le resultat fiscal! <a href=\"/ResultatFiscalParameters/Index\"> Paramétrage Resultat fiscal  </a>";
                return View("FileError");
            }
            return View();
        }

    }
}
