const express = require('express');
const multer = require('multer');
const cors = require('cors');
const xlsx = require('xlsx');

const app = express();
app.use(cors());
app.use(express.static(__dirname));

const upload = multer({ storage: multer.memoryStorage() });

app.post('/api/calculate-profit', upload.fields([
    { name: 'achat', maxCount: 1 },
    { name: 'vente', maxCount: 1 },
    { name: 'cession', maxCount: 1 }
]), (req, res) => {
    try {
        if (!req.files || !req.files['achat'] || !req.files['vente'] || !req.files['cession']) {
            return res.status(400).json({ error: "Veuillez sélectionner les 3 fichiers (Achat, Vente, Cession)." });
        }

        const { year, month, startDate, endDate } = req.body;

        const cleanNumber = (val) => {
            if (val === undefined || val === null || val === '') return 0;
            const parsed = parseFloat(String(val).replace(/\s/g, '').replace(',', '.'));
            return isNaN(parsed) ? 0 : parsed;
        };

        const parseExcelBuffer = (fileBuffer, type) => {
            const workbook = xlsx.read(fileBuffer, { type: 'buffer', cellDates: true, dateNF: 'yyyy-mm-dd hh:mm:ss' });
            const sheetName = workbook.SheetNames[0];
            const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { raw: false });
            
            return data
                .filter(row => row['Statut'] !== 'Annulée' && row['Statut'] !== 'En cours')
                .map(row => {
                    let dateField = type === 'CESSION' ? row["Date/Heure d'initiation de la cession"] : row['Date & Heure Bordereau'];
                    let quantiteField = (type === 'ACHAT' || type === 'VENTE') ? row['Montant en devises'] : row['Montant devises cédées'];
                    let tauxField = type === 'CESSION' ? row['Cours appliqué'] : row['Cours de change accordé'];
                    let cvMadField = row['Contrevaleur en MAD'];
                    let bordereau = type === 'CESSION' ? row['Numéro Bordereau de cession provisoire'] : row['Numéro Bordereau'];

                    const qte = cleanNumber(quantiteField);
                    const taux = cleanNumber(tauxField);
                    const cvMad = cleanNumber(cvMadField) || (qte * taux);
                    const dateObj = new Date(dateField);

                    const yyyy = String(dateObj.getFullYear());
                    const mm = String(dateObj.getMonth() + 1).padStart(2, '0');

                    return {
                        Bordereau: bordereau || '-',
                        Date: dateObj,
                        Year: yyyy,
                        MonthNum: dateObj.getMonth() + 1,
                        YearMonth: `${yyyy}-${mm}`,
                        Type: type,
                        Devise: row['Devise'],
                        Quantite: qte,
                        Taux: taux,
                        ContrevaleurMAD: cvMad
                    };
                })
                .filter(row => row.Quantite > 0 && !isNaN(row.Date));
        };

        const achats = parseExcelBuffer(req.files['achat'][0].buffer, 'ACHAT');
        const ventes = parseExcelBuffer(req.files['vente'][0].buffer, 'VENTE');
        const cessions = parseExcelBuffer(req.files['cession'][0].buffer, 'CESSION');

        const allOperations = [...achats, ...ventes, ...cessions].sort((a, b) => a.Date - b.Date);

        const startFilter = startDate ? new Date(startDate + 'T00:00:00') : null;
        const endFilter = endDate ? new Date(endDate + 'T23:59:59') : null;

        let stockQuantite = {};
        let stockValeurMad = {};
        let totalProfitMad = 0.0;
        let operationsInPeriod = 0;
        let detailsParDevise = {};
        let resumeMensuel = {};
        let journalOperations = [];

        allOperations.forEach(op => {
            const { Bordereau, Devise: dev, Quantite: qte, Taux: taux, ContrevaleurMAD: cvMad, Type: type, Date: dateObj, Year: yr, MonthNum: mNum, YearMonth: ym } = op;

            if (stockQuantite[dev] === undefined) {
                stockQuantite[dev] = 0.0;
                stockValeurMad[dev] = 0.0;
            }
            if (detailsParDevise[dev] === undefined) {
                detailsParDevise[dev] = {
                    stock_initial: 0, cv_stock_initial: 0,
                    achats: 0, cv_achats: 0,
                    ventes: 0, cv_ventes: 0,
                    cessions: 0, cv_cessions: 0,
                    reste: 0, profit: 0
                };
            }
            if (resumeMensuel[ym] === undefined) {
                resumeMensuel[ym] = { year: yr, monthNum: mNum, ops: 0, cv_achats: 0, cv_ventes: 0, cv_cessions: 0, profit: 0 };
            }

            let isBeforePeriod = false;
            let inPeriod = true;

            if (year && year !== 'ALL' && yr !== year) {
                if (yr < year) isBeforePeriod = true;
                inPeriod = false;
            }
            if (month && month !== 'ALL') {
                if (ym < month) isBeforePeriod = true;
                if (ym !== month) inPeriod = false;
            }
            if (startFilter && dateObj < startFilter) {
                isBeforePeriod = true;
                inPeriod = false;
            }
            if (endFilter && dateObj > endFilter) {
                inPeriod = false;
            }

            let opProfit = 0;

            if (type === 'ACHAT') {
                stockQuantite[dev] += qte;
                stockValeurMad[dev] += (qte * taux);

                resumeMensuel[ym].ops += 1;
                resumeMensuel[ym].cv_achats += cvMad;

                if (inPeriod) {
                    operationsInPeriod += 1;
                    detailsParDevise[dev].achats += qte;
                    detailsParDevise[dev].cv_achats += cvMad;
                }
            } 
            else if (type === 'VENTE' || type === 'CESSION') {
                let cumpActuel = stockQuantite[dev] > 0 ? stockValeurMad[dev] / stockQuantite[dev] : taux;
                opProfit = (taux - cumpActuel) * qte;
                
                stockQuantite[dev] -= qte;
                stockValeurMad[dev] -= (qte * cumpActuel);

                resumeMensuel[ym].ops += 1;
                resumeMensuel[ym].profit += opProfit;
                if (type === 'VENTE') resumeMensuel[ym].cv_ventes += cvMad;
                if (type === 'CESSION') resumeMensuel[ym].cv_cessions += cvMad;

                if (inPeriod) {
                    operationsInPeriod += 1;
                    totalProfitMad += opProfit;
                    detailsParDevise[dev].profit += opProfit;

                    if (type === 'VENTE') {
                        detailsParDevise[dev].ventes += qte;
                        detailsParDevise[dev].cv_ventes += cvMad;
                    }
                    if (type === 'CESSION') {
                        detailsParDevise[dev].cessions += qte;
                        detailsParDevise[dev].cv_cessions += cvMad;
                    }
                }
            }

            if (isBeforePeriod) {
                detailsParDevise[dev].stock_initial = stockQuantite[dev];
                detailsParDevise[dev].cv_stock_initial = stockValeurMad[dev];
            }

            if (isBeforePeriod || inPeriod) {
                detailsParDevise[dev].reste = stockQuantite[dev];
            }

            if (inPeriod) {
                journalOperations.push({
                    date: dateObj.toISOString().replace('T', ' ').substring(0, 19),
                    bordereau: Bordereau,
                    type: type,
                    devise: dev,
                    quantite: qte,
                    cours: taux,
                    cv_mad: cvMad,
                    profit: opProfit
                });
            }
        });

        res.json({
            message: "Mise à jour effectuée avec succès !",
            operations_traitees: operationsInPeriod,
            benefice_total_mad: totalProfitMad.toFixed(2),
            par_devise: detailsParDevise,
            par_mois: resumeMensuel,
            journal: journalOperations.slice(-500).reverse() // Les 500 dernières opérations de la période
        });

    } catch (error) {
        console.error(error);
        res.status(500).json({ error: "Erreur lors de l'analyse des fichiers." });
    }
});

// On garde le listen pour vos tests en local (quand on tape 'node server.js')
if (process.env.NODE_ENV !== 'production') {
    const PORT = 3000;
    app.listen(PORT, () => {
        console.log(`✅ Serveur API en ligne sur http://localhost:${PORT}`);
    });
}

// Ligne CRITIQUE pour que Vercel puisse exécuter l'API en ligne
module.exports = app;