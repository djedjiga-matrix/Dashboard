const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const cors = require('cors');
const fs = require('fs');

const app = express();
const port = 5000;

app.use(cors());
app.use(express.json());

const upload = multer({ dest: 'uploads/' });

// Fonction pour convertir le format "HH:MM:SS" en décimal
const convertTimeToDecimal = (timeString) => {
    if (!timeString) return 0;
    
    // Si c'est déjà un nombre
    if (typeof timeString === 'number') return timeString;
    
    if (typeof timeString === 'string') {
        const parts = timeString.split(':');
        if (parts.length >= 2) {
            const hours = parseInt(parts[0], 10);
            const minutes = parseInt(parts[1], 10);
            return Number(`${hours}.${minutes.toString().padStart(2, '0')}`);
        }
    }
    return 0;
};

// 1. Updated normalization function
const normalizeQualification = (qualif) => {
    if (!qualif) return '';
    qualif = qualif.toLowerCase().trim();
    
    const mapping = {
        'pa en ligne': 'pa en ligne',
        'pa': 'pa',
        'don en ligne': 'don en ligne',
        'don avec montant': 'don avec montant',
        'don montant': 'don avec montant',
        'indecis don': 'indecis Don',
        'indécis don': 'indecis Don',
        'refus argumente': 'refus argumente',
        'refus argumenté': 'refus argumente'
    };
    
    // Log unknown qualifications for debugging
    if (!mapping[qualif]) {
        console.log(`Qualification non mappée: ${qualif}`);
    }
    
    return mapping[qualif] || qualif;
};

// 2. Update processData function
const processData = (importData, extractWorkbook) => {
    let agentStats = {};
    
    // Debug: Log unique qualifications
    const qualifications = new Set();
    importData.forEach(row => {
        if (row.contact_qualif1) {
            qualifications.add(row.contact_qualif1);
        }
    });
    console.log('Qualifications uniques:', Array.from(qualifications));
    
    // 1. Traitement des données d'import
    importData.forEach(row => {
        const Agents = row.Agents;
        const qualif = normalizeQualification(row.contact_qualif1);
        
        if (!agentStats[agent]) {
            agentStats[agent] = {
                agent,
                "Total général": 0,
                "don avec montant": 0,
                "don en ligne": 0,
                "Total_Cu+": 0,
                "Tx Accord_don": "0%",
                "pa": 0,
                "Pa en ligne": 0,
                "Tx Accord_Pal": "0%",
                "indecis Don": 0,
                "refus argumente": 0,
                "Cu's/h": 0,
                "Durée production": 0,
                "Durée présence": 0,
                "Pause Brief": 0,
                "Pauses non productives": 0,
                "Nbr/J Travailler": 0
            };
        }

        if (qualif) {
            // Incrémenter le compteur spécifique
            agentStats[agent][qualif.toLowerCase()] = (agentStats[agent][qualif.toLowerCase()] || 0) + 1;
            
            // Incrémenter le total général pour chaque qualification pertinente
            if (["don avec montant", "refus argumente", "pa", "pa en ligne", "indecis don", "don en ligne"].includes(qualif.toLowerCase())) {
                agentStats[agent]["Total général"]++;
            }
        }
    });

    // 2. Traitement des données de la feuille Resume
    const resumeSheet = extractWorkbook.Sheets['Resume'];
    if (resumeSheet) {
        const resumeData = xlsx.utils.sheet_to_json(resumeSheet);
        resumeData.forEach(row => {
            const agentName = row.Agents; // Le nom de la colonne est "Agents" dans le fichier
            if (agentStats[agentName]) {
                agentStats[agentName]["Durée production"] = convertTimeToDecimal(row["Durée production"]);
                agentStats[agentName]["Durée présence"] = convertTimeToDecimal(row["Durée présence"]);
            }
        });
    }

    // 3. Comptage des jours travaillés et pauses
    const dailySheets = extractWorkbook.SheetNames.filter(name => 
        name !== 'Resume' && name !== 'Worksheet 1' && /^\d{4}-\d{2}-\d{2}$/.test(name)
    );

    dailySheets.forEach(sheetName => {
        const daySheet = extractWorkbook.Sheets[sheetName];
        const dayData = xlsx.utils.sheet_to_json(daySheet);
        
        dayData.forEach(row => {
            const agentName = row.Agents; // Le nom de la colonne est "Agents" dans le fichier
            if (agentStats[agentName]) {
                agentStats[agentName]["Nbr/J Travailler"]++;
                agentStats[agentName]["Pause Brief"] += convertTimeToDecimal(row["Pause : Pause Brief"]);
                agentStats[agentName]["Pauses non productives"] += convertTimeToDecimal(row["Pauses non productives"]);
            }
        });
    });

    // 4. Calculs finaux
    for (let agent in agentStats) {
        const stats = agentStats[agent];
        
        // Calcul du Total_Cu+
        stats["Total_Cu+"] = stats["don avec montant"] + stats["don en ligne"];
        
        // Calcul des taux
        if (stats["Total général"] > 0) {
            // Tx Accord_don
            const txAccordDon = (stats["Total_Cu+"] / stats["Total général"]) * 100;
            stats["Tx Accord_don"] = txAccordDon.toFixed(2) + "%";
            
            // Tx Accord_Pal
            const txAccordPal = (stats["pa en ligne"] / stats["Total général"]) * 100;
            stats["Tx Accord_Pal"] = txAccordPal.toFixed(2) + "%";
        }
        
        // Calcul du Cu's/h
        if (stats["Durée production"] > 0) {
            stats["Cu's/h"] = (stats["Total général"] / stats["Durée production"]).toFixed(2);
        }
    }

    return Object.values(agentStats);
};

// 3. Update upload endpoint with better validation
app.post('/upload', upload.fields([{ name: 'import' }, { name: 'extract' }]), (req, res) => {
    if (!req.files || !req.files['import'] || !req.files['extract']) {
        return res.status(400).json({ error: 'Fichiers manquants' });
    }

    try {
        const importFile = req.files['import'][0].path;
        const extractFile = req.files['extract'][0].path;

        // Lecture des fichiers avec les options appropriées
        const importWorkbook = xlsx.readFile(importFile, {
            cellDates: true,
            cellNF: true,
            cellText: true
        });
        const extractWorkbook = xlsx.readFile(extractFile, {
            cellDates: true,
            cellNF: true,
            cellText: true
        });

        const importSheet = importWorkbook.Sheets[importWorkbook.SheetNames[0]];
        const importData = xlsx.utils.sheet_to_json(importSheet);

        const processedData = processData(importData, extractWorkbook);

        // Nettoyage des fichiers temporaires
        fs.unlinkSync(importFile);
        fs.unlinkSync(extractFile);

        res.json({ processedData });
    } catch (error) {
        console.error('Erreur de traitement:', error);
        res.status(500).json({ 
            error: 'Erreur lors du traitement des fichiers',
            details: error.message 
        });
    }
});

app.listen(port, () => {
    console.log(`Serveur démarré sur http://localhost:${port}`);
});