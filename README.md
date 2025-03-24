# EZproject
# EZproject

🧾 Script d'automatisation du reporting commercial avec extraction de données via Selenium.

Ce projet permet de :
- Accéder automatiquement à une interface de gestion de commandes
- Extraire les données pertinentes (clients, dates, montants)
- Générer des numéros de factures formatés
- Ajouter les nouvelles données dans un fichier Excel local

---

## Fonctionnalités

- Accès automatisé à un site de gestion (via Selenium WebDriver)
- Lecture et écriture dans un fichier Excel existant (`.xlsx`)
- Conversion automatique de numéros de commande → facture
- Calcul des montants HT (20%) et TTC

---

## ⚙️ Prérequis

- Python 3.7+
- Google Chrome installé
- ChromeDriver compatible avec ta version de Chrome
