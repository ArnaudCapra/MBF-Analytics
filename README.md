# 💰 MBF-Analytics
### Classe VBA de méthodes financières basées sur RapidAPI

---

## 🚀 Objectif du projet

**MBF Analytics** est un projet collaboratif visant à créer une **classe VBA** regroupant un ensemble de **méthodes financières fiables, simples et vérifiées**, s’appuyant sur des **données issues de RapidAPI**.

Le but est d’offrir une bibliothèque **facile à utiliser** et **robuste**, permettant aux utilisateurs d’Excel de :

- récupérer des **données financières externes** (actions, devises, indices, etc.),
- tout en restant **100 % VBA natif**, sans dépendances externes lourdes.

---

## ⚙️ Fonctionnalités principales

| Type de méthode | Exemple | Description |
|------------------|----------|--------------|
| Indicateurs financiers | `Call m.bloomberg_financials(sheetname:="Orange Soc", symbol:="ORA:FP", currencyname:="EUR")` |  Télécharge les indicateurs classiques annuels et trimestriels |
| Historique des cotations | `Call m.real_time_quotes1(sheetname:="Orange Sto", interval:="4hour", symbol:="ORA.PA", fromdt:="2025-06-01", untildt:="2025-11-01")` | Télécharge l'historique récent des cotations |

Toutes les méthodes sont :
- 🔍 **Simples à utiliser** (appel direct depuis Excel, illustré d'un exemple complet)
- 🔍 **Robustes** (gestion d’erreurs et d’API incluse)
- 🔍 **Vérifiées par un tiers** avant validation
- 🔍 **Documentées** avec un lien vers la page officielle de l'API et un exemple fonctionnel

---
## 📥 Téléchargement

Obtenez la dernière version stable de **MBF-Analytics** :

### Dernière release
[MBF-Analytics](https://github.com/MbfColab/MBF-Analytics/releases/tag/Latest)
![Dernière Release](https://img.shields.io/github/v/release/MbfColab/MBF-Analytics?label=latest)

### Autres releases
- [Toutes les releases](https://github.com/MbfColab/MBF-Analytics/releases)
- Téléchargement ZIP : [Download ZIP](https://github.com/MbfColab/MBF-Analytics/archive/refs/heads/main.zip)
---

## 🚀 Exemple d’utilisation

```vba
Sub test_MBFanalytics()

    Dim m As mbfAnalytics
    Set m = New mbfAnalytics
    
    ' Inscrire sa propre clé
    m.initKey "XXXXXXX"
    
    ' Test Bloomberg Financial
    Call m.bloomberg_financials(sheetname:="Orange Soc", symbol:="ORA:FP", currencyname:="EUR")
        
    ' Test Real Time Quotes1
    Call m.real_time_quotes1(sheetname:="Orange Sto", interval:="4hour", symbol:="ORA.PA", fromdt:="2025-06-01", untildt:="2025-11-01")

    'Test Yahou finance historical (Théo, Mathieu, Noé)
    Call m.yhf_historical(sheetname:="TSLA Hist", ticker:="TSLA", sdate:="2024-01-01", edate:="2024-12-31")

'test financial data (elio armarnd )
 Call m.cash_flow(sheetname:="cashflow", symbol:="AAPL:NASDAQ", period:="QUARTERLY")

    ' Test invocation Financial Modeling (Lilou, Juliette S)
    Call m.financial_modeling_prep(sheetname:="Cashflow statement", period:="FY", limit:="50", symbol:="AAPL")

    'Test BabounSkoff Sentiment stock api (Yanis, Pierre-Alexandre, Nicolas)
    Call m.Sentiment_Stock(sheetname:="Tesla", ticker:="TSLA")

  ' Exemple invocation Bloomberg Financial ( quentin, baptiste, kyan, tugdual)
    Call m.CryptoMajors(sheetname:="Vide", limit:=10)

 ' Exemple currency extracter (Selyana, Alissia)
   Call m.currency_extractor(sheetname:="currency_extractor", start_date:="2019-01-01", end_date:="2019-06-01", base:="USD", symbols:="EUR, AFN")

  ' Exemple extraction d'options (Phuong-Linh NHU, Agathe CANEL, Vanessa JIN)
    Call m.Options("Options", "NVDA", "en-US", "US")

' Exemple MacroCalendar(Valentin Grange, Marion Neveu Perret)
  Call m.MacroCalendar("Calendrier", "2025-11-17", "2025-11-23", "US", "GMT+1", "NONE", "500")

'  Exemple invocation insider_reporting (Alexandre Castagna, Louis Simon De Grimaudet De Rochebouët, Arnaud Capra)
    Call m.insider_reporting(sheetname:="NVDA insiders", symbol:="NVDA")



End Sub


End Sub
```

<p align="center">© MBF Assas — 2025</p>
