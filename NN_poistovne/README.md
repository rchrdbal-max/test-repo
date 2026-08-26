# NN poisťovňa — export z ActiveCampaign

Skripty na vytiahnutie kontaktov a ich aktivít z ActiveCampaign účtu do Excelu/CSV.

> **Dáta nie sú v tomto repozitári a ani tu nikdy byť nesmú.**
> Exporty obsahujú mená, emaily, telefónne čísla, IP adresy a odpovede
> z dotazníkov reálnych ľudí. Tento repozitár je verejný.
> `.gitignore` blokuje `exports/`, `*.csv`, `*.xlsx` aj `.ac_cache*/`.

## Nastavenie

```bash
pip install requests python-dotenv openpyxl
```

Prístup do API sa načítava z `.env` (gitignorované):

```
AC_API_URL=https://<ucet>.api-us1.com
AC_API_KEY=<kluc zo Settings -> Developer>
```

## Použitie

```bash
python3 export_contacts.py     # kontakty, custom polia, lead scores, tagy
python3 export_activities.py   # aktivity -> hárok "Activities"
```

Vznikne `Active_campaign-all data.xlsx` s hárkami `export final` a `Activities`.

## Ako to funguje

`export_contacts.py` neťahá kontakty po jednom, ale hromadne po endpointoch
(`contacts`, `fieldValues`, `contactTags`, `scoreValues`, …) a spojí ich lokálne.
`export_activities.py` rovnako — poskladá aktivity z `logs`, `contactAutomations`,
`contactLists`, `contactTags`, `linkData` a `bounceLogs`.

Odpovede z API sa cachujú do `.ac_cache/`, takže opakované spustenie je okamžité
a nezaťažuje API. Pre čerstvé stiahnutie priečinok zmaž.

Pár vecí, ktoré API vracia inak, než čakáš — skripty ich prevádzajú:

- IP adresa príde ako 32-bit integer → `66.249.93.205`
- `Email Activity` / `Website Scoring` / `Score 3` nie sú custom polia, ale
  lead scores cez `/api/3/scoreValues`
- účet si drží aktivity aj pre už zmazané kontakty; tie sa preskakujú

## Pozor

`archive_contacts.py` a `archive_contacts_v2.py` **nenávratne mažú všetky
kontakty** cez `DELETE /api/3/contacts/{id}`. Nespúšťať omylom.
