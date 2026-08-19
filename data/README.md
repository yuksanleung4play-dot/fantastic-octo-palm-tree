# Holiday list sources

`holidays.csv` is **not invented**. It is compiled from:

1. UK bank holidays (England and Wales): https://www.gov.uk/bank-holidays.json (2026–2028).
2. US federal / bank holidays: Nager.Date public holiday API, `types` containing `Bank` (2026–2029).
3. Cross-checked against LME Notice 26/131 (12 May 2026), *Banking Holidays, Non-Valid Prompt Dates and LME Business Days: 2026 up to and including 2037*. Dates marked **No (for any currencies)** are UK and/or US holidays that are not valid LME prompt dates. JPY-only / EUR-only closures are **not** included, because USD/GBP prompt dates remain valid.

## LME Saturday-observation convention

When a US statutory holiday falls on **Saturday**, US banks typically close the preceding Friday. LME Notice 26/131 does **not** treat that Friday as a non-valid prompt date. Those Fridays are therefore omitted here:

- 2026-07-03 (Independence Day observed) — statutory holiday 2026-07-04 Saturday
- 2027-06-18 (Juneteenth observed) — statutory holiday 2027-06-19 Saturday
- 2027-12-24 (Christmas observed) — statutory holiday 2027-12-25 Saturday
- 2027-12-31 (New Year observed) — statutory holiday 2028-01-01 Saturday
- 2028-11-10 (Veterans Day observed) — statutory holiday 2028-11-11 Saturday

Sunday statutory holidays that US banks observe on Monday **are** included (and listed in the LME notice), e.g. 2027-07-05 Independence Day observed, 2029-11-12 Veterans Day observed.

## 2029 UK dates

gov.uk had not published 2029 England-and-Wales bank holidays at the time this file was built. 2029 UK rows follow LME Notice 26/131 and the standard UK bank-holiday rules (they match).

## Replacing this file

If you have a Bloomberg `CDR <GO>` London export, replace `holidays.csv` with that list (mapped to `Date,Region,Holiday_Name`). The updater will not run without a readable CSV — it never falls back to an empty or hard-coded holiday set.
