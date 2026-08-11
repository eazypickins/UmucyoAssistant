# EazyPickins Prospect Engineering Guide
## How to extract the full procurement contact universe from BigQuery

Purpose: This guide lets you (or any AI assistant) replicate the 50-company prospect + cross-industry test-list work across the ENTIRE BigQuery dataset.

## PART 1 - THE PROMPT (copy-paste this to an AI)

I am the CEO of EazyPickins Ltd, a Rwandan tech company building UmucyoAssistant, a procurement intelligence platform for Rwandan SMEs. We have 14 years of RPPA procurement data in Google BigQuery. Build a COMPLETE prospect database from it.

Data access: query the data through the deployed UmucyoAssistant web app via HTTP POST:
Endpoint: https://script.google.com/macros/s/AKfycbyYFaHR8rVFMh06y0VineXguO2KIcN50KkV5E4jMwsVkBYMWykfaUXx27Ftt9aYNeIW/exec
Method: POST, Content-Type: application/json
Body: {action: action, params: params, uid: admin}

Available actions:
- getIndustryContacts (industry, limit) - returns BUYERS with verified email/phone/contact name, filterable by industry keyword
- getSupplierAwards (name, limit) - returns award count, total value, recent wins for a supplier
- getSupplierBids (name, limit) - returns bid count + parties record (contact person, email, phone) for a supplier
- debugBigQueryAccess - returns the list of all tables in the dataset

## PART 2 - TECHNICAL REFERENCE

### BigQuery location
- Project: rppa-umucyo
- Dataset: rppa_historical
- Tables: rppc_main, rppc_awards, rppc_award_suppliers, rppc_buyer_contacts, rppc_buyer_risk, rppc_buyer_competitors, rppc_prediction_signals, rppc_contracts, rppc_contracts_amendments, rppc_contracts_documents, rppc_contracts_items, rppc_parties, rppc_parties_details_classifications, rppc_related_processes, rppc_tender_documents, rppc_tender_items, rppc_tender_lots, rppc_tender_tenderers, umucyo_live_tenders

### Key table roles
- rppc_buyer_contacts: BUYERS (procuring entities) with entity, contact_name, contact_email, contact_phone - the industry-filterable contact source
- rppc_parties: ALL parties (buyers + suppliers) with name, identifier_legalName, contactPoint_name, contactPoint_email, contactPoint_telephone, roles (e.g. tenderer)
- rppc_tender_tenderers: Who bid on each tender (bid activity)
- rppc_awards + rppc_award_suppliers: Who won awards, with values
- rppc_main: The tender records themselves

### Drive CSV source (if reload needed)
The raw OCDS CSVs live in a Google Drive folder (id 13xnj-6JzfK8PGyyOJXl46tdy8QENPRlu). Mapping: main.csv to rppc_main, awards.csv to rppc_awards, awards_suppliers.csv to rppc_award_suppliers, contracts.csv to rppc_contracts, contracts_amendments.csv to rppc_contracts_amendments, parties.csv to rppc_parties, tender_tenderers.csv to rppc_tender_tenderers, and similar for the rest.

## PART 3 - METHODOLOGY (what worked, what to watch for)

The full extraction approach, pitfalls, and data quality rules from the original work:

### Proven extraction approach
1. Buyers: getIndustryContacts with keyword iteration is the fastest route. Each call returns up to N entities with verified email+phone. Iterate ~30 keywords, dedupe by name+email.
2. Suppliers: getSupplierBids returns bid count + the parties contact record in one call. getSupplierAwards adds award value. Batch these across your target name list.
3. The parties table is gold: rppc_parties holds named contacts for entities with roles = tenderer.

### Known pitfalls
- Partial-name false matches (FAIR matched MAYFAIR INSURANCE; TIGER matched TIGER SOUND STUDIO). Use full names, verify parties.name returned.
- Name normalization needed (LTD vs Ltd vs .LTD; AND vs and).
- Rate limits: space calls ~600ms apart.
- Zero-match does not mean no activity - record as No procurement records found.

### Data quality rules
- Never fabricate contact info.
- Mark unverifiable fields exactly as: Not found
- Label evidence SOURCE-VERIFIED (from RPPA records) or Not found.
- Keep duplicates out; preserve official spellings.
- Directory presence is NOT proof of bidding - use bid/award counts as evidence.
## PART 4 - OUTPUT TEMPLATE

### Master table columns
1. Organization
2. Priority
3. Sector/Industry
4. RPPA Category (if applicable)
5. Procurement Evidence (bid count, award count, total value)
6. Website
7. Email
8. Phone
9. Decision Maker / Role
10. Best Outreach Route
11. Why Approach
12. Source
13. Verification Status

### CEO summary (always include)
- Count of A / B / C prospects
- Count with verified email / phone / named decision-maker
- Count with only general contact route
- Recommended First 50 for immediate outreach (ranked by activity + contactability)

### Example outreach opener (proven format)
We noticed {COMPANY} has submitted {N} bids / won {N} contracts totaling {X} RWF in public procurement. UmucyoAssistant helps contractors like you discover and win more tenders. Could we show you a 15-minute demo?

## Quick-start checklist
1. Test the endpoint with a simple getIndustryContacts call (industry: hospital)
2. Run debugBigQueryAccess to confirm the table list
3. Iterate buyer keywords to build BUYERS_MASTER.csv
4. Build the supplier name list, batch getSupplierBids + getSupplierAwards to build SUPPLIERS_MASTER.csv
5. Dedupe + normalize names
6. Prioritize A/B/C
7. Write the master report + CEO summary
8. Save all CSVs + the report into the project folder

---
Prepared by EazyPickins Ltd - UmucyoAssistant. Deployment v1.40. Guide version 1.0.

