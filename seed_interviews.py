"""Seed script to insert hypothetical interview records into the database.

Run:
    python seed_interviews.py
"""

from database import SessionLocal
from models import OnboardingSession


def main() -> None:
    session = SessionLocal()

    record_1_original = {
        "How many Active Directory domains are in scope, and what are their names?": "Abbiamo 2 domini: acmecorp.local e acmecorp.cloud",
        "What is the current process for account creation, modification, and disabling in Active Directory?": "La creazione avviene tramite ticket ServiceNow, la modifica è gestita manualmente dall'IT e la disabilitazione è automatica alla data di fine contratto",
        "What is the current OU hierarchy structure in Active Directory, and how is it organized (e.g., by country, organization, function)?": "La struttura è organizzata per paese al primo livello e poi per dipartimento: IT, HR, Finance, Operations",
        "Are there dedicated OUs for employees and separate OUs for external users?": "Sì, abbiamo una OU separata chiamata 'External_Users' per i consulenti e contractor",
        "Are there dedicated OUs for different countries or geographic regions?": "Sì, abbiamo OU dedicate per Italia, Germania, Francia e UK",
        "Is there a dedicated OU (or set of OUs) for groups, and how is it structured?": "Sì, esiste una OU 'Groups' suddivisa in 'Security_Groups' e 'Distribution_Lists'",
        "Are Office 365 licenses assigned via group membership in Active Directory or Azure AD?": "Le licenze O365 sono assegnate tramite gruppi di Azure AD con assegnazione automatica basata sul dipartimento",
        "What is the approximate number of groups defined, and are meaningful descriptions available for these groups?": "Circa 450 gruppi, di cui circa il 70% ha una descrizione significativa nel campo description",
        "Are there standard groups that must be automatically assigned to specific user classes (e.g., by organization, country, role)?": "Sì, ogni utente riceve automaticamente un gruppo paese e un gruppo dipartimento alla creazione",
        "What is the current Azure AD Connect synchronization schedule (frequency, timing)?": "La sincronizzazione avviene ogni 30 minuti tramite Azure AD Connect v2",
        "Are there any customizations or particular/critical scenarios in Active Directory or Azure AD that we should be aware of for the IGA integration?": "Abbiamo un sistema legacy HR che sincronizza attributi custom ogni notte e non deve essere sovrascritto dall'IGA",
        "What is the approximate number of users to be managed per country?": "Italia: 1200, Germania: 800, Francia: 500, UK: 300",
        "Which environments are available for Active Directory (e.g., PROD only, or also DEV/TEST/UAT)?": "Disponibili PROD e TEST, il TEST è un mirror mensile del PROD",
        "Where is Active Directory located in the network, and which firewall rules or network constraints are relevant for the IGA integration?": "AD è nella rete interna LAN, raggiungibile dalla DMZ tramite porta LDAPS 636 e HTTPS 443",
    }

    record_1_english = {
        "How many Active Directory domains are in scope, and what are their names?": "We have 2 domains: acmecorp.local and acmecorp.cloud",
        "What is the current process for account creation, modification, and disabling in Active Directory?": "Account creation is handled via ServiceNow tickets, modifications are managed manually by IT, and disabling is automated based on contract end date",
        "What is the current OU hierarchy structure in Active Directory, and how is it organized (e.g., by country, organization, function)?": "The structure is organized by country at the first level and then by department: IT, HR, Finance, Operations",
        "Are there dedicated OUs for employees and separate OUs for external users?": "Yes, we have a separate OU called 'External_Users' for consultants and contractors",
        "Are there dedicated OUs for different countries or geographic regions?": "Yes, we have dedicated OUs for Italy, Germany, France, and UK",
        "Is there a dedicated OU (or set of OUs) for groups, and how is it structured?": "Yes, there is a 'Groups' OU divided into 'Security_Groups' and 'Distribution_Lists'",
        "Are Office 365 licenses assigned via group membership in Active Directory or Azure AD?": "O365 licenses are assigned via Azure AD groups with automatic assignment based on department",
        "What is the approximate number of groups defined, and are meaningful descriptions available for these groups?": "Approximately 450 groups, of which about 70% have a meaningful description in the description field",
        "Are there standard groups that must be automatically assigned to specific user classes (e.g., by organization, country, role)?": "Yes, each user automatically receives a country group and a department group upon creation",
        "What is the current Azure AD Connect synchronization schedule (frequency, timing)?": "Synchronization occurs every 30 minutes via Azure AD Connect v2",
        "Are there any customizations or particular/critical scenarios in Active Directory or Azure AD that we should be aware of for the IGA integration?": "We have a legacy HR system that synchronizes custom attributes nightly and must not be overwritten by IGA",
        "What is the approximate number of users to be managed per country?": "Italy: 1200, Germany: 800, France: 500, UK: 300",
        "Which environments are available for Active Directory (e.g., PROD only, or also DEV/TEST/UAT)?": "PROD and TEST are available, TEST is a monthly mirror of PROD",
        "Where is Active Directory located in the network, and which firewall rules or network constraints are relevant for the IGA integration?": "AD is located in the internal LAN, reachable from the DMZ via LDAPS port 636 and HTTPS 443",
    }

    record_2_original = {
        "Is the target system deployed on‑premises?": "Sì, il sistema è completamente on-premises nel datacenter di Milano",
        "Where is the target system located in the network (e.g., LAN, Secure Zone, DMZ)?": "Si trova nella Secure Zone, accessibile solo dalla rete interna tramite VPN",
        "Does the target system use a local database as its primary authentication mechanism?": "Sì, l'autenticazione primaria è gestita da un database locale Oracle",
        "Does the target system use a database (e.g., SQL) as its authorization system?": "Sì, l'autorizzazione è gestita tramite tabelle Oracle dedicate con ruoli e profili",
        "Which database technology is used by the target system (e.g., Oracle, Microsoft SQL Server, other)?": "Oracle 19c Enterprise Edition",
        "Is there a separate instance of the target system for each country or business unit?": "No, è un'istanza unica centralizzata con segregazione logica per business unit tramite attributi",
        "Are there currently any relevant activities or projects ongoing that could impact the IGA integration for this target system?": "È in corso un upgrade a Oracle 21c previsto per Q3 2026, ma non dovrebbe impattare le tabelle utenti",
        "Is it necessary to create user records in the local database in order to grant access to the application?": "Sì, è necessario creare un record nella tabella USERS per ogni utente che deve accedere",
        "What is the key/username that must be used to uniquely identify users in the local database?": "Il campo USER_ID che corrisponde alla matricola aziendale dell'utente",
        "Is it possible to create more than one record in the local database for the same user (e.g., multiple profiles or accounts)?": "No, ogni utente ha un unico record, i profili multipli sono gestiti tramite la tabella USER_ROLES",
        "Is there a basic security group in Active Directory that is required to allow access to the application?": "Sì, è richiesta l'appartenenza al gruppo AD 'APP_HRPORTAL_ACCESS' per l'accesso base",
        "Is user profiling (e.g., roles, permissions) managed directly in the local database?": "Sì, il profiling è gestito interamente nel database locale tramite ruoli assegnati nella tabella USER_ROLES",
        "Is there a separate table in the database dedicated to access rights/permissions?": "Sì, la tabella ROLES contiene tutti i diritti di accesso disponibili",
        "Which attributes are present in the access‑rights table in the database (e.g., code, description, type, scope)?": "ROLE_ID, ROLE_NAME, ROLE_DESCRIPTION, ROLE_TYPE (admin/standard/readonly), BUSINESS_UNIT",
        "Is there a user‑friendly description stored in the database for each access right?": "Sì, il campo ROLE_DESCRIPTION contiene una descrizione leggibile per ogni ruolo",
        "Is profiling in the database managed via specific attributes on the user table, or via a separate table that manages assignments between access rights and users?": "Tramite una tabella separata USER_ROLES che gestisce l'associazione N:M tra utenti e ruoli",
        "Will a dedicated service account be created for the database connector with read/write permissions, and what are the expected privileges for this account?": "Sì, verrà creato l'account SVC_IGA_CONNECTOR con permessi SELECT, INSERT, UPDATE e DELETE sulle tabelle USERS e USER_ROLES",
        "Is the target system available only in the PROD environment, or are there additional environments (e.g., DEV, TEST, UAT)?": "Disponibili DEV, UAT e PROD. DEV viene aggiornato settimanalmente da PROD",
    }

    record_2_english = {
        "Is the target system deployed on‑premises?": "Yes, the system is fully on-premises in the Milan datacenter",
        "Where is the target system located in the network (e.g., LAN, Secure Zone, DMZ)?": "It is located in the Secure Zone, accessible only from the internal network via VPN",
        "Does the target system use a local database as its primary authentication mechanism?": "Yes, primary authentication is managed by a local Oracle database",
        "Does the target system use a database (e.g., SQL) as its authorization system?": "Yes, authorization is managed via dedicated Oracle tables with roles and profiles",
        "Which database technology is used by the target system (e.g., Oracle, Microsoft SQL Server, other)?": "Oracle 19c Enterprise Edition",
        "Is there a separate instance of the target system for each country or business unit?": "No, it is a single centralized instance with logical segregation by business unit via attributes",
        "Are there currently any relevant activities or projects ongoing that could impact the IGA integration for this target system?": "An upgrade to Oracle 21c is planned for Q3 2026, but it should not impact user tables",
        "Is it necessary to create user records in the local database in order to grant access to the application?": "Yes, a record in the USERS table must be created for each user who needs access",
        "What is the key/username that must be used to uniquely identify users in the local database?": "The USER_ID field which corresponds to the employee's corporate ID number",
        "Is it possible to create more than one record in the local database for the same user (e.g., multiple profiles or accounts)?": "No, each user has a single record; multiple profiles are managed via the USER_ROLES table",
        "Is there a basic security group in Active Directory that is required to allow access to the application?": "Yes, membership in the AD group 'APP_HRPORTAL_ACCESS' is required for basic access",
        "Is user profiling (e.g., roles, permissions) managed directly in the local database?": "Yes, profiling is managed entirely in the local database via roles assigned in the USER_ROLES table",
        "Is there a separate table in the database dedicated to access rights/permissions?": "Yes, the ROLES table contains all available access rights",
        "Which attributes are present in the access‑rights table in the database (e.g., code, description, type, scope)?": "ROLE_ID, ROLE_NAME, ROLE_DESCRIPTION, ROLE_TYPE (admin/standard/readonly), BUSINESS_UNIT",
        "Is there a user‑friendly description stored in the database for each access right?": "Yes, the ROLE_DESCRIPTION field contains a human-readable description for each role",
        "Is profiling in the database managed via specific attributes on the user table, or via a separate table that manages assignments between access rights and users?": "Via a separate USER_ROLES table that manages the N:M association between users and roles",
        "Will a dedicated service account be created for the database connector with read/write permissions, and what are the expected privileges for this account?": "Yes, the account SVC_IGA_CONNECTOR will be created with SELECT, INSERT, UPDATE, and DELETE permissions on the USERS and USER_ROLES tables",
        "Is the target system available only in the PROD environment, or are there additional environments (e.g., DEV, TEST, UAT)?": "DEV, UAT, and PROD are available. DEV is refreshed weekly from PROD",
    }

    try:
        rec1 = OnboardingSession(
            company="Acme Corp",
            target_system="Corporate Active Directory",
            system_type="AD-Azure",
            collected_data_original=record_1_original,
            collected_data_english=record_1_english,
        )
        session.add(rec1)
        session.commit()
        session.refresh(rec1)
        print(f"Inserted record 1 (id={getattr(rec1, 'id', 'n/a')}) for company='Acme Corp'")

        rec2 = OnboardingSession(
            company="Beta Industries",
            target_system="HR Portal Database",
            system_type="Target DB",
            collected_data_original=record_2_original,
            collected_data_english=record_2_english,
        )
        session.add(rec2)
        session.commit()
        session.refresh(rec2)
        print(f"Inserted record 2 (id={getattr(rec2, 'id', 'n/a')}) for company='Beta Industries'")

    except Exception as exc:
        session.rollback()
        raise
    finally:
        session.close()


if __name__ == "__main__":
    main()
