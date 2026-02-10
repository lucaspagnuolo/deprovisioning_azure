import streamlit as st
import pandas as pd
from io import StringIO

# =========================
# Utility
# =========================
def _normalize_colname(name: str) -> str:
    return str(name).strip().lower()

def _require_columns(df: pd.DataFrame, required: list[str], label: str) -> tuple[bool, list[str]]:
    norm_cols = {_normalize_colname(c): c for c in df.columns}
    missing = [c for c in required if c not in norm_cols]
    return (len(missing) == 0, missing)

def _get(df: pd.DataFrame, colname: str):
    # Prende la colonna originale rispettando il nome reale (case-insensitive)
    mapping = {_normalize_colname(c): c for c in df.columns}
    return df[mapping[colname]]

# --- NEW: supporto colonne alternative (EN/IT) ---
def _resolve_any_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """
    Ritorna il nome colonna reale presente in df tra i candidati (case-insensitive),
    oppure None se nessuna è trovata.
    """
    mapping = {_normalize_colname(c): c for c in df.columns}
    for cand in candidates:
        key = _normalize_colname(cand)
        if key in mapping:
            return mapping[key]
    return None

def _get_any(df: pd.DataFrame, candidates: list[str]) -> pd.Series:
    """
    Ritorna la Series della prima colonna trovata tra i candidati.
    Solleva KeyError se nessuna è presente.
    """
    col = _resolve_any_column(df, candidates)
    if col is None:
        raise KeyError(f"Nessuna delle colonne attese trovata: {', '.join(candidates)}")
    return df[col]

def _require_any(df: pd.DataFrame, logical_required: dict[str, list[str]], label: str) -> tuple[bool, list[str]]:
    """
    logical_required: { 'NomeLogico': ['colA', 'colB', ...], ... }
    OK se per ogni NomeLogico esiste almeno una colonna tra le alternative.
    missing contiene i NomiLogici mancanti.
    """
    missing = []
    for logical_name, candidates in logical_required.items():
        if _resolve_any_column(df, candidates) is None:
            missing.append(logical_name)
    return (len(missing) == 0, missing)

def _read_excel(file, label: str) -> pd.DataFrame | None:
    try:
        return pd.read_excel(file, engine="openpyxl", dtype=str)
    except Exception as e:
        st.error(f"Errore nel leggere il file '{label}': {e}")
        return None

def _clean_series_to_list(s: pd.Series) -> list[str]:
    return sorted(set([str(x).strip() for x in s.dropna().astype(str) if str(x).strip() != ""]))

# =========================
# Estrazione dai file
# =========================
def estrai_da_utenti_azure(upn: str, df: pd.DataFrame) -> tuple[str | None, str | None]:
    """
    Ritorna (Display name, Manager display name) per l'UPN, oppure (None, None) se non trovati.
    """
    req = ["user principal name", "display name", "manager display name"]
    ok, missing = _require_columns(df, req, "Utenti_Azure")
    if not ok:
        st.error(f"Nel file 'Utenti_Azure' mancano le colonne: {', '.join(missing)}")
        return None, None

    upn_col = _get(df, "user principal name").astype(str).str.strip().str.lower()
    mask = upn_col == upn
    if not mask.any():
        st.warning("Nessuna corrispondenza trovata in 'Utenti_Azure' per l'UPN specificato.")
        return None, None

    row = df.loc[mask].iloc[0]
    dn_col = _get(df, "display name").name
    mdn_col = _get(df, "manager display name").name
    display_name = str(row[dn_col]).strip() if pd.notna(row[dn_col]) else None
    manager_display_name = str(row[mdn_col]).strip() if pd.notna(row[mdn_col]) else None
    return (display_name or None, manager_display_name or None)

def estrai_shared_mailboxes(upn: str, df: pd.DataFrame) -> list[str]:
    """
    Ritorna elenco di EmailAddress dove Member == UPN.
    """
    req = ["member", "emailaddress"]
    ok, missing = _require_columns(df, req, "SharedMailboxesDetails")
    if not ok:
        st.error(f"Nel file 'SharedMailboxesDetails' mancano le colonne: {', '.join(missing)}")
        return []

    member_col = _get(df, "member").astype(str).str.strip().str.lower()
    mask = member_col == upn
    if not mask.any():
        return []

    emails = _clean_series_to_list(_get(df.loc[mask], "emailaddress"))
    return emails

def estrai_group_members(upn: str, df: pd.DataFrame) -> list[str]:
    """
    Ritorna elenco di GroupName/NomeGruppo dove:
    - MemberUserPrincipalName (EN) oppure UserPrincipalNameMembro (IT) == UPN
    Supporta intestazioni EN/IT.
    """
    required = {
        "member_upn": ["memberuserprincipalname", "userprincipalnamemembro"],
        "group_name": ["groupname", "nomegruppo"],
    }
    ok, missing = _require_any(df, required, "EntraGroupMembers")
    if not ok:
        st.error(f"Nel file 'EntraGroupMembers' mancano i campi: {', '.join(missing)}")
        return []

    member_col = _get_any(df, required["member_upn"]).astype(str).str.strip().str.lower()
    mask = member_col == upn
    if not mask.any():
        return []

    groups = _clean_series_to_list(_get_any(df.loc[mask], required["group_name"]))
    return groups

def estrai_user_mailbox_exists(upn: str, df: pd.DataFrame) -> bool:
    """
    Verifica se esiste una mailbox associata all'UPN cercando in 'ObjectKey'.
    """
    req = ["objectkey"]
    ok, missing = _require_columns(df, req, "UserMailboxes")
    if not ok:
        st.error(f"Nel file 'UserMailboxes' mancano le colonne: {', '.join(missing)}")
        return False

    obj_col = _get(df, "objectkey").astype(str).str.strip().str.lower()
    return (obj_col == upn).any()

def estrai_group_owners_for_user(upn: str, df: pd.DataFrame) -> list[str]:
    """
    Ritorna i GroupName per i quali l'utente (OwnerEmail == upn) risulta Owner.
    """
    req = ["owneremail", "groupname"]
    ok, missing = _require_columns(df, req, "EntraGroupOwners")
    if not ok:
        st.error(f"Nel file 'EntraGroupOwners' mancano le colonne: {', '.join(missing)}")
        return []

    owners = _get(df, "owneremail").astype(str).str.strip().str.lower()
    mask = owners == upn
    if not mask.any():
        return []
    groups = _clean_series_to_list(_get(df.loc[mask], "groupname"))
    return groups

# =========================
# Analisi per avvisi
# =========================
def build_owner_group_warnings(owner_groups: list[str], df_groups: pd.DataFrame | None, upn: str) -> list[str]:
    """
    Per i gruppi dove l'utente è Owner, genera avvisi:
    - elenco dei gruppi per cui è owner
    - se unico utente registrato
    - oppure elenco altri utenti registrati (MemberEmail/EmailMembro preferito; fallback UPN membro),
      escludendo l'UPN
    """
    warnings = []
    if not owner_groups:
        return warnings

    warnings.append(f"Per i seguenti Gruppi {owner_groups} utente indicato è Owner")

    if df_groups is None:
        warnings.append("Impossibile verificare il numero di membri: file 'EntraGroupMembers' non caricato.")
        return warnings

    # Supporto colonne EN/IT su EntraGroupMembers
    group_candidates = ["groupname", "nomegruppo"]
    member_email_candidates = ["memberemail", "emailmembro"]
    member_upn_candidates = ["memberuserprincipalname", "userprincipalnamemembro"]

    # Group column obbligatoria
    grp_col_name = _resolve_any_column(df_groups, group_candidates)
    if grp_col_name is None:
        warnings.append("Nel file 'EntraGroupMembers' manca la colonna del nome gruppo (GroupName/NomeGruppo).")
        return warnings
    grp_col = df_groups[grp_col_name].astype(str).str.strip()

    # Preferisci email membro se presente, altrimenti UPN membro
    mem_col_name = _resolve_any_column(df_groups, member_email_candidates) or _resolve_any_column(df_groups, member_upn_candidates)
    if mem_col_name is None:
        warnings.append("Nel file 'EntraGroupMembers' non sono presenti colonne membri attese (MemberEmail/EmailMembro o MemberUserPrincipalName/UserPrincipalNameMembro).")
        return warnings
    mem_col = df_groups[mem_col_name].astype(str).str.strip()

    # Per ogni gruppo owner, valuta quanti membri e chi
    for grp in owner_groups:
        mask_grp = grp_col == str(grp).strip()
        if not mask_grp.any():
            warnings.append(f"Per il gruppo {grp} non sono presenti membri in 'EntraGroupMembers'.")
            continue

        members_all = _clean_series_to_list(mem_col.loc[mask_grp])
        members_all_lower = [m.lower() for m in members_all]

        # Escludi l'UPN dalle liste da mostrare
        others = [m for m in members_all if m.lower() != upn]

        if len(members_all_lower) == 1 and members_all_lower[0] == upn:
            warnings.append(f"Per il gruppo che è owner {grp} è unico utente registrato")
        elif len(others) > 0:
            warnings.append(f"Per il gruppo che è owner {grp} risultano registrati anche gli utenti {others}")
        else:
            warnings.append(f"Per il gruppo che è owner {grp} non sono emersi altri utenti oltre all'UPN indicato.")
    return warnings

def build_shared_mailbox_last_user_warnings(shared_mailboxes: list[str], df_sm: pd.DataFrame | None, upn: str) -> list[str]:
    """
    Per ciascuna SM trovata, verifica se l'UPN è l'unico membro.
    Se sì, genera l'avviso richiesto.
    """
    warnings = []
    if not shared_mailboxes or df_sm is None:
        return warnings

    if not _require_columns(df_sm, ["member", "emailaddress"], "SharedMailboxesDetails")[0]:
        return warnings

    member_col = _get(df_sm, "member").astype(str).str.strip().str.lower()
    email_col = _get(df_sm, "emailaddress").astype(str).str.strip()

    for sm in shared_mailboxes:
        mask_sm = email_col == sm
        if not mask_sm.any():
            continue
        members_for_sm = _clean_series_to_list(member_col.loc[mask_sm])
        if len(members_for_sm) == 1 and members_for_sm[0] == upn:
            warnings.append(f"Utente {upn} risulta essere ultimo per la Shared {sm}")
    return warnings

# =========================
# Generazione Template
# =========================
def genera_template_deprovisioning(
    upn: str,
    ticket: str | None,
    display_name: str | None,
    manager_display_name: str | None,
    shared_mailboxes: list[str],
    group_names: list[str],
    has_user_mailbox: bool,
) -> list[str]:
    # Soggetto
    if ticket and ticket.strip():
        title = f"[Consip – SR][{ticket.strip()}] Deprovisioning - {display_name or upn}"
    else:
        title = f"Consip – SR Deprovisioning - {display_name or upn}"

    lines = []
    lines.append("Ciao,")
    lines.append(f"per {upn}")

    # Costruiamo dinamicamente gli step numerati
    step_items: list[str] = []
    step_items.append("Disabilitare l’account di Azure")
    step_items.append(f"Impostazione Manager con: {manager_display_name or '—'}")
    step_items.append("Impostare Hide dalla Rubrica")

    # Punto PST se l'utente ha mailbox
    if has_user_mailbox:
        step_items.append(
            f"Estrarre il PST (O365 eDiscovery) da archiviare in "
            r"\nasconsip2....\backuppst\03 - backup email cancellate"
            f"\{upn} (in z7 con psw condivisa)"
        )

    step_items.append("Rimuovere le appartenenze dall’utenza Azure")
    step_items.append("Rimuovere le applicazioni dall’utenza Azure")
    step_items.append("Rimozione ruoli")

    # Aggiungi gli step numerati
    for idx, item in enumerate(step_items, start=1):
        lines.append(f"{idx}. {item}")

    step = len(step_items) + 1

    # Sezioni dinamiche successive
    if shared_mailboxes:
        lines.append(f"{step}. Rimozione abilitazione da SM:")
        for sm in shared_mailboxes:
            lines.append(f"   - {sm}")
        step += 1

    if group_names:
        lines.append(f"{step}. Rimozione gruppi Azure:")
        for g in group_names:
            lines.append(f"   - {g}")
        step += 1

    # Finali (aggiornati: rimosse foto Azure e Wi-Fi)
    lines.append(f"{step}. Rimozione licenze"); step += 1

    return [title] + lines

# =========================
# UI Streamlit
# =========================
def main():
    st.set_page_config(page_title="Deprovisioning Consip", layout="centered")
    st.title("Deprovisioning Risorsa Azure")

    # --- Input ---
    upn_input = st.text_input("UserPrincipalName", "nome.cognome.ext@consip.it").strip().lower()
    tt_input = st.text_input("Inserire il numero TT (opzionale)", "").strip()

    st.markdown("### Carica i file Excel richiesti")
    f_utenti  = st.file_uploader("Carica file **Utenti_Azure** (Excel)", type="xlsx", key="utenti")
    f_sm      = st.file_uploader("Carica file **SharedMailboxesDetails** (Excel)", type="xlsx", key="smbx")
    f_groups  = st.file_uploader("Carica file **EntraGroupMembers** (Excel)", type="xlsx", key="groups")
    f_umbx    = st.file_uploader("Carica file **UserMailboxes** (Excel)", type="xlsx", key="user_mailboxes")
    f_owners  = st.file_uploader("Carica file **EntraGroupOwners** (Excel)", type="xlsx", key="group_owners")

    st.markdown("---")

    if st.button("Genera Template di Deprovisioning"):
        if not upn_input:
            st.error("Inserisci un UserPrincipalName valido.")
            return

        # Lettura file
        df_utenti = _read_excel(f_utenti, "Utenti_Azure") if f_utenti else None
        df_sm = _read_excel(f_sm, "SharedMailboxesDetails") if f_sm else None
        df_groups = _read_excel(f_groups, "EntraGroupMembers") if f_groups else None
        df_umbx = _read_excel(f_umbx, "UserMailboxes") if f_umbx else None
        df_owners = _read_excel(f_owners, "EntraGroupOwners") if f_owners else None

        # Estrazione dati
        display_name, manager_display_name = (None, None)
        shared_mailboxes = []
        group_names = []
        has_user_mailbox = False
        owner_groups = []

        if df_utenti is not None:
            display_name, manager_display_name = estrai_da_utenti_azure(upn_input, df_utenti)
        else:
            st.info("File 'Utenti_Azure' non caricato: il Display name/Manager non saranno popolati.")

        if df_sm is not None:
            shared_mailboxes = estrai_shared_mailboxes(upn_input, df_sm)
        else:
            st.info("File 'SharedMailboxesDetails' non caricato: nessuna SM sarà elencata.")

        if df_groups is not None:
            group_names = estrai_group_members(upn_input, df_groups)
        else:
            st.info("File 'EntraGroupMembers' non caricato: nessun gruppo sarà elencato.")

        if df_umbx is not None:
            has_user_mailbox = estrai_user_mailbox_exists(upn_input, df_umbx)
        else:
            st.info("File 'UserMailboxes' non caricato: non sarà aggiunto il punto PST nel template.")

        if df_owners is not None:
            owner_groups = estrai_group_owners_for_user(upn_input, df_owners)
        else:
            st.info("File 'EntraGroupOwners' non caricato: non verranno mostrati avvisi su gruppi owner.")

        # Generazione template
        steps = genera_template_deprovisioning(
            upn=upn_input,
            ticket=tt_input,
            display_name=display_name,
            manager_display_name=manager_display_name,
            shared_mailboxes=shared_mailboxes,
            group_names=group_names,
            has_user_mailbox=has_user_mailbox,
        )

        # Visualizzazione Template
        for i, line in enumerate(steps):
            if i == 0:
                st.subheader(line)
            else:
                st.text(line)

        # Area testo + download
        st.markdown("---")
        full_text = steps[0] + "\n\n" + "\n".join(steps[1:])
        st.text_area("Anteprima completa", value=full_text, height=320)
        st.download_button(
            label="Scarica come TXT",
            data=full_text.encode("utf-8"),
            file_name=f"deprovisioning_{(display_name or upn_input).replace(' ', '_')}.txt",
            mime="text/plain",
        )

        # =========================
        # Avvisi
        # =========================
        avvisi = []

        # Avvisi Owner Gruppi
        avvisi += build_owner_group_warnings(owner_groups, df_groups, upn_input)

        # Avvisi ultima utenza per SM
        avvisi += build_shared_mailbox_last_user_warnings(shared_mailboxes, df_sm, upn_input)

        if avvisi:
            st.markdown("### Avvisi")
            for msg in avvisi:
                st.warning(msg)

if __name__ == "__main__":
    main()
