import streamlit as st
import pandas as pd
from io import StringIO

# =========================
# Utility
# =========================
def _normalize_colname(name: str) -> str:
    return str(name).strip().lower()

def _require_columns(df: pd.DataFrame, required: list[str], label: str) -> tuple[bool, list[str]]:
    norm_cols = { _normalize_colname(c): c for c in df.columns }
    missing = [c for c in required if c not in norm_cols]
    return (len(missing) == 0, missing)

def _get(df: pd.DataFrame, colname: str):
    # Prende la colonna originale rispettando il nome reale (case-insensitive)
    mapping = { _normalize_colname(c): c for c in df.columns }
    return df[mapping[colname]]

def _read_excel(file, label: str) -> pd.DataFrame | None:
    try:
        return pd.read_excel(file, engine="openpyxl", dtype=str)
    except Exception as e:
        st.error(f"Errore nel leggere il file '{label}': {e}")
        return None

def _clean_series_to_list(s: pd.Series) -> list[str]:
    return sorted(set([str(x).strip() for x in s.dropna().astype(str) if str(x).strip() != ""]))

# =========================
# Estrazione dai 3 file
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
    display_name = str(row[_get(df, "display name").name]).strip() if pd.notna(row[_get(df, "display name").name]) else None
    manager_display_name = str(row[_get(df, "manager display name").name]).strip() if pd.notna(row[_get(df, "manager display name").name]) else None
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
    Ritorna elenco di GroupName dove MemberUserPrincipalName == UPN.
    """
    req = ["memberuserprincipalname", "groupname"]
    ok, missing = _require_columns(df, req, "EntraGroupMembers")
    if not ok:
        st.error(f"Nel file 'EntraGroupMembers' mancano le colonne: {', '.join(missing)}")
        return []

    member_col = _get(df, "memberuserprincipalname").astype(str).str.strip().str.lower()
    mask = member_col == upn
    if not mask.any():
        return []

    groups = _clean_series_to_list(_get(df.loc[mask], "groupname"))
    return groups

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
) -> list[str]:
    # Soggetto
    if ticket and ticket.strip():
        title = f"[Consip – SR][{ticket.strip()}] Deprovisioning - {display_name or upn}"
    else:
        title = f"Consip – SR Deprovisioning - {display_name or upn}"

    lines = []
    lines.append("Ciao,")
    lines.append(f"per {upn}")
    lines.append("1. Disabilitare l’account di Azure")
    lines.append(f"2. Impostazione Manager con: {manager_display_name or '—'}")
    lines.append("3. Impostare Hide dalla Rubrica")
    lines.append("4. Rimuovere le appartenenze dall’utenza Azure")
    lines.append("5. Rimuovere le applicazioni dall’utenza Azure")
    lines.append("6. Rimozione ruoli")

    step = 7
    # Sezioni dinamiche
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

    # Finali
    lines.append(f"{step}. Rimozione licenze"); step += 1
    lines.append(f"{step}. Cancellare la foto da Azure"); step += 1
    lines.append(f"{step}. Rimozione Wi-Fi")

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
    f_utenti = st.file_uploader("Carica file **Utenti_Azure** (Excel)", type="xlsx", key="utenti")
    f_sm = st.file_uploader("Carica file **SharedMailboxesDetails** (Excel)", type="xlsx", key="smbx")
    f_groups = st.file_uploader("Carica file **EntraGroupMembers** (Excel)", type="xlsx", key="groups")

    st.markdown("---")

    if st.button("Genera Template di Deprovisioning"):
        if not upn_input:
            st.error("Inserisci un UserPrincipalName valido.")
            return

        # Lettura file (obbligatori per le rispettive sezioni)
        df_utenti = _read_excel(f_utenti, "Utenti_Azure") if f_utenti else None
        df_sm = _read_excel(f_sm, "SharedMailboxesDetails") if f_sm else None
        df_groups = _read_excel(f_groups, "EntraGroupMembers") if f_groups else None

        # Estrazione dati (con grazie se alcuni file non sono presenti)
        display_name, manager_display_name = (None, None)
        shared_mailboxes = []
        group_names = []

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

        # Generazione template
        steps = genera_template_deprovisioning(
            upn=upn_input,
            ticket=tt_input,
            display_name=display_name,
            manager_display_name=manager_display_name,
            shared_mailboxes=shared_mailboxes,
            group_names=group_names,
        )

        # Visualizzazione
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

if __name__ == "__main__":
    main()
