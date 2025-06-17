import streamlit as st
import pandas as pd

# Funzione testuale di deprovisioning
def genera_deprovisioning(email: str, ticket: str, cognome: str ,nome: str , manager: str, sm_df: pd.DataFrame) -> list:
    email_lower = email.strip().lower()
    title = f"[Consip – SR][{ticket}] Deprovisioning - {cognome} {nome} (esterno)"
    lines = ["Ciao,", f"per {email_lower}: "]

    step = 1
    # Fixed steps
    fixed = [
        "Disabilitare l’account di Azure",
        f"Impostazione Manager con: {manager}",
        "Impostare Hide dalla Rubrica",
        "Rimuovere le appartenenze dall’utenza Azure",
        "Rimuovere le applicazioni dall’utenza Azure",
        "Rimozione ruoli"
    ]
    for desc in fixed:
        lines.append(f"{step}. {desc}")
        step += 1

    # Step for SM
    sm_list = []
    if not sm_df.empty and sm_df.shape[1] > 2:
        mask = sm_df.iloc[:, 2].astype(str).str.lower() == email_lower
        sm_list = sm_df.loc[mask, sm_df.columns[0]].dropna().tolist()
    if sm_list:
        lines.append(f"{step}. Rimozione abilitazione da SM:")
        for sm in sm_list:
            lines.append(f"   - {sm}")
        step += 1

    # Remaining fixed
    final = [
        "Rimozione licenze",
        "Cancellare la foto da Azure",
        "Rimozione Wi-Fi"
    ]
    for desc in final:
        lines.append(f"{step}. {desc}")
        step += 1

    return [title] + lines

# Streamlit UI

def main():
    st.set_page_config(page_title="Deprovisioning Consip", layout="centered")
    st.title("Deprovisioning Risorsa Azure")

    # Input fields
    nome = st.text_input("Nome", "").strip()
    cognome = st.text_input("Cognome", "").strip()
    email = st.text_input("Email della risorsa Azure", "@consip.it").strip()
    manager = st.text_input("Manager", "").strip()
    ticket = st.text_input("Numero di riferimento Ticket", "").strip()
    st.markdown("---")

    # File uploader for SM
    sm_file = st.file_uploader("Carica file SM (Excel)", type="xlsx")

    if st.button("Genera Template di Deprovisioning"):
        if not email or not manager:
            st.error("Inserisci sia l'email della risorsa che il manager")
            return

        # Read SM dataframe
        sm_df = pd.read_excel(sm_file) if sm_file else pd.DataFrame()

        # Generate steps
        steps = genera_deprovisioning(email, manager, sm_df)

        # Display result
        for line in steps:
            if line.startswith("[") and line.endswith("]"):
                st.subheader(line)
            else:
                st.text(line)

if __name__ == "__main__":
    main()
