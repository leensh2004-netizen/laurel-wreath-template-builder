from __future__ import annotations

import io

import pandas as pd
import streamlit as st

from document_engine import (
    POLICY_SECTIONS,
    NOTE_SECTIONS,
    ALL_SECTIONS,
    TEMPLATE_PATH,
    extract_table_data,
    extract_policy_section_items,
    generate_document,
)
from trial_balance import (
    build_note_rows,
    read_trial_balance,
    rows_to_financial_values,
    build_ppe_movement_cells,
    rows_to_cell_values,
)

st.set_page_config(
    page_title="Laurel Wreath Template Builder",
    page_icon="🌿",
    layout="wide",
)

st.markdown("""
<style>
/* Hide Streamlit's default small running/loading widget so the app feels branded. */
[data-testid="stStatusWidget"] {
    visibility: hidden;
}

/* Soft Laurel brand styling */
.laurel-hero {
    display: flex;
    align-items: center;
    gap: 18px;
    padding: 18px 22px;
    border: 1px solid rgba(176, 143, 67, 0.35);
    border-radius: 20px;
    background: linear-gradient(135deg, rgba(176, 143, 67, 0.12), rgba(255, 255, 255, 0.02));
    margin-bottom: 18px;
}
.laurel-logo-static {
    width: 74px;
    height: 74px;
    border-radius: 50%;
    display: grid;
    place-items: center;
    position: relative;
    background: radial-gradient(circle, rgba(176,143,67,0.22) 0%, rgba(176,143,67,0.08) 58%, transparent 59%);
    border: 1px solid rgba(176, 143, 67, 0.45);
    flex: 0 0 auto;
}
.laurel-logo-static .leaf-left,
.laurel-logo-static .leaf-right {
    position: absolute;
    font-size: 28px;
    color: #b08f43;
    top: 22px;
}
.laurel-logo-static .leaf-left {
    left: 11px;
    transform: rotate(-35deg);
}
.laurel-logo-static .leaf-right {
    right: 11px;
    transform: scaleX(-1) rotate(-35deg);
}
.laurel-logo-static .seal {
    font-size: 22px;
}
.laurel-hero-title {
    font-size: 34px;
    font-weight: 800;
    line-height: 1.15;
    margin: 0;
}
.laurel-hero-subtitle {
    margin-top: 6px;
    opacity: 0.8;
    font-size: 15px;
}

/* Unique Laurel Wreath loading card */
.laurel-loader-card {
    position: relative;
    overflow: hidden;
    display: flex;
    align-items: center;
    gap: 18px;
    padding: 18px 20px;
    border-radius: 18px;
    border: 1px solid rgba(176, 143, 67, 0.45);
    background:
        radial-gradient(circle at 20% 20%, rgba(176, 143, 67, 0.18), transparent 28%),
        linear-gradient(135deg, rgba(176, 143, 67, 0.14), rgba(176, 143, 67, 0.04));
    margin: 12px 0 18px 0;
}
.laurel-loader-card:before {
    content: "";
    position: absolute;
    width: 120px;
    height: 100%;
    left: -140px;
    top: 0;
    background: linear-gradient(90deg, transparent, rgba(255,255,255,0.18), transparent);
    animation: laurel-shine 1.7s ease-in-out infinite;
}
.laurel-loader-mark {
    width: 76px;
    height: 76px;
    position: relative;
    border-radius: 50%;
    display: grid;
    place-items: center;
    background: radial-gradient(circle, rgba(176,143,67,0.25) 0%, rgba(176,143,67,0.09) 58%, transparent 59%);
    border: 1px solid rgba(176, 143, 67, 0.55);
    flex: 0 0 auto;
}
.laurel-loader-mark .doc {
    position: absolute;
    font-size: 26px;
    animation: laurel-doc-pulse 1.25s ease-in-out infinite;
}
.laurel-loader-mark .branch {
    position: absolute;
    font-size: 31px;
    color: #b08f43;
    top: 21px;
    animation: laurel-branch-sway 1.25s ease-in-out infinite;
}
.laurel-loader-mark .branch.left {
    left: 8px;
    transform-origin: 100% 80%;
}
.laurel-loader-mark .branch.right {
    right: 8px;
    transform-origin: 0% 80%;
    animation-name: laurel-branch-sway-right;
}
.laurel-loader-text strong {
    display: block;
    font-size: 18px;
    margin-bottom: 3px;
}
.laurel-loader-text span {
    opacity: 0.78;
    font-size: 14px;
}
.laurel-dots::after {
    content: "";
    animation: laurel-dots 1.2s steps(4, end) infinite;
}

@keyframes laurel-shine {
    0% { left: -140px; }
    55% { left: 110%; }
    100% { left: 110%; }
}
@keyframes laurel-doc-pulse {
    0%, 100% { transform: scale(1); opacity: 0.88; }
    50% { transform: scale(1.12); opacity: 1; }
}
@keyframes laurel-branch-sway {
    0%, 100% { transform: rotate(-36deg); }
    50% { transform: rotate(-18deg); }
}
@keyframes laurel-branch-sway-right {
    0%, 100% { transform: scaleX(-1) rotate(-36deg); }
    50% { transform: scaleX(-1) rotate(-18deg); }
}
@keyframes laurel-dots {
    0% { content: ""; }
    25% { content: "."; }
    50% { content: ".."; }
    75%, 100% { content: "..."; }
}
</style>
""", unsafe_allow_html=True)


def laurel_loader_html(message: str = "Generating your Word document") -> str:
    return f"""
    <div class="laurel-loader-card">
        <div class="laurel-loader-mark" aria-hidden="true">
            <span class="branch left">🌿</span>
            <span class="doc">📄</span>
            <span class="branch right">🌿</span>
        </div>
        <div class="laurel-loader-text">
            <strong>{message}<span class="laurel-dots"></span></strong>
            <span>Laurel Wreath is assembling the Arabic الإيضاحات document and cleaning the formatting.</span>
        </div>
    </div>
    """


st.markdown("""
<div class="laurel-hero">
    <div class="laurel-logo-static" aria-hidden="true">
        <span class="leaf-left">🌿</span>
        <span class="seal">📄</span>
        <span class="leaf-right">🌿</span>
    </div>
    <div>
        <div class="laurel-hero-title">Laurel Wreath for Auditing</div>
        <div class="laurel-hero-subtitle">Document Builder · Fill, edit, extract Excel numbers, review movement tables, and generate the إيضاحات document.</div>
    </div>
</div>
""", unsafe_allow_html=True)

st.caption("Fill the form, upload a trial balance if available, review/edit extracted numbers, choose sections, edit policy wording, and download the generated document.")

if not TEMPLATE_PATH.exists():
    st.error("Template file is missing. Put نموذج الايضاحات.docx inside the templates folder.")
    st.stop()


@st.cache_data(show_spinner=False)
def load_default_policy_items():
    return extract_policy_section_items()


def normalize_text_for_compare(value: str) -> str:
    return " ".join(str(value or "").split())


def join_policy_items(items):
    return "\n\n".join(str(item or "").strip() for item in items if str(item or "").strip())


default_policy_items = load_default_policy_items()


# ------------------------------------------------------------
# Temporary policy mapping by company/client type
# A version: stored only in Streamlit session for testing
# ------------------------------------------------------------

COMPANY_TYPES = [
    "أخرى",
    "صناعية",
    "تجارية",
    "خدمية",
    "بنوك",
    "تأمين",
    "جمعيات خيرية و تعاونية",
    "مستشفيات",
    "مستوصفات",
    "فنادق",
    "استثمارات",
    "مقاولات",
    "عقارية",
]


def policy_types_key(policy_key: str) -> str:
    return f"policy_types_{policy_key}"


def policy_type_checkbox_key(policy_key: str, company_type: str) -> str:
    safe_type = company_type.replace(" ", "_")
    return f"policy_type_checkbox_{policy_key}_{safe_type}"

def custom_policy_type_checkbox_key(index: int, company_type: str) -> str:
    safe_type = company_type.replace(" ", "_")
    return f"custom_policy_type_checkbox_{index}_{safe_type}"


def initialize_policy_type_mapping() -> None:
    """
    Temporary A-version:
    By default, every policy applies to every company type.
    The manager can change the mapping using checkbox boxes.
    Later we will save this mapping permanently in JSON.
    """
    for sec in POLICY_SECTIONS:
        list_key = policy_types_key(sec.key)

        if list_key not in st.session_state:
            st.session_state[list_key] = COMPANY_TYPES.copy()

        selected_types = set(st.session_state.get(list_key, []))

        for company_type_value in COMPANY_TYPES:
            checkbox_key = policy_type_checkbox_key(sec.key, company_type_value)
            if checkbox_key not in st.session_state:
                st.session_state[checkbox_key] = company_type_value in selected_types


def sync_policy_type_list_from_checkboxes(policy_key: str) -> None:
    selected_types = []

    for company_type_value in COMPANY_TYPES:
        checkbox_key = policy_type_checkbox_key(policy_key, company_type_value)
        if st.session_state.get(checkbox_key, False):
            selected_types.append(company_type_value)

    st.session_state[policy_types_key(policy_key)] = selected_types


def get_policy_keys_for_company_type(company_type_value: str) -> set:
    selected_policy_keys = set()

    for sec in POLICY_SECTIONS:
        selected_types = st.session_state.get(policy_types_key(sec.key), [])
        if company_type_value in selected_types:
            selected_policy_keys.add(sec.key)

    return selected_policy_keys


def apply_policy_type_mapping(company_type_value: str) -> None:
    selected_policy_keys = get_policy_keys_for_company_type(company_type_value)

    for sec in POLICY_SECTIONS:
        st.session_state[f"sec_{sec.key}"] = sec.key in selected_policy_keys
initialize_policy_type_mapping()

selected_company_type_for_policies = st.session_state.get("company_type_select", "أخرى")

if selected_company_type_for_policies not in COMPANY_TYPES:
    selected_company_type_for_policies = "أخرى"

should_apply_mapping = False

if st.session_state.get("last_policy_mapping_company_type") != selected_company_type_for_policies:
    should_apply_mapping = True

if st.session_state.pop("pending_apply_policy_mapping", False):
    should_apply_mapping = True

if should_apply_mapping:
    apply_policy_type_mapping(selected_company_type_for_policies)
    st.session_state["last_policy_mapping_company_type"] = selected_company_type_for_policies


with st.sidebar:
    st.header("Document sections")
    st.caption("Uncheck any section that should be removed from the final Word document.")

    st.subheader("Accounting policies")
    included = set()

    st.caption(f"Selected based on company type: {selected_company_type_for_policies}")

    if st.button("Apply policy mapping for this company type"):
        st.session_state["pending_apply_policy_mapping"] = True
        st.rerun()

    for sec in POLICY_SECTIONS:
        keep = st.checkbox(sec.title, key=f"sec_{sec.key}")
        if keep:
            included.add(sec.key)

    st.subheader("Financial notes")
    all_notes = st.checkbox("Keep all financial note sections", value=True)
    for sec in NOTE_SECTIONS:
        value = True if all_notes else sec.default
        keep = st.checkbox(sec.title, value=value, key=f"sec_{sec.key}")
        if keep:
            included.add(sec.key)

    clear_replaced_format = st.checkbox("Remove red/highlight from filled fields", value=True)


main_tab, partners_tab, policies_tab, financial_tables_tab, generate_tab = st.tabs([
    "1) Basic info",
    "2) Partners",
    "3) Accounting policies",
    "4) Financial tables",
    "5) Generate",
])


with main_tab:
    st.subheader("Company information")
    c1, c2 = st.columns(2)
    with c1:
        company_name = st.text_input("Company name / اسم الشركة")

        company_type = st.selectbox(
            "Company type / نوع الشركة",
            COMPANY_TYPES,
            key="company_type_select",
            help="This selects the related accounting policy sections based on the policy mapping. You can still manually adjust them from the sidebar.",
        )

        city = st.text_input("City / المدينة")
        country = st.text_input("Country / الدولة", value="الأردن")
        financial_year = st.text_input("Financial year phrase / السنة المالية")
        registration_number = st.text_input("Registration number / رقم التسجيل")
        registration_date = st.text_input("Registration date / تاريخ التسجيل")
    with c2:
        capital = st.text_input("Capital / رأس المال")
        currency = st.text_input("Currency / العملة", value="دينار أردني")
        po_box = st.text_input("P.O. Box / صندوق البريد")
        postal_code = st.text_input("Postal code / الرمز البريدي")
        approval_date = st.text_input("Financial statements approval date / تاريخ الموافقة")
        current_year = st.text_input("Current year shown in tables", value="2024")
        previous_year = st.text_input("Comparative year shown in tables", value="2023")

    st.subheader("Auditor information")
    c3, c4 = st.columns(2)
    with c3:
        audit_logo_text = st.text_input("Audit logo text / شعار مكتب التدقيق", value="شعار مكتب التدقيق")
        audit_office = st.text_input("Audit office name / اسم مكتب التدقيق")
        audit_partner = st.text_input("Audit partner name / اسم الشريك")
    with c4:
        audit_license = st.text_input("License number / اجازة رقم")
        audit_date = st.text_input("Audit report date / التاريخ")
        st.info("The app avoids replacing every 'اسم الشريك' globally because that phrase is also a partners-table header.")


with partners_tab:
    st.subheader("Partners table")
    num_partners = st.number_input("Number of partners", min_value=0, max_value=30, value=2, step=1)
    partners = []
    for i in range(int(num_partners)):
        st.markdown(f"**Partner {i + 1}**")
        p1, p2, p3, p4 = st.columns(4)
        with p1:
            name = st.text_input("Name", key=f"partner_name_{i}")
        with p2:
            shares = st.text_input("Number of shares", key=f"partner_shares_{i}")
        with p3:
            value = st.text_input("Share value", key=f"partner_value_{i}")
        with p4:
            percentage = st.text_input("Percentage", key=f"partner_pct_{i}")
        partners.append({"name": name, "shares": shares, "value": value, "percentage": percentage})

# These will be populated in the Accounting policies tab and used in Generate.
policy_text_edits = {}
custom_sections = []


with policies_tab:
    st.subheader("Editable accounting policies / السياسات المحاسبية القابلة للتعديل")
    st.info(
        "For each accounting policy, choose which company/client types it applies to, "
    "then edit the policy wording if needed. "
    "After changing the checkboxes, click 'Update included policies from checkboxes' "
    "to refresh the sidebar."
    )

    if st.button("Update included policies from checkboxes"):
        st.session_state["pending_apply_policy_mapping"] = True
        st.rerun()

    st.divider()

    changed_count = 0

    for sec in POLICY_SECTIONS:
        included_now = sec.key in included
        default_items = default_policy_items.get(sec.key, [sec.title])
        if not default_items:
            default_items = [sec.title]

        with st.expander(
            f"{sec.title}" + ("" if included_now else "  — removed for selected company type"),
            expanded=(sec.key == "policy_basis"),
        ):
            st.markdown("**1) Choose which company/client types this policy applies to:**")

            c1, c2, c3, c4, c5 = st.columns(5)

            for idx, company_type_value in enumerate(COMPANY_TYPES):
                checkbox_key = policy_type_checkbox_key(sec.key, company_type_value)

                if idx % 5 == 0:
                    col = c1
                elif idx % 5 == 1:
                    col = c2
                elif idx % 5 == 2:
                    col = c3
                elif idx % 5 == 3:
                    col = c4
                else:
                    col = c5

                with col:
                    st.checkbox(
                        company_type_value,
                        key=checkbox_key,
                    )

            sync_policy_type_list_from_checkboxes(sec.key)

            selected_types = st.session_state.get(policy_types_key(sec.key), [])
            if selected_company_type_for_policies in selected_types:
                st.success(f"This policy applies to: {selected_company_type_for_policies}")
            else:
                st.warning(f"This policy does not apply to: {selected_company_type_for_policies}")

            st.divider()

            st.markdown("**2) Edit policy text:**")

            col_a, col_b, col_c = st.columns([3, 1, 1])
            with col_a:
                st.caption("Leave unchanged to preserve original Word formatting.")
            with col_b:
                if st.button("Reset text", key=f"reset_{sec.key}"):
                    for j, item in enumerate(default_items):
                        st.session_state[f"policy_item_{sec.key}_{j}"] = item
            with col_c:
                st.write("Included ✅" if included_now else "Removed ❌")

            edited_items = []

            for j, default_item in enumerate(default_items):
                item_key = f"policy_item_{sec.key}_{j}"

                if item_key not in st.session_state:
                    st.session_state[item_key] = default_item

                if j == 0:
                    edited_value = st.text_input(
                        "Section heading / عنوان القسم",
                        key=item_key,
                        disabled=not included_now,
                        help="This is separate from the policy points below.",
                    )
                else:
                    height = 115 if len(str(default_item)) < 450 else 190
                    edited_value = st.text_area(
                        f"Point / paragraph {j}",
                        key=item_key,
                        height=height,
                        disabled=not included_now,
                    )

                edited_items.append(edited_value)

            default_joined = join_policy_items(default_items)
            edited_joined = join_policy_items(edited_items)

            if included_now and normalize_text_for_compare(edited_joined) != normalize_text_for_compare(default_joined):
                policy_text_edits[sec.key] = edited_joined
                changed_count += 1

    st.success(f"Edited policy sections ready to apply: {changed_count}")
    st.divider()
    st.subheader("Add new accounting policies / إضافة سياسات محاسبية جديدة")

    if "custom_policy_count" not in st.session_state:
        st.session_state["custom_policy_count"] = 0

    if st.button("➕ Add new accounting policy"):
        st.session_state["custom_policy_count"] += 1
        st.rerun()

    for i in range(st.session_state["custom_policy_count"]):
        with st.expander(f"New accounting policy {i + 1}", expanded=True):
            st.markdown("**1) Choose which company/client types this new policy applies to:**")

            c1, c2, c3, c4, c5 = st.columns(5)

            for idx, company_type_value in enumerate(COMPANY_TYPES):
                checkbox_key = custom_policy_type_checkbox_key(i, company_type_value)

                if checkbox_key not in st.session_state:
                    st.session_state[checkbox_key] = True

                if idx % 5 == 0:
                    col = c1
                elif idx % 5 == 1:
                    col = c2
                elif idx % 5 == 2:
                    col = c3
                elif idx % 5 == 3:
                    col = c4
                else:
                    col = c5

                with col:
                    st.checkbox(
                        company_type_value,
                        key=checkbox_key,
                    )

            custom_selected_types = [
                company_type_value
                for company_type_value in COMPANY_TYPES
                if st.session_state.get(custom_policy_type_checkbox_key(i, company_type_value), False)
            ]

            custom_included_now = selected_company_type_for_policies in custom_selected_types

            if custom_included_now:
                st.success(f"This new policy applies to: {selected_company_type_for_policies}")
            else:
                st.warning(f"This new policy does not apply to: {selected_company_type_for_policies}")

            st.divider()
            st.markdown("**2) Write the new policy text:**")

            custom_title = st.text_input(
                "Policy title / عنوان السياسة",
                key=f"custom_policy_title_{i}",
                disabled=not custom_included_now,
            )

            custom_body = st.text_area(
                "Policy description / نص السياسة",
                key=f"custom_policy_body_{i}",
                height=160,
                disabled=not custom_included_now,
            )

            if custom_included_now and custom_title.strip():
                custom_sections.append({
                    "title": custom_title,
                    "body": custom_body,
                })

# These will be populated in the Excel tab and used in Generate.
# These will be populated in the Financial tables tab and used in Generate.
financial_note_values = {}
cell_values = {}
table_updates = {}


with financial_tables_tab:
    st.subheader("Financial tables / الجداول المالية")
    st.info(
        "Build tables manually and choose which existing Word table they should replace. "
        "The table will stay in the same place in the Word template, but its rows and cells will be replaced."
    )

    table_infos = extract_table_data()

    if "financial_table_replacement_count" not in st.session_state:
        st.session_state["financial_table_replacement_count"] = 0

    if st.button("➕ Add table replacement"):
        st.session_state["financial_table_replacement_count"] += 1
        st.rerun()

    if st.session_state["financial_table_replacement_count"] == 0:
        st.info("Click ➕ Add table replacement to start building a financial table.")

    table_options = {
        f"{t['index']}: {t['label']}": t["index"]
        for t in table_infos
    }

    used_table_indexes = set()

    
    for i in range(st.session_state["financial_table_replacement_count"]):
        with st.expander(f"Table replacement {i + 1}", expanded=True):
            selected_option = st.selectbox(
                "Choose existing Word table to replace",
                list(table_options.keys()),
                key=f"financial_table_target_{i}",
            )

            target_table_index = table_options[selected_option]
            used_table_indexes.add(target_table_index)

            selected_table_info = table_infos[target_table_index]

            show_preview = st.checkbox(
                "Show original Word table preview",
                value=False,
                key=f"show_original_table_preview_{i}",
            )

            if show_preview:
                original_data = selected_table_info.get("data", [])
                if original_data:
                    original_max_cols = max(len(row) for row in original_data)
                    normalized_original = [
                        row + [""] * (original_max_cols - len(row))
                        for row in original_data
                    ]
                    st.dataframe(
                        pd.DataFrame(normalized_original),
                        use_container_width=True,
                        hide_index=True,
                    )
                else:
                    st.write("No preview available for this table.")

            st.markdown("**Build the replacement table:**")
            st.caption("Write the table exactly as you want it to appear. Include the header row too.")

            rows = st.number_input(
                "Number of rows",
                min_value=1,
                max_value=50,
                value=4,
                step=1,
                key=f"financial_table_rows_{i}",
            )

            cols = st.number_input(
                "Number of columns",
                min_value=1,
                max_value=10,
                value=3,
                step=1,
                key=f"financial_table_cols_{i}",
            )

            default_df = pd.DataFrame(
                [["" for _ in range(int(cols))] for _ in range(int(rows))],
                columns=[f"Column {c + 1}" for c in range(int(cols))],
            )

            edited_df = st.data_editor(
                default_df,
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                key=f"financial_table_editor_{i}",
            )

            replacement_data = edited_df.fillna("").astype(str).values.tolist()

            if any(any(cell.strip() for cell in row) for row in replacement_data):
                table_updates[target_table_index] = replacement_data
                st.success(f"This will replace Word table {target_table_index}.")
            else:
                st.warning("This replacement table is empty, so it will not be applied.")



with generate_tab:
    st.subheader("Generate final Word document")
    form = {
        "company_name": company_name,
        "company_type": company_type,
        "city": city,
        "country": country,
        "financial_year": financial_year,
        "registration_number": registration_number,
        "registration_date": registration_date,
        "capital": capital,
        "currency": currency,
        "po_box": po_box,
        "postal_code": postal_code,
        "approval_date": approval_date,
        "current_year": current_year,
        "previous_year": previous_year,
        "audit_logo_text": audit_logo_text,
        "audit_office": audit_office,
        "audit_partner": audit_partner,
        "audit_license": audit_license,
        "audit_date": audit_date,
    }

    st.write("Sections included:", len(included), "of", len(ALL_SECTIONS))
    st.write("Policy sections edited:", len(policy_text_edits))
    st.write("Financial note rows ready to update:", len(financial_note_values))
    st.write("Movement-table cells ready to update:", len(cell_values))
    st.write("Manual table replacements ready:", len(table_updates))
    missing_required = []
    if not company_name:
        missing_required.append("Company name")
    if not financial_year:
        missing_required.append("Financial year")
    if missing_required:
        st.warning("Recommended fields still empty: " + ", ".join(missing_required))

    if st.button("Generate final Word document", type="primary"):
        loader_placeholder = st.empty()
        loader_placeholder.markdown(
            laurel_loader_html("Generating your Word document"),
            unsafe_allow_html=True,
        )

        try:
            output = generate_document(
                form=form,
                partners=partners,
                included_section_keys=included,
                custom_sections=custom_sections,
                table_updates=table_updates,
                financial_note_values=financial_note_values,
                cell_values=cell_values,
                policy_text_edits=policy_text_edits,
                clear_replaced_format=clear_replaced_format,
            )
        finally:
            loader_placeholder.empty()

        safe_name = company_name.strip() or "company"
        file_name = f"ايضاحات_{safe_name}.docx"
        st.success("Document generated.")
        st.download_button(
            label="Download final Word document",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
