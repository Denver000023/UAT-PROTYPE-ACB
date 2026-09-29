import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


def clean_value(value):
    """Normalize Excel values for comparison."""
    if pd.isna(value):
        return ""

    value = str(value).strip()

    if value.endswith(".0"):
        value = value[:-2]

    value = " ".join(value.split())

    return value.upper()


def process_hs_codes(reference_df, shipment_df):

    required_reference = [
        "Client_Internal_tracking",
        "Goods_Description",
        "Client_HS_code",
        "Adjusted_HS_code"
    ]

    required_shipment = [
        "Client_Internal_tracking",
        "Goods_Description",
        "HS_code"
    ]

    missing_reference = [
        col for col in required_reference
        if col not in reference_df.columns
    ]

    missing_shipment = [
        col for col in required_shipment
        if col not in shipment_df.columns
    ]

    if missing_reference:
        raise ValueError(
            "First file is missing: "
            + ", ".join(missing_reference)
        )

    if missing_shipment:
        raise ValueError(
            "Second file is missing: "
            + ", ".join(missing_shipment)
        )

    result = shipment_df.copy()

    # Keep original HS code
    result.insert(
        result.columns.get_loc("HS_code"),
        "Original_HS_code",
        result["HS_code"].fillna("").astype(str)
    )

    # Add audit columns
    hs_position = result.columns.get_loc("HS_code")

    result.insert(
        hs_position,
        "Reference_Client_HS_code",
        ""
    )

    result.insert(
        hs_position,
        "Reference_Adjusted_HS_code",
        ""
    )

    result.insert(
        hs_position,
        "HS_Code_Status",
        ""
    )

    # ---------------------------------------------------------
    # FIRST FILE LOOKUP
    #
    # Tracking + Goods Description
    # ---------------------------------------------------------

    reference_lookup = {}

    for _, row in reference_df.iterrows():

        tracking = clean_value(
            row["Client_Internal_tracking"]
        )

        description = clean_value(
            row["Goods_Description"]
        )

        client_hs = clean_value(
            row["Client_HS_code"]
        )

        adjusted_hs = clean_value(
            row["Adjusted_HS_code"]
        )

        if not tracking or not description:
            continue

        key = (
            tracking,
            description
        )

        reference_lookup.setdefault(
            key,
            []
        ).append({
            "client_hs": client_hs,
            "adjusted_hs": adjusted_hs
        })

    changed_flags = []

    # ---------------------------------------------------------
    # PROCESS SECOND FILE
    # ---------------------------------------------------------

    for index, row in result.iterrows():

        tracking = clean_value(
            row["Client_Internal_tracking"]
        )

        description = clean_value(
            row["Goods_Description"]
        )

        current_hs = clean_value(
            row["Original_HS_code"]
        )

        key = (
            tracking,
            description
        )

        matches = reference_lookup.get(
            key,
            []
        )

        # -----------------------------------------------------
        # NO PRODUCT MATCH
        # -----------------------------------------------------

        if not matches:

            result.at[
                index,
                "HS_Code_Status"
            ] = "No Match / Review"

            changed_flags.append(False)

            continue

        # -----------------------------------------------------
        # PRIORITY 1
        #
        # Check Client_HS_code FIRST
        #
        # Client_HS_code = INVALID
        # -----------------------------------------------------

        invalid_match = None

        for match in matches:

            if current_hs == match["client_hs"]:

                invalid_match = match
                break

        if invalid_match:

            result.at[
                index,
                "Reference_Client_HS_code"
            ] = invalid_match["client_hs"]

            result.at[
                index,
                "Reference_Adjusted_HS_code"
            ] = invalid_match["adjusted_hs"]

            adjusted_hs = invalid_match["adjusted_hs"]

            if adjusted_hs:

                result.at[
                    index,
                    "HS_code"
                ] = adjusted_hs

                result.at[
                    index,
                    "HS_Code_Status"
                ] = "Invalid → Adjusted"

                changed_flags.append(
                    current_hs != adjusted_hs
                )

            else:

                result.at[
                    index,
                    "HS_Code_Status"
                ] = "Invalid / No Adjustment"

                changed_flags.append(False)

            continue

        # -----------------------------------------------------
        # PRIORITY 2
        #
        # Check Adjusted_HS_code
        #
        # Already VALID
        # -----------------------------------------------------

        valid_match = None

        for match in matches:

            if current_hs == match["adjusted_hs"]:

                valid_match = match
                break

        if valid_match:

            result.at[
                index,
                "Reference_Client_HS_code"
            ] = valid_match["client_hs"]

            result.at[
                index,
                "Reference_Adjusted_HS_code"
            ] = valid_match["adjusted_hs"]

            result.at[
                index,
                "HS_Code_Status"
            ] = "Valid"

            changed_flags.append(False)

            continue

        # -----------------------------------------------------
        # PRIORITY 3
        #
        # Unknown HS code
        # -----------------------------------------------------

        if len(matches) == 1:

            result.at[
                index,
                "Reference_Client_HS_code"
            ] = matches[0]["client_hs"]

            result.at[
                index,
                "Reference_Adjusted_HS_code"
            ] = matches[0]["adjusted_hs"]

        result.at[
            index,
            "HS_Code_Status"
        ] = "No Match / Review"

        changed_flags.append(False)

    result["_hs_changed"] = changed_flags

    return result


def create_excel(result_df):

    output = BytesIO()

    export_df = result_df.drop(
        columns=["_hs_changed"],
        errors="ignore"
    ).copy()

    # Force HS codes to text
    hs_columns = [
        "Original_HS_code",
        "HS_code",
        "Reference_Client_HS_code",
        "Reference_Adjusted_HS_code"
    ]

    for col in hs_columns:

        if col in export_df.columns:

            export_df[col] = (
                export_df[col]
                .fillna("")
                .astype(str)
            )

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        export_df.to_excel(
            writer,
            index=False,
            sheet_name="HS Validation"
        )

    output.seek(0)

    workbook = load_workbook(output)

    worksheet = workbook["HS Validation"]

    headers = {
        cell.value: cell.column
        for cell in worksheet[1]
    }

    hs_col = headers.get("HS_code")
    status_col = headers.get("HS_Code_Status")

    green_fill = PatternFill(
        fill_type="solid",
        fgColor="C6EFCE"
    )

    yellow_fill = PatternFill(
        fill_type="solid",
        fgColor="FFEB9C"
    )

    red_fill = PatternFill(
        fill_type="solid",
        fgColor="FFC7CE"
    )

    green_font = Font(
        bold=True,
        color="006100"
    )

    bold_font = Font(
        bold=True
    )

    red_font = Font(
        bold=True,
        color="9C0006"
    )

    # ---------------------------------------------------------
    # FORMAT RESULT ROWS
    # ---------------------------------------------------------

    for row_number in range(
        2,
        worksheet.max_row + 1
    ):

        status = worksheet.cell(
            row=row_number,
            column=status_col
        ).value

        hs_cell = worksheet.cell(
            row=row_number,
            column=hs_col
        )

        status_cell = worksheet.cell(
            row=row_number,
            column=status_col
        )

        if status == "Invalid → Adjusted":

            hs_cell.fill = green_fill
            hs_cell.font = green_font

            status_cell.fill = green_fill
            status_cell.font = green_font

        elif status == "Valid":

            hs_cell.font = bold_font
            status_cell.font = bold_font

        elif status == "No Match / Review":

            hs_cell.fill = yellow_fill
            status_cell.fill = yellow_fill
            status_cell.font = bold_font

        elif status == "Invalid / No Adjustment":

            hs_cell.fill = red_fill
            status_cell.fill = red_fill
            status_cell.font = red_font

    # ---------------------------------------------------------
    # HEADER
    # ---------------------------------------------------------

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )

    for cell in worksheet[1]:

        cell.fill = header_fill

        cell.font = Font(
            bold=True,
            color="FFFFFF"
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    worksheet.freeze_panes = "A2"

    worksheet.auto_filter.ref = worksheet.dimensions

    # ---------------------------------------------------------
    # COLUMN WIDTH
    # ---------------------------------------------------------

    for column_cells in worksheet.columns:

        max_length = 0

        column_letter = get_column_letter(
            column_cells[0].column
        )

        for cell in column_cells:

            try:

                max_length = max(
                    max_length,
                    len(str(cell.value))
                )

            except Exception:
                pass

        worksheet.column_dimensions[
            column_letter
        ].width = min(
            max_length + 2,
            50
        )

    final_output = BytesIO()

    workbook.save(final_output)

    final_output.seek(0)

    return final_output


# ============================================================
# STREAMLIT MODULE ENTRY POINT
# ============================================================

def run():

    st.title("🔎 APC HS Code Validation")

    st.markdown(
        """
        ### Validation Rules

        **Priority 1 — Invalid**

        The `HS_code` from the second file is checked against
        `Client_HS_code` from the first file.

        If found, it is considered **Invalid** and is replaced
        with the corresponding `Adjusted_HS_code`.

        **Priority 2 — Valid**

        If it is not found in `Client_HS_code`, the application
        checks `Adjusted_HS_code`.

        If found, it is considered **Valid** and remains unchanged.

        **Priority 3 — Review**

        If neither is found, the original HS code is retained
        and marked **No Match / Review**.
        """
    )

    st.divider()

    # --------------------------------------------------------
    # UPLOAD FILES
    # --------------------------------------------------------

    col1, col2 = st.columns(2)

    with col1:

        st.subheader(
            "1️⃣ First File — HS Reference"
        )

        first_file = st.file_uploader(
            "Upload First Excel File",
            type=["xlsx", "xls"],
            key="apc_hs_first_file"
        )

        st.caption(
            "Required: Client_Internal_tracking, "
            "Goods_Description, Client_HS_code, "
            "Adjusted_HS_code"
        )

    with col2:

        st.subheader(
            "2️⃣ Second File — Shipment"
        )

        second_file = st.file_uploader(
            "Upload Second Excel File",
            type=["xlsx", "xls"],
            key="apc_hs_second_file"
        )

        st.caption(
            "Required: Client_Internal_tracking, "
            "Goods_Description, HS_code"
        )

    st.divider()

    # --------------------------------------------------------
    # RUN VALIDATION
    # --------------------------------------------------------

    if first_file and second_file:

        if st.button(
            "🚀 Validate HS Codes",
            type="primary",
            use_container_width=True
        ):

            try:

                with st.spinner(
                    "Checking HS codes..."
                ):

                    reference_df = pd.read_excel(
                        first_file,
                        dtype=str
                    )

                    shipment_df = pd.read_excel(
                        second_file,
                        dtype=str
                    )

                    reference_df.columns = (
                        reference_df.columns
                        .str.strip()
                    )

                    shipment_df.columns = (
                        shipment_df.columns
                        .str.strip()
                    )

                    result = process_hs_codes(
                        reference_df,
                        shipment_df
                    )

                    excel_file = create_excel(
                        result
                    )

                st.success(
                    "HS code validation completed successfully."
                )

                # ------------------------------------------------
                # SUMMARY
                # ------------------------------------------------

                status_counts = (
                    result["HS_Code_Status"]
                    .value_counts()
                )

                total = len(result)

                adjusted = status_counts.get(
                    "Invalid → Adjusted",
                    0
                )

                valid = status_counts.get(
                    "Valid",
                    0
                )

                review = status_counts.get(
                    "No Match / Review",
                    0
                )

                no_adjustment = status_counts.get(
                    "Invalid / No Adjustment",
                    0
                )

                st.subheader(
                    "📊 Validation Summary"
                )

                c1, c2, c3, c4, c5 = st.columns(5)

                c1.metric(
                    "Total",
                    total
                )

                c2.metric(
                    "Adjusted",
                    adjusted
                )

                c3.metric(
                    "Valid",
                    valid
                )

                c4.metric(
                    "Review",
                    review
                )

                c5.metric(
                    "No Adjustment",
                    no_adjustment
                )

                st.divider()

                # ------------------------------------------------
                # PREVIEW
                # ------------------------------------------------

                st.subheader(
                    "🔎 Result Preview"
                )

                preview_columns = [
                    "Client_Internal_tracking",
                    "Goods_Description",
                    "Original_HS_code",
                    "Reference_Client_HS_code",
                    "Reference_Adjusted_HS_code",
                    "HS_code",
                    "HS_Code_Status"
                ]

                preview_columns = [
                    col
                    for col in preview_columns
                    if col in result.columns
                ]

                st.dataframe(
                    result[
                        preview_columns
                    ].head(100),
                    use_container_width=True
                )

                st.divider()

                # ------------------------------------------------
                # DOWNLOAD
                # ------------------------------------------------

                st.subheader(
                    "⬇️ Download"
                )

                st.download_button(
                    label="📥 Download Validated Excel",
                    data=excel_file,
                    file_name=(
                        "APC_HS_Code_Validated_Output.xlsx"
                    ),
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    ),
                    use_container_width=True
                )

            except Exception as e:

                st.error(
                    "Failed to process the Excel files."
                )

                st.exception(e)

    else:

        st.info(
            "Upload both Excel files to begin."
        )
