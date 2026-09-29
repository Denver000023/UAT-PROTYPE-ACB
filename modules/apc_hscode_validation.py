import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="HS Code Validator",
    page_icon="📦",
    layout="wide"
)


# ============================================================
# HELPER
# ============================================================

def clean_value(value):
    """
    Clean Excel values for reliable comparison.
    """
    if pd.isna(value):
        return ""

    value = str(value).strip()

    # Excel sometimes converts numeric HS codes to:
    # 3304990000.0
    if value.endswith(".0"):
        value = value[:-2]

    # Normalize multiple spaces
    value = " ".join(value.split())

    return value.upper()


# ============================================================
# VALIDATE / PROCESS
# ============================================================

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

    # --------------------------------------------------------
    # Check columns
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # Copy data
    # --------------------------------------------------------

    result = shipment_df.copy()

    # Preserve original HS code
    result.insert(
        result.columns.get_loc("HS_code"),
        "Original_HS_code",
        result["HS_code"].astype(str)
    )

    # Add reference columns
    result.insert(
        result.columns.get_loc("HS_code"),
        "Reference_Client_HS_code",
        ""
    )

    result.insert(
        result.columns.get_loc("HS_code"),
        "Reference_Adjusted_HS_code",
        ""
    )

    result.insert(
        result.columns.get_loc("HS_code"),
        "HS_Code_Status",
        ""
    )

    # Internal formatting flag
    changed_flags = []

    # --------------------------------------------------------
    # Build reference lookup
    #
    # KEY:
    # Client_Internal_tracking
    # +
    # Goods_Description
    # --------------------------------------------------------

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

        if key not in reference_lookup:
            reference_lookup[key] = []

        reference_lookup[key].append({
            "client_hs": client_hs,
            "adjusted_hs": adjusted_hs
        })

    # --------------------------------------------------------
    # Process shipment rows
    # --------------------------------------------------------

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

        # ----------------------------------------------------
        # No matching product
        # ----------------------------------------------------

        if not matches:

            result.at[
                index,
                "HS_Code_Status"
            ] = "No Match / Review"

            changed_flags.append(False)

            continue

        # ----------------------------------------------------
        # PRIORITY 1:
        #
        # Check current HS against Client_HS_code
        #
        # Client_HS_code = INVALID
        # ----------------------------------------------------

        invalid_match = None

        for match in matches:

            if current_hs == match["client_hs"]:

                invalid_match = match
                break

        if invalid_match:

            client_hs = invalid_match["client_hs"]
            adjusted_hs = invalid_match["adjusted_hs"]

            result.at[
                index,
                "Reference_Client_HS_code"
            ] = client_hs

            result.at[
                index,
                "Reference_Adjusted_HS_code"
            ] = adjusted_hs

            # Only change when adjusted value exists
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

        # ----------------------------------------------------
        # PRIORITY 2:
        #
        # Current HS is not an invalid Client_HS_code.
        # Check whether it is already a valid Adjusted_HS_code.
        # ----------------------------------------------------

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

        # ----------------------------------------------------
        # PRIORITY 3:
        #
        # HS exists neither as invalid nor valid code.
        # Do not automatically change.
        # ----------------------------------------------------

        # If there is only one reference record, show it
        # for review.
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

    # --------------------------------------------------------
    # Internal flags
    # --------------------------------------------------------

    result["_hs_changed"] = changed_flags

    return result


# ============================================================
# EXCEL FORMAT
# ============================================================

def create_excel(result_df):

    output = BytesIO()

    export_df = result_df.drop(
        columns=["_hs_changed"],
        errors="ignore"
    ).copy()

    # Convert HS code columns to strings
    # so leading zeros are preserved.
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

    # --------------------------------------------------------
    # Header map
    # --------------------------------------------------------

    headers = {
        cell.value: cell.column
        for cell in worksheet[1]
    }

    hs_col = headers.get("HS_code")
    status_col = headers.get("HS_Code_Status")

    # --------------------------------------------------------
    # Colors
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # Row formatting
    # --------------------------------------------------------

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

        # --------------------------------------------
        # INVALID → ADJUSTED
        # --------------------------------------------

        if status == "Invalid → Adjusted":

            hs_cell.fill = green_fill
            hs_cell.font = green_font

            status_cell.fill = green_fill
            status_cell.font = green_font

        # --------------------------------------------
        # VALID
        # --------------------------------------------

        elif status == "Valid":

            hs_cell.font = bold_font
            status_cell.font = bold_font

        # --------------------------------------------
        # REVIEW
        # --------------------------------------------

        elif status == "No Match / Review":

            hs_cell.fill = yellow_fill
            status_cell.fill = yellow_fill

            status_cell.font = bold_font

        # --------------------------------------------
        # INVALID WITHOUT ADJUSTMENT
        # --------------------------------------------

        elif status == "Invalid / No Adjustment":

            hs_cell.fill = red_fill
            status_cell.fill = red_fill

            status_cell.font = red_font

    # --------------------------------------------------------
    # Header formatting
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # Freeze header
    # --------------------------------------------------------

    worksheet.freeze_panes = "A2"

    # --------------------------------------------------------
    # Auto filter
    # --------------------------------------------------------

    worksheet.auto_filter.ref = worksheet.dimensions

    # --------------------------------------------------------
    # Column widths
    # --------------------------------------------------------

    for column_cells in worksheet.columns:

        max_length = 0

        column_letter = get_column_letter(
            column_cells[0].column
        )

        for cell in column_cells:

            try:
                length = len(
                    str(cell.value)
                )

                if length > max_length:
                    max_length = length

            except Exception:
                pass

        worksheet.column_dimensions[
            column_letter
        ].width = min(
            max_length + 2,
            50
        )

    # --------------------------------------------------------
    # Final file
    # --------------------------------------------------------

    final_output = BytesIO()

    workbook.save(final_output)

    final_output.seek(0)

    return final_output


# ============================================================
# STREAMLIT UI
# ============================================================

st.title("📦 HS Code Validation & Adjustment")

st.markdown(
    """
This tool validates the **HS_code from the second file**
against the first-file reference.

### Validation priority

**1. `Client_HS_code` = Invalid**

If the second-file `HS_code` is found here,
it will be replaced by the corresponding
`Adjusted_HS_code`.

**2. `Adjusted_HS_code` = Valid**

If the second-file `HS_code` is already here,
it will remain unchanged.

**3. No match**

The HS code will remain unchanged and the row
will be marked **No Match / Review**.
"""
)

st.divider()


# ============================================================
# FILE UPLOAD
# ============================================================

col1, col2 = st.columns(2)

with col1:

    st.subheader("1️⃣ First File — HS Reference")

    first_file = st.file_uploader(
        "Upload reference Excel",
        type=["xlsx", "xls"],
        key="first"
    )

    st.caption(
        "Client_Internal_tracking + "
        "Goods_Description + "
        "Client_HS_code + "
        "Adjusted_HS_code"
    )


with col2:

    st.subheader("2️⃣ Second File — Shipment")

    second_file = st.file_uploader(
        "Upload shipment Excel",
        type=["xlsx", "xls"],
        key="second"
    )

    st.caption(
        "Client_Internal_tracking + "
        "Goods_Description + HS_code"
    )


st.divider()


# ============================================================
# PROCESS
# ============================================================

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

                # Clean headers
                reference_df.columns = (
                    reference_df.columns
                    .str.strip()
                )

                shipment_df.columns = (
                    shipment_df.columns
                    .str.strip()
                )

                # Process
                result = process_hs_codes(
                    reference_df,
                    shipment_df
                )

                # Create Excel
                excel_file = create_excel(
                    result
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

            invalid_no_adjustment = status_counts.get(
                "Invalid / No Adjustment",
                0
            )

            st.success(
                "HS code validation completed."
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
                "Invalid → Adjusted",
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
                "Invalid / No Adjustment",
                invalid_no_adjustment
            )

            st.divider()

            # ------------------------------------------------
            # PREVIEW
            # ------------------------------------------------

            st.subheader(
                "🔎 Validation Preview"
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
                    "HS_Code_Validated_Output.xlsx"
                ),
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),
                use_container_width=True
            )

        except Exception as error:

            st.error(
                f"Error: {error}"
            )

else:

    st.info(
        "Upload both Excel files to start."
    )
