import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from io import BytesIO
import re
import zipfile

# XML HEADER EXTRACTOR
def normalize_name(name):
    
    name = str(name).lower().strip()
    name = re.sub(r"[,.&()\-]", "", name)
    name = " ".join(name.split())
    return name

def extract_header(root):
    
    header = root.find("manifestHeader")
    data = {}
    
    if header is None:
        return data
    
    # Simple fields
    for child in header:
        if len(child) == 0:
            data[child.tag] = child.text
    
    # shipFromAddress grouped by role
    for addr in header.findall("shipFromAddress"):
        role = (addr.attrib.get("AddressType", "")).lower()
        data[role] = {
            "name": addr.findtext("name", ""),
            "addressLine1": addr.findtext("addressLine1", ""),
            "city": addr.findtext("city", ""),
            "zip": addr.findtext("zip", ""),
            "stateProvince": addr.findtext(".//stateProvince", ""),
            "countryCode": addr.findtext("countryCode", "")
        }

    return data

# ITEM EXTRACTOR
def extract_items(root):

    items = []

    for item in root.findall(".//shipmentPackageItemDetail"):
        qty_node = item.find(".//quantity")
        weight_node = item.find(".//weightValue")
        money_node = item.find(".//monetaryAmount")

        items.append({
            "asin": item.findtext("asin", ""),
            "itemID": item.findtext("itemID", ""),
            
            "hs_code": item.findtext("destinationHTSCode", "").replace(".", ""),
            "description": item.findtext("harmonizedTariffDescription", ""),
            "country": item.findtext("countryOfOrigin", ""),
            "eccn": item.findtext("ECCN", ""),
            "quantity": qty_node.text if qty_node is not None else "",
            "quantity_uom": qty_node.attrib.get("unitOfMeasure", "") if qty_node is not None else "",
            "weight": weight_node.text if weight_node is not None else "",
            "weight_uom": weight_node.attrib.get("unitOfMeasure", "") if weight_node is not None else "",
            "unit_price": money_node.text if money_node is not None else "",
            "currency": money_node.attrib.get("currencyISOCode", "") if money_node is not None else "",
            "total_value": item.findtext(".//totalUnitValue/monetaryAmount", ""),
            
        })

    return items

# BUILD CANADA ROW
def build_row(header, item, mapping_dict, hs_mapping, mawb_number, program_scope):

    seller = header.get("seller", {})
    receiver = header.get("receiver", {})
    shipper = header.get("shipper", {})
    biller = header.get("biller", {})
    
    original_seller_name = seller.get("name", "")
    merchant_id = normalize_name(header.get("merchantId", ""))

    lookup = mapping_dict.get(merchant_id, {})

    importer_number = lookup.get("importer_number", "")
    importer_party_id = lookup.get("BroderEze Account", "")

    hs_code = item.get("hs_code", "")
    hs_lookup = hs_mapping.get(hs_code, {})

    return {
        
        # ---------------- HEADER INFO ----------------
        "Inco_term": header.get("incoterms", ""),
        "Mode_of_transport": "2",
        
        # ---------------- SELLER ----------------
        "Seller_code": "",
        "Seller_name": str(original_seller_name).strip(),
        "Seller_address": seller.get("addressLine1", ""),
        "Seller_city": seller.get("city", ""),
        "Seller_postal_code": seller.get("zip", ""),
        "Seller_state": seller.get("stateProvince", ""),
        "Seller_country": seller.get("countryCode", ""),
        "Seller_phone_number": "555-555-5555",
        "Seller_email": "email@email.com",
        
        # ---------------- PICKUP (SHIPPER) ----------------
        "Pickup_code": "",
        "Pickup_name": shipper.get("name", ""),
        "Pickup_address": shipper.get("addressLine1", ""),
        "Pickup_city": shipper.get("city", ""),
        "Pickup_postal_code": shipper.get("zip", ""),
        "Pickup_state": shipper.get("stateProvince", ""),
        "Pickup_country": shipper.get("countryCode", ""),
        
        # ---------------- BUYER (RECEIVER) ----------------
        "Buyer_code": "",
        "Buyer_name": receiver.get("name", ""),
        "Buyer_address": receiver.get("addressLine1", ""),
        "Buyer_city": receiver.get("city", ""),
        "Buyer_postal_code": receiver.get("zip", ""),
        "Buyer_province": receiver.get("stateProvince", ""),
        "Buyer_country": receiver.get("countryCode", ""),
        "Buyer_phone_number": "555-555-5555",
        "Buyer_email": "email@email.com",
        
        # ---------------- ORDER ----------------
        "Order_number": header.get("invoiceNumber", ""),
        
        #"Reliable_tracking": (
            #str(header.get("CCN", "")).replace("1BML", "", 1)
            #if str(header.get("CCN", "")).startswith("1BML")
            #else header.get("CCN", "")
        #), original reliable tracking

        "Reliable_tracking": (
            str(header.get("CCN", ""))[4:]
        ),

        "Client_Internal_tracking": header.get("trackingID", ""),
        
        # ---------------- PACKAGE / ITEM ----------------
        "Parcel_item_weight": item.get("weight", ""),
        "Parcel_item_weight_UOM": (
            "LBR" if item.get("weight_uom", "").upper() == "LB"
            else "KGM" if item.get("weight_uom", "").upper() == "KG"
            else item.get("weight_uom", "")
        ),

        "Width": "",
        "Length": "",
        "Height": "",
        "Width_Length_Height_UOM": "",
        
        # ---------------- PRODUCT ----------------
        "Product_code": item.get("asin", ""),
        #"AMAZON_FNSKU": item.get("itemID", ""), # Purpose on this line was reported by requested Sidreck Nisay
        "Currency_code": item.get("currency", ""),
        "Package_no": item.get("quantity", ""),
        "Quantity": item.get("quantity", ""),
        "Quantity_UOM": "PK",
        "Unit_price": item.get("unit_price", ""),
        "UNDG": "",
        "Total_value_of_item": item.get("total_value", ""),
        "Total_value_of_parcel": "",
        
        # ---------------- CUSTOMS ----------------
        "HS_code": item.get("hs_code", ""),
        "Goods_Description": item.get("description", "") + " | " + item.get("itemID", ""),
        "Country_of_origin": (
            "U" + str(seller.get("stateProvince", "")).strip()[:2].upper()
            if item.get("country", "").upper() == "US"
            else item.get("country", "")
        ),
        
        # ---------------- OTHERS ----------------
        "Url": "",
        "Importer_number": importer_number,
        "Importer_party_id": importer_party_id,
        
        # ✅ DEFAULT VALUES (WILL BE OVERWRITTEN IF INPUT PROVIDED)
        "AutoCalc_Provincial_Rate": "C",
        "CBSA_Port_of_Release": st.session_state.get("cbsa_port", "0440"),
        "CBSA_Warehouse_Sub_Location_Code": st.session_state.get("cbsa_wh", ""),
        "Port_of_Discharge": st.session_state.get("cbsa_discharge", ""),
        "Port_of_Discharge_Sublocation Code": st.session_state.get("cbsa_subloc", ""),
        
        "IID_Y/N": "Y",
        "PGA Flag": "CFIA",
        "Category": "HVS",
        "MAWB #": mawb_number,
        #"Carrier code": "1BML", Original Carrier Code

        "Carrier code": (
            str(header.get("CCN", ""))[:4]
        ), 
        
        "Manifest Only": "",
        "Movement Type": "",
        "TARIFF_TREATMENT_CODE": "2",
        "External Reference 2":re.sub(r"-\d{3}$","",str(header.get("invoiceTitle", "")).strip()), # New updated logic to use invoiceTitle
        "GST CODE": "",
        "UOM_Quantity": "", # New column for UOM based on HS code mapping
        "Converted_Quantity": "", # New column for converted quantity based on UOM
        "Container": header.get("PONumber", ""),  # New updated logic to use container as TrackingID

        "_program_scope": program_scope, #identify S-IOR & A-IOR for summary metrics

        #"Container": re.sub(r"-\d{3}$","",str(header.get("invoiceTitle", "")).strip()), #Original logic, but changed to POnumber for better accuracy
        #"External Reference 2": header.get("PONumber", ""), #Original logic, but changed to invoiceTitle for better accuracy
    } 

# CREATE EXCEL
def create_excel(df):

    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(
            writer,
            sheet_name="CANDATA_AMAZON_B2B",
            index=False
        )

        worksheet = writer.sheets["CANDATA_AMAZON_B2B"]

        for i, col in enumerate(df.columns):
            max_len = max(
                df[col].astype(str).map(len).max(),
                len(col)
            )
            worksheet.set_column(i, i, max_len + 5)

    output.seek(0)
    return output


# CREATE AIOR / SIOR EXCEL
def create_aior_sior_excel(df):

    output = BytesIO()

    # Filter using Program_Scope
    sior_df = df[
        df["Program_Scope"]
        .astype(str)
        .str.upper()
        .str.contains("S-IOR", na=False)
    ].copy()

    aior_df = df[
        df["Program_Scope"]
        .astype(str)
        .str.upper()
        .str.contains("A-IOR", na=False)
    ].copy()

    # WHOLE = all records
    whole_df = df.copy()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

        # ---------------- WHOLE ----------------
        whole_df.to_excel(
            writer,
            sheet_name="CANDATA_AMAZON_B2B",
            index=False
        )

        whole_worksheet = writer.sheets["CANDATA_AMAZON_B2B"]

        for i, col in enumerate(whole_df.columns):
            max_len = max(
                whole_df[col].astype(str).map(len).max()
                if not whole_df.empty else 0,
                len(col)
            )
            whole_worksheet.set_column(
                i,
                i,
                max_len + 5
            )

        # ---------------- SIOR ----------------
        sior_df.to_excel(
            writer,
            sheet_name="SIOR",
            index=False
        )

        sior_worksheet = writer.sheets["SIOR"]

        for i, col in enumerate(sior_df.columns):
            max_len = max(
                sior_df[col].astype(str).map(len).max()
                if not sior_df.empty else 0,
                len(col)
            )
            sior_worksheet.set_column(
                i,
                i,
                max_len + 5
            )

        # ---------------- AIOR ----------------
        aior_df.to_excel(
            writer,
            sheet_name="AIOR",
            index=False
        )

        aior_worksheet = writer.sheets["AIOR"]

        for i, col in enumerate(aior_df.columns):
            max_len = max(
                aior_df[col].astype(str).map(len).max()
                if not aior_df.empty else 0,
                len(col)
            )
            aior_worksheet.set_column(
                i,
                i,
                max_len + 5
            )

    output.seek(0)

    return output


# CREATE SELLER FILES ZIP
def create_seller_files_zip(df):

    zip_output = BytesIO()

    seller_df = df.copy()

    
    # CHECK REQUIRED COLUMNS
    required_columns = [
        "Program_Scope",
        "Client_Internal_tracking",
        "Reliable_tracking",
        "Seller_name"
    ]

    missing_columns = [
        col for col in required_columns
        if col not in seller_df.columns
    ]

    if missing_columns:
        raise ValueError(
            f"Seller ZIP generation is missing columns: {missing_columns}"
        )

    
    # REMOVE BLANK SELLER
    seller_df = seller_df[
        seller_df["Seller_name"]
        .fillna("")
        .astype(str)
        .str.strip()
        != ""
    ].copy()

    
    # NORMALIZE PROGRAM SCOPE
    seller_df["Program_Scope"] = (
        seller_df["Program_Scope"]
        .fillna("")
        .astype(str)
        .str.upper()
        .str.strip()
    )

    # EXTRACT MERCHANT ID
    # Client_Internal_tracking example:
    # 116KJKXZL-A207H0TGYJH8M
    # Result:
    # A207H0TGYJH8M
    
    def extract_merchant_id(value):

        value = str(value).strip()

        if "-" in value:
            return value.rsplit("-", 1)[-1].strip()

        return value

    seller_df["Merchant_ID"] = (
        seller_df["Client_Internal_tracking"]
        .apply(extract_merchant_id)
    )

    
    # CCN
    seller_df["CCN"] = (
        seller_df["Reliable_tracking"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # KEEP ONLY AIOR / SIOR
    
    seller_df = seller_df[
        seller_df["Program_Scope"].isin([
            "A-IOR",
            "S-IOR"
        ])
    ].copy()

    # DEBUG DISPLAY
    # This lets us SEE exactly what the ZIP function is using.
    st.subheader("🔍 Seller File Grouping Preview")

    grouping_preview = seller_df[
        [
            "Program_Scope",
            "Seller_name",
            "Client_Internal_tracking",
            "Merchant_ID",
            "Reliable_tracking",
            
        ]
    ].drop_duplicates()

    st.dataframe(
        grouping_preview,
        use_container_width=True
    )

    # GROUPING KEY
    # Program Scope + Merchant ID + CCN
    
    group_columns = [
        "Program_Scope",
        "Merchant_ID",
        "CCN"
    ]

    # CREATE ZIP
    with zipfile.ZipFile(
        zip_output,
        mode="w",
        compression=zipfile.ZIP_DEFLATED
    ) as zip_file:

        for group_key, group_data in seller_df.groupby(
            group_columns,
            sort=True,
            dropna=False
        ):

            program_scope, merchant_id, ccn = group_key

            # FOLDER
            if program_scope == "A-IOR":
                folder_name = "AIOR"

            elif program_scope == "S-IOR":
                folder_name = "SIOR"

            else:
                continue

            # SELLER NAME
            # Only used for filename.
            # NOT used for grouping.
            
            seller_name = (
                group_data.iloc[0]["Seller_name"]
            )

            seller_name = str(
                seller_name
            ).strip()

            # CLEAN SELLER NAME
            safe_seller_name = re.sub(
                r'[<>:"/\\|?*]',
                "_",
                seller_name
            )

            safe_seller_name = (
                safe_seller_name[:150]
                .strip()
            )

            if not safe_seller_name:
                safe_seller_name = "Unknown_Seller"

            
            # CLEAN MERCHANT ID
            safe_merchant_id = re.sub(
                r'[<>:"/\\|?*]',
                "_",
                str(merchant_id).strip()
            )

            
            # CLEAN CCN
            safe_ccn = re.sub(
                r'[<>:"/\\|?*]',
                "_",
                str(ccn).strip()
            )

            
            # FINAL FILE NAME
            # Seller_MerchantID_CCN.xlsx
            
            filename = (
                f"{safe_seller_name}_"
                f"{safe_merchant_id}_"
                f"{safe_ccn}.xlsx"
            )

            
            # CREATE EXCEL
            seller_output = BytesIO()

            # Remove temporary helper columns
            output_data = group_data.drop(
                columns=[
                    "Merchant_ID",
                    "CCN"
                ],
                errors="ignore"
            ).copy()

            with pd.ExcelWriter(
                seller_output,
                engine="xlsxwriter"
            ) as writer:

                output_data.to_excel(
                    writer,
                    sheet_name="CANDATA",
                    index=False
                )

                worksheet = writer.sheets["CANDATA"]

                for i, col in enumerate(
                    output_data.columns
                ):

                    if output_data.empty:
                        max_len = len(col)

                    else:
                        max_len = max(
                            output_data[col]
                            .astype(str)
                            .map(len)
                            .max(),
                            len(col)
                        )

                    worksheet.set_column(
                        i,
                        i,
                        max_len + 5
                    )

            seller_output.seek(0)

            
            # ADD FILE TO ZIP
            zip_file.writestr(
                f"{folder_name}/{filename}",
                seller_output.getvalue()
            )

    zip_output.seek(0)

    return zip_output

def extract_bol_mapping(root, bol_filename):
    """
    Returns:
        {
            loadNumber: {
                "bill_number": billOfLadingNumber,
                "file_name": BOL XML filename
            }
        }
    """

    header = root.find("manifestHeader")

    if header is None:
        return {}

    load_number = header.findtext("loadNumber", "").strip()
    bill_number = header.findtext("billOfLadingNumber", "").strip()

    if not load_number:
        return {}

    return {
        load_number: {
            "bill_number": bill_number,
            "file_name": bol_filename
        }
    }


# STREAMLIT APP
def run():

    st.subheader("📄 XML → CANDATA UPLOAD FILE")
    st.caption("Amazon XML to CANDATA UPLOAD FILE")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        uploaded_files = st.file_uploader(
            "**Upload CI XML Files**",
            type=["xml"],
            accept_multiple_files=True
        )

    with col2:
        bol_files = st.file_uploader(
            "**Upload BOL XML Files**",
            type=["xml"],
            accept_multiple_files=True
        )

    with col3:
        mapping_file = st.file_uploader(
            "**Upload Account Setup Tracker**",
            type=["xlsx", "xls"]
        )

    with col4:
        hs_file = st.file_uploader(
            "**Upload HS Code Master File**",
            type=["xlsx", "xls"]
        )
    
    st.info("Note: Please update your Account Setup Tracker Excel file using the latest online template before uploading. Account Setup Tracker is based on normalized seller names; minor variations may be accepted, but significant differences may cause mapping failures. Please also ensure accurate data entry. CANDATA is strict about formatting, including spaces, special characters (e.g., commas and periods), and spelling. Careful attention to these details will help prevent errors and ensure smoother processing..")

    st.markdown("---")

    # 🔥 NEW INPUT BOXES

    st.subheader("⚙️ CBSA Overwrite Defaults")
    st.info("Optionally overwrite default CBSA values for Port of Release, Warehouse Sub Location Code, Port of Discharge, and Port of Discharge Sublocation Code. If left blank, defaults will be used in the output file.")

    col1, col2 = st.columns(2)

    with col1:
        st.session_state.cbsa_port = st.text_input("**CBSA Port of Release**", "0440")
        st.session_state.cbsa_wh = st.text_input("**CBSA Warehouse Sub Location Code**", "")

    with col2:
        st.session_state.cbsa_discharge = st.text_input("**Port of Discharge**", "")
        st.session_state.cbsa_subloc = st.text_input("**Port of Discharge Sublocation Code**", "")

    st.markdown("---")
    st.caption("© 2026 ACB Toolkit | Developed by IT Department")


    if not uploaded_files:
        return
    
    # LOAD HS MASTER
    hs_mapping = {}

    if hs_file:

        hs_df = pd.read_excel(hs_file)

        hs_df.columns = (
            hs_df.columns
            .str.strip()
            .str.upper()
        )

        for _, row in hs_df.iterrows():

            hs_code = (
                str(row.get("HS_CODE", ""))
                .replace(".", "")
                .strip()
            )

            hs_mapping[hs_code] = {
                "uom": str(row.get("UOM", "")).strip().upper()
            }
    
    # BOL Mapping
    bol_mapping = {}

    if bol_files:

        for bol_file in bol_files:

            try:
                tree = ET.parse(bol_file)
                root = tree.getroot()

                bol_mapping.update(
                    extract_bol_mapping(
                        root,
                        bol_file.name
                    )
                )

            except Exception as e:
                st.error(
                    f"BOL XML Error: {bol_file.name} - {e}"
                )  
        
    # LOAD MAPPING
    mapping_dict = {}

    if mapping_file is not None:
        mapping_df = pd.read_excel(mapping_file)
        mapping_df.columns = mapping_df.columns.str.strip()

        for _, row in mapping_df.iterrows():
            token_id = normalize_name(row.get("Token ID", ""))

            mapping_dict[token_id] = {
                "importer_number": str(row.get("Importer Number", "")).strip(),
                "BroderEze Account": str(row.get("BroderEze Account", "")).strip(),
                "program_scope": str(row.get("Program / Scope", "")).strip()
            }

    all_rows = []
    missing_mapping_validation = []
    missing_bol = []
    missing_hs_validation = []

    with st.status("Processing files...", expanded=False) as status:
        for uploaded_file in uploaded_files:
            try:
                tree = ET.parse(uploaded_file)
                root = tree.getroot()

                header = extract_header(root)

                #BOL Mapping Validation
                tracking_id = header.get("trackingID", "").strip()

                bol_info = bol_mapping.get(tracking_id, {})

                mawb_number = bol_info.get("bill_number", "")
                bol_filename = bol_info.get("file_name", "")

                if not mawb_number:
                    missing_bol.append({
                        "Tracking ID": tracking_id,
                        "CI - XML File Name": uploaded_file.name,
                        "BOL - XML File Name": bol_filename
                    })

                # Merchant ID Mapping Validation
                merchant_id = str(header.get("merchantId", "")).strip()

                reliable_tracking = (
                    str(header.get("CCN", "")).replace("1BML", "", 1)
                    if str(header.get("CCN", "")).startswith("1BML")
                    else header.get("CCN", "")
                )

                seller_name = header.get("seller", {}).get("name", "")

                if merchant_id:

                    normalized_merchant_id = normalize_name(merchant_id)

                    if normalized_merchant_id not in mapping_dict:

                        missing_mapping_validation.append({
                            "Merchant ID": merchant_id,
                            "Seller Name": seller_name,
                            "Reliable Tracking": reliable_tracking,
                            "File Name": uploaded_file.name,
                            "Issue": "Merchant ID not found in Account Setup Mapping file"
                        })

                items = extract_items(root)

                merchant_id = normalize_name(header.get("merchantId", ""))
                lookup = mapping_dict.get(merchant_id, {})
                program_scope = lookup.get("program_scope", "")

                for item in items:

                    row = build_row(header, item, mapping_dict, hs_mapping, mawb_number, program_scope)
                    #row["_program_scope"] = program_scope
                    all_rows.append(row)

                status.write(f"Completed: {uploaded_file.name}")

            except Exception as e:
                status.write(f"Error: {uploaded_file.name} → {str(e)}")


    status.write(f"✅ **Total Loaded {len(mapping_dict)} Account Setup Tracker Mapping.**")

    status.update(label="**Processing complete**", state="complete")

    if missing_bol:

        st.warning("**Some CI XML files did not find a matching BOL XML**.")

        st.dataframe(
            pd.DataFrame(missing_bol).drop_duplicates(),
            use_container_width=True
        )

    if missing_mapping_validation:

        st.warning(
            "**Validation issues found. These records will still be included in the final output.**"
        )

        validation_df = (
            pd.DataFrame(missing_mapping_validation)
            .drop_duplicates()
        )

        st.subheader("📄 **Merchant ID Mapping Validation**")

        st.dataframe(
            validation_df,
            use_container_width=True
        )


    df = pd.DataFrame(all_rows)

    df["HS_code"] = (
        df["HS_code"]
        .astype(str)
        .str.replace(".", "", regex=False)
        .str.strip()
    )

    df["_uom"] = df["HS_code"].map(
        lambda x: hs_mapping.get(x, {}).get("uom", "")
    )

    # Default blank / "-" UOM to NMB
    df["_uom"] = (
        df["_uom"]
        .replace(["", "-", None], "NMB")
        .fillna("NMB")
    )

    df["UOM_Quantity"] = df["_uom"]

    def converted_quantity(row):

        uom = str(row["_uom"]).upper().strip()

        try:
            qty = float(row["Quantity"])
        except:
            qty = 0

        try:
            weight = float(row["Parcel_item_weight"])
        except:
            weight = 0


        if uom == "DZN":
            return round(qty / 12, 2)

        elif uom == "GRO":
            return round(qty / 144, 2)

        elif uom == "GRM":
            return round(weight * 453.592, 2)

        elif uom == "KGM":
            return round(weight * 0.453592, 2)

        elif uom == "TNE":
            #return round((weight * 0.453592) / 1000, 6)
            return max(round((weight * 0.453592) / 1000,2), 0.01)

        elif uom in [
            "MIL",
            "LTR",
            "NMB",
            "PAR",
            "HLT",
            "LPA",
            "MTQ",
            "TMQ",
            "MWH",
            "KNS",
            "GBQ",
            "MTK",
            "MTR",
            "KSD",
            "TSD",
            "CTM",
            "NAP",
        ]:
            return qty

        return ""

    df["Converted_Quantity"] = df.apply(
        converted_quantity,
        axis=1
    )

    # remove helper column
    df = df.drop(
        columns=["_uom"],
        errors="ignore"
    )

    st.subheader("📊 **Summary Metrics**")

    df["_program_scope"] = df["_program_scope"].astype(str).str.upper().str.strip()
    
    # FILTERS
    sior_entries = df["_program_scope"].str.contains("S-IOR", na=False).sum()
    aior_entries = df["_program_scope"].str.contains("A-IOR", na=False).sum()

    # UNIQUE COUNTS
    sior_unique = df[df["_program_scope"].str.contains("S-IOR", na=False)]["Reliable_tracking"].nunique()
    aior_unique = df[df["_program_scope"].str.contains("A-IOR", na=False)]["Reliable_tracking"].nunique()

    col1, col2 = st.columns(2)

    with col1:
        st.metric("**SIOR Entries**", f"{sior_entries:,}")
        st.metric("**AIOR Entries**", f"{aior_entries:,}")
        
    with col2:
        st.metric("**SIOR Unique**", f"{sior_unique:,}")
        st.metric("**AIOR Unique**", f"{aior_unique:,}")

    unique_tracking_count = df["Reliable_tracking"].nunique()

    duplicate_tracking_count = (
        df["Reliable_tracking"]
        .duplicated(keep=False)
        .sum()
    )

    col1, col2 = st.columns(2)

    with col1:
        st.metric(
            "**Duplicate Reliable Tracking Rows**",
            f"{duplicate_tracking_count:,}"
        )
        
    with col2:
        st.metric(
            "**Unique Reliable Tracking**",
            f"{unique_tracking_count:,}"
        )
        
    df = df.sort_values(by="Reliable_tracking", ascending=True)

    df['Unit_price'] = pd.to_numeric(df['Unit_price'], errors='coerce').fillna(0)
    df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)

    df['Total_value_of_item'] = (df['Unit_price'] * df['Quantity']).round(2)

    df['Total_value_of_parcel'] = df.groupby('Reliable_tracking')['Total_value_of_item'].transform('sum').round(2)

    preview_df = df.copy()

    st.dataframe(preview_df, use_container_width=True)

    # Rename for final output (recommended)
    df = df.rename(columns={"_program_scope": "Program_Scope"})

    #preview_df = df.drop(columns=["_program_scope"], errors="ignore")

    #st.dataframe(preview_df, use_container_width=True)

    #df = df.drop(columns=["_program_scope"])

    excel_data = create_excel(df)

    aior_sior_excel = create_aior_sior_excel(df)

    seller_files_zip = create_seller_files_zip(df)

    col1, col2, col3 = st.columns(3)

    # FULL CANDATA
    with col1:

        st.download_button(
            label="⬇ **Download Canada Upload File**",
            data=excel_data,
            file_name=(
                f"AMAZON_B2B_CANDATA_UPLOAD_"
                f"{pd.Timestamp.now().strftime('%Y%m%d%H%M%S')}.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )

    # AIOR / SIOR
    with col2:

        st.download_button(
            label="⬇ **Download AIOR / SIOR Files**",
            data=aior_sior_excel,
            file_name=(
                f"AMAZON_B2B_AIOR&SIOR_"
                f"{pd.Timestamp.now().strftime('%Y%m%d%H%M%S')}.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )

    # SELLER FILES
    with col3:

        st.download_button(
            label="⬇ **Download Seller Files**",
            data=seller_files_zip,
            file_name=(
                f"AMAZON_B2B_SELLER_FILES_"
                f"{pd.Timestamp.now().strftime('%Y%m%d%H%M%S')}.zip"
            ),
            mime="application/zip"
        )

# ENTRY
if __name__ == "__main__":
    run()