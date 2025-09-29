import streamlit as st
import pandas as pd
from collections import defaultdict, Counter
from io import BytesIO
import re  # <-- voor het saneren van tabelnamen

st.set_page_config(page_title="BOM Vergelijker", layout="wide")

st.title("BOM Vergelijker")
st.write("Upload een BOM CSV-bestand uit Teamcenter en een D365-exportbestand om verschillen te analyseren.")

# Houd gebruikte tabelnamen bij om uniek te blijven (voor de zekerheid)
_USED_TABLE_NAMES = set()

def _safe_table_name(sheet_name: str) -> str:
    """
    Maakt een geldige en (bij voorkeur) unieke Excel-tabelnaam:
    - Alleen [A-Za-z0-9_]
    - Begint met een letter of underscore
    """
    name = f"T_{sheet_name}"
    # vervang alles wat geen letter/cijfer/underscore is door underscore
    name = re.sub(r'[^A-Za-z0-9_]', '_', name)
    # mag niet met een cijfer beginnen
    if re.match(r'^\d', name):
        name = f"_{name}"
    # beperk extreme lengte (Excel kan veel aan, maar we houden het netjes)
    if len(name) > 128:
        name = name[:128]
    # uniek maken indien nodig
    base = name
    i = 1
    while name in _USED_TABLE_NAMES:
        suffix = f"_{i}"
        name = (base[:128 - len(suffix)]) + suffix
        i += 1
    _USED_TABLE_NAMES.add(name)
    return name

def _write_df_as_table(writer, df: pd.DataFrame, sheet_name: str):
    """
    Schrijf df als ECHTE Excel-tabel (betrouwbaar in alle Excel-versies):
    - Eerst data zonder headers schrijven vanaf rij 2 (startrow=1)
    - Daarna een Excel-tabel plaatsen die rij 1 als header gebruikt
    - Met autofilter, banded rows en eenvoudige 'autofit'
    - Tabelnaam wordt gesaneerd (geen spaties/streepjes e.d.)
    """
    df = df.copy()
    # Zorg dat werkblad bestaat met data, maar zonder header-rij
    df.to_excel(writer, index=False, header=False, sheet_name=sheet_name, startrow=1, startcol=0)
    worksheet = writer.sheets[sheet_name]

    nrows, ncols = df.shape
    columns = [{"header": str(col)} for col in df.columns]

    # Voeg de tabel toe: rij 0 = header, data begint op rij 1
    worksheet.add_table(
        0, 0, max(nrows, 0), max(ncols - 1, 0),
        {
            "columns": columns,
            "name": _safe_table_name(sheet_name),
            "style": "Table Style Medium 9",
            "header_row": True,
            "autofilter": True,
            "banded_rows": True,
        }
    )

    # Schrijf de zichtbare headers (tabel zet ze ook, maar dit maakt het expliciet)
    for idx, col in enumerate(df.columns):
        worksheet.write(0, idx, str(col))

    # Eenvoudig 'autofit' op basis van de eerste ~200 rijen
    for idx, col in enumerate(df.columns):
        max_len = max((len(str(col)),) + tuple(len(str(v)) for v in df[col].head(200).tolist())) if ncols else len(str(col))
        worksheet.set_column(idx, idx, min(max_len + 2, 60))

# Upload Teamcenter BOM-bestand
st.header("📤 Stap 1: Upload Teamcenter BOM-bestand")
uploaded_file = st.file_uploader("Kies een BOM-bestand (CSV met '#' als scheidingsteken)", type=["csv"], key="teamcenter")

# Upload D365-bestand
st.header("📤 Stap 2: Upload D365-bestand")
d365_file = st.file_uploader("Kies een D365-bestand (Excel)", type=["xlsx"], key="d365")

if uploaded_file is not None:
    try:
        # Inlezen Teamcenter CSV
        df_raw = pd.read_csv(uploaded_file, sep=r'\(#\)', engine='python', encoding='ISO-8859-1')
        df_raw.columns = df_raw.columns.str.strip().str.lower().str.replace(r'[^\w]', '', regex=True)

        df = df_raw.rename(columns={
            'parentpart': 'parentpart',
            'qtyper': 'qtyper',
            'item': 'item',
            'template': 'template',
            'makebuy': 'makebuy',
            'linetype': 'linetype',
            'productname': 'productname',
            'level': 'level'
        })

        df['qtyper'] = df['qtyper'].astype(str).str.replace(',', '.').astype(float)

        # Classificatie (ongewijzigd)
        def classify(row):
            makebuy = str(row.get('makebuy', '')).strip().lower()
            linetype = str(row.get('linetype', '')).strip().lower()
            if 'purch' in makebuy:
                return 'buy'
            elif 'production' in makebuy:
                if 'phantom' in makebuy or 'phantom' in linetype:
                    return 'phantom'
                else:
                    return 'make'
            return 'unknown'

        df['type'] = df.apply(classify, axis=1)

        def is_length_item(row):
            return 'mm' in str(row.get('template', '')).lower()

        # Root bepalen
        root_rows = df[df['level'] == 0]
        if root_rows.empty:
            st.error("Geen root item gevonden (level == 0 ontbreekt).")
            st.stop()
        root_item = root_rows['item'].iloc[0]

        # Logs
        trace_log = defaultdict(list)
        length_log = defaultdict(list)
        seen_paths = set()

        # Traverse (ongewijzigd; lengte-items in length_log i.p.v. final_results)
        def traverse(item, multiplier=1, path=[]):
            matches = df[df['parentpart'] == item]
            if matches.empty:
                return

            for _, row in matches.iterrows():
                child = row['item']
                qty = float(row['qtyper'])
                type_ = row['type']
                is_length = is_length_item(row)
                new_path = path + [(item, qty)]

                path_key = tuple(new_path + [(child, qty)])
                if path_key in seen_paths:
                    continue
                seen_paths.add(path_key)

                total_qty = multiplier * qty

                if type_ in ['buy', 'make']:
                    trace_log[child].append((total_qty, new_path + [(child, qty)]))
                    if is_length:
                        length_log[child].append((total_qty, new_path + [(child, qty)]))
                    else:
                        final_results[child] += total_qty
                elif type_ == 'phantom':
                    traverse(child, total_qty, new_path)

        final_results = Counter()
        traverse(root_item, 1, [])

        # Bestellijst (Teamcenter) – zonder lengte-artikelen
        result_df = pd.DataFrame(final_results.items(), columns=['item', 'total_quantity'])
        result_df = result_df.merge(df[['item', 'productname']].drop_duplicates(), on='item', how='left')
        result_df = result_df.groupby(['item', 'productname'], as_index=False)['total_quantity'].sum()
        result_df = result_df.sort_values(by='item')

        st.success("✅ Bestellijst gegenereerd uit Teamcenter")
        st.dataframe(result_df, use_container_width=True)

        # Trace UI (laatste item zonder ×q)
        st.subheader("🔍 Traceer herkomst per artikel")
        trace_item = st.selectbox("Kies een itemnummer om het berekeningspad te zien:", sorted(trace_log.keys()))
        if trace_item:
            for idx, (qty, path) in enumerate(trace_log[trace_item], 1):
                st.markdown(f"**Pad {idx}: totaal {qty} stuks**")
                parts = []
                for j, (i, q) in enumerate(path):
                    if j == len(path) - 1:
                        parts.append(f"{i}")   # laatste zonder (×q)
                    else:
                        parts.append(f"{i} (×{q})")
                path_str = " → ".join(parts)
                st.code(path_str)

        # 📏 Lengte-artikelen uit length_log aggregeren
        length_totals = [(itm, sum(q for q, _ in paths)) for itm, paths in length_log.items()]
        length_df = pd.DataFrame(length_totals, columns=['item', 'total_quantity'])
        if not length_df.empty:
            length_df = length_df.merge(
                df[['item', 'productname', 'template']].drop_duplicates(),
                on='item', how='left'
            )
            length_df = length_df[['item', 'productname', 'total_quantity', 'template']]

        st.subheader("📏 Lengte-artikelen (Teamcenter)")
        if length_df.empty:
            st.info("Geen lengte-artikelen gevonden in de bestellijst.")
        else:
            st.dataframe(length_df, use_container_width=True)

        # Vergelijking met D365
        if d365_file is not None:
            d365_df = pd.read_excel(d365_file, engine='openpyxl')
            d365_df.columns = d365_df.columns.str.strip().str.lower()

            d365_df = d365_df.rename(columns={
                'item number': 'item',
                'product name': 'productname',
                'quantity': 'total_quantity'
            })

            d365_df['total_quantity'] = pd.to_numeric(d365_df['total_quantity'], errors='coerce').fillna(0)
            d365_df = d365_df.groupby(['item', 'productname'], as_index=False)['total_quantity'].sum()

            merged = pd.merge(result_df, d365_df, on='item', how='outer', suffixes=('_teamcenter', '_d365'))

            def compare_rows(row):
                if pd.isna(row['total_quantity_teamcenter']):
                    return '❌ Alleen in D365'
                elif pd.isna(row['total_quantity_d365']):
                    return '❌ Alleen in Teamcenter'
                elif abs(row['total_quantity_teamcenter'] - row['total_quantity_d365']) > 0.01:
                    return '⚠️ Hoeveelheid verschilt'
                elif str(row['productname_teamcenter']).strip() != str(row['productname_d365']).strip():
                    return '⚠️ Naam verschilt'
                else:
                    return '✅ Match'

            merged['status'] = merged.apply(compare_rows, axis=1)
            merged = merged.sort_values(by='item')

            st.subheader("📊 Vergelijking Teamcenter vs D365")
            st.dataframe(merged[['item', 'productname_teamcenter', 'total_quantity_teamcenter',
                                 'productname_d365', 'total_quantity_d365',
                                 'status']], use_container_width=True)

            # Downloadknop (Excel met tabellen)
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                _write_df_as_table(writer, merged, sheet_name='Vergelijking')
                _write_df_as_table(writer, result_df, sheet_name='Bestellijst TeamCenter')

                # Lengte-artikelen altijd als Excel-tabel (ook als leeg)
                if length_df is not None and not length_df.empty:
                    _write_df_as_table(writer, length_df, sheet_name='Lengte-artikelen')
                else:
                    empty_df = pd.DataFrame(columns=['item', 'productname', 'total_quantity', 'template'])
                    _write_df_as_table(writer, empty_df, sheet_name='Lengte-artikelen')

            st.download_button(
                label="📥 Download vergelijking als Excel",
                data=output.getvalue(),
                file_name="BOM_Comparison_Result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Er ging iets mis: {e}")
