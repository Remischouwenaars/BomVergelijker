import streamlit as st
import pandas as pd
from collections import defaultdict, Counter
from io import BytesIO

st.set_page_config(page_title="BOM Vergelijker", layout="wide")

st.title("BOM Generator & Vergelijker")
st.write("Upload een BOM CSV-bestand uit Teamcenter en een D365-exportbestand om verschillen te analyseren.")

# Upload Teamcenter BOM-bestand
st.header("📤 Stap 1: Upload Teamcenter BOM-bestand")
uploaded_file = st.file_uploader("Kies een BOM-bestand (CSV met '#' als scheidingsteken)", type=["csv"], key="teamcenter")

# Upload D365-bestand
st.header("📤 Stap 2: Upload D365-bestand")
d365_file = st.file_uploader("Kies een D365-bestand (Excel)", type=["xlsx"], key="d365")

if uploaded_file is not None:
    try:
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

        root_rows = df[df['level'] == 0]
        if root_rows.empty:
            st.error("Geen root item gevonden (level == 0 ontbreekt).")
            st.stop()
        root_item = root_rows['item'].iloc[0]

        trace_log = defaultdict(list)
        length_log = defaultdict(list)
        seen_paths = set()

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

        result_df = pd.DataFrame(final_results.items(), columns=['item', 'total_quantity'])
        result_df = result_df.merge(df[['item', 'productname']].drop_duplicates(), on='item', how='left')
        result_df = result_df.groupby(['item', 'productname'], as_index=False)['total_quantity'].sum()
        result_df = result_df.sort_values(by='item')

        st.success("✅ Bestellijst gegenereerd uit Teamcenter")
        st.dataframe(result_df, use_container_width=True)

        st.subheader("🔍 Traceer herkomst per artikel")
        trace_item = st.selectbox("Kies een itemnummer om het berekeningspad te zien:", sorted(trace_log.keys()))
        if trace_item:
            for idx, (qty, path) in enumerate(trace_log[trace_item], 1):
                st.markdown(f"**Pad {idx}: totaal {qty} stuks**")
                path_str = " → ".join([f"{i} (×{q})" for i, q in path])
                st.code(path_str)

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
                                 'productname_d365', 'total_quantity_d365', 'status']], use_container_width=True)

            # Downloadknop
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                merged.to_excel(writer, index=False, sheet_name='Vergelijking')

            st.download_button(
                label="📥 Download vergelijking als Excel",
                data=output.getvalue(),
                file_name="BOM_Comparison_Result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Er ging iets mis: {e}")
