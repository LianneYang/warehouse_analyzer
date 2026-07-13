import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image
import io

# Configure the browser tab title and visual layout
st.set_page_config(page_title="Material Volume Analyzer", page_icon="📊", layout="wide")

st.title("📊 Material Volume Share Analyzer & Chart Generator")
st.markdown(
    "Upload your raw data Excel file. The system will automatically aggregate data, generate a collision-free stacked bar chart, and output an engineered Excel summary.")

# 1. Drag-and-Drop File Uploader Component
uploaded_file = st.file_uploader("Upload your raw data Excel file (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Read the file directly into memory
        df = pd.read_excel(uploaded_file)
        st.success("File uploaded successfully! Processing dataset...")

        # --- Mandatory Column Validation ---
        required_cols = ['Material', 'Material Description', 'Customer Name', 'Customer Group', 'QTY']
        missing_cols = [col for col in required_cols if col not in df.columns]

        if missing_cols:
            st.error(
                f"The Excel file is missing required headers: {missing_cols}. Please fix your spreadsheet and try again.")
        else:
            # --- 2. [CORE DATA RESTRUCTURING LOGIC] ---
            group_counts = df.groupby('Material')['Customer Group'].nunique()
            target_materials = group_counts[group_counts > 1].index.tolist()
            filtered_df = df[df['Material'].isin(target_materials)].copy()

            aggregation_cols = ['Material', 'Material Description', 'Customer Name', 'Customer Group']
            aggregated_df = filtered_df.groupby(aggregation_cols, as_index=False)['QTY'].sum()

            stats_df = aggregated_df.groupby('Material').agg(
                Total_Customers=('Customer Name', 'nunique'),
                Global_QTY=('QTY', 'sum')
            ).reset_index()

            final_df = aggregated_df.merge(stats_df, on='Material', how='left')
            final_df = final_df.sort_values(by=['Global_QTY', 'Material', 'QTY'], ascending=[False, True, False])
            sorted_target_materials = final_df['Material'].unique().tolist()

            grouped_rows = []
            output_cols = ['Material', 'Material Description', 'Customer Name', 'Customer Group', 'QTY',
                           'Total Customers for Material']

            for mat in sorted_target_materials:
                mat_df = final_df[final_df['Material'] == mat]
                for _, row in mat_df.iterrows():
                    row_dict = {col: row[col] if col != 'Total Customers for Material' else row['Total_Customers'] for
                                col in output_cols}
                    grouped_rows.append(row_dict)

                total_qty_row = {
                    'Material': f"Total for {mat}", 'Material Description': '', 'Customer Name': '',
                    'Customer Group': '', 'QTY': mat_df['Global_QTY'].iloc[0], 'Total Customers for Material': ''
                }
                grouped_rows.append(total_qty_row)
                grouped_rows.append({col: '' for col in output_cols})

            spaced_df = pd.DataFrame(grouped_rows)

            plot_data_raw = final_df.groupby(['Material', 'Customer Group'])['QTY'].sum().unstack(fill_value=0)
            plot_data = plot_data_raw.reindex(sorted_target_materials)
            material_totals = plot_data.sum(axis=1)

            # --- 3. [CHART GENERATION ENGINE WITH OPTIMIZATIONS] ---
            fig, ax = plt.subplots(figsize=(17, 10))
            customer_groups = plot_data.columns.tolist()
            num_groups = len(customer_groups)
            x_indexes = np.arange(len(sorted_target_materials))

            # Strictly configured Narrow Bars
            bar_width = 0.12
            distinct_colors = plt.cm.tab20(np.linspace(0, 1, num_groups))
            bottom_heights = np.zeros(len(sorted_target_materials))
            raw_labels = {i: [] for i in range(len(sorted_target_materials))}

            for idx, group in enumerate(customer_groups):
                qtys = plot_data[group].values
                bars = ax.bar(x_indexes, qtys, bar_width, bottom=bottom_heights, label=group,
                              color=distinct_colors[idx], edgecolor='#333333', linewidth=0.7)

                for i, qty in enumerate(qtys):
                    if qty > 0:
                        total_for_mat = material_totals.iloc[i]
                        percentage = (qty / total_for_mat) * 100
                        center_y = bottom_heights[i] + (qty / 2)
                        label_text = f"{percentage:.1f}%\n({group})"
                        raw_labels[i].append(
                            {'initial_y': center_y, 'current_y': center_y, 'text': label_text, 'group': group})
                bottom_heights += qtys

            # Vertical Anti-Overlap Relaxer (Locked strictly to the right edge)
            min_vertical_clearance = max(material_totals) * 0.038
            for i in range(len(sorted_target_materials)):
                material_labels = raw_labels[i]
                if not material_labels: continue
                x_pin = x_indexes[i] + (bar_width / 2)
                x_text = x_pin + 0.04

                material_labels.sort(key=lambda x: x['initial_y'])
                for _ in range(5):
                    for idx_l in range(1, len(material_labels)):
                        diff = material_labels[idx_l]['current_y'] - material_labels[idx_l - 1]['current_y']
                        if diff < min_vertical_clearance: material_labels[idx_l]['current_y'] += (
                                    min_vertical_clearance - diff)
                    for idx_l in range(len(material_labels) - 2, -1, -1):
                        diff = material_labels[idx_l + 1]['current_y'] - material_labels[idx_l]['current_y']
                        if diff < min_vertical_clearance: material_labels[idx_l]['current_y'] -= (
                                    min_vertical_clearance - diff)

                for lbl in material_labels:
                    final_y = max(lbl['current_y'], 5)
                    ax.annotate(lbl['text'], xy=(x_pin, lbl['initial_y']), xytext=(x_text, final_y), ha='left',
                                va='center', fontsize=6.5, weight='bold', color='#111111',
                                arrowprops=dict(arrowstyle="-", color="#555555", lw=0.6, shrinkA=0, shrinkB=2,
                                                connectionstyle="arc3,rad=0"))

            for i, total_qty in enumerate(material_totals):
                ax.text(x_indexes[i], total_qty + (max(material_totals) * 0.02), f"{total_qty:,.0f}", ha='center',
                        va='bottom', fontsize=11.5, weight='bold', color='black')

            ax.set_title("Material Volume Share by Customer Group", fontsize=14, weight='bold', pad=25)
            ax.set_ylabel("Quantity (QTY)", fontsize=11, weight='bold')
            ax.set_xlabel("Material Description", fontsize=11, weight='bold', labelpad=15)

            # Restored and safe ticker assignment
            x_tick_labels = [f"{final_df[final_df['Material'] == mat]['Material Description'].iloc[0]}" for mat in
                             sorted_target_materials]
            ax.set_xticks(x_indexes)
            ax.set_xticklabels(x_tick_labels, fontsize=9, weight='bold', rotation=45, ha='right')
            ax.set_xlim(-0.5, len(sorted_target_materials) + 1.1)
            ax.set_ylim(0, max(material_totals) * 1.18)

            # Amplified Legend Design
            ax.legend(title="Customer Group", title_fontproperties={'weight': 'bold', 'size': 11.5},
                      prop={'size': 10.5, 'weight': 'normal'},
                      loc='upper right', labelspacing=1.2, handletextpad=0.8, borderpad=1.0, framealpha=1.0,
                      edgecolor='#333333')

            # --- 4. DISPLAY RUNTIME WEB REVIEWS ---
            st.subheader("📈 Live Visual Analytics Preview")
            st.pyplot(fig)

            # --- 5. STREAM EXCEL EXPORTS OUT OF IN-MEMORY STRING BUFFERS ---
            img_buffer = io.BytesIO()
            plt.savefig(img_buffer, format='png', dpi=130)
            img_buffer.seek(0)
            plt.close()

            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                spaced_df.to_excel(writer, sheet_name='Data Summary', index=False)
                workbook = writer.book
                ws_charts = workbook.create_sheet(title='Visual Charts')
                img = Image(img_buffer)
                ws_charts.add_image(img, 'A1')
            excel_buffer.seek(0)

            # --- 6. RENDER INTERACTIVE INTERFACE DOWNLOAD BUTTONS ---
            st.subheader("📥 Export Final Assets")
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="Download Final Structured Excel Report",
                    data=excel_buffer,
                    file_name="final_narrow_bars_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col2:
                img_buffer.seek(0)
                st.download_button(
                    label="Download Chart Graphic File (PNG)",
                    data=img_buffer,
                    file_name="narrow_bars_stacked_bar_chart.png",
                    mime="image/png"
                )

    except Exception as e:
        st.error(f"An unexpected data layout mismatch occurred: {str(e)}")
