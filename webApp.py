import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image
import io

# 1. Configure the browser tab title and visual layout
st.set_page_config(page_title="Material Volume Analyzer", page_icon="📊", layout="wide")

st.title("📊 Material Volume Share Analyzer & Chart Generator")
st.markdown(
    "Upload your raw data Excel file. The system will automatically aggregate data, generate a collision-free stacked bar chart, and output an engineered Excel summary.")

# 2. Drag-and-Drop File Uploader Component
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
            # --- 3. [CORE DATA RESTRUCTURING LOGIC] ---
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
                    'Material': f"Total for {mat}",
                    'Material Description': '',
                    'Customer Name': '',
                    'Customer Group': '',
                    'QTY': mat_df['Global_QTY'].iloc[0],
                    'Total Customers for Material': ''
                }
                grouped_rows.append(total_qty_row)
                grouped_rows.append({col: '' for col in output_cols})

            spaced_df = pd.DataFrame(grouped_rows)

            plot_data_raw = final_df.groupby(['Material', 'Customer Group'])['QTY'].sum().unstack(fill_value=0)
            plot_data = plot_data_raw.reindex(sorted_target_materials)
            material_totals = plot_data.sum(axis=1)

            # --- 4. [CHART GENERATION ENGINE WITH OPTIMIZATIONS] ---
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
            ax.set_xlabel("Material Summary Description", fontsize=11, weight='bold', labelpad=15)

            # Safe ticker assignment
            x_tick_labels = [f"{final_df[final_df['Material'] == mat]['Material Description'].iloc[0]}" for mat in
                             sorted_target_materials]
            ax.set_xticks(x_indexes)
            ax.set_xticklabels(x_tick_labels, fontsize=9, weight='bold', rotation=45, ha='right')

            # [UPDATED FIX]: 使用 tight_layout 动态优化边距，配合 subplots_adjust 补充底部空间，防截断
            plt.tight_layout()
            plt.subplots_adjust(bottom=0.32)

            ax.set_xlim(-0.5, len(sorted_target_materials) + 1.1)
            ax.set_ylim(0, max(material_totals) * 1.18)

            # Amplified Legend Design
            ax.legend(title="Customer Group", title_fontproperties={'weight': 'bold', 'size': 11.5},
                      prop={'size': 10.5, 'weight': 'normal'},
                      loc='upper right', labelspacing=1.2, handletextpad=0.8, borderpad=1.0, framealpha=1.0,
                      edgecolor='#333333')

            # --- 5. [STREAMLIT DISPLAY & DOWNLOAD WORKFLOW] ---
            # 在网页端渲染图表
            st.header("📊 Generated Visualization")
            st.pyplot(fig)

            # 内存处理：将图表和数据保存为可供用户下载的 Excel 汇总表格
            st.header("💾 Download Summary Report")

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                spaced_df.to_excel(writer, sheet_name='Summary Data', index=False)

                # 将图表作为图片插入到 Excel 的指定位置
                img_buf = io.BytesIO()
                fig.savefig(img_buf, format='png', dpi=150)
                img_buf.seek(0)

                workbook = writer.book
                worksheet = writer.sheets['Summary Data']
                xl_img = Image(img_buf)

                # 在第 H 列，第 2 行开始插入柱状图图片
                worksheet.add_image(xl_img, 'H2')

            processed_data = output.getvalue()

            st.download_button(
                label="📥 Download Engineered Excel Summary",
                data=processed_data,
                file_name="material_volume_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"An unexpected error occurred during execution: {e}")
