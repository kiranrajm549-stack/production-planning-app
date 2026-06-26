import pandas as pd
import streamlit as st
from planning import run_planning

st.set_page_config(page_title="Production Planning", layout="wide")
st.title("Production Planning Dashboard")

uploaded_file = st.file_uploader("Upload planning input file", type=["xlsx"])

target_utilisation_pct = st.slider("Target utilisation %", 50, 100, 85)
running_thickness_l1 = st.selectbox(
    "Running thickness - Line 1",
    [None, "30MM", "40MM", "50MM", "PUF"],
    index=1,
)
running_thickness_l2 = st.selectbox(
    "Running thickness - Line 2",
    [None, "30MM", "40MM", "50MM", "PUF"],
    index=2,
)

if uploaded_file and st.button("Generate Plan"):
    input_bytes = uploaded_file.read()

    output_bytes, stats = run_planning(
        input_bytes=input_bytes,
        target_utilisation_pct=target_utilisation_pct,
        running_thickness_l1=running_thickness_l1,
        running_thickness_l2=running_thickness_l2,
    )

    st.success("Plan generated successfully.")

    st.subheader("Overall plan")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("L1 Runtime (h)", f"{stats['total_runtime_l1']:.2f}")
    c2.metric("L2 Runtime (h)", f"{stats['total_runtime_l2']:.2f}")
    c3.metric("L1 Utilisation %", f"{stats['util_l1_pct']:.1f}")
    c4.metric("L2 Utilisation %", f"{stats['util_l2_pct']:.1f}")

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("L1 SQM", f"{stats['total_sqm_l1']:.2f}")
    c6.metric("L2 SQM", f"{stats['total_sqm_l2']:.2f}")
    c7.metric("L1 Changeovers", f"{stats['thickness_changeovers_l1']}")
    c8.metric("L2 Changeovers", f"{stats['thickness_changeovers_l2']}")

    st.subheader("Day-wise plan")

    daywise_df = pd.DataFrame([
        {
            "Day": "Day 1",
            "L1 Runtime (h)": round(stats["day_stats_l1"]["day1_runtime"], 2),
            "L1 SQM": round(stats["day_stats_l1"]["day1_sqm"], 2),
            "L1 Rows": stats["day_stats_l1"]["day1_rows"],
            "L2 Runtime (h)": round(stats["day_stats_l2"]["day1_runtime"], 2),
            "L2 SQM": round(stats["day_stats_l2"]["day1_sqm"], 2),
            "L2 Rows": stats["day_stats_l2"]["day1_rows"],
        },
        {
            "Day": "Day 2",
            "L1 Runtime (h)": round(stats["day_stats_l1"]["day2_runtime"], 2),
            "L1 SQM": round(stats["day_stats_l1"]["day2_sqm"], 2),
            "L1 Rows": stats["day_stats_l1"]["day2_rows"],
            "L2 Runtime (h)": round(stats["day_stats_l2"]["day2_runtime"], 2),
            "L2 SQM": round(stats["day_stats_l2"]["day2_sqm"], 2),
            "L2 Rows": stats["day_stats_l2"]["day2_rows"],
        },
        {
            "Day": "Day 3",
            "L1 Runtime (h)": round(stats["day_stats_l1"]["day3_runtime"], 2),
            "L1 SQM": round(stats["day_stats_l1"]["day3_sqm"], 2),
            "L1 Rows": stats["day_stats_l1"]["day3_rows"],
            "L2 Runtime (h)": round(stats["day_stats_l2"]["day3_runtime"], 2),
            "L2 SQM": round(stats["day_stats_l2"]["day3_sqm"], 2),
            "L2 Rows": stats["day_stats_l2"]["day3_rows"],
        },
    ])

    st.dataframe(daywise_df, use_container_width=True)

    st.subheader("Line-wise day cards")
    for i, day_label in enumerate(["Day 1", "Day 2", "Day 3"], start=1):
        key = f"day{i}"
        col1, col2 = st.columns(2)

        with col1:
            st.markdown(f"**{day_label} - Line 1**")
            st.write(f"Runtime: {stats['day_stats_l1'][f'{key}_runtime']:.2f} h")
            st.write(f"SQM: {stats['day_stats_l1'][f'{key}_sqm']:.2f}")
            st.write(f"Rows: {stats['day_stats_l1'][f'{key}_rows']}")

        with col2:
            st.markdown(f"**{day_label} - Line 2**")
            st.write(f"Runtime: {stats['day_stats_l2'][f'{key}_runtime']:.2f} h")
            st.write(f"SQM: {stats['day_stats_l2'][f'{key}_sqm']:.2f}")
            st.write(f"Rows: {stats['day_stats_l2'][f'{key}_rows']}")

    st.download_button(
        label="Download Planned File",
        data=output_bytes,
        file_name=f"PLANNING-OUTPUT-{target_utilisation_pct}pc-L1-{running_thickness_l1}-L2-{running_thickness_l2}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
