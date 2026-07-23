"""
streamlit_app.py

Streamlit front-end that ties stats_processor.py and pptx_generator.py
together into a single report-generation workflow.

Inputs collected from the user:
  - Crosses CSV for Brooklyn FC (BKFC)
  - All Events CSV for Brooklyn FC (BKFC)
  - Crosses CSV for the opponent
  - All Events CSV for the opponent
  - Manual PPDA value for BKFC and the opponent
  - Average position PNG for BKFC
  - Team name (as it should appear in the deck) and brand color for
    each side

Outputs:
  - team_comparison.csv (downloadable)
  - <team>_players.csv / <opponent>_players.csv (downloadable)
  - Final .pptx report (downloadable)
"""

import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

import stats_processor as sp
import pptx_generator as pg


st.set_page_config(page_title="Match Report Generator", layout="wide")

DEFAULT_BKFC_NAME = "Brooklyn FC"
DEFAULT_BKFC_COLOR = "#0B3D91"
DEFAULT_OPPONENT_COLOR = "#F2A900"


def _hex_is_valid(value: str) -> bool:
    h = value.strip().lstrip("#")
    if len(h) != 6:
        return False
    try:
        int(h, 16)
        return True
    except ValueError:
        return False


def main():
    st.title("⚽ Match Report Generator")
    st.caption(
        "Upload StatsBomb CSV exports for both teams, fill in a couple of "
        "manual values, and generate a CSV + PowerPoint report."
    )

    # ------------------------------------------------------------
    # 1. File uploads
    # ------------------------------------------------------------
    st.header("1. Upload CSV files")

    col_bkfc, col_opp = st.columns(2)

    with col_bkfc:
        st.subheader("Brooklyn FC")
        bkfc_events_file = st.file_uploader(
            "All Events CSV — Brooklyn FC", type="csv", key="bkfc_events"
        )
        bkfc_crosses_file = st.file_uploader(
            "Crosses CSV — Brooklyn FC", type="csv", key="bkfc_crosses"
        )

    with col_opp:
        st.subheader("Opponent")
        opp_events_file = st.file_uploader(
            "All Events CSV — Opponent", type="csv", key="opp_events"
        )
        opp_crosses_file = st.file_uploader(
            "Crosses CSV — Opponent", type="csv", key="opp_crosses"
        )

    # ------------------------------------------------------------
    # 2. Manual PPDA input
    # ------------------------------------------------------------
    st.header("2. PPDA (manual entry)")
    col_ppda_1, col_ppda_2 = st.columns(2)
    with col_ppda_1:
        bkfc_ppda = st.number_input("Input BKFC ppda", min_value=0.0, step=0.1, format="%.2f")
    with col_ppda_2:
        opponent_ppda = st.number_input("Input opponent ppda", min_value=0.0, step=0.1, format="%.2f")

    # ------------------------------------------------------------
    # 3. Team names & colors
    # ------------------------------------------------------------
    st.header("3. Team names & colors (as they'll appear in the deck)")
    col_name_1, col_name_2 = st.columns(2)
    with col_name_1:
        bkfc_display_name = st.text_input("BKFC team name", value=DEFAULT_BKFC_NAME)
        bkfc_color = st.color_picker("BKFC color", value=DEFAULT_BKFC_COLOR)
    with col_name_2:
        opponent_display_name = st.text_input("Opponent team name", value="Opponent")
        opponent_color = st.color_picker("Opponent color", value=DEFAULT_OPPONENT_COLOR)

    # ------------------------------------------------------------
    # 4. Average position image
    # ------------------------------------------------------------
    st.header("4. Average position image (BKFC)")
    avg_position_file = st.file_uploader(
        "Average Position PNG — Brooklyn FC", type=["png", "jpg", "jpeg"], key="avg_position"
    )

    # ------------------------------------------------------------
    # 5. Optional report metadata
    # ------------------------------------------------------------
    st.header("5. Report details")
    col_meta_1, col_meta_2 = st.columns(2)
    with col_meta_1:
        kicker = st.text_input("Eyebrow / kicker text", value="STATSBOMB MATCH RECAP")
        match_date = st.text_input("Match date (footer, e.g. 7-18-2026)", value="")
    with col_meta_2:
        subtitle = st.text_input(
            "Subtitle",
            value="STATISTICAL MATCH ANALYSIS  \u2022  PLAYER & TEAM PERFORMANCE BREAKDOWN",
        )
        include_supplementary_table = st.checkbox(
            "Include supplementary KPI table (PPDA + secondary stats)", value=True
        )

    st.divider()

    required_files_present = all([
        bkfc_events_file, bkfc_crosses_file, opp_events_file, opp_crosses_file
    ])

    if not required_files_present:
        st.info("Upload all four CSV files above to enable report generation.")
        return

    if not _hex_is_valid(bkfc_color) or not _hex_is_valid(opponent_color):
        st.error("Team colors must be valid hex colors.")
        return

    if not bkfc_display_name.strip() or not opponent_display_name.strip():
        st.error("Both team names are required.")
        return

    if bkfc_display_name.strip() == opponent_display_name.strip():
        st.error("BKFC and opponent team names must be different.")
        return

    generate = st.button("Generate report", type="primary")

    if not generate:
        return

    # ------------------------------------------------------------
    # 6. Run the stats pipeline
    # ------------------------------------------------------------
    with st.spinner("Crunching the numbers..."):
        try:
            bkfc_events_df, bkfc_crosses_df = sp.load_team_data(bkfc_events_file, bkfc_crosses_file)
            opp_events_df, opp_crosses_df = sp.load_team_data(opp_events_file, opp_crosses_file)
        except Exception as e:
            st.error(f"Failed to process the uploaded CSVs: {e}")
            return

        bkfc_player_stats = sp.calculate_player_stats(bkfc_events_df, bkfc_crosses_df)
        bkfc_team_stats = sp.calculate_team_stats(bkfc_player_stats)
        bkfc_goals = sp.count_goals(bkfc_events_df)

        opp_player_stats = sp.calculate_player_stats(opp_events_df, opp_crosses_df)
        opp_team_stats = sp.calculate_team_stats(opp_player_stats)
        opp_goals = sp.count_goals(opp_events_df)

        comparison_df = sp.build_team_comparison(
            bkfc_team_stats,
            opp_team_stats,
            bkfc_display_name,
            opponent_display_name,
            team_ppda=bkfc_ppda,
            opponent_ppda=opponent_ppda,
        )

        bkfc_players_df = sp.player_stats_to_csv_df(pd.DataFrame(bkfc_player_stats).T)
        opponent_players_df = sp.player_stats_to_csv_df(pd.DataFrame(opp_player_stats).T)

    st.success("Stats processed.")

    st.subheader("Final Score")
    st.markdown(f"**{bkfc_display_name} {bkfc_goals} — {opp_goals} {opponent_display_name}**")

    st.subheader("Team KPI Comparison")
    st.dataframe(comparison_df, use_container_width=True)

    tab_bkfc, tab_opp = st.tabs([f"{bkfc_display_name} players", f"{opponent_display_name} players"])
    with tab_bkfc:
        st.dataframe(bkfc_players_df, use_container_width=True)
    with tab_opp:
        st.dataframe(opponent_players_df, use_container_width=True)

    # ------------------------------------------------------------
    # 7. Write CSVs + build the pptx in a temp workspace
    # ------------------------------------------------------------
    with st.spinner("Building the slide deck..."):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)

            comparison_csv_path = tmp / "team_comparison.csv"
            bkfc_players_csv_path = tmp / f"{bkfc_display_name}_players.csv"
            opponent_players_csv_path = tmp / f"{opponent_display_name}_players.csv"

            comparison_df.to_csv(comparison_csv_path, index=False)
            bkfc_players_df.to_csv(bkfc_players_csv_path, index=False)
            opponent_players_df.to_csv(opponent_players_csv_path, index=False)

            avg_position_path = None
            if avg_position_file is not None:
                avg_position_path = tmp / f"avg_position{Path(avg_position_file.name).suffix}"
                avg_position_path.write_bytes(avg_position_file.getvalue())

            pptx_output_path = tmp / "match_report.pptx"

            try:
                pg.generate_report(
                    comparison_csv=comparison_csv_path,
                    team_name=bkfc_display_name,
                    team_color=bkfc_color,
                    opponent_name=opponent_display_name,
                    opponent_color=opponent_color,
                    output_path=pptx_output_path,
                    team_score=bkfc_goals,
                    opponent_score=opp_goals,
                    date_label=match_date or None,
                    kicker=kicker,
                    subtitle=subtitle,
                    team_player_stats_csv=bkfc_players_csv_path,
                    opponent_player_stats_csv=opponent_players_csv_path,
                    avg_position_image=avg_position_path,
                    include_supplementary_table=include_supplementary_table,
                )
            except Exception as e:
                st.error(f"Failed to build the PowerPoint report: {e}")
                return

            pptx_bytes = pptx_output_path.read_bytes()
            comparison_csv_bytes = comparison_csv_path.read_bytes()
            bkfc_players_csv_bytes = bkfc_players_csv_path.read_bytes()
            opponent_players_csv_bytes = opponent_players_csv_path.read_bytes()

    st.success("Report generated!")

    st.header("6. Downloads")
    col_dl_1, col_dl_2 = st.columns(2)
    with col_dl_1:
        st.download_button(
            "⬇️ Download PowerPoint report",
            data=pptx_bytes,
            file_name="match_report.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary",
        )
    with col_dl_2:
        st.download_button(
            "⬇️ Download team comparison CSV",
            data=comparison_csv_bytes,
            file_name="team_comparison.csv",
            mime="text/csv",
        )

    col_dl_3, col_dl_4 = st.columns(2)
    with col_dl_3:
        st.download_button(
            f"⬇️ Download {bkfc_display_name} player stats CSV",
            data=bkfc_players_csv_bytes,
            file_name=f"{bkfc_display_name}_players.csv",
            mime="text/csv",
        )
    with col_dl_4:
        st.download_button(
            f"⬇️ Download {opponent_display_name} player stats CSV",
            data=opponent_players_csv_bytes,
            file_name=f"{opponent_display_name}_players.csv",
            mime="text/csv",
        )


if __name__ == "__main__":
    main()
