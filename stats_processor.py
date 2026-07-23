"""
stats_processor.py

Processes StatsBomb CSV exports (all-events + crosses) and calculates
player- and team-level stats & KPIs for a single team.

Also includes a helper (build_team_comparison) for assembling a two-team
comparison table (used by the Streamlit app / pptx generator) that folds
in a manually-entered PPDA value for each side, since PPDA is not
derivable from the standard StatsBomb CSV export.
"""

import pandas as pd
from collections import defaultdict


# ============================================================
# CONSTANTS
# ============================================================

TOUCH_EVENTS = [
    "ball_receipt",
    "pass",
    "carry",
    "dribble",
    "shot",
    "clearance"
]

PENALTY_ENTRY_EVENTS = [
    "carry",
    "dribble"
]

# The 12 "headline" metrics used for the Team Totals slide, the
# per-metric player-breakdown slides, and the Full-Time summary slide.
# Order matches the reference deck design.
MAIN_METRICS = [
    ("% Pass Forward", "forward_pass_percentage"),
    ("Completed Passes to A 1/3", "passes_to_attacking_third"),
    ("PA Entry \u2013 passes", "passes_to_penalty_area"),
    ("Open play crosses", "open_play_crosses"),
    ("Successful open play crosses", "successful_open_play_crosses"),
    ("A1/2 Recoveries", "attacking_half_recoveries"),
    ("PA Entry \u2013 dribble", "penalty_entries"),
    ("RZ Shots", "red_zone_shots"),
    ("Attempted RZ Passes", "red_zone_passes_attempted"),
    ("Completed RZ Passes", "red_zone_passes_completed"),
    ("RZ touches", "red_zone_touches"),
    ("A3 Touches", "touches_attacking_third"),
]

# Extra metrics that don't get their own slide but are useful in the
# supplementary KPI table (along with the manually-entered PPDA).
SUPPLEMENTARY_METRICS = [
    ("Total Passes", "total_passes"),
    ("Forward Passes (Outside Box)", "forward_passes_outside_box"),
    ("Possession Lost", "possession_lost"),
    ("Possession Lost (Defensive Half)", "possession_lost_defensive_half"),
]

# Ordered list of (Display Label, team_stats key) used to build the
# full team comparison table. "PPDA" is handled separately since it is
# a manually-entered value, not something derived from the CSVs.
COMPARISON_METRICS = MAIN_METRICS + SUPPLEMENTARY_METRICS


# ============================================================
# DATA LOADING
# ============================================================

def load_team_data(events_file, crosses_file):
    """
    Load StatsBomb CSV exports.

    Parameters:
        events_file: All Events CSV path (or file-like object)
        crosses_file: Crosses CSV path (or file-like object)

    Returns:
        events dataframe
        crosses dataframe
    """

    events = pd.read_csv(events_file)
    crosses = pd.read_csv(crosses_file)

    events = clean_booleans(events)
    crosses = clean_booleans(crosses)

    return events, crosses


# ============================================================
# BOOLEAN CLEANING
# ============================================================

def clean_booleans(df):

    boolean_cols = [
        "Forward Pass",
        "In Attacking Half",
        "In Defensive Half",
        "In Attacking Third",
        "Into Attacking Third",
        "Into Opposition Penalty Box",
        "Open Play"
    ]

    for col in boolean_cols:

        if col in df.columns:

            df[col] = (
                df[col]
                .fillna(False)
                .astype(str)
                .str.lower()
                .eq("true")
            )

    return df


# ============================================================
# FIELD DEFINITIONS
# ============================================================

def is_penalty_area(x, y):

    return (
        pd.notna(x)
        and pd.notna(y)
        and x > 102
        and 18 < y < 62
    )


def is_red_zone(x, y):

    return (
        pd.notna(x)
        and pd.notna(y)
        and 100 < x < 110
        and 20 < y < 60
    )


# ============================================================
# PLAYER STAT CALCULATION
# ============================================================

def calculate_player_stats(events, crosses):

    stats = defaultdict(lambda: {

        "forward_passes": 0,
        "total_passes": 0,
        "forward_pass_percentage": 0,

        "forward_passes_outside_box": 0,

        "attacking_half_recoveries": 0,

        "touches_attacking_third": 0,

        "passes_to_attacking_third": 0,

        "passes_to_penalty_area": 0,

        "penalty_entries": 0,

        "red_zone_touches": 0,

        "red_zone_shots": 0,

        "failed_passes": 0,

        "dispossessed": 0,

        "possession_lost": 0,

        "failed_passes_defensive_half": 0,

        "dispossessed_defensive_half": 0,

        "possession_lost_defensive_half": 0,

        "open_play_crosses": 0,

        "successful_open_play_crosses": 0,

        "red_zone_passes_attempted": 0,

        "red_zone_passes_completed": 0

    })

    # ----------------------------
    # PASSING
    # ----------------------------

    passes = events[
        events["Event Type"] == "pass"
    ]

    completed_passes = passes[
        passes["Outcome"] == "complete"
    ]

    for player, count in passes.groupby("Player").size().items():
        stats[player]["total_passes"] = count

    forward_passes = completed_passes[
        completed_passes["Forward Pass"]
    ]

    for player, count in forward_passes.groupby("Player").size().items():
        stats[player]["forward_passes"] = count

    for _, row in forward_passes.iterrows():

        if not is_penalty_area(
            row["Start X"],
            row["Start Y"]
        ):

            stats[row["Player"]]["forward_passes_outside_box"] += 1

    # ----------------------------
    # RECOVERIES
    # ----------------------------

    recoveries = events[
        (events["Event Type"] == "ball_recovery")
        &
        (events["In Attacking Half"])
    ]

    for player, count in recoveries.groupby("Player").size().items():
        stats[player]["attacking_half_recoveries"] = count

    # ----------------------------
    # TOUCHES
    # ----------------------------

    touches = events[
        (events["Event Type"].isin(TOUCH_EVENTS))
        &
        (events["In Attacking Third"])
    ]

    for player, count in touches.groupby("Player").size().items():
        stats[player]["touches_attacking_third"] = count

    # ----------------------------
    # PROGRESSION
    # ----------------------------

    final_third = completed_passes[
        completed_passes["Into Attacking Third"]
    ]

    for player, count in final_third.groupby("Player").size().items():
        stats[player]["passes_to_attacking_third"] = count

    box_passes = completed_passes[
        completed_passes["Into Opposition Penalty Box"]
    ]

    for player, count in box_passes.groupby("Player").size().items():
        stats[player]["passes_to_penalty_area"] = count

    entries = events[
        (events["Event Type"].isin(PENALTY_ENTRY_EVENTS))
        &
        (events["Into Opposition Penalty Box"])
    ]

    for player, count in entries.groupby("Player").size().items():
        stats[player]["penalty_entries"] = count

    # ----------------------------
    # RED ZONE
    # ----------------------------

    for _, row in events[
        events["Event Type"].isin(TOUCH_EVENTS)
    ].iterrows():

        if is_red_zone(
            row["Start X"],
            row["Start Y"]
        ):

            stats[row["Player"]]["red_zone_touches"] += 1

    for _, row in events[
        events["Event Type"] == "shot"
    ].iterrows():

        if is_red_zone(
            row["Start X"],
            row["Start Y"]
        ):

            stats[row["Player"]]["red_zone_shots"] += 1

    for _, row in passes.iterrows():

        if is_red_zone(
            row["Start X"],
            row["Start Y"]
        ):

            stats[row["Player"]]["red_zone_passes_attempted"] += 1

    for _, row in completed_passes.iterrows():

        if is_red_zone(
            row["Start X"],
            row["Start Y"]
        ):

            stats[row["Player"]]["red_zone_passes_completed"] += 1

    # ----------------------------
    # POSSESSION LOST
    # ----------------------------

    failed_passes = passes[
        passes["Outcome"] != "complete"
    ]

    for player, count in failed_passes.groupby("Player").size().items():
        stats[player]["failed_passes"] = count

    dispossessed = events[
        events["Event Type"] == "dispossessed"
    ]

    for player, count in dispossessed.groupby("Player").size().items():
        stats[player]["dispossessed"] = count

    failed_def = failed_passes[
        failed_passes["In Defensive Half"]
    ]

    for player, count in failed_def.groupby("Player").size().items():
        stats[player]["failed_passes_defensive_half"] = count

    dispossessed_def = dispossessed[
        dispossessed["In Defensive Half"]
    ]

    for player, count in dispossessed_def.groupby("Player").size().items():
        stats[player]["dispossessed_defensive_half"] = count

    # ----------------------------
    # CROSSES
    # ----------------------------

    crosses_open = crosses[
        crosses["Open Play"]
    ]

    for player, count in crosses_open.groupby("Player").size().items():
        stats[player]["open_play_crosses"] = count

    crosses_success = crosses[
        (crosses["Open Play"])
        &
        (crosses["Outcome"] == "complete")
    ]

    for player, count in crosses_success.groupby("Player").size().items():
        stats[player]["successful_open_play_crosses"] = count

    # ----------------------------
    # FINAL CALCULATIONS
    # ----------------------------

    for player in stats:

        total = stats[player]["total_passes"]

        if total:
            stats[player]["forward_pass_percentage"] = (
                stats[player]["forward_passes"]
                /
                total
                *
                100
            )

        stats[player]["possession_lost"] = (
            stats[player]["failed_passes"]
            +
            stats[player]["dispossessed"]
        )

        stats[player]["possession_lost_defensive_half"] = (
            stats[player]["failed_passes_defensive_half"]
            +
            stats[player]["dispossessed_defensive_half"]
        )

    return dict(stats)


# ============================================================
# GOALS (for the title-slide scoreline)
# ============================================================

def count_goals(events):
    """Count goals scored by a team from its All Events export."""
    shots = events[events["Event Type"] == "shot"]
    return int((shots["Outcome"] == "goal").sum())


# ============================================================
# TEAM CALCULATION
# ============================================================

def calculate_team_stats(player_stats):

    if not player_stats:
        # No player events found - return a zeroed-out stat block so
        # downstream code (comparison tables, pptx) doesn't crash.
        return {key: 0 for _, key in COMPARISON_METRICS}

    df = pd.DataFrame(player_stats).T

    team_stats = {}

    for column in df.columns:

        if column != "forward_pass_percentage":

            team_stats[column] = df[column].sum()

    if team_stats["total_passes"]:

        team_stats["forward_pass_percentage"] = (
            team_stats["forward_passes"]
            /
            team_stats["total_passes"]
            *
            100
        )

    else:

        team_stats["forward_pass_percentage"] = 0

    return team_stats


# ============================================================
# MAIN PIPELINE FUNCTION
# ============================================================

def analyze_team(events_file, crosses_file):
    """
    Complete StatsBomb processing pipeline.

    Returns:
        {
            "player_stats": dataframe,
            "team_stats": dictionary
        }
    """

    events, crosses = load_team_data(
        events_file,
        crosses_file
    )

    player_stats = calculate_player_stats(
        events,
        crosses
    )

    team_stats = calculate_team_stats(
        player_stats
    )

    return {

        "player_stats": pd.DataFrame(player_stats).T,

        "team_stats": team_stats

    }


# ============================================================
# TWO-TEAM COMPARISON TABLE (for the pptx report)
# ============================================================

def build_team_comparison(
    team_stats,
    opponent_stats,
    team_name,
    opponent_name,
    team_ppda=None,
    opponent_ppda=None,
    metrics=None,
    include_ppda=True,
):
    """
    Build a tidy comparison dataframe with one row per KPI and one
    column per team, ready to be written to CSV and consumed by the
    pptx generator.

    PPDA is passed in manually (it is not present in the StatsBomb
    CSV export) and, if included, is always the first row of the
    table. Pass metrics=stats_processor.MAIN_METRICS (etc.) to build a
    table scoped to a subset of KPIs.
    """

    if metrics is None:
        metrics = COMPARISON_METRICS

    rows = []
    if include_ppda:
        rows.append({
            "Metric": "PPDA",
            team_name: team_ppda,
            opponent_name: opponent_ppda,
        })

    for label, key in metrics:
        rows.append({
            "Metric": label,
            team_name: round(float(team_stats.get(key, 0)), 2),
            opponent_name: round(float(opponent_stats.get(key, 0)), 2),
        })

    return pd.DataFrame(rows)


def player_stats_to_csv_df(player_stats_df):
    """
    Normalize a player_stats dataframe (indexed by player name) into a
    flat dataframe with 'Player' as a regular column, suitable for
    writing to CSV.
    """

    df = player_stats_df.copy()
    df.index.name = "Player"
    return df.reset_index()
