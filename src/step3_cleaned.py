"""
Step 3: Build Trip Cleaned Dataset

This module replicates the BuildTripCleaned VBA macro.
It filters out flagged trips and generates summary statistics.

Input:  trip_behavior DataFrame with behavior classifications
Output: trip_cleaned DataFrame (clean trips only) and summary statistics dict
"""

import pandas as pd
from typing import Tuple


def build_trip_cleaned(trip_behavior_df: pd.DataFrame, config: dict) -> Tuple[pd.DataFrame, dict]:
    """
    Filter clean trips and generate summary statistics.

    Removes all trips where BEH_ANY_FLAG is "TRUE" (problematic trips).
    Returns only the data columns (Person_ID through Dest_Lng) for clean trips.
    Also computes comprehensive summary statistics.

    Args:
        trip_behavior_df: DataFrame from step2 with behavior classifications
        config: Configuration dictionary (reserved for future use)

    Returns:
        tuple containing:
            - trip_cleaned_df: DataFrame with only clean trips and data columns A-M:
                Person_ID, Full_Name, Branch, Trip_Start, Trip_End, Duration_min,
                Origin, Destination, Distance_m, Origin_Lat, Origin_Lng,
                Dest_Lat, Dest_Lng
            - summary_dict: Dictionary with statistics:
                - total_before: total rows before cleaning
                - removed_forget: count of BEH_Forget_Checkout==TRUE rows
                - removed_newday: count of BEH_New_Day_Work==TRUE rows
                - total_removed: total rows removed
                - total_kept: clean rows remaining
                - total_distance_m: sum of Distance_m for clean rows
                - total_distance_km: total_distance_m / 1000
                - avg_distance_m: average Distance_m for clean rows (None if no clean rows)
                - avg_duration_min: average Duration_min for clean rows (None if no clean rows)
                - per_person: list of per-person summaries, each with:
                    - person_id
                    - full_name
                    - branch
                    - trip_count
                    - total_dist_m
                    - total_dist_km
                    - avg_dist_m
    """

    # Define data columns A-M (13 columns: Person_ID through Dest_Lng)
    data_columns = [
        'Person_ID',
        'Full_Name',
        'Branch',
        'Trip_Start',
        'Trip_End',
        'Duration_min',
        'Origin',
        'Destination',
        'Distance_m',
        'Origin_Lat',
        'Origin_Lng',
        'Dest_Lat',
        'Dest_Lng'
    ]

    # Record statistics before cleaning
    total_before = len(trip_behavior_df)

    # Count trips to be removed by each flag
    removed_forget = (trip_behavior_df['BEH_Forget_Checkout'] == 1).sum()
    removed_newday = (trip_behavior_df['BEH_New_Day_Work'] == 1).sum()
    removed_same_loc = (trip_behavior_df['BEH_Same_Loc'] == 1).sum()
    removed_gps_borderline = (trip_behavior_df['BEH_GPS_Borderline'] == 1).sum()

    # Filter to keep only rows where BEH_ANY_FLAG is NOT "TRUE"
    # This means keeping rows where BEH_ANY_FLAG is empty string or missing
    clean_mask = trip_behavior_df['BEH_ANY_FLAG'] != 1
    trip_cleaned_df = trip_behavior_df[clean_mask][data_columns].copy()

    # Calculate removal statistics
    total_removed = total_before - len(trip_cleaned_df)
    total_kept = len(trip_cleaned_df)

    # Initialize summary dictionary
    summary_dict = {
        'total_before': total_before,
        'removed_forget': removed_forget,
        'removed_newday': removed_newday,
        'removed_same_loc': removed_same_loc,
        'removed_gps_borderline': removed_gps_borderline,
        'total_removed': total_removed,
        'total_kept': total_kept,
    }

    # Calculate aggregate distance and duration statistics
    if len(trip_cleaned_df) > 0:
        # Total and average distances
        total_distance_m = trip_cleaned_df['Distance_m'].sum()
        total_distance_km = total_distance_m / 1000.0
        avg_distance_m = trip_cleaned_df['Distance_m'].mean()
        avg_duration_min = trip_cleaned_df['Duration_min'].mean()

        summary_dict['total_distance_m'] = total_distance_m
        summary_dict['total_distance_km'] = total_distance_km
        summary_dict['avg_distance_m'] = avg_distance_m
        summary_dict['avg_duration_min'] = avg_duration_min

        # Per-person statistics
        per_person_list = []
        grouped = trip_cleaned_df.groupby('Person_ID', as_index=False)

        for person_id, person_group in grouped:
            full_name = person_group['Full_Name'].iloc[0]
            branch = person_group['Branch'].iloc[0]
            trip_count = len(person_group)
            total_dist_m = person_group['Distance_m'].sum()
            total_dist_km = total_dist_m / 1000.0
            avg_dist_m = person_group['Distance_m'].mean()

            per_person_list.append({
                'person_id': person_id,
                'full_name': full_name,
                'branch': branch,
                'trip_count': trip_count,
                'total_dist_m': total_dist_m,
                'total_dist_km': total_dist_km,
                'avg_dist_m': avg_dist_m
            })

        summary_dict['per_person'] = per_person_list

    else:
        # No clean trips
        summary_dict['total_distance_m'] = 0
        summary_dict['total_distance_km'] = 0.0
        summary_dict['avg_distance_m'] = None
        summary_dict['avg_duration_min'] = None
        summary_dict['per_person'] = []

    return trip_cleaned_df, summary_dict
