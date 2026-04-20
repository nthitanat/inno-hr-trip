"""
step1_flags.py

Faithful Python rewrite of the VBA BuildTripFlags macro.
Generates trip flags from raw GPS/location data and builds trip chains.
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from math import radians, sin, cos, sqrt, atan2


def haversine(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """
    Calculate the great-circle distance between two points on Earth.

    Args:
        lat1, lon1: Latitude and longitude of first point (degrees)
        lat2, lon2: Latitude and longitude of second point (degrees)

    Returns:
        Distance in metres
    """
    # Handle NaN or invalid values
    if pd.isna(lat1) or pd.isna(lon1) or pd.isna(lat2) or pd.isna(lon2):
        return np.nan

    # Earth's radius in metres
    R = 6371000

    # Convert to radians
    lat1_rad = radians(lat1)
    lon1_rad = radians(lon1)
    lat2_rad = radians(lat2)
    lon2_rad = radians(lon2)

    # Differences
    dlat = lat2_rad - lat1_rad
    dlon = lon2_rad - lon1_rad

    # Haversine formula
    a = sin(dlat / 2) ** 2 + cos(lat1_rad) * cos(lat2_rad) * sin(dlon / 2) ** 2
    c = 2 * atan2(sqrt(a), sqrt(1 - a))

    return R * c


def build_trip_flags(rawdata_df: pd.DataFrame, config: dict) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Build trip flags from raw data and generate trip chains.

    Args:
        rawdata_df: Input DataFrame with raw GPS/location records
        config: Configuration dict with keys:
            - SAME_LOC_M (float, default 30): threshold for "same location" in metres
            - BORDERLINE_M (float, default 50): threshold for "borderline GPS" in metres
            - SHORT_STOP_MIN (float, default 5): threshold for "short stop" in minutes
            - NIGHT_HR_START (int, default 0): start hour of "night" window (0-23)
            - NIGHT_HR_END (int, default 5): end hour of "night" window (0-23)
            - NEXTDAY_MIN (float, default 720): threshold for "next day" gap in minutes

    Returns:
        Tuple of (rawdata_flagged_df, trips_df)
    """

    # Configuration defaults
    SAME_LOC_M = config.get('SAME_LOC_M', 30)
    BORDERLINE_M = config.get('BORDERLINE_M', 50)
    SHORT_STOP_MIN = config.get('SHORT_STOP_MIN', 5)
    NIGHT_HR_START = config.get('NIGHT_HR_START', 0)
    NIGHT_HR_END = config.get('NIGHT_HR_END', 5)
    NEXTDAY_MIN = config.get('NEXTDAY_MIN', 720)

    # Make a copy to avoid modifying the original
    df = rawdata_df.copy()

    # =========================================================================
    # STEP 1: Build Full_Name and parse DateTime
    # =========================================================================

    # Column indices (0-based Python)
    COL_TITLE = 7      # Col H
    COL_FNAME = 8      # Col I
    COL_LNAME = 9      # Col J
    COL_DATETIME = 12  # Col M
    COL_LAT = 17       # Col R
    COL_LNG = 18       # Col S
    COL_GPS_ACC = 19   # Col T
    COL_PERSON_ID = 6  # Col G
    COL_BRANCH = 1     # Col B

    # Build Full_Name: Title + First_Name + Last_Name (space-separated)
    df['Full_Name'] = (
        df.iloc[:, COL_TITLE].fillna('').astype(str) + ' ' +
        df.iloc[:, COL_FNAME].fillna('').astype(str) + ' ' +
        df.iloc[:, COL_LNAME].fillna('').astype(str)
    ).str.strip()
    # Clean up multiple spaces
    df['Full_Name'] = df['Full_Name'].str.replace(r'\s+', ' ', regex=True)

    # Parse DateTime (format: "dd/mm/yyyy hh:mm:ss")
    df['Parsed_DateTime'] = pd.to_datetime(
        df.iloc[:, COL_DATETIME],
        format='%d/%m/%Y %H:%M:%S',
        errors='coerce'
    )

    # =========================================================================
    # STEP 2: Calculate distance and time differences from previous row
    # =========================================================================

    # Group by Person_ID and sort by DateTime within each group
    df_sorted = df.sort_values(
        by=[df.columns[COL_PERSON_ID], 'Parsed_DateTime'],
        na_position='last'
    ).reset_index(drop=True)

    df_sorted['Dist_From_Prev_m'] = np.nan
    df_sorted['TimeDiff_From_Prev_min'] = np.nan

    for person_id in df_sorted.iloc[:, COL_PERSON_ID].unique():
        if pd.isna(person_id):
            continue

        person_mask = df_sorted.iloc[:, COL_PERSON_ID] == person_id
        person_indices = df_sorted.index[person_mask].tolist()

        if len(person_indices) < 2:
            continue

        for i in range(1, len(person_indices)):
            curr_idx = person_indices[i]
            prev_idx = person_indices[i - 1]

            # Calculate distance
            lat_curr = df_sorted.loc[curr_idx, df_sorted.columns[COL_LAT]]
            lng_curr = df_sorted.loc[curr_idx, df_sorted.columns[COL_LNG]]
            lat_prev = df_sorted.loc[prev_idx, df_sorted.columns[COL_LAT]]
            lng_prev = df_sorted.loc[prev_idx, df_sorted.columns[COL_LNG]]

            dist = haversine(lat_prev, lng_prev, lat_curr, lng_curr)
            df_sorted.loc[curr_idx, 'Dist_From_Prev_m'] = dist

            # Calculate time difference
            dt_curr = df_sorted.loc[curr_idx, 'Parsed_DateTime']
            dt_prev = df_sorted.loc[prev_idx, 'Parsed_DateTime']

            if pd.notna(dt_curr) and pd.notna(dt_prev):
                tdiff = (dt_curr - dt_prev).total_seconds() / 60  # Convert to minutes
                df_sorted.loc[curr_idx, 'TimeDiff_From_Prev_min'] = tdiff

    # =========================================================================
    # STEP 3: Create flags
    # =========================================================================

    # FLAG_EXACT_DUPE: dist < 0.01m AND tdiff == 0
    df_sorted['FLAG_EXACT_DUPE'] = (
        (df_sorted['Dist_From_Prev_m'] < 0.01) &
        (df_sorted['TimeDiff_From_Prev_min'] == 0)
    )

    # FLAG_MIDNIGHT_PING: hour is 0-5 (inclusive)
    df_sorted['FLAG_MIDNIGHT_PING'] = (
        df_sorted['Parsed_DateTime'].dt.hour.between(NIGHT_HR_START, NIGHT_HR_END)
    )

    # FLAG_SAME_LOC_under30m: dist < 30m
    df_sorted['FLAG_SAME_LOC_under30m'] = (
        df_sorted['Dist_From_Prev_m'] < SAME_LOC_M
    )

    # FLAG_SAME_LOC_SHORT_under5min: dist < 30m AND tdiff < 5min
    df_sorted['FLAG_SAME_LOC_SHORT_under5min'] = (
        (df_sorted['Dist_From_Prev_m'] < SAME_LOC_M) &
        (df_sorted['TimeDiff_From_Prev_min'] < SHORT_STOP_MIN)
    )

    # FLAG_SAME_LOC_LONG_over5min: dist < 30m AND tdiff >= 5min
    df_sorted['FLAG_SAME_LOC_LONG_over5min'] = (
        (df_sorted['Dist_From_Prev_m'] < SAME_LOC_M) &
        (df_sorted['TimeDiff_From_Prev_min'] >= SHORT_STOP_MIN)
    )

    # FLAG_GPS_BORDERLINE_30to50m: 30 <= dist < 50m
    df_sorted['FLAG_GPS_BORDERLINE_30to50m'] = (
        (df_sorted['Dist_From_Prev_m'] >= SAME_LOC_M) &
        (df_sorted['Dist_From_Prev_m'] < BORDERLINE_M)
    )

    # FLAG_CROSS_DAY_BOUNDARY: different calendar date from previous row
    df_sorted['FLAG_CROSS_DAY_BOUNDARY'] = False
    for person_id in df_sorted.iloc[:, COL_PERSON_ID].unique():
        if pd.isna(person_id):
            continue

        person_mask = df_sorted.iloc[:, COL_PERSON_ID] == person_id
        person_indices = df_sorted.index[person_mask].tolist()

        for i in range(1, len(person_indices)):
            curr_idx = person_indices[i]
            prev_idx = person_indices[i - 1]

            dt_curr = df_sorted.loc[curr_idx, 'Parsed_DateTime']
            dt_prev = df_sorted.loc[prev_idx, 'Parsed_DateTime']

            if pd.notna(dt_curr) and pd.notna(dt_prev):
                if dt_curr.date() != dt_prev.date():
                    df_sorted.loc[curr_idx, 'FLAG_CROSS_DAY_BOUNDARY'] = True

    # FLAG_OPEN_END_LAST_ROW: last row of each person
    df_sorted['FLAG_OPEN_END_LAST_ROW'] = False
    for person_id in df_sorted.iloc[:, COL_PERSON_ID].unique():
        if pd.isna(person_id):
            continue

        person_mask = df_sorted.iloc[:, COL_PERSON_ID] == person_id
        person_indices = df_sorted.index[person_mask].tolist()

        if len(person_indices) > 0:
            last_idx = person_indices[-1]
            df_sorted.loc[last_idx, 'FLAG_OPEN_END_LAST_ROW'] = True

    # FLAG_SINGLE_RECORD_PERSON: only 1 row for this person
    person_counts = df_sorted.iloc[:, COL_PERSON_ID].value_counts()
    df_sorted['FLAG_SINGLE_RECORD_PERSON'] = (
        df_sorted.iloc[:, COL_PERSON_ID].map(person_counts) == 1
    )

    # FLAG_CANNOT_COMBINE: EXACT_DUPE OR MIDNIGHT_PING OR SAME_LOC_SHORT OR SINGLE_RECORD
    df_sorted['FLAG_CANNOT_COMBINE'] = (
        df_sorted['FLAG_EXACT_DUPE'] |
        df_sorted['FLAG_MIDNIGHT_PING'] |
        df_sorted['FLAG_SAME_LOC_SHORT_under5min'] |
        df_sorted['FLAG_SINGLE_RECORD_PERSON']
    )

    # =========================================================================
    # STEP 4: Convert all flag columns to VBA-style "TRUE"/"FALSE" strings
    # =========================================================================

    flag_cols = [
        'FLAG_EXACT_DUPE',
        'FLAG_MIDNIGHT_PING',
        'FLAG_SAME_LOC_under30m',
        'FLAG_SAME_LOC_SHORT_under5min',
        'FLAG_SAME_LOC_LONG_over5min',
        'FLAG_GPS_BORDERLINE_30to50m',
        'FLAG_CROSS_DAY_BOUNDARY',
        'FLAG_OPEN_END_LAST_ROW',
        'FLAG_SINGLE_RECORD_PERSON',
        'FLAG_CANNOT_COMBINE'
    ]

    for col in flag_cols:
        df_sorted[col] = df_sorted[col].apply(lambda x: 1 if x else 0)

    # =========================================================================
    # STEP 5: Build trips by chaining consecutive non-skip rows
    # =========================================================================

    trips_list = []

    for person_id in df_sorted.iloc[:, COL_PERSON_ID].unique():
        if pd.isna(person_id):
            continue

        person_mask = df_sorted.iloc[:, COL_PERSON_ID] == person_id
        person_df = df_sorted[person_mask].reset_index(drop=True)

        # Filter out rows where FLAG_CANNOT_COMBINE is TRUE
        valid_rows = person_df[person_df['FLAG_CANNOT_COMBINE'] == 0].reset_index(drop=True)

        if len(valid_rows) == 0:
            continue

        # Build trips by chaining: each valid row becomes origin for next valid row
        from_row = None
        from_idx = None

        for idx in range(len(valid_rows)):
            current_row = valid_rows.iloc[idx]

            if from_row is None:
                # Start a new trip
                from_row = current_row.copy()
                from_idx = idx
            else:
                # We have a pending "from" row, so create a trip
                to_row = current_row

                # Extract trip data
                person_id_trip = from_row.iloc[COL_PERSON_ID]
                full_name_trip = from_row['Full_Name']
                branch_trip = from_row.iloc[COL_BRANCH]

                trip_start = from_row['Parsed_DateTime']
                trip_end = to_row['Parsed_DateTime']

                # Duration in minutes
                if pd.notna(trip_start) and pd.notna(trip_end):
                    duration_min = (trip_end - trip_start).total_seconds() / 60
                else:
                    duration_min = np.nan

                origin = from_row['Location_Name'] if 'Location_Name' in from_row.index else from_row.iloc[20]
                destination = to_row['Location_Name'] if 'Location_Name' in to_row.index else to_row.iloc[20]
                distance_m = to_row['Dist_From_Prev_m']

                origin_lat = from_row.iloc[COL_LAT]
                origin_lng = from_row.iloc[COL_LNG]
                dest_lat = to_row.iloc[COL_LAT]
                dest_lng = to_row.iloc[COL_LNG]

                # Determine gap flags
                flag_same_loc_long = to_row['FLAG_SAME_LOC_LONG_over5min']
                flag_gps_borderline = to_row['FLAG_GPS_BORDERLINE_30to50m']

                # Calculate gap between from and to
                if pd.notna(trip_start) and pd.notna(trip_end):
                    gap_min = (trip_end - trip_start).total_seconds() / 60
                    cross_day = trip_start.date() != trip_end.date()
                    flag_cross_day_overnight = 1 if (0 < gap_min < NEXTDAY_MIN) and cross_day else 0
                    flag_cross_day_nextday = 1 if gap_min >= NEXTDAY_MIN and cross_day else 0
                else:
                    flag_cross_day_overnight = 0
                    flag_cross_day_nextday = 0

                flag_open_end = 0  # Not an open-end trip (we have both start and end)

                # Copy over flags from the to_row (distance-based flags belong to the arriving row)
                flag_exact_dupe = to_row['FLAG_EXACT_DUPE']
                flag_midnight_ping = from_row['FLAG_MIDNIGHT_PING']
                flag_same_loc_under30m = to_row['FLAG_SAME_LOC_under30m']
                flag_same_loc_short_under5min = to_row['FLAG_SAME_LOC_SHORT_under5min']
                flag_cross_day_boundary = to_row['FLAG_CROSS_DAY_BOUNDARY']
                flag_single_record_person = from_row['FLAG_SINGLE_RECORD_PERSON']
                flag_cannot_combine = to_row['FLAG_CANNOT_COMBINE']

                trip_dict = {
                    'Person_ID': person_id_trip,
                    'Full_Name': full_name_trip,
                    'Branch': branch_trip,
                    'Trip_Start': trip_start,
                    'Trip_End': trip_end,
                    'Duration_min': duration_min,
                    'Origin': origin,
                    'Destination': destination,
                    'Distance_m': distance_m,
                    'Origin_Lat': origin_lat,
                    'Origin_Lng': origin_lng,
                    'Dest_Lat': dest_lat,
                    'Dest_Lng': dest_lng,
                    'FLAG_SAME_LOC_LONG': flag_same_loc_long,
                    'FLAG_GPS_BORDERLINE': flag_gps_borderline,
                    'FLAG_CROSS_DAY_OVERNIGHT': flag_cross_day_overnight,
                    'FLAG_CROSS_DAY_NEXTDAY': flag_cross_day_nextday,
                    'FLAG_OPEN_END': flag_open_end,
                    'FLAG_EXACT_DUPE': flag_exact_dupe,
                    'FLAG_MIDNIGHT_PING': flag_midnight_ping,
                    'FLAG_SAME_LOC_under30m': flag_same_loc_under30m,
                    'FLAG_SAME_LOC_SHORT_under5min': flag_same_loc_short_under5min,
                    'FLAG_CROSS_DAY_BOUNDARY': flag_cross_day_boundary,
                    'FLAG_SINGLE_RECORD_PERSON': flag_single_record_person,
                    'FLAG_CANNOT_COMBINE': flag_cannot_combine,
                }

                trips_list.append(trip_dict)

                # Chain: the destination becomes the next origin
                from_row = to_row.copy()
                from_idx = idx

        # Handle open-end trip (last row has no pairing)
        if from_row is not None:
            origin = from_row['Full_Name']
            trip_start = from_row['Parsed_DateTime']

            trip_dict = {
                'Person_ID': from_row.iloc[COL_PERSON_ID],
                'Full_Name': from_row['Full_Name'],
                'Branch': from_row.iloc[COL_BRANCH],
                'Trip_Start': trip_start,
                'Trip_End': None,
                'Duration_min': np.nan,
                'Origin': from_row['Location_Name'] if 'Location_Name' in from_row.index else from_row.iloc[20],
                'Destination': None,
                'Distance_m': np.nan,
                'Origin_Lat': from_row.iloc[COL_LAT],
                'Origin_Lng': from_row.iloc[COL_LNG],
                'Dest_Lat': np.nan,
                'Dest_Lng': np.nan,
                'FLAG_SAME_LOC_LONG': 0,
                'FLAG_GPS_BORDERLINE': 0,
                'FLAG_CROSS_DAY_OVERNIGHT': 0,
                'FLAG_CROSS_DAY_NEXTDAY': 0,
                'FLAG_OPEN_END': 1,
                'FLAG_EXACT_DUPE': from_row['FLAG_EXACT_DUPE'],
                'FLAG_MIDNIGHT_PING': from_row['FLAG_MIDNIGHT_PING'],
                'FLAG_SAME_LOC_under30m': from_row['FLAG_SAME_LOC_under30m'],
                'FLAG_SAME_LOC_SHORT_under5min': from_row['FLAG_SAME_LOC_SHORT_under5min'],
                'FLAG_CROSS_DAY_BOUNDARY': from_row['FLAG_CROSS_DAY_BOUNDARY'],
                'FLAG_SINGLE_RECORD_PERSON': from_row['FLAG_SINGLE_RECORD_PERSON'],
                'FLAG_CANNOT_COMBINE': from_row['FLAG_CANNOT_COMBINE'],
            }

            trips_list.append(trip_dict)

    trips_df = pd.DataFrame(trips_list)

    # =========================================================================
    # STEP 6: Prepare rawdata_flagged output (return to original row order)
    # =========================================================================

    # Re-apply the original row order by merging back to the input df structure
    # Build a flagged version with all original columns + new flag/derived columns

    output_cols = list(df_sorted.columns)
    rawdata_flagged_df = df_sorted[output_cols].copy()

    return rawdata_flagged_df, trips_df
