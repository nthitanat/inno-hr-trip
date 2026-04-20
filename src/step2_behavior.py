"""
Step 2: Build Trip Behavior Classification

This module replicates the BuildTripBehavior VBA macro.
It classifies trips based on flag conditions and appends behavior columns.

Input:  trips DataFrame with flag columns
Output: trip_behavior DataFrame with behavior classifications
"""

import pandas as pd


def build_trip_behavior(trips_df: pd.DataFrame, config: dict) -> pd.DataFrame:
    """
    Classify trip behaviors based on flag conditions.

    Takes the input trips DataFrame and adds behavior classification columns.
    Behavior flags identify problematic trips that should be filtered in step 3.

    Args:
        trips_df: DataFrame with columns including:
            - FLAG_CROSS_DAY_OVERNIGHT
            - FLAG_CROSS_DAY_NEXTDAY
            - FLAG_OPEN_END
            - FLAG_SAME_LOC_LONG
            - FLAG_CROSS_DAY_BOUNDARY
            Plus data columns: Person_ID, Full_Name, Branch, Trip_Start, Trip_End,
                              Duration_min, Origin, Destination, Distance_m,
                              Origin_Lat, Origin_Lng, Dest_Lat, Dest_Lng
        config: Configuration dictionary (reserved for future use)

    Returns:
        trip_behavior_df: DataFrame with original data columns (A-M) plus behavior columns:
            - BEH_Forget_Checkout: "TRUE" if FLAG_OPEN_END == "TRUE"
            - BEH_New_Day_Work: "TRUE" if FLAG_CROSS_DAY_OVERNIGHT=="TRUE" OR FLAG_CROSS_DAY_NEXTDAY=="TRUE"
            - BEH_Same_Loc: "TRUE" if FLAG_SAME_LOC_LONG == "TRUE"
            - BEH_GPS_Borderline: "TRUE" if FLAG_GPS_BORDERLINE == "TRUE"
            - BEH_Double_Checkin: Always blank (filtered in step 1)
            - BEH_Double_Checkout: Always blank (filtered in step 1)
            - BEH_ANY_FLAG: "TRUE" if any behavior flag is "TRUE"
    """

    # Create a copy to avoid modifying the input DataFrame
    trip_behavior_df = trips_df.copy()

    # Initialize behavior columns with empty string (blank)
    trip_behavior_df['BEH_Forget_Checkout'] = 0
    trip_behavior_df['BEH_New_Day_Work'] = 0
    trip_behavior_df['BEH_Same_Loc'] = 0
    trip_behavior_df['BEH_GPS_Borderline'] = 0
    trip_behavior_df['BEH_Double_Checkin'] = 0
    trip_behavior_df['BEH_Double_Checkout'] = 0

    # BEH_Forget_Checkout: Set to TRUE if FLAG_OPEN_END is TRUE
    forget_checkout_mask = trip_behavior_df['FLAG_OPEN_END'] == 1
    trip_behavior_df.loc[forget_checkout_mask, 'BEH_Forget_Checkout'] = 1

    # BEH_New_Day_Work: Set to TRUE if either cross-day overnight or next-day flag is TRUE,
    # then remove those rows from the DataFrame
    new_day_work_mask = (
        (trip_behavior_df['FLAG_CROSS_DAY_OVERNIGHT'] == 1) |
        (trip_behavior_df['FLAG_CROSS_DAY_NEXTDAY'] == 1)
    )
    trip_behavior_df.loc[new_day_work_mask, 'BEH_New_Day_Work'] = 1
    trip_behavior_df = trip_behavior_df[~new_day_work_mask].reset_index(drop=True)

    # BEH_Same_Loc: Set to TRUE if FLAG_SAME_LOC_LONG is TRUE, then remove those rows
    same_loc_mask = trip_behavior_df['FLAG_SAME_LOC_LONG'] == 1
    trip_behavior_df.loc[same_loc_mask, 'BEH_Same_Loc'] = 1
    trip_behavior_df = trip_behavior_df[~same_loc_mask].reset_index(drop=True)

    # BEH_GPS_Borderline: Set to TRUE if FLAG_GPS_BORDERLINE is TRUE, then remove those rows
    gps_borderline_mask = trip_behavior_df['FLAG_GPS_BORDERLINE'] == 1
    trip_behavior_df.loc[gps_borderline_mask, 'BEH_GPS_Borderline'] = 1
    trip_behavior_df = trip_behavior_df[~gps_borderline_mask].reset_index(drop=True)

    # BEH_ANY_FLAG: Set to TRUE if any of the behavior flags is TRUE
    any_flag_mask = (
        (trip_behavior_df['BEH_Forget_Checkout'] == 1) |
        (trip_behavior_df['BEH_New_Day_Work'] == 1) |
        (trip_behavior_df['BEH_Same_Loc'] == 1) |
        (trip_behavior_df['BEH_GPS_Borderline'] == 1) |
        (trip_behavior_df['BEH_Double_Checkin'] == 1) |
        (trip_behavior_df['BEH_Double_Checkout'] == 1)
    )
    trip_behavior_df.loc[any_flag_mask, 'BEH_ANY_FLAG'] = 1

    # Fill any missing BEH_ANY_FLAG values with 0
    trip_behavior_df['BEH_ANY_FLAG'] = trip_behavior_df['BEH_ANY_FLAG'].fillna(0)
    return trip_behavior_df
