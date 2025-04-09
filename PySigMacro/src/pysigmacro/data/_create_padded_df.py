#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-07 15:10:30 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/data/_create_padded_df.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/data/_create_padded_df.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------

import numpy as np
import pandas as pd

def create_padded_df(*args, filler=np.nan):
    """
    Create a padded pandas DataFrame from multiple dictionaries, DataFrames, or data convertible to DataFrames.
    Inputs can be provided individually or as lists/tuples. When multiple inputs
    are provided, they are concatenated row-wise with padding. For dictionaries, values
    of different lengths are padded with a filler value to ensure a rectangular data structure.

    Args:
        *args (dict, pandas.DataFrame, or list/tuple of these or convertible objects): Data to be concatenated.
        filler (any, optional): Value used to pad shorter lists. Defaults to np.nan.

    Returns:
        pandas.DataFrame: A DataFrame resulting from concatenating the input data.

    Notes:
    - Single values (str, int, float) are treated as length-1 lists.
    - If a list of pandas Series is provided, they'll be converted to dictionaries.
    - Mixing dicts and DataFrames is acceptable; inputs provided in lists/tuples are flattened.
    - Other objects convertible to DataFrame (via pd.DataFrame) are also accepted.
    """

    ACCEPTABLE_DATATYPES = (int, float, str, dict, list, tuple, np.ndarray, pd.DataFrame)

    # Extract acceptable datatypes
    items = []
    for arg in args:
        if isinstance(arg, ACCEPTABLE_DATATYPES):
            items.append(arg)

    items = _ensure_1dims(items)

    # Calculate max lengths
    max_length = 0
    for ii, item in enumerate(items):
        if isinstance(item, dict):
            _max_length = max([len(v) for v in item.values()])
        if isinstance(item, pd.DataFrame):
            _max_length = len(item)
        max_length = max(max_length, _max_length)

    # Padding
    for ii, item in enumerate(items):
        items[ii] = pad_item(item, max_length, filler)

    # To df
    dfs = []
    for ii, item in enumerate(items):
        if isinstance(item, dict):
            item = pd.DataFrame(item) # This fails when dict with different length of items
        dfs.append(item)

    df = pd.concat(dfs, axis=1)

    return df

def _ensure_1dims(items):
    # Ensure dict values are 1-d list, tuple, or dict of 1-d list, tuple values
    for ii, item in enumerate(items):
        if isinstance(item, (int, float, str)):
            items[ii] = [item]
        if isinstance(item, (list, tuple)):
            assert np.array(item).ndim == 1
            items[ii] = item
        if isinstance(item, np.ndarray):
            assert item.ndim == 1
            items[ii] = item.tolist()
        if isinstance(item, dict):
            for k, v in item.items():
                if isinstance(v, (int, float, str)):
                    item[k] = [v]
                elif isinstance(v, (list, tuple)):
                    assert np.array(v).ndim == 1
                    item[k] = v
                elif isinstance(v, (np.ndarray, pd.DataFrame)):
                    assert v.ndim == 1
                    item[k] = v.tolist()
            # Aggregation
            items[ii] = item
    return items


def pad_item(item, target_length, filler):
    def _pad_int_float_str(val, target_length, filler):
        # Pads a single int, float, or str to a list of target_length
        return [val] + [filler] * (target_length - 1)

    def _pad_list_tuple(seq, target_length, filler):
        # Pads a list or tuple to target_length
        result = list(seq)
        result += [filler] * (target_length - len(result))
        return result

    def _pad_array(arr, target_length, filler):
        # Pads a 1d array (or array-like) to target_length and returns a list
        arr = np.array(arr)
        if arr.ndim != 1:
            raise ValueError("pad_array: Expected 1D array")
        if len(arr) < target_length:
            extra = np.full(target_length - len(arr), filler)
            arr = np.concatenate([arr, extra])
        return arr.tolist()

    def _pad_df(df, target_length, filler):
        # Pads a DataFrame with extra rows filled with filler to reach target_length rows
        if len(df) < target_length:
            n_rows_pad = target_length - len(df)
            df_pad = pd.DataFrame(np.full((n_rows_pad, df.shape[1]), filler), columns=df.columns)
            return pd.concat([df, df_pad], ignore_index=True)
        return df

    # Main
    if isinstance(item, (int, float, str)):
        return _pad_int_float_str(item, target_length, filler)
    elif isinstance(item, (list, tuple)):
        return _pad_list_tuple(item, target_length, filler)
    elif isinstance(item, np.ndarray):
        return _pad_array(item, target_length, filler)
    elif isinstance(item, pd.DataFrame):
        return _pad_df(item, target_length, filler)
    elif isinstance(item, dict):
        return {k: pad_item(v, target_length, filler) for k,v in item.items()}
    else:
        raise TypeError("Unsupported datatype for padding")

# EOF