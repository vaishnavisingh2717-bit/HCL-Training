"""
src/data_loader.py - SQLite Data Access & Caching Layer
Provides centralized queries into data/app.db with @st.cache_data memoization.
Supports pandas when present, with a lightweight fallback for headless environments.
"""

import os
import sqlite3
from pathlib import Path

# Path Resolution using Path(__file__)
BASE_DIR = Path(__file__).resolve().parent.parent
DB_PATH = BASE_DIR / "data" / "app.db"

try:
    import pandas as pd
except ImportError:
    pd = None

try:
    import streamlit as st
    cache_decorator = st.cache_data
except ImportError:
    def cache_decorator(func):
        return func

class SimpleDataFrame:
    """Lightweight fallback representing a tabular result when pandas is not installed."""
    def __init__(self, data: list, columns: list):
        self._data = data
        self.columns = columns
        self.shape = (len(data), len(columns))
        
    def __len__(self):
        return len(self._data)
        
    def __getitem__(self, col):
        if isinstance(col, str):
            idx = self.columns.index(col)
            return [row[idx] for row in self._data]
        elif isinstance(col, list):
            indices = [self.columns.index(c) for c in col]
            sub_data = [[row[i] for i in indices] for row in self._data]
            return SimpleDataFrame(sub_data, col)
        return self._data[col]
        
    def head(self, n=5):
        return SimpleDataFrame(self._data[:n], self.columns)
        
    def to_dict(self, orient="records"):
        return [dict(zip(self.columns, row)) for row in self._data]

def get_connection(read_only: bool = True):
    """Establishes an SQLite connection with Foreign Keys enabled."""
    db_uri = f"file:{DB_PATH.as_posix()}?mode=ro" if read_only else str(DB_PATH)
    conn = sqlite3.connect(db_uri, uri=read_only)
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn

def _execute_query(sql: str, conn):
    if pd is not None:
        return pd.read_sql_query(sql, conn)
    cur = conn.cursor()
    cur.execute(sql)
    rows = cur.fetchall()
    cols = [desc[0] for desc in cur.description]
    return SimpleDataFrame(rows, cols)

@cache_decorator
def load_master_data():
    """Loads all master dimension tables into a dictionary of DataFrames."""
    conn = get_connection()
    try:
        dim_channel = _execute_query("SELECT * FROM dim_channel;", conn)
        dim_campaign = _execute_query("SELECT * FROM dim_campaign;", conn)
        dim_customer = _execute_query("SELECT * FROM dim_customer;", conn)
        return {
            "dim_channel": dim_channel,
            "dim_campaign": dim_campaign,
            "dim_customer": dim_customer
        }
    finally:
        conn.close()

@cache_decorator
def load_analytical_view():
    """Loads the denormalized reporting view (vw_marketing_analysis)."""
    conn = get_connection()
    try:
        df = _execute_query("SELECT * FROM vw_marketing_analysis;", conn)
        if pd is not None and "date" in df.columns:
            df["date"] = pd.to_datetime(df["date"])
        return df
    finally:
        conn.close()

@cache_decorator
def load_touchpoints():
    """Loads customer multi-touch journey touchpoints."""
    conn = get_connection()
    try:
        df = _execute_query("SELECT * FROM fact_customer_touchpoint;", conn)
        if pd is not None and "touchpoint_timestamp" in df.columns:
            df["touchpoint_timestamp"] = pd.to_datetime(df["touchpoint_timestamp"])
        return df
    finally:
        conn.close()

@cache_decorator
def load_daily_facts():
    """Loads raw daily campaign performance fact table."""
    conn = get_connection()
    try:
        df = _execute_query("SELECT * FROM fact_daily_campaign_performance;", conn)
        if pd is not None and "date" in df.columns:
            df["date"] = pd.to_datetime(df["date"])
        return df
    finally:
        conn.close()
