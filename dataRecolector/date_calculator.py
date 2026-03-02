"""
Date calculation module for transit and return dates.

Handles business day calculations considering holidays and weekends.
"""
import datetime as dt
from functools import lru_cache
from typing import List, Optional, FrozenSet
import pandas as pd
import numpy as np


class DateCalculator:
    """
    Calculates transit days and return dates considering non-working days.
    
    Uses LRU cache for performance optimization on repeated calculations.
    """
    
    def __init__(self, non_working_days: List[dt.datetime] = None):
        """
        Initialize the calculator with a list of non-working days.
        
        Args:
            non_working_days: List of dates that are holidays/non-working days.
        """
        # Convert to frozenset for hashability (needed for caching)
        self._non_working_days: FrozenSet[dt.datetime] = frozenset(
            non_working_days or []
        )
    
    def calculate_transit_days(
        self,
        ship_date: dt.datetime,
        delivery_date: dt.datetime
    ) -> int:
        """
        Calculate the number of transit days between ship and delivery dates.
        
        Args:
            ship_date: The date the shipment is sent.
            delivery_date: The date the shipment should arrive.
            
        Returns:
            Number of transit days (minimum 1).
        """
        if self._is_invalid_date(ship_date) or self._is_invalid_date(delivery_date):
            return 1
        
        return self._cached_transit_days(ship_date, delivery_date, self._non_working_days)
    
    def calculate_return_date(
        self,
        has_return: bool,
        delivery_date: dt.datetime,
        transit_days: int
    ) -> Optional[dt.datetime]:
        """
        Calculate the return date based on delivery date and transit days.
        
        Args:
            has_return: Whether a return is required.
            delivery_date: The delivery date.
            transit_days: Number of transit days.
            
        Returns:
            The calculated return date, or None if no return is needed.
        """
        if not has_return:
            return None
        
        if self._is_invalid_date(delivery_date) or self._is_invalid_transit(transit_days):
            return None
        
        return self._cached_return_date(delivery_date, transit_days, self._non_working_days)
    
    @staticmethod
    @lru_cache(maxsize=512)
    def _cached_transit_days(
        ship_date: dt.datetime,
        delivery_date: dt.datetime,
        non_working_days: FrozenSet[dt.datetime]
    ) -> int:
        """Cached calculation of transit days."""
        next_day = ship_date + dt.timedelta(days=1)
        days_to_working = DateCalculator._days_until_working_day(next_day, non_working_days)
        
        total_days = (delivery_date - ship_date).days
        transit = total_days - days_to_working
        
        return max(transit, 1)
    
    @staticmethod
    @lru_cache(maxsize=512)
    def _cached_return_date(
        delivery_date: dt.datetime,
        transit_days: int,
        non_working_days: FrozenSet[dt.datetime]
    ) -> dt.datetime:
        """Cached calculation of return date."""
        # First working day after delivery
        next_working = DateCalculator._next_working_day(delivery_date, non_working_days)
        
        # Add transit days and find the next working day
        target_date = next_working + dt.timedelta(days=transit_days)
        return DateCalculator._next_working_day(target_date, non_working_days)
    
    @staticmethod
    def _is_working_day(date: dt.datetime, non_working_days: FrozenSet[dt.datetime]) -> bool:
        """Check if a date is a working day (weekday and not a holiday)."""
        is_weekday = date.weekday() <= 4  # Monday = 0, Friday = 4
        is_not_holiday = date not in non_working_days
        return is_weekday and is_not_holiday
    
    @staticmethod
    def _days_until_working_day(
        date: dt.datetime,
        non_working_days: FrozenSet[dt.datetime]
    ) -> int:
        """Count days until the next working day (0 if already a working day)."""
        days = 0
        while not DateCalculator._is_working_day(date + dt.timedelta(days=days), non_working_days):
            days += 1
        return days
    
    @staticmethod
    def _next_working_day(
        date: dt.datetime,
        non_working_days: FrozenSet[dt.datetime]
    ) -> dt.datetime:
        """Get the next working day on or after the given date."""
        days_to_add = DateCalculator._days_until_working_day(date, non_working_days)
        return date + dt.timedelta(days=days_to_add)
    
    @staticmethod
    def _is_invalid_date(date) -> bool:
        """Check if a date value is invalid (None or NaT)."""
        return date is None or pd.isna(date)
    
    @staticmethod
    def _is_invalid_transit(transit_days) -> bool:
        """Check if transit days value is invalid."""
        try:
            return np.isnan(transit_days)
        except (TypeError, ValueError):
            return True
    
    def clear_cache(self):
        """Clear the calculation caches."""
        self._cached_transit_days.cache_clear()
        self._cached_return_date.cache_clear()
