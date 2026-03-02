"""
UI Events.

Defines event types for communication between Model and View,
replacing magic strings with strongly-typed events.
"""
from dataclasses import dataclass
from enum import Enum, auto
from typing import Any, Optional
import pandas as pd


class UIEventType(Enum):
    """Types of UI events that can be emitted."""
    BLOCK_WIDGETS = auto()
    UNBLOCK_WIDGETS = auto()
    UPDATE_ORDERS = auto()
    UPDATE_ROW = auto()  # Update a single row in the orders table
    SHOW_MESSAGE = auto()
    SHOW_ERROR = auto()


@dataclass(frozen=True)
class UIEvent:
    """
    Represents an event to be processed by the View.
    
    Attributes:
        event_type: The type of event
        data: Optional payload data for the event
    """
    event_type: UIEventType
    data: Optional[Any] = None
    
    @classmethod
    def block_widgets(cls) -> 'UIEvent':
        """Create a BLOCK_WIDGETS event."""
        return cls(UIEventType.BLOCK_WIDGETS)
    
    @classmethod
    def unblock_widgets(cls) -> 'UIEvent':
        """Create an UNBLOCK_WIDGETS event."""
        return cls(UIEventType.UNBLOCK_WIDGETS)
    
    @classmethod
    def update_orders(cls, dataframe: pd.DataFrame) -> 'UIEvent':
        """Create an UPDATE_ORDERS event with the orders dataframe."""
        return cls(UIEventType.UPDATE_ORDERS, dataframe)
    
    @classmethod
    def update_row(cls, row_data: dict) -> 'UIEvent':
        """Create an UPDATE_ROW event with the row data."""
        return cls(UIEventType.UPDATE_ROW, row_data)
    
    @classmethod
    def show_message(cls, message: str) -> 'UIEvent':
        """Create a SHOW_MESSAGE event."""
        return cls(UIEventType.SHOW_MESSAGE, message)
    
    @classmethod
    def show_error(cls, error: str) -> 'UIEvent':
        """Create a SHOW_ERROR event."""
        return cls(UIEventType.SHOW_ERROR, error)
