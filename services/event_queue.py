"""
Event Queue.

Provides an encapsulated queue for UI events with proper typing
and backward compatibility with legacy string-based events.
"""
import queue
from typing import Union, Optional, Any
import pandas as pd

from .ui_events import UIEvent, UIEventType


class EventQueue:
    """
    Encapsulated event queue for Model-View communication.
    
    This class wraps a queue.Queue and provides:
    - Type-safe event emission
    - Backward compatibility with legacy string events
    - Encapsulation (no direct queue access)
    """
    
    # Legacy string mappings for backward compatibility
    _LEGACY_STRINGS = {
        "BLOCK MAIN USERFORM WIDGETS": UIEventType.BLOCK_WIDGETS,
        "UNBLOCK MAIN USERFORM WIDGETS": UIEventType.UNBLOCK_WIDGETS,
    }
    
    def __init__(self):
        self._queue: queue.Queue = queue.Queue()
    
    def emit(self, event: UIEvent) -> None:
        """
        Emit a typed UI event.
        
        Args:
            event: The UIEvent to emit
        """
        self._queue.put(event)
    
    def emit_block_widgets(self) -> None:
        """Emit a block widgets event."""
        self.emit(UIEvent.block_widgets())
    
    def emit_unblock_widgets(self) -> None:
        """Emit an unblock widgets event."""
        self.emit(UIEvent.unblock_widgets())
    
    def emit_update_orders(self, dataframe: pd.DataFrame) -> None:
        """Emit an update orders event with the dataframe."""
        self.emit(UIEvent.update_orders(dataframe))
    
    def emit_message(self, message: str) -> None:
        """Emit a message event."""
        self.emit(UIEvent.show_message(message))
    
    def emit_error(self, error: str) -> None:
        """Emit an error event."""
        self.emit(UIEvent.show_error(error))
    
    def get(self, block: bool = True, timeout: Optional[float] = None) -> UIEvent:
        """
        Get an event from the queue.
        
        Handles both new UIEvent types and legacy string/DataFrame events
        for backward compatibility.
        
        Args:
            block: Whether to block waiting for an event
            timeout: Timeout in seconds
            
        Returns:
            UIEvent: The next event from the queue
            
        Raises:
            queue.Empty: If non-blocking and queue is empty
        """
        item = self._queue.get(block=block, timeout=timeout)
        return self._convert_to_event(item)
    
    def get_nowait(self) -> UIEvent:
        """
        Get an event without blocking.
        
        Returns:
            UIEvent: The next event from the queue
            
        Raises:
            queue.Empty: If queue is empty
        """
        return self.get(block=False)
    
    def empty(self) -> bool:
        """Check if the queue is empty."""
        return self._queue.empty()
    
    def _convert_to_event(self, item: Any) -> UIEvent:
        """
        Convert legacy items to UIEvent.
        
        Handles:
        - UIEvent: returned as-is
        - str: converted to appropriate event type
        - DataFrame: converted to UPDATE_ORDERS event
        - dict: converted to UPDATE_ROW event
        """
        if isinstance(item, UIEvent):
            return item
        
        if isinstance(item, str):
            event_type = self._LEGACY_STRINGS.get(item)
            if event_type == UIEventType.BLOCK_WIDGETS:
                return UIEvent.block_widgets()
            elif event_type == UIEventType.UNBLOCK_WIDGETS:
                return UIEvent.unblock_widgets()
            else:
                return UIEvent.show_message(item)
        
        if isinstance(item, pd.DataFrame):
            return UIEvent.update_orders(item)
        
        if isinstance(item, dict):
            return UIEvent.update_row(item)
        
        # Unknown type, wrap as message
        return UIEvent.show_message(str(item))
    
    # Legacy compatibility - allows putting raw items
    def put_legacy(self, item: Any) -> None:
        """
        Put a raw item in the queue (legacy compatibility).
        
        Prefer using emit() methods for new code.
        """
        self._queue.put(item)
    
    # Alias for backward compatibility with code that uses queue.put()
    def put(self, item: Any) -> None:
        """
        Backward compatible put method.
        
        Automatically converts items to UIEvent when possible.
        """
        if isinstance(item, UIEvent):
            self._queue.put(item)
        elif isinstance(item, pd.DataFrame):
            self.emit_update_orders(item)
        elif isinstance(item, dict):
            self.emit(UIEvent.update_row(item))
        elif isinstance(item, str):
            event_type = self._LEGACY_STRINGS.get(item)
            if event_type == UIEventType.BLOCK_WIDGETS:
                self.emit_block_widgets()
            elif event_type == UIEventType.UNBLOCK_WIDGETS:
                self.emit_unblock_widgets()
            else:
                self.emit_message(item)
        else:
            self._queue.put(item)
