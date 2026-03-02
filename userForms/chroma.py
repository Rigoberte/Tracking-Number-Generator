"""
Modern Theme System for the application.

Provides a centralized color palette with dark/light mode support
and modern UI styling constants.
"""
from dataclasses import dataclass
from typing import Tuple


@dataclass(frozen=True)
class ThemeColors:
    """Immutable color palette for a theme."""
    # Main colors
    body: str
    sidebar: str
    surface: str
    surface_hover: str
    
    # Primary brand colors
    primary: str
    primary_hover: str
    primary_light: str
    
    # Accent colors
    accent: str
    accent_hover: str
    
    # Text colors
    text_primary: str
    text_secondary: str
    text_on_primary: str
    text_muted: str
    
    # Status colors
    success: str
    success_light: str
    error: str
    error_light: str
    warning: str
    warning_light: str
    
    # Table colors
    table_row_odd: str
    table_row_even: str
    table_row_done: str
    table_row_done_alt: str
    table_row_error: str
    table_row_error_alt: str
    table_header: str
    table_border: str
    
    # Input colors
    input_bg: str
    input_border: str
    input_focus: str
    
    # Shadow (for modern depth)
    shadow: str


# Modern Light Theme - Keeping red as primary brand color
LIGHT_THEME = ThemeColors(
    # Main colors
    body="#F1F1F1",
    sidebar="#DC2626",  # Modern red (slightly adjusted from #E71316)
    surface="#FFFFFF",
    surface_hover="#F3F4F6",
    
    # Primary brand colors (red tones - buttons darker than sidebar for contrast)
    primary="#B91C1C",
    primary_hover="#991B1B",
    primary_light="#FEE2E2",
    
    # Accent colors
    accent="#059669",  # Emerald green for contrast
    accent_hover="#047857",
    
    # Text colors
    text_primary="#111827",
    text_secondary="#4B5563",
    text_on_primary="#FFFFFF",
    text_muted="#9CA3AF",
    
    # Status colors
    success="#059669",
    success_light="#D1FAE5",
    error="#DC2626",
    error_light="#FEE2E2",
    warning="#D97706",
    warning_light="#FEF3C7",
    
    # Table colors
    table_row_odd="#FFFFFF",
    table_row_even="#F9FAFB",
    table_row_done="#D1FAE5",
    table_row_done_alt="#A7F3D0",
    table_row_error="#FEE2E2",
    table_row_error_alt="#FECACA",
    table_header="#F3F4F6",
    table_border="#E5E7EB",
    
    # Input colors
    input_bg="#FFFFFF",
    input_border="#D1D5DB",
    input_focus="#DC2626",
    
    # Shadow
    shadow="#00000015",
)


# Modern Dark Theme - Using purple accent like original
DARK_THEME = ThemeColors(
    # Main colors
    body="#0F172A",  # Slate 900
    sidebar="#1E293B",  # Slate 800
    surface="#1E293B",
    surface_hover="#334155",
    
    # Primary brand colors (purple tones for dark mode)
    primary="#8B5CF6",  # Violet
    primary_hover="#7C3AED",
    primary_light="#2E1065",
    
    # Accent colors
    accent="#10B981",  # Emerald
    accent_hover="#059669",
    
    # Text colors
    text_primary="#F8FAFC",
    text_secondary="#CBD5E1",
    text_on_primary="#FFFFFF",
    text_muted="#64748B",
    
    # Status colors
    success="#10B981",
    success_light="#064E3B",
    error="#EF4444",
    error_light="#450A0A",
    warning="#F59E0B",
    warning_light="#451A03",
    
    # Table colors
    table_row_odd="#1E293B",
    table_row_even="#0F172A",
    table_row_done="#2D4F4F",      # Soft teal-gray
    table_row_done_alt="#3D5F5F",  # Slightly lighter
    table_row_error="#4F2D3D",     # Soft mauve-gray
    table_row_error_alt="#5F3D4D", # Slightly lighter
    table_header="#334155",
    table_border="#475569",
    
    # Input colors
    input_bg="#1E293B",
    input_border="#475569",
    input_focus="#8B5CF6",
    
    # Shadow
    shadow="#00000040",
)


@dataclass
class UIConstants:
    """Modern UI styling constants."""
    # Border radius
    radius_sm: int = 4
    radius_md: int = 8
    radius_lg: int = 12
    radius_xl: int = 16
    radius_full: int = 9999
    
    # Spacing
    spacing_xs: int = 4
    spacing_sm: int = 8
    spacing_md: int = 16
    spacing_lg: int = 24
    spacing_xl: int = 32
    
    # Font sizes
    font_xs: int = 11
    font_sm: int = 12
    font_md: int = 14
    font_lg: int = 16
    font_xl: int = 20
    font_2xl: int = 24
    font_3xl: int = 30
    
    # Font families
    font_family: str = "Segoe UI"
    font_family_mono: str = "Consolas"
    
    # Button sizes
    btn_height_sm: int = 32
    btn_height_md: int = 40
    btn_height_lg: int = 48
    btn_width_sm: int = 80
    btn_width_md: int = 120
    btn_width_lg: int = 160
    
    # Animation duration (ms)
    animation_fast: int = 150
    animation_normal: int = 250
    animation_slow: int = 350


# Singleton constants
UI = UIConstants()


class Chroma:
    """
    Theme manager with dark/light mode toggle.
    
    Maintains backward compatibility while providing modern theme support.
    """
    
    def __init__(self):
        self.dark = False
        self._theme: ThemeColors = LIGHT_THEME
        
        # Legacy compatibility attributes
        self.body_color = None
        self.sidebar_color = None
        self.primary_color = None
        self.primary_color_light = None
        self.toggle_color_for_buttons = None
        self.text_color = None
        
        self.toggle()
    
    @property
    def theme(self) -> ThemeColors:
        """Get the current theme colors."""
        return self._theme
    
    def toggle(self) -> None:
        """Toggle between dark and light mode."""
        self.dark = not self.dark
        self._theme = DARK_THEME if self.dark else LIGHT_THEME
        self._update_legacy_colors()
    
    def _update_legacy_colors(self) -> None:
        """Update legacy color attributes for backward compatibility."""
        self.body_color = self._theme.body
        self.sidebar_color = self._theme.sidebar
        self.primary_color = self._theme.primary
        self.primary_color_light = self._theme.primary_hover
        self.toggle_color_for_buttons = self._theme.text_on_primary
        self.text_color = self._theme.text_primary
    
    # Legacy getters (kept for backward compatibility)
    def getBodyColor(self) -> str:
        return self._theme.body
    
    def getSidebarColor(self) -> str:
        return self._theme.sidebar
    
    def getPrimaryColor(self) -> str:
        return self._theme.primary
    
    def getPrimaryColorLight(self) -> str:
        return self._theme.primary_hover
    
    def getTextColorForButtons(self) -> str:
        return self._theme.text_on_primary
    
    def getTextColor(self) -> str:
        return self._theme.text_primary
    
    def getDarkMode(self) -> bool:
        return self.dark
    
    # New modern getters
    def get_surface(self) -> str:
        return self._theme.surface
    
    def get_surface_hover(self) -> str:
        return self._theme.surface_hover
    
    def get_text_secondary(self) -> str:
        return self._theme.text_secondary
    
    def get_text_muted(self) -> str:
        return self._theme.text_muted
    
    def get_success(self) -> str:
        return self._theme.success
    
    def get_error(self) -> str:
        return self._theme.error
    
    def get_accent(self) -> str:
        return self._theme.accent
    
    def get_input_bg(self) -> str:
        return self._theme.input_bg
    
    def get_input_border(self) -> str:
        return self._theme.input_border
    
    def get_table_colors(self) -> dict:
        """Get all table-related colors."""
        return {
            'odd': self._theme.table_row_odd,
            'even': self._theme.table_row_even,
            'odd_done': self._theme.table_row_done,
            'even_done': self._theme.table_row_done_alt,
            'odd_error': self._theme.table_row_error,
            'even_error': self._theme.table_row_error_alt,
            'header': self._theme.table_header,
            'border': self._theme.table_border,
        }
