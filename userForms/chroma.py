class Chroma:
    def __init__(self):
        """
        Class constructor for Chroma
        """
        self.dark = False

        self.body_color = None
        self.sidebar_color = None
        self.primary_color = None
        self.primary_color_light = None
        self.toggle_color = None
        self.text_color = None
        self.toggle()

    def toggle(self):
        """
        Toggles between dark and light mode
        """
        self.dark = not self.dark
        if self.dark:
            self.body_color = '#121212'
            self.sidebar_color = '#2B2B2B'
            self.primary_color = '#3700B3' 
            self.primary_color_light = '#BB86FC'
            self.toggle_color = '#3700B3'
            self.text_color = '#CCC'
        else:
            self.body_color = '#CCC'
            self.sidebar_color = '#E71316'
            self.primary_color = '#c00000'
            self.primary_color_light = '#E83713'
            self.toggle_color = '#BE2F31'
            self.text_color = '#000'

    def getBodyColor(self):
        return self.body_color
    
    def getSidebarColor(self):
        return self.sidebar_color
    
    def getPrimaryColor(self):
        return self.primary_color
    
    def getPrimaryColorLight(self):
        return self.primary_color_light
    
    def getToggleColor(self):
        return self.toggle_color
    
    def getTextColor(self):
        return self.text_color

    def getDarkMode(self):
        return self.dark
