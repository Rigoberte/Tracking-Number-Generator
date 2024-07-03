from Model import Model
from View import View
from Controller import Controller

if __name__ == "__main__":
    view = View()
    model = Model()
    controller = Controller(model, view)
    controller.show_mainUserForm()