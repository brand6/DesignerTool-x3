import xlwings as xl


class WorkApp(xl.App):

    def __init__(self, visible=False, addBook=False):
        for app in xl.apps:
            if app.visible is False:
                app.quit()
        super().__init__(visible, add_book=addBook)

    def __enter__(self):
        self.display_alerts = False
        self.screen_updating = False
        return self

    def __exit__(self, *exc_info):
        self.display_alerts = True
        self.screen_updating = True
        self.quit()
