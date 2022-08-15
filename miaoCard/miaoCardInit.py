import xlwings as xw
from miaoCard.miaoCard import MiaoCard


def main():
    app = xw.apps.active
    try:
        app.screen_updating = False
        miaoCard = MiaoCard()
        miaoCard.GameBegin()
        miaoCard.ShowChange()
    finally:
        app.screen_updating = True
