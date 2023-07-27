from card import Card
from cardGacha import Gacha
from dollCatcher import DollCatcher
from hangUp import HangUp
from itemSpawn import ItemSpawn
from miaoGacha import MiaoGacha
from player import Player
from stage import Stage


def testDollCatcher():
    catcher = DollCatcher(2101, 201)
    dollMap = {20020: 6, 20130: 6}
    dolls = catcher.GetDoll()
    print("getDolls", dolls)


def testMiaoGacha():
    gacha = MiaoGacha(1001)
    for i in range(28):
        gacha.Draw()
    print(gacha.getList)


def testHangUp():
    player1 = Player("test")
    result = player1.HangUpExplore(1000, 10)
    print(result)


def testGacha():
    ItemSpawn.ItemInit()
    Card.CardInit()
    player1 = Player("test")
    result = player1.DrawCard(101, 1, True)
    print(result)


def testCard():
    resMap = {"1=1": 50000, "210=0": 10000, "205=100101": 1000}
    playerLev = 20
    card1 = Card(111113)
    for i in range(10):
        result, info = card1.LevelUp(resMap)
        print("levelUp", i, result, info)
    print(f"card1 lev:{card1.cardLev}")
    for i in range(2):
        result, info = card1.StarUp(playerLev, resMap)
        print("starUp", i, result, info)
    print(f"card1 star:{card1.cardStar}")


def testStage():
    resMap = {}
    stage1 = Stage(40102)
    stage1.GetReward(resMap, 1)
    print(resMap)


testGacha()
