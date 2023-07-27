import json
import re

nameMap = {
    '2201': '跳过',
    '2202': '冰冻',
    '2203': '否决',
    '2204': '发财',
    '2205': '贪心',
    '2206': '过河拆喵',
    '2207': '顺手牵喵',
    '2208': '变色',
    '2209': '破产',
    '2210': '变小',
    '2211': '拆迁',
    '2212': '兴奋'
}

numList = [1, 2, 3, 4, 5, 6]


def len_zh(value):
    temp = re.findall('[^a-zA-Z0-9.]+', str(value))
    count = 0
    for i in temp:
        count += len(i)
    return count


def adjust(value, length):
    zh = len_zh(value)
    return str(value).ljust(length - zh)


def getCardName(cardMap: map):
    returnMap = {}
    for card in cardMap.keys():
        if card in nameMap:
            returnMap[nameMap[card]] = cardMap[card]
        else:
            if card[-1] in returnMap:
                returnMap[int(card[-1])] += cardMap[card]
            else:
                returnMap[int(card[-1])] = cardMap[card]

    for num in numList:
        if num not in returnMap:
            returnMap[num] = 0
    for name in nameMap.values():
        if name not in returnMap:
            returnMap[name] = 0
    return returnMap


file1 = r'F:\x3-obt-dev\Program\Tools\LuaServerCheck\lua_server\p1_card_record.json'
file2 = r'F:\x3-obt-dev\Program\Tools\LuaServerCheck\lua_server\p2_card_record.json'

with open(file2) as f:
    content = json.loads(str.replace(f.read(), '\x10', ''))
    manDrawMap = getCardName(content[0])
    manPutMap = getCardName(content[1])

with open(file1) as f:
    content = json.loads(str.replace(f.read(), '\x10', ''))
    playerDrawMap = getCardName(content[0])
    playerPutMap = getCardName(content[1])

for n in numList:
    print('[数字牌:{}]男主摸牌:{} 男主出牌:{}玩家摸牌:{} 玩家出牌:{}'.format(adjust(n, 8), adjust(manDrawMap[n], 4), adjust(manPutMap[n], 4),
                                                          adjust(playerDrawMap[n], 4), adjust(playerPutMap[n], 4)))

for n in nameMap.values():
    print('[功能牌:{}]男主摸牌:{} 男主出牌:{}玩家摸牌:{} 玩家出牌:{}'.format(adjust(n, 8), adjust(manDrawMap[n], 4), adjust(manPutMap[n], 4),
                                                          adjust(playerDrawMap[n], 4), adjust(playerPutMap[n], 4)))
