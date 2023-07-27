class Common:
    @classmethod
    def toStr(cls, content) -> str:
        """转为字符串，None会转为''

        Args:
            content (_type_): _description_

        Returns:
            _type_: _description_
        """
        if content is None:
            return ""
        elif cls.isNumber(content):
            intC = cls.toInt(content)
            if intC == round(float(content), 10):
                return str(intC)
            else:
                return str(content)
        else:
            return str(content)

    @classmethod
    def toNum(cls, content) -> float:
        """转为数值，非数值返回0

        Args:
            content (_type_): _description_

        Returns:
            _type_: _description_
        """
        if cls.isNumber(content):
            return float(content)
        else:
            return 0.0

    @classmethod
    def toInt(cls, content) -> int:
        """转为整数，四舍五入(处理excel数据莫名其妙变成很长小数的问题)，非数值返回-1

        Args:
            content (_type_): _description_

        Returns:
            _type_: _description_
        """
        if cls.isNumber(content):
            return round(float(content))
        else:
            return -1

    @classmethod
    def isNumberValid(cls, content, checkNum=0) -> bool:
        """判断数字是否有效，>checkNum为有效

        Args:
            content (_type_): 支持字符串格式的数字
            checkNum (int, optional): 有效的条件. Defaults to 0.

        Returns:
            bool: _description_
        """
        if not cls.isNumber(content):
            return False
        elif float(content) > checkNum:
            return True
        else:
            return False

    @classmethod
    def isNumber(cls, content) -> bool:
        """判断是否数字，None不是数字

        Args:
            content (_type_): _description_

        Returns:
            _type_: _description_
        """
        try:
            float(content)
            return True
        except TypeError:
            return False

    @classmethod
    def isEmpty(cls, content) -> bool:
        """判断是否为空对象或空字符串

        Args:
            content (_type_): _description_

        Returns:
            _type_: _description_
        """
        if content == "" or content is None:
            return True
        else:
            return False

    @classmethod
    def split(cls, content, sep) -> list[str]:
        """重载分隔操作，空对象会转为空列表

        Args:
            content (_type_): 处理对象
            sep (_type_): 分隔符

        Returns:
            list[str]: _description_
        """
        if content is None:
            return []
        else:
            return str.split(cls.toStr(content), sep)
