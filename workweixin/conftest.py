#!/usr/bin/python
# -*- coding: UTF-8 -*-
from typing import List
import pytest

@pytest.fixture(scope='module',autouse=True)
def status():
    print("开始测试--打开浏览器")
    yield

    print("结束测试--关闭浏览器")
def pytest_collection_modifyitems(
    session: "Session", config: "Config", items: List["Item"]
) -> None:
    """
    用例结果显示汉子
    """
    for item in items:

        item.name = item.name.encode('utf-8').decode('unicode-escape')
        item._nodeid = item.nodeid.encode('utf8').decode('unicode-escape')
