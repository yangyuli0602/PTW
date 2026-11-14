# -*- coding: utf-8 -*-
"""版本真相：只有这里定义版本号与更新源。

改版本只改这一个文件（打包脚本会读它，server 也引用它）。
UPDATE_REPO 形如 "yourname/PTW"——填了才会去检查更新；留空则「检查更新」直接报未配置。
"""
import os

#: 当前版本（发布时改这里，然后打包 → 建 Release → 最后改仓库 version.json）
APP_VERSION = "1.3.3"

#: GitHub 仓库（owner/repo）。留空 = 不检查更新（开发版）
UPDATE_REPO = "yangyuli0602/PTW"

#: 下载页（用户点「打开下载页」时跳这里）
GITHUB_PAGE = ("https://github.com/%s/releases/latest" % UPDATE_REPO) if UPDATE_REPO else ""

APP_NAME = "PTW平台"
