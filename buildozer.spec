[app]

title = VocabApp
package.name = vocabapp
package.domain = org.test

source.dir = .
source.include_exts = py,png,jpg,kv,atlas,xlsx,xlsm,mp3,ttc,ttf

version = 0.1
# 只保留核心套件 + Android 原生 TTS 語音工具 (plyer)
requirements = python3,kivy,openpyxl,plyer

orientation = portrait
fullscreen = 0

[buildozer]

log_level = 2
warn_on_root = 1

[android]

android.permissions = INTERNET
android.api = 33
android.minapi = 21
android.accept_sdk_license = True
# 只編譯 64 位元，避免雙架構衝突
android.archs = arm64-v8a
