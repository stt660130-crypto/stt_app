[app]

title = VocabApp
package.name = vocabapp
package.domain = org.test

source.dir = .
source.include_exts = py,png,jpg,kv,atlas,xlsx,xlsm,mp3,ttc,ttf,kv

version = 0.1
requirements = python3,kivy,openpyxl,plyer

orientation = portrait
fullscreen = 0

[buildozer]

log_level = 2
warn_on_root = 1

[android]

android.permissions = INTERNET, MODIFY_AUDIO_SETTINGS
android.api = 33
android.minapi = 24
android.accept_sdk_license = True
android.archs = arm64-v8a
