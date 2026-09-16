[app]

title = VocabularyApp
package.name = vocabapp
package.domain = org.test

source.dir = .
source.include_exts = py,png,jpg,kv,atlas,xlsx,xlsm,mp3,ttc,ttf

version = 0.1
requirements = python3,kivy,openpyxl,edge-tts,aiohttp,attrs,multidict,yarl,async-timeout,idna,charset-normalizer

orientation = portrait
fullscreen = 0

[buildozer]

log_level = 2
warn_on_root = 1

[android]

android.permissions = INTERNET
android.api = 33
android.minapi = 21
android.ndk_path = 
android.sdk_path = 
android.accept_sdk_license = True
android.archs = arm64-v8a, armeabi-v7a
