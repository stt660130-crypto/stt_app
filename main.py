import asyncio
import io
import os
import re
import shutil
import sys
import threading
import edge_tts
import openpyxl

# ----------------- Kivy 模組 -----------------
from kivy.app import App
from kivy.core.window import Window
from kivy.core.text import LabelBase
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView

# 設定視窗背景顏色為淺灰
Window.clearcolor = (0.95, 0.95, 0.96, 1)

# ----------------- 跨平台中文字型處理 -----------------
def init_custom_font():
    """自動適應 Windows 與 Android 系統字型"""
    font_paths = [
        "/system/fonts/NotoSansCJK-Regular.ttc",  # Android 常用
        "/system/fonts/DroidSansFallback.ttf",    # 舊版 Android
        "C:/Windows/Fonts/msjh.ttc",              # Windows 微軟正黑體
        "C:/Windows/Fonts/simhei.ttf"             # Windows 黑體
    ]
    for font_path in font_paths:
        if os.path.exists(font_path):
            try:
                LabelBase.register(name="Roboto", fn_regular=font_path)
                LabelBase.register(name="default_font", fn_regular=font_path)
                break
            except Exception as e:
                print(f"字型載入失敗 {font_path}: {e}")

init_custom_font()

# Android 音訊處理優化
try:
    import pygame
    pygame.mixer.init()
    HAS_PYGAME = True
except Exception:
    HAS_PYGAME = False


def clean_ipa_text(ipa_str):
    """將容易缺字的特殊音標符號轉為相近安全字元"""
    if not ipa_str:
        return ""
    
    ipa_map = {
        'ɛ': 'e', 'ə': 'e', 'æ': 'ae', 'ɑ': 'a', 'ɔ': 'o', 
        'ʊ': 'u', 'i': 'i', 'ɪ': 'i', 'ʌ': 'u', 'ɜ': 'er',
        'ʃ': 'sh', 'ʒ': 'zh', 'θ': 'th', 'ð': 'th', 'ŋ': 'ng',
        'ˈ': "'", 'ˌ': ",", '`': "'"
    }
    
    cleaned = str(ipa_str)
    for char, replacement in ipa_map.items():
        cleaned = cleaned.replace(char, replacement)
    
    cleaned = re.sub(r'[^\x00-\x7F]+', '', cleaned)
    return cleaned


class CustomLabel(Label):
    """支援自動換行與中文字型檔的自訂 Label"""
    def __init__(self, **kwargs):
        kwargs.setdefault('font_name', 'Roboto')
        super().__init__(**kwargs)
        self.bind(size=self._update_text_size)

    def _update_text_size(self, *args):
        self.text_size = (self.width, None)


class VocabularyApp(App):

    def build(self):
        self.title = "單字記憶學習機"
        self.voice_en = "en-US-AvaNeural"
        self.current_index = 0

        # --- Android Excel 檔案準備流程 ---
        self.excel_path = self.get_android_excel_path()
        self.words_data = self.load_excel_data(self.excel_path)

        self.pos_order = [
            ("V", "V.", 4, 5),
            ("N", "N.", 6, 7),
            ("Adj", "Adj.", 8, 9),
            ("Adv", "Adv.", 10, 11),
        ]

        # 主畫面佈局
        root = BoxLayout(orientation="vertical", padding=15, spacing=15)

        # 頂部導覽列與進度標示
        top_nav = BoxLayout(orientation="horizontal", size_hint_y=None, height="50dp", spacing=10)

        btn_prev = Button(
            text="← Prev", font_size="20sp", font_name="Roboto", bold=True,
            background_normal="", background_color=(0.4, 0.45, 0.5, 1), color=(1, 1, 1, 1)
        )
        btn_prev.bind(on_release=lambda x: self.prev_word())
        top_nav.add_widget(btn_prev)

        self.lbl_lesson_info = Label(
            text="", font_size="20sp", font_name="Roboto", color=(0.1, 0.1, 0.1, 1), bold=True
        )
        top_nav.add_widget(self.lbl_lesson_info)

        btn_next = Button(
            text="Next →", font_size="20sp", font_name="Roboto", bold=True,
            background_normal="", background_color=(0.12, 0.43, 0.72, 1), color=(1, 1, 1, 1)
        )
        btn_next.bind(on_release=lambda x: self.next_word())
        top_nav.add_widget(btn_next)

        root.add_widget(top_nav)

        # 滾動內容區域
        scroll = ScrollView(size_hint=(1, 1))
        content_box = BoxLayout(orientation="vertical", spacing=20, size_hint_y=None)
        content_box.bind(minimum_height=content_box.setter("height"))

        # 1. 主要單字區
        top_word_box = BoxLayout(orientation="horizontal", size_hint_y=None, height="70dp", spacing=12)
        
        self.lbl_main_word = Label(
            text="", font_size="36sp", font_name="Roboto", bold=True, color=(0.85, 0.1, 0.1, 1), size_hint_x=None
        )
        self.lbl_main_word.bind(texture_size=self.lbl_main_word.setter("size"))

        self.lbl_main_cn = Label(
            text="", font_size="20sp", font_name="Roboto", bold=True, color=(0.1, 0.1, 0.1, 1), size_hint_x=None
        )
        self.lbl_main_cn.bind(texture_size=self.lbl_main_cn.setter("size"))

        self.lbl_main_pos = Label(
            text="", font_size="20sp", font_name="Roboto", bold=True, color=(0.3, 0.3, 0.3, 1), size_hint_x=None
        )
        self.lbl_main_pos.bind(texture_size=self.lbl_main_pos.setter("size"))

        self.lbl_main_ipa = Label(
            text="", font_size="20sp", font_name="Roboto", bold=True, color=(0.2, 0.2, 0.2, 1), size_hint_x=None
        )
        self.lbl_main_ipa.bind(texture_size=self.lbl_main_ipa.setter("size"))

        top_word_box.add_widget(self.lbl_main_word)
        top_word_box.add_widget(self.lbl_main_cn)
        top_word_box.add_widget(self.lbl_main_pos)
        top_word_box.add_widget(self.lbl_main_ipa)

        btn_main_pronounce = Button(
            text="發音", font_size="18sp", font_name="Roboto", bold=True, size_hint_x=None, width="90dp",
            background_normal="", background_color=(0.2, 0.65, 0.35, 1), color=(1, 1, 1, 1)
        )
        btn_main_pronounce.bind(on_release=lambda x: self.play_main_word_audio())
        top_word_box.add_widget(btn_main_pronounce)

        content_box.add_widget(top_word_box)

        # 2. 詞性列表區域
        table_grid = GridLayout(cols=3, size_hint_y=None, spacing=10)
        table_grid.bind(minimum_height=table_grid.setter("height"))

        table_grid.add_widget(Label(text="詞性", font_size="18sp", font_name="Roboto", color=(0.1, 0.1, 0.1, 1), bold=True, size_hint_y=None, height="40dp"))
        table_grid.add_widget(Label(text="英文單字", font_size="18sp", font_name="Roboto", color=(0.1, 0.1, 0.1, 1), bold=True, size_hint_y=None, height="40dp"))
        table_grid.add_widget(Label(text="中文釋義", font_size="18sp", font_name="Roboto", color=(0.1, 0.1, 0.1, 1), bold=True, size_hint_y=None, height="40dp"))

        self.pos_en_buttons = {}
        self.pos_cn_widgets = {}

        for key, label_text, _, _ in self.pos_order:
            lbl_pos = Label(text=label_text, font_size="20sp", font_name="Roboto", color=(0.05, 0.25, 0.5, 1), bold=True, size_hint_y=None, height="55dp")
            
            btn_en = Button(
                text="—", font_size="20sp", font_name="Roboto", bold=True, size_hint_y=None, height="55dp",
                background_normal="", background_color=(0.88, 0.92, 0.98, 1), color=(0.05, 0.35, 0.75, 1)
            )
            btn_en.bind(on_release=lambda btn, k=key: self.play_pos_audio(k))

            lbl_cn = Label(text="—", font_size="20sp", font_name="Roboto", bold=True, color=(0.15, 0.15, 0.15, 1), size_hint_y=None, height="55dp")

            table_grid.add_widget(lbl_pos)
            table_grid.add_widget(btn_en)
            table_grid.add_widget(lbl_cn)

            self.pos_en_buttons[key] = btn_en
            self.pos_cn_widgets[key] = lbl_cn

        content_box.add_widget(table_grid)

        # 3. 字根欄位
        root_box = BoxLayout(orientation="horizontal", size_hint_y=None, height="45dp", spacing=5)
        root_title = Label(text="字根:", font_size="18sp", font_name="Roboto", bold=True, color=(0.1, 0.1, 0.1, 1), size_hint_x=None, width="70dp")
        self.lbl_root_val = CustomLabel(text="", font_size="18sp", bold=True, color=(0.2, 0.2, 0.2, 1))
        root_box.add_widget(root_title)
        root_box.add_widget(self.lbl_root_val)
        content_box.add_widget(root_box)

        # 4. 情境例句區
        ex_box = BoxLayout(orientation="vertical", size_hint_y=None, spacing=10)
        ex_box.bind(minimum_height=ex_box.setter("height"))

        ex_header = BoxLayout(orientation="horizontal", size_hint_y=None, height="45dp")
        lbl_ex_title = Label(text="情境例句", font_size="20sp", font_name="Roboto", bold=True, color=(0.1, 0.1, 0.1, 1), size_hint_x=None, width="120dp")
        
        btn_play_example = Button(
            text="播放例句", font_size="16sp", font_name="Roboto", bold=True, size_hint_x=None, width="110dp",
            background_normal="", background_color=(0.12, 0.5, 0.5, 1), color=(1, 1, 1, 1)
        )
        btn_play_example.bind(on_release=lambda x: self.play_sentence_audio())
        ex_header.add_widget(lbl_ex_title)
        ex_header.add_widget(btn_play_example)

        self.lbl_ex_en = CustomLabel(text="", font_size="18sp", bold=True, color=(0.05, 0.05, 0.05, 1), size_hint_y=None)
        self.lbl_ex_en.bind(texture_size=self.lbl_ex_en.setter("size"))

        self.lbl_ex_zh = CustomLabel(text="", font_size="18sp", bold=True, color=(0.25, 0.25, 0.25, 1), size_hint_y=None)
        self.lbl_ex_zh.bind(texture_size=self.lbl_ex_zh.setter("size"))

        ex_box.add_widget(ex_header)
        ex_box.add_widget(self.lbl_ex_en)
        ex_box.add_widget(self.lbl_ex_zh)
        content_box.add_widget(ex_box)

        scroll.add_widget(content_box)
        root.add_widget(scroll)

        # 載入初始單字
        if self.words_data:
            self.display_word(self.current_index)

        return root

    def get_android_excel_path(self):
        """處理 Android 與 PC 的 Excel 檔案路徑與複製機制"""
        filename = "單字.xlsx"  # 建議打包時統一轉為 .xlsx
        
        # 1. 取得使用者可寫入目錄
        user_dir = self.user_data_dir
        target_path = os.path.join(user_dir, filename)

        # 2. 如果目標路徑檔案不存在，由打包資源區複製過去
        if not os.path.exists(target_path):
            source_path = os.path.join(os.path.dirname(__file__), filename)
            if not os.path.exists(source_path):
                # 嘗試 .xlsm
                source_path = os.path.join(os.path.dirname(__file__), "單字.xlsm")
            
            if os.path.exists(source_path):
                try:
                    shutil.copy(source_path, target_path)
                except Exception as e:
                    print(f"複製 Excel 失敗: {e}")
                    return source_path
        
        return target_path if os.path.exists(target_path) else filename

    def load_excel_data(self, path):
        """讀取 Excel 資料"""
        data_list = []
        if os.path.exists(path):
            try:
                wb = openpyxl.load_workbook(path, data_only=True)
                sheet = wb.active
                for row in sheet.iter_rows(values_only=True):
                    clean_row = [str(cell).strip() if cell is not None else "" for cell in row]
                    if any(clean_row):
                        data_list.append(clean_row)
                if data_list and len(data_list[0]) > 0 and (
                    "leason" in data_list[0][0].lower() or "主要" in data_list[0][1]
                ):
                    data_list.pop(0)
                return data_list
            except Exception as e:
                print(f"讀取 Excel 失敗: {e}")
        return []

    def get_col(self, row, idx):
        return row[idx].strip() if idx < len(row) else ""

    def display_word(self, index):
        if not self.words_data or index >= len(self.words_data):
            return

        row = self.words_data[index]
        lesson_name = self.get_col(row, 0)

        current_count = sum(
            1 for idx, r in enumerate(self.words_data)
            if self.get_col(r, 0) == lesson_name and idx <= index
        )
        total_count = sum(
            1 for r in self.words_data if self.get_col(r, 0) == lesson_name
        )

        self.lbl_lesson_info.text = f"{lesson_name}: {current_count}/{total_count}"

        main_word = self.get_col(row, 1)
        main_cn = self.get_col(row, 2)
        raw_ipa = self.get_col(row, 3)
        main_ipa = clean_ipa_text(raw_ipa)
        main_pos = ""

        if main_word:
            for key, pos_label, w_idx, _ in self.pos_order:
                if self.get_col(row, w_idx).lower() == main_word.lower():
                    main_pos = pos_label
                    break

        self.lbl_main_word.text = main_word
        self.lbl_main_cn.text = main_cn
        self.lbl_main_pos.text = main_pos
        self.lbl_main_ipa.text = f"/{main_ipa}/" if main_ipa else ""

        for key, _, w_idx, c_idx in self.pos_order:
            w_val = self.get_col(row, w_idx)
            c_val = self.get_col(row, c_idx)
            btn_en = self.pos_en_buttons[key]
            btn_en.text = w_val if (w_val and w_val != "—") else "—"
            self.pos_cn_widgets[key].text = c_val if c_val else "—"

        self.lbl_ex_en.text = self.get_col(row, 12)
        self.lbl_ex_zh.text = self.get_col(row, 13)
        self.lbl_root_val.text = self.get_col(row, 15)

    def prev_word(self):
        if self.words_data:
            self.current_index = (self.current_index - 1) % len(self.words_data)
            self.display_word(self.current_index)

    def next_word(self):
        if self.words_data:
            self.current_index = (self.current_index + 1) % len(self.words_data)
            self.display_word(self.current_index)

    def play_main_word_audio(self):
        if self.lbl_main_word.text:
            self.speak_text(self.lbl_main_word.text)

    def play_pos_audio(self, pos_key):
        word = self.pos_en_buttons[pos_key].text
        if word and word != "—":
            self.speak_text(word)

    def play_sentence_audio(self):
        if self.lbl_ex_en.text:
            self.speak_text(self.lbl_ex_en.text)

    def speak_text(self, text):
        threading.Thread(
            target=lambda: asyncio.run(self._generate_and_play(text)), daemon=True
        ).start()

    async def _generate_and_play(self, text):
        try:
            # 存成臨時 MP3 檔案
            temp_file = os.path.join(self.user_data_dir, "temp_speech.mp3")
            communicate = edge_tts.Communicate(text, self.voice_en, rate="-15%")
            await communicate.save(temp_file)
            
            # 使用 Kivy 內建的 SoundLoader 播放
            from kivy.core.audio import SoundLoader
            sound = SoundLoader.load(temp_file)
            if sound:
                sound.play()
        except Exception as e:
            print(f"TTS 播放失敗: {e}")


if __name__ == "__main__":
    VocabularyApp().run()
