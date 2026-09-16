import os
import shutil
import openpyxl
from plyer import tts

from kivy.app import App
from kivy.core.window import Window
from kivy.core.text import LabelBase
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView

Window.clearcolor = (0.95, 0.95, 0.96, 1)

def init_custom_font():
    font_paths = [
        "/system/fonts/NotoSansCJK-Regular.ttc",
        "/system/fonts/DroidSansFallback.ttf",
        "C:/Windows/Fonts/msjh.ttc"
    ]
    for font_path in font_paths:
        if os.path.exists(font_path):
            try:
                LabelBase.register(name="Roboto", fn_regular=font_path)
                return
            except Exception:
                pass

init_custom_font()

class CustomLabel(Label):
    def __init__(self, **kwargs):
        kwargs.setdefault('font_name', 'Roboto')
        super().__init__(**kwargs)
        self.bind(size=self._update_text_size)

    def _update_text_size(self, *args):
        self.text_size = (self.width, None)

class VocabularyApp(App):

    def build(self):
        self.title = "單字學習機"
        self.current_index = 0

        self.excel_path = self.get_android_excel_path()
        self.words_data = self.load_excel_data(self.excel_path)

        self.pos_order = [
            ("V", "V.", 4, 5),
            ("N", "N.", 6, 7),
            ("Adj", "Adj.", 8, 9),
            ("Adv", "Adv.", 10, 11),
        ]

        root = BoxLayout(orientation="vertical", padding=15, spacing=15)

        # 頂部導覽列
        top_nav = BoxLayout(orientation="horizontal", size_hint_y=None, height="50dp", spacing=10)
        btn_prev = Button(text="← Prev", font_size="18sp", font_name="Roboto")
        btn_prev.bind(on_release=lambda x: self.prev_word())
        
        self.lbl_lesson_info = Label(text="", font_size="18sp", font_name="Roboto", color=(0.1, 0.1, 0.1, 1))
        
        btn_next = Button(text="Next →", font_size="18sp", font_name="Roboto")
        btn_next.bind(on_release=lambda x: self.next_word())

        top_nav.add_widget(btn_prev)
        top_nav.add_widget(self.lbl_lesson_info)
        top_nav.add_widget(btn_next)
        root.add_widget(top_nav)

        scroll = ScrollView(size_hint=(1, 1))
        content_box = BoxLayout(orientation="vertical", spacing=15, size_hint_y=None)
        content_box.bind(minimum_height=content_box.setter("height"))

        # 主要單字
        top_word_box = BoxLayout(orientation="horizontal", size_hint_y=None, height="60dp", spacing=10)
        self.lbl_main_word = Label(text="", font_size="32sp", font_name="Roboto", bold=True, color=(0.85, 0.1, 0.1, 1))
        self.lbl_main_cn = Label(text="", font_size="20sp", font_name="Roboto", color=(0.1, 0.1, 0.1, 1))
        
        btn_main_pronounce = Button(text="發音", font_size="18sp", font_name="Roboto", size_hint_x=None, width="80dp")
        btn_main_pronounce.bind(on_release=lambda x: self.speak_text(self.lbl_main_word.text))

        top_word_box.add_widget(self.lbl_main_word)
        top_word_box.add_widget(self.lbl_main_cn)
        top_word_box.add_widget(btn_main_pronounce)
        content_box.add_widget(top_word_box)

        # 詞性列表
        table_grid = GridLayout(cols=3, size_hint_y=None, spacing=10)
        table_grid.bind(minimum_height=table_grid.setter("height"))
        self.pos_en_buttons = {}
        self.pos_cn_widgets = {}

        for key, label_text, _, _ in self.pos_order:
            lbl_pos = Label(text=label_text, font_size="18sp", font_name="Roboto", color=(0, 0, 0, 1), size_hint_y=None, height="45dp")
            btn_en = Button(text="—", font_size="18sp", font_name="Roboto", size_hint_y=None, height="45dp")
            btn_en.bind(on_release=lambda btn, k=key: self.speak_text(self.pos_en_buttons[k].text))
            lbl_cn = Label(text="—", font_size="18sp", font_name="Roboto", color=(0, 0, 0, 1), size_hint_y=None, height="45dp")

            table_grid.add_widget(lbl_pos)
            table_grid.add_widget(btn_en)
            table_grid.add_widget(lbl_cn)
            self.pos_en_buttons[key] = btn_en
            self.pos_cn_widgets[key] = lbl_cn

        content_box.add_widget(table_grid)

        # 例句
        ex_box = BoxLayout(orientation="vertical", size_hint_y=None, spacing=10)
        ex_box.bind(minimum_height=ex_box.setter("height"))
        
        btn_play_example = Button(text="播放例句", font_size="16sp", font_name="Roboto", size_hint_y=None, height="40dp")
        btn_play_example.bind(on_release=lambda x: self.speak_text(self.lbl_ex_en.text))
        
        self.lbl_ex_en = CustomLabel(text="", font_size="18sp", color=(0, 0, 0, 1), size_hint_y=None)
        self.lbl_ex_zh = CustomLabel(text="", font_size="18sp", color=(0.3, 0.3, 0.3, 1), size_hint_y=None)

        ex_box.add_widget(btn_play_example)
        ex_box.add_widget(self.lbl_ex_en)
        ex_box.add_widget(self.lbl_ex_zh)
        content_box.add_widget(ex_box)

        scroll.add_widget(content_box)
        root.add_widget(scroll)

        if self.words_data:
            self.display_word(self.current_index)

        return root

    def get_android_excel_path(self):
        filename = "單字.xlsx"
        target_path = os.path.join(self.user_data_dir, filename)
        if not os.path.exists(target_path):
            source_path = os.path.join(os.path.dirname(__file__), filename)
            if os.path.exists(source_path):
                try:
                    shutil.copy(source_path, target_path)
                except Exception:
                    pass
        return target_path if os.path.exists(target_path) else filename

    def load_excel_data(self, path):
        data_list = []
        if os.path.exists(path):
            try:
                wb = openpyxl.load_workbook(path, data_only=True)
                sheet = wb.active
                for row in sheet.iter_rows(values_only=True):
                    clean_row = [str(cell).strip() if cell is not None else "" for cell in row]
                    if any(clean_row):
                        data_list.append(clean_row)
                if data_list and ("leason" in data_list[0][0].lower() or "主要" in data_list[0][1]):
                    data_list.pop(0)
                return data_list
            except Exception:
                pass
        return []

    def get_col(self, row, idx):
        return row[idx].strip() if idx < len(row) else ""

    def display_word(self, index):
        if not self.words_data:
            return
        row = self.words_data[index]
        self.lbl_lesson_info.text = f"進度: {index + 1}/{len(self.words_data)}"
        self.lbl_main_word.text = self.get_col(row, 1)
        self.lbl_main_cn.text = self.get_col(row, 2)

        for key, _, w_idx, c_idx in self.pos_order:
            self.pos_en_buttons[key].text = self.get_col(row, w_idx) or "—"
            self.pos_cn_widgets[key].text = self.get_col(row, c_idx) or "—"

        self.lbl_ex_en.text = self.get_col(row, 12)
        self.lbl_ex_zh.text = self.get_col(row, 13)

    def prev_word(self):
        if self.words_data:
            self.current_index = (self.current_index - 1) % len(self.words_data)
            self.display_word(self.current_index)

    def next_word(self):
        if self.words_data:
            self.current_index = (self.current_index + 1) % len(self.words_data)
            self.display_word(self.current_index)

    def speak_text(self, text):
        if text and text.strip() and text != "—":
            try:
                tts.speak(text.strip())
            except Exception as e:
                print(f"TTS 失敗: {e}")

if __name__ == "__main__":
    VocabularyApp().run()
