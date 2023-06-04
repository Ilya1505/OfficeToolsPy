import uno
import unohelper
import requests
import json
import Window

from directory import token_translate
from com.sun.star.awt.FontWeight import (NORMAL, BOLD)
from com.sun.star.awt.FontUnderline import (SINGLE, NONE)
from com.sun.star.awt.FontSlant import (NONE, ITALIC)

from com.sun.star.awt import XActionListener

from com.sun.star.task import XJobExecutor

# главный класс перевода текста
class Translate_Text( unohelper.Base, XActionListener ):
    def __init__(self, ctx):
        self.ctx = ctx# контекст документа
        self.CONTENT_TYPE = "application/json"
        self.TOKEN = "Api-Key " + token_translate
        self.source_language = 'ru'
        self.target_language = 'en'
        self.selected_text = ''# выделенный текст
        self.translated_text = ''# переведенный текст
        self.url_yandex = 'https://translate.api.cloud.yandex.net/translate/v2/translate'
        self.dialog = None

    # слушатель на событие нажатия кнопки
    def actionPerformed(self, actionEvent):
        if actionEvent.Source.Model.Name == 'change_but':
            self.change_language()
        else:
            self.translate()

    # метод смены языка перевода
    def change_language(self):
        if self.source_language == 'ru':
            self.source_language = 'en'
            self.target_language = 'ru'
            self.dialog.language_label.Label = 'Английский -> Русский'
        else:
            self.source_language = 'ru'
            self.target_language = 'en'
            self.dialog.language_label.Label = 'Русский -> Английский'

    # метод перевода
    def translate(self):
        # Получение объекта рабочего стола
        desktop = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
        # Получение объекта текущего документа
        document = desktop.getCurrentComponent()
        # Проверка возможности доступа к тексту документа
        if not hasattr(document, "Text"):
            Window.errorbox('Нет доступа к документу', 'Ошибка')
            return
        
        cursor = document.getCurrentController().getViewCursor()
        self.selected_text = cursor.getString()
        ###########validator###############
        if self.selected_text == '':
            Window.errorbox('Нет выделенного текста', 'Ошибка')
            return
        self.translated_text = self.request(self.selected_text)# метод запроса
        if self.translated_text == '':
            Window.errorbox('Ошибка перевода, проверьте подключение к сети или попробуйте позже', 'Ошибка')
            return
        # Вставляем переведенный текст
        self.dialog.text_box.Text = self.translated_text
        return
    
    # метод post-запроса
    def request(self, text):
        body = {
        	"targetLanguageCode": self.target_language,
        	"sourceLanguageCode": self.source_language,
        	"texts": text,
        }

        headers = {
        	"Content-Type": self.CONTENT_TYPE,
        	"Authorization": self.TOKEN
        }
        # post-запрос на сервис перевода
        try:
            response = requests.post(self.url_yandex,
        	    json = body,
        	    headers = headers
            )
        except:
            return ''
        if response.status_code != 200:
            return ''
        
        parsed_string = json.loads(response.text)# парсинг строки в json
        translated_text = parsed_string['translations'][0]['text']# выборка переведенного текста
        return translated_text
 
# класс обертка, инициализиурющий класс подсистемы
class Translate_Context(unohelper.Base, XJobExecutor):
    
    #конструктор класса
    def __init__(self, ctx):
        self.translate = Translate_Text(ctx)
        self.dialog = self.create_translate_dialog()
        self.translate.dialog = self.dialog
    
    #обработчик события
    def trigger(self, event):
        self.dialog.open()

    #создание пользовательской формы
    def create_translate_dialog(self):
        args= {
            'Name': 'translate_window',
            'Title': 'Перевод текста',
            'Width': 200,
            'Height': 220,
        }
        dialog = Window.create_dialog(args)

        args = {
            'Type': 'Label',
            'Name': 'text_label',
            'Width': 150,
            'Height': 20,
            'Align': 1,
            'VerticalAlign': 1,
            'Step': 10,
            'FontHeight': 16,
            'Label': 'Текст: выделенный фрагмент',
        }
        dialog.add_control(args)# добавление компонентов на форму
        dialog.center(dialog.text_label, y=5)# центрирование компонентов

        args = {
            'Type': 'Label',
            'Name': 'language_label',
            'Width': 150,
            'Height': 20,
            'Align': 1,
            'VerticalAlign': 1,
            'Step': 10,
            'FontHeight': 16,
            'Label': 'Русский -> Английский',
        }
        dialog.add_control(args)
        dialog.center(dialog.language_label, y=25)

        args = {
            'Type': 'Button',
            'Name': 'change_but',
            'Label': 'Поменять язык',
            'Width': 100,
            'Height': 20,
            'Step': 10,
            'Align': 1,
            'FontHeight': 16,
        }
        ctl = dialog.add_control(args)
        ctl.addActionListener(self.translate)
        dialog.center(dialog.change_but, y=50)

        args = {
            'Type': 'Button',
            'Name': 'translate_but',
            'Label': 'Перевести',
            'Width': 100,
            'Height': 20,
            'Step': 10,
            'Align': 1,
            'FontHeight': 16,
        }
        ctl = dialog.add_control(args)
        ctl.addActionListener(self.translate)# добавление слушателей на кнопку
        dialog.center(dialog.translate_but, y=75)

        args = {
            'Type': 'Text',
            'Name': 'text_box',
            'Text': '',
            'Width': 190,
            'Height': 110,
            'Align': 0,
            'VerticalAlign': 0,
            'Step': 10,
            'FontHeight': 14,
            'TextColor': Window.get_color('black'),
            'AutoVScroll': True,
            'BackgroundColor': Window.get_color('white'),
            'MultiLine': True,
        }
        dialog.add_control(args)
        dialog.center(dialog.text_box, y=-10)
        return dialog

# Регистрация реализации службы
g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation( \
		Translate_Context,                                # Имя класса UNO
		"org.openoffice.comp.pyuno.exp.Translate_Context",# Имя реализации
		("com.sun.star.task.Job",),)                      # Список реализованных служб
