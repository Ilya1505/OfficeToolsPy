from __future__ import unicode_literals
import io
import html
import os.path
import codecs

import uno
import unohelper
import requests
import json
import Window

from token_yandex_api import token
from com.sun.star.awt.FontWeight import (NORMAL, BOLD)
from com.sun.star.awt.FontUnderline import (SINGLE, NONE)
from com.sun.star.awt.FontSlant import (NONE, ITALIC)

from com.sun.star.awt import XActionListener

from com.sun.star.task import XJobExecutor

class Translate_Text( unohelper.Base, XActionListener ):
    def __init__(self, ctx):
        self.ctx = ctx
        self.CONTENT_TYPE = "application/json"
        self.TOKEN = "Api-Key " + token
        self.source_language = 'ru'
        self.target_language = 'en'
        self.url_yandex = 'https://translate.api.cloud.yandex.net/translate/v2/translate'
        self.dialog = None

    def actionPerformed(self, actionEvent):
        if actionEvent.Source.Model.Name == 'change_but':
            self.change_language()
        else:
            self.translate()

    def change_language(self):
        if self.source_language == 'ru':
            self.source_language = 'en'
            self.target_language = 'ru'
            self.dialog.language_label.Label = 'Английский -> Русский'
        else:
            self.source_language = 'ru'
            self.target_language = 'en'
            self.dialog.language_label.Label = 'Русский -> Английский'

    def translate(self):
        # Получение объекта рабочего стола
        desktop = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
        # Получение объекта текущего документа
        document = desktop.getCurrentComponent()
        # Проверка возможности доступа к тексту документа
        if not hasattr(document, "Text"):
            Window.errorbox('Нет доступа к документу', 'Ошибка')
            return
        # controller = document.getCurrentController()
        # select = controller.getSelection()
        
        cursor = document.getCurrentController().getViewCursor()
        selected_text = cursor.getString()
        if selected_text == '':
            Window.errorbox('Нет выделенного текста', 'Ошибка')
            return
        
        # if self.dialog.text_box.Text != '':
        #     Window.errorbox('Переведенный текст уже существует', 'Ошибка')
        #     return
        body = {
        	"targetLanguageCode": self.target_language,
        	"sourceLanguageCode": self.source_language,
        	"texts": selected_text,
        }

        headers = {
        	"Content-Type": self.CONTENT_TYPE,
        	"Authorization": self.TOKEN
        }

        response = requests.post(self.url_yandex,
        	json = body,
        	headers = headers
        )
        
        if response.status_code != 200:
            Window.errorbox('Ошибка перевода, попробуйте позже', 'Ошибка')
            return
        
        parsed_string = json.loads(response.text)
        translated_text = parsed_string['translations'][0]['text']
        
        # Вставляем переведенный текст вместо исходного
        self.dialog.text_box.Text = translated_text
 
class Translate_Context(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        '''Конструктор класса'''
        # Сохранение контекста компонента для последующего использования 
        self.translate = Translate_Text(ctx)
        self.dialog = self.create_translate_dialog()
        self.translate.dialog = self.dialog
    
    def trigger(self, event):
        '''Обработчик события'''
        self.dialog.open()

    def create_translate_dialog(self):
        BUTTON_WH = 20
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
        dialog.add_control(args)
        dialog.center(dialog.text_label, y=5)

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
        ctl.addActionListener(self.translate)
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
        # dialog.lst_log.visible = False

        return dialog

# Регистрация реализации службы
g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation( \
		Translate_Context,                                # Имя класса UNO
		"org.openoffice.comp.pyuno.exp.Translate_Context",# Имя реализации
		("com.sun.star.task.Job",),)                # Список реализованных служб
