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

from com.sun.star.task import XJobExecutor

class TranslateText(unohelper.Base, XJobExecutor):
	def __init__(self, ctx):
		'''Конструктор класса'''
		# Сохранение контекста компонента для последующего использования 
		self.ctx = ctx
		self.CONTENT_TYPE = "application/json"
		self.TOKEN = "Api-Key " + token
		self.target_language = 'en'
		self.source_language = 'ru'
		self.url_yandex = 'https://translate.api.cloud.yandex.net/translate/v2/translate'
	
	
	def trigger(self, event):
		'''Обработчик события'''
		# Получение объекта рабочего стола
		desktop = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
		# Получение объекта текущего документа
		document = desktop.getCurrentComponent()
		# Проверка возможности доступа к тексту документа
		# if not hasattr(document, "Text"):
		# 	return
		# controller = document.getCurrentController()
		# select = controller.getSelection()
		
		# cursor = document.getCurrentController().getViewCursor()
		# selected_text = cursor.getString()

		# body = {
		# 	"targetLanguageCode": self.target_language,
		# 	"sourceLanguageCode": self.source_language,
		# 	"texts": selected_text,
		# }

		# headers = {
		# 	"Content-Type": self.CONTENT_TYPE,
		# 	"Authorization": self.TOKEN
		# }

		# response = requests.post(self.url_yandex,
		# 	json = body,
		# 	headers = headers
		# )
		# parsed_string = json.loads(response.text)
		# translated_text = parsed_string['translations'][0]['text']
		
		# # Вставляем переведенный текст вместо исходного
		# cursor.setString(translated_text)
		open_translate_dialog()
		

def _create_translate_dialog():
    BUTTON_WH = 20
    args= {
        'Name': 'translate_window',
        'Title': 'Перевод текста',
        'Width': 200,
        'Height': 220,
    }
    dialog = Window.create_dialog(args)
    #dialog.id = ID_EXTENSION
    #dialog.events = Controllers

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
    dialog.add_control(args)
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
    dialog.add_control(args)
    dialog.center(dialog.translate_but, y=75)

    args = {
        'Type': 'Text',
        'Name': 'text_box',
        'Text': 'hellohello',
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

def open_translate_dialog():
    dialog = _create_translate_dialog()

    dialog.open()
    return

# Регистрация реализации службы
g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation( \
		TranslateText,                                # Имя класса UNO
		"org.openoffice.comp.pyuno.exp.TranslateText",# Имя реализации
		("com.sun.star.task.Job",),)                # Список реализованных служб
