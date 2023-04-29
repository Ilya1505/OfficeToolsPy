import io
import html
import os.path
import codecs

import uno
import unohelper

from com.sun.star.awt.FontWeight import (NORMAL, BOLD)
from com.sun.star.awt.FontUnderline import (SINGLE, NONE)
from com.sun.star.awt.FontSlant import (NONE, ITALIC)

from com.sun.star.task import XJobExecutor

class TranslateText(unohelper.Base, XJobExecutor):
	def __init__(self, ctx):
		'''Конструктор класса'''
		# Сохранение контекста компонента для последующего использования 
		self.ctx = ctx
	
	
	def trigger(self, event):
		'''Обработчик события'''
		# Получение объекта рабочего стола
		desktop = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
		# Получение объекта текущего документа
		document = desktop.getCurrentComponent()
		# Проверка возможности доступа к тексту документа
		if not hasattr(document, "Text"):
			return
		controller = document.getCurrentController()
		select = controller.getSelection()
		
		cursor = document.getCurrentController().getViewCursor()
		selected_text = cursor.getString()
		# # Задаем язык и направление перевода
		# src_lang = 'en'
		# dst_lang = 'ru'

		# # Используем сервис перевода Яндекса
		# translate_url = 'https://translate.yandex.net/api/v1.5/tr/translate'
		# translate_key = 'YOUR_API_KEY'
		# params = {
		# 	'key': translate_key,
		# 	'text': selected_text,
		# 	'lang': f'{src_lang}-{dst_lang}'
		# }
		# response = requests.get(translate_url, params=params)

		# # Извлекаем переведенный текст из ответа
		# xml_content = ElementTree.fromstring(response.content)
		# translated_text = xml_content[0][0].text

		# # Вставляем переведенный текст вместо исходного
		cursor.setString("'" + selected_text + "'")

		
# Регистрация реализации службы
g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation( \
		TranslateText,                                # Имя класса UNO
		"org.openoffice.comp.pyuno.exp.TranslateText",# Имя реализации
		("com.sun.star.task.Job",),)                # Список реализованных служб
