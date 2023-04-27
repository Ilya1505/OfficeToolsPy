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

class RusLinuxJob(unohelper.Base, XJobExecutor):
	def __init__(self, ctx):
		'''Конструктор класса'''
		# Сохранение контекста компонента для последующего использования 
		self.ctx = ctx
	
	def replaceSpecialSymbol(self, error):
		'''Функция замены специальных символов'''
		# Специальные символы и их эквиваленты
		specialsymbols = {'-' : '-', '«' : '"', '»' : '"'}
		# Результирующий фрагмент строки
		outfragment = ''
		# Обход проблемных символов
		for symnum in range(error.end - error.start):
			symbol = str(error.object)[error.start + symnum]
			if (symbol in specialsymbols):
				# Замена символа в том случае, если он известен
				outfragment = outfragment + specialsymbols[symbol]
			else:
				# Замена неизвестного символа на символ пробела
				outfragment = outfragment + ' '
		# Возврат результирующего фрагмента строки и позиции для продолжения перекодировки
		return (outfragment, error.end)
	
	def convertDate(self, date):
		'''Функция преобразования даты в строковый формат'''
		# Список названий месяцев
		months = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря']
		if date is not None:
			try:
				# Преобразование переданной даты
				resdate = str(date.Day) + ' ' + months[date.Month - 1] + ' ' + str(date.Year)  + ' г.'
			except:
				# Преобразование текущей даты 
				date = datetime.datetime.now()
				resdate = str(date.day) + ' ' + months[date.month - 1] + ' ' + str(date.year)  + ' г.'
		else:
			# Преобразование текущей даты
			date = datetime.datetime.now()
			resdate = str(date.day) + ' ' + months[date.month - 1] + ' ' + str(date.year)  + ' г.'
		
		return resdate
	
	def stringEqualsWithoutCase(self, string1, string2):
		'''Функция сравнения строк без учета регистра'''
		try:
			# Сравнение строк в нижнем регистре
			return string1.lower() == string2.lower()
		except AttributeError:
			# Простое сравнение в случае неудачи
			return string1 == string2
	
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
		count = select.getCount()
		for i in range(count) :
			symbol = select.getByIndex(i)
			theString = symbol.getString()
			if len(theString)!=0 :
				#get the XText interface
				text = document.Text
				#create an XTextRange at the end of the document
				tRange = text.End
				#and set the string
				tRange.String = str(count)
		
		
# Регистрация реализации службы
g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation( \
		RusLinuxJob,                                # Имя класса UNO
		"org.openoffice.comp.pyuno.exp.RusLinux",   # Имя реализации
		("com.sun.star.task.Job",),)                # Список реализованных служб
