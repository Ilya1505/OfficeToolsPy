import uno
import unohelper
import Window
import codecs


from com.sun.star.awt.FontWeight import (NORMAL, BOLD)
from com.sun.star.awt.FontUnderline import (SINGLE, NONE)
from com.sun.star.awt.FontSlant import (NONE, ITALIC)




class Validator(unohelper.Base):
    def __init__(self):
        '''Конструктор класса'''
		# Сохранение контекста компонента для последующего использования 

    def validate_rough(self, ctx):
        # Получение объекта рабочего стола
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        
        # Получение объекта текущего документа
        document = desktop.getCurrentComponent()
        
        # Проверка возможности доступа к тексту документа
        if not hasattr(document, "Text"):
            Window.errorbox('Нет доступа к документу', 'Ошибка')
            return
        
        selected_text = ''
        
        # Получение текущей выделенной области текста
        selection = document.CurrentSelection
        if selection.supportsService("com.sun.star.text.TextRanges"):
            # Получение выделенного фрагмента текста
            text_range = selection.getByIndex(0)

            
            # Проверка, что выделенный фрагмент не пуст
            if text_range.String != '':
                # Обход параграфов выделенного фрагмента
                paragraphs = text_range.createEnumeration()
                while paragraphs.hasMoreElements():
                    paragraph = paragraphs.nextElement()
                    
                    # Проверка типа параграфа
                    if paragraph.supportsService("com.sun.star.text.Paragraph"):
                        # Обход фрагментов параграфа
                        elements = paragraph.createEnumeration()
                        while elements.hasMoreElements():
                            element = elements.nextElement()
                            
                            # Гиперссылка
                            if hasattr(element, "HyperLinkURL"):
                                if element.HyperLinkURL != '':
                                    selected_text += ''

                            # Формула
                            if hasattr(element, "Formula"):
                                if element.Formula != '':
                                    selected_text += ''

                        
                            # Ole объекты
                            if hasattr(element, "EmbeddedObject"):
                                if element.EmbeddedObject != '':
                                    selected_text += ''
                            
                            # Рисунок ( не робит)
                            if hasattr(element, "GraphicObject"):
                                graphic_object = element.GraphicObject
                                if graphic_object.supportsService("com.sun.star.text.TextGraphicObject"):
                                    selected_text += ''

                            # Вставка текста                 
                            selected_text += element.String

                              
                            # Закрытие тегов
                            # Гиперссылка
                            if hasattr(element, "HyperLinkURL"):
                                if element.HyperLinkURL != '':
                                    selected_text += ' '

                            
                    # Таблица
                    elif paragraph.supportsService("com.sun.star.text.TextTable"):
                        selected_text += ''
            else:
                Window.errorbox('Нет выделенного текста', 'Ошибка')
        
        return selected_text


   
    def validate_easy(self, ctx):
        # Получение объекта рабочего стола
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        
        # Получение объекта текущего документа
        document = desktop.getCurrentComponent()
        
        # Проверка возможности доступа к тексту документа
        if not hasattr(document, "Text"):
            Window.errorbox('Нет доступа к документу', 'Ошибка')
            return
        
        selected_text = ''
        
        # Получение текущей выделенной области текста
        selection = document.CurrentSelection
        if selection.supportsService("com.sun.star.text.TextRanges"):
            # Получение выделенного фрагмента текста
            text_range = selection.getByIndex(0)

            
            # Проверка, что выделенный фрагмент не пуст
            if text_range.String != '':
                # Обход параграфов выделенного фрагмента
                paragraphs = text_range.createEnumeration()
                while paragraphs.hasMoreElements():
                    paragraph = paragraphs.nextElement()
                    
                    # Проверка типа параграфа
                    if paragraph.supportsService("com.sun.star.text.Paragraph"):
                        # Обход фрагментов параграфа
                        elements = paragraph.createEnumeration()
                        while elements.hasMoreElements():
                            element = elements.nextElement()
                            
                            # Гиперссылка
                            if hasattr(element, "HyperLinkURL"):
                                if element.HyperLinkURL != '':
                                    selected_text += ''

                            # Формула
                            if hasattr(element, "Formula"):
                                if element.Formula != '':
                                    selected_text += '<formula>'

                            # Ole объекты
                            if hasattr(element, "EmbeddedObject"):
                                if element.EmbeddedObject != '':
                                    selected_text += '<EmbeddedObject>'
                            
                            # Рисунок ( не робит)
                            if hasattr(element, "GraphicObject"):
                                graphic_object = element.GraphicObject
                                if graphic_object.supportsService("com.sun.star.text.TextGraphicObject"):
                                    selected_text += '<image>'
                                   
                            
                            # Список (не робит, но есть идея как сделать)
                            if hasattr(element, "ParaListLevel"):
                                selected_text += '<list>'
                            
                            
                            selected_text += element.String
                            
                            # Закрытие тегов форматирования
                            # Гиперссылка
                            if hasattr(element, "HyperLinkURL"):
                                if element.HyperLinkURL != '':
                                    selected_text += ' '

                            
                    # Таблица
                    elif paragraph.supportsService("com.sun.star.text.TextTable"):
                        selected_text += '\n<table>\n'
                    
            else:
                Window.errorbox('Нет выделенного текста', 'Ошибка')
        
        return selected_text


  



