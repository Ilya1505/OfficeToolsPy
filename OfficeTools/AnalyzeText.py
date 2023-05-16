import uno
import unohelper
import Window
from string import punctuation #сборник символов пунктуации
from razdel import tokenize
from directory import stopwords, token_synonym, url_synonym
from pymystem3 import Mystem
import requests
import json

from com.sun.star.awt.FontWeight import (NORMAL, BOLD)
from com.sun.star.awt.FontUnderline import (SINGLE, NONE)
from com.sun.star.awt.FontSlant import (NONE, ITALIC)
from com.sun.star.awt import XActionListener
from com.sun.star.task import XJobExecutor

class Analyze_Text( unohelper.Base, XActionListener ):
    def __init__(self, ctx):
        # Сохранение контекста компонента для последующего использования 
        self.ctx = ctx
        self.dialog = None
        self.freq_lemmas = None

    def actionPerformed(self, actionEvent):
        if actionEvent.Source.Model.Name == 'analyze_text_but':
            self.text_normalization()
        else:
            self.search_synonyms()
        return
    
    def text_normalization(self):
        # Получение объекта рабочего стола
        desktop = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
        # Получение объекта текущего документа
        document = desktop.getCurrentComponent()
        # Проверка возможности доступа к тексту документа
        if not hasattr(document, "Text"):
            Window.errorbox('Нет доступа к документу', 'Ошибка')
            return
        
        cursor = document.getCurrentController().getViewCursor()
        selected_text = cursor.getString()
        ##########validator###############
        if selected_text == '':
            Window.errorbox('Нет выделенного текста', 'Ошибка')
            return
        selected_text = selected_text.lower()
        token_text = list(tokenize(selected_text))
        punct = stopwords + [p for p in punctuation]
        clean_token = [word.text for word in token_text if word.text not in punct]
        clean_token = ' '.join(clean_token)
        mystem = Mystem()
        try:
            lemmas = mystem.lemmatize(clean_token)
        except:
            Window.errorbox('Ошибка лемматизации, проверьте исходный текст', 'Ошибка')
            return
        punct = [' ', '\n']
        lemmas = [word for word in lemmas if word not in punct]
        if len(lemmas) == 0:
            Window.errorbox('Не найдено подходящих слов', 'Ошибка')
            return

        self.freq_lemmas = self.sorted_quantity(self.counter_word(lemmas))
        count = 0
        for lem in self.freq_lemmas:
            self.dialog.source_word_list.insert(lem + " ( " + str(self.freq_lemmas.get(lem)) + " )")
            count+=1
            if count > 10:
                break
        self.dialog.source_word_list.select()
        self.show_synonyms_control()
        return
        
    def counter_word(self, lemmas):
        freq_word = {}
        for word in lemmas:
            if word in freq_word:
                freq_word.update({word: freq_word.get(word)+1})
            else:
                freq_word[word] = 1
        return freq_word

    def sorted_quantity(self, dictionary):
        return dict(sorted(dictionary.items(), key=lambda value: value[1], reverse=True))
        
    def search_synonyms(self):
        self.dialog.synonyms_word_list.clear()
        select_word = self.get_select_lemm(self.dialog.source_word_list.value)
        synonyms = self.post_reguest(select_word)
        if len(synonyms) == 1:
            Window.msgbox('Для выбранного слова синонимы не найдены', 'Результат')
            return
        for word in synonyms:
            self.dialog.synonyms_word_list.insert(word)
        self.dialog.synonyms_word_list.select()
        return
    
    def get_select_lemm(self, lemmas):
        return lemmas.split(' ')[0]

    def post_reguest(self, word):
        
        json_synonym = {
            'token': token_synonym, 
            'c': 'syns',
            'query': word,
            'top': 10,
            'lang': 'ru',
            'format': 'json',
            'forms': 0,
            'scores': 0,
        }

        try:
            response = requests.post(url_synonym, json_synonym)
        except:
            Window.errorbox('Проверьте подключение к сети или попробуйте позже', 'Ошибка соединения')
            return
        if response.status_code != 200:
            Window.errorbox('Проверьте подключение к сети или попробуйте позже', 'Ошибка')
            return
        parsed_string = json.loads(response.text)
        synonyms = parsed_string['response']['1']['syns']
        return synonyms
    
    def show_synonyms_control(self):
        self.dialog.frequence_label.visible = True
        self.dialog.source_word_list.visible = True
        self.dialog.search_synonyms_but.visible = True
        self.dialog.synonyms_word_list.visible = True
        self.dialog.analyze_text_but.visible = False



class Analyze_Context(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        '''Конструктор класса'''
        self.analyze = Analyze_Text(ctx)
        self.dialog = self.create_analyze_dialog()
        self.analyze.dialog = self.dialog
	

    def trigger(self, event):
        '''Обработчик события'''
        self.dialog.open()

    def create_analyze_dialog(self):
        args= {
            'Name': 'analyze_window',
            'Title': 'Поиск синонимов',
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
            'Name': 'frequence_label',
            'Width': 150,
            'Height': 20,
            'Align': 1,
            'VerticalAlign': 1,
            'Step': 10,
            'FontHeight': 16,
            'Label': 'Чаще всего встречаются слова:',
        }
        dialog.add_control(args)
        dialog.center(dialog.frequence_label, y=25)
        dialog.frequence_label.visible = False

        args = {
            'Type': 'Listbox',
            'Name': 'source_word_list',
            'Width': 100,
            'Height': 45,
            'Step': 10,
            'FontHeight': 14,
        }
        dialog.add_control(args)
        dialog.center(dialog.source_word_list, y=50)
        dialog.source_word_list.visible = False
        
        args = {
            'Type': 'Button',
            'Name': 'search_synonyms_but',
            'Label': 'Найти синонимы',
            'Width': 100,
            'Height': 20,
            'Step': 10,
            'Align': 1,
            'FontHeight': 16,
        }
        ctl = dialog.add_control(args)
        ctl.addActionListener(self.analyze)
        dialog.center(dialog.search_synonyms_but, y=105)
        dialog.search_synonyms_but.visible = False

        args = {
            'Type': 'Listbox',
            'Name': 'synonyms_word_list',
            'Width': 100,
            'Height': 45,
            'Step': 10,
            'FontHeight': 14,
        }
        dialog.add_control(args)
        dialog.center(dialog.synonyms_word_list, y=135)
        dialog.synonyms_word_list.visible = False

        args = {
            'Type': 'Button',
            'Name': 'analyze_text_but',
            'Label': 'Произвести анализ',
            'Width': 100,
            'Height': 20,
            'Step': 10,
            'Align': 1,
            'FontHeight': 16,
        }
        ctl = dialog.add_control(args)
        ctl.addActionListener(self.analyze)
        dialog.center(dialog.analyze_text_but, y=50)

        return dialog

# Регистрация реализации службы
g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation( \
        Analyze_Context,                                # Имя класса UNO
        "org.openoffice.comp.pyuno.exp.Analyze_Context",# Имя реализации
        ("com.sun.star.task.Job",),)                    # Список реализованных служб

