import uno
import unohelper
import Window
from string import punctuation #сборник символов пунктуации
from razdel import tokenize
from directory import stopwords, token_synonym, url_synonym
from pymystem3 import Mystem
import requests
import json
from Validator import *

from com.sun.star.awt.FontWeight import (NORMAL, BOLD)
from com.sun.star.awt.FontUnderline import (SINGLE, NONE)
from com.sun.star.awt.FontSlant import (NONE, ITALIC)
from com.sun.star.awt import XActionListener
from com.sun.star.task import XJobExecutor

#главный класс нормализации текста и поиска синонимов
class Analyze_Text( unohelper.Base, XActionListener ):
    #конструктор класса
    def __init__(self, ctx):
        self.ctx = ctx# контекст документа
        self.dialog = None
        self.freq_lemmas = None# словарь слов с их количеством вхождений
        self.selected_text = ''# выделенный текст
        self.validator = Validator()

    # слушатель события нажатия на кнопку
    def actionPerformed(self, actionEvent):
        if actionEvent.Source.Model.Name == 'analyze_text_but':
            self.text_normalization()
        else:
            self.search_synonyms()
        return
    
    # метод нормализации текста
    def text_normalization(self):
        self.selected_text = self.validator.validate_rough(self.ctx)
        if self.selected_text == '':
            return
        self.selected_text = self.selected_text.lower()# перевод текста в нижний регистр
        token_text = list(tokenize(self.selected_text))# токенизация текста
        punct = stopwords + [p for p in punctuation]# список стоп-слов и знаков пунктуации
        # удаление стоп-слов и знаков пунктуации
        clean_token = [word.text for word in token_text if word.text not in punct]
        clean_token = ' '.join(clean_token)
        mystem = Mystem()# объект для лемматизации слов
        try:
            lemmas = mystem.lemmatize(clean_token)# лемматизация слов
        except:
            Window.errorbox('Ошибка лемматизации, проверьте исходный текст', 'Ошибка')
            return
        punct = [' ', '\n']
        # удаление пробелов и символов новой строки
        lemmas = [word for word in lemmas if word not in punct]
        if len(lemmas) == 0:
            Window.errorbox('Не найдено подходящих слов', 'Ошибка')
            return

        # подсчет числа вхождений каждого слова в текст и сортировка по убыванию
        self.freq_lemmas = self.sorted_quantity(self.counter_word(lemmas))
        # вывод списка часто встречаымх слов в интерфейс пользователя
        count = 0
        for lem in self.freq_lemmas:
            self.dialog.source_word_list.insert(lem + " ( " + str(self.freq_lemmas.get(lem)) + " )")
            count+=1
            if count > 10:
                break
        self.dialog.source_word_list.select()
        self.show_synonyms_control()# вывод компонентов формы
        return

    # подсчет числа вхождений каждого слова в текст    
    def counter_word(self, lemmas):
        freq_word = {}
        for word in lemmas:
            if word in freq_word:
                freq_word.update({word: freq_word.get(word)+1})
            else:
                freq_word[word] = 1
        return freq_word

    # сортировка по убыванию
    def sorted_quantity(self, dictionary):
        return dict(sorted(dictionary.items(), key=lambda value: value[1], reverse=True))

    # метод поиска синонимов    
    def search_synonyms(self):
        self.dialog.synonyms_word_list.clear()# очистка списка синонимов
        # получение выбранного слова
        select_word = self.get_select_lemm(self.dialog.source_word_list.value)
        # пост-запрос на сервис поиска синонимов
        synonyms = self.post_reguest(select_word)
        # если синонимы не найдены
        if len(synonyms) == 1:
            Window.msgbox('Для выбранного слова синонимы не найдены', 'Результат')
            return
        # вывод всех синонимов списком в интерфейс пользователя
        for word in synonyms:
            self.dialog.synonyms_word_list.insert(word)
        self.dialog.synonyms_word_list.select()
        return
    
    # получение выбранного слова
    def get_select_lemm(self, lemmas):
        return lemmas.split(' ')[0]

    # метод post-запроса на сервис поиска синонимов
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

        # post-запрос
        try:
            response = requests.post(url_synonym, json_synonym)
        except:
            Window.errorbox('Проверьте подключение к сети или попробуйте позже', 'Ошибка соединения')
            return
        if response.status_code != 200:
            Window.errorbox('Проверьте подключение к сети или попробуйте позже', 'Ошибка')
            return
        parsed_string = json.loads(response.text)# парсинг строки в json
        synonyms = parsed_string['response']['1']['syns']# выборка всех синонимов
        return synonyms
    
    # вывод компонентов формы
    def show_synonyms_control(self):
        self.dialog.frequence_label.visible = True
        self.dialog.source_word_list.visible = True
        self.dialog.search_synonyms_but.visible = True
        self.dialog.synonyms_word_list.visible = True
        self.dialog.analyze_text_but.visible = False


# класс обертка, инициализирующий класс подсистемы
class Analyze_Context(unohelper.Base, XJobExecutor):
    # конструктор класса
    def __init__(self, ctx):
        self.analyze = Analyze_Text(ctx)# создание класса обработки текста и поиска синонимов
        self.dialog = self.create_analyze_dialog()# создание пользовательской формы
        self.analyze.dialog = self.dialog
	
    # обработчик события
    def trigger(self, event):
        self.dialog.open()

    # создание пользовательской формы
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
        dialog.add_control(args)# добавление компонентов на форму
        dialog.center(dialog.text_label, y=5)# центрирование компонентов

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
        ctl.addActionListener(self.analyze)# добавление слушателей на кнопку
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

