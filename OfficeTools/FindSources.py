from bs4 import BeautifulSoup
import uno
import unohelper
import requests
import json
import Window
import webbrowser
import time
from Validator import *

from directory import token_google_search, token_google_id


from com.sun.star.awt.FontWeight import (NORMAL, BOLD)
from com.sun.star.awt.FontUnderline import (SINGLE, NONE)
from com.sun.star.awt.FontSlant import (NONE, ITALIC)

from com.sun.star.awt import XActionListener

from com.sun.star.task import XJobExecutor


class Find_Sources( unohelper.Base, XActionListener ):
    def __init__(self, ctx):
        self.ctx = ctx
        self.api_key = token_google_search
        self.cx = token_google_id
        self.url = "https://www.google.com/search?q="
        self.user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
        self.dialog = None
        self.list_of_links = None
        

            
  
    def actionPerformed(self, actionEvent):
        if actionEvent.Source.Model.Name == 'find_source_but':
            validator = Validator()
            selected_text = validator.validate_rough(self.ctx)
        
            if selected_text == '':
                return
            
            selected_text = selected_text.replace(" ", "+")

            self.list_of_links = self.search_api(selected_text)
            
            count = 0
            for title in self.list_of_links:
                self.dialog.link_list.insert(title)
                count+=1
                if count > 30:
                    break
            
            self.dialog.link_list.select()
            self.show_source_control()
  
        else:
            select_link = self.list_of_links[self.dialog.link_list.value]
            webbrowser.open(select_link)

    def show_source_control(self):
        self.dialog.link_list.visible = True
        self.dialog.text_label.visible = False
        self.dialog.find_source_but.visible = False
        self.dialog.follow_link_but.visible = True

    def search_api(self, text):
        if not text:
            print("Ошибка: пустой текст")
            Window.errorbox('Ошибка: пустой текст')
            return []

        url = f"https://www.googleapis.com/customsearch/v1"
        params = {
            "key": self.api_key,
            "cx": self.cx,
            "q": text
        }

        try:
            response = requests.get(url, params=params)
            response.raise_for_status()
            data = response.json()

            extracted_results = {}
            if "items" in data:
                for item in data["items"]:
                    link = item["link"]
                    title = item["title"]
                    extracted_results[title] = link

            return extracted_results

        except requests.exceptions.RequestException as e:
            print("Произошла ошибка при запросе к API:", str(e))
        except KeyError:
            print("Ошибка при обработке результатов поиска.")
        except json.JSONDecodeError:
            print("Ошибка при декодировании JSON-ответа.")

        return {}



    def search(self, text):

        # формируем URL для поиска в Google
        self.url += text

        headers = {
            "User-Agent": self.user_agent
        }

        try:
            # отправляем GET-запрос на сервер Google
            response = requests.get(self.url, headers=headers)

        except ConnectionError:
            # Обработка исключения ConnectionError
            Window.errorbox('Ошибка подключения')
            

        except Exception as e:

            # Обработка других исключений
            Window.errorbox('Произошла ошибка' + str(e))

        else:
            # Work -----------------------------------
            # Блок кода, который будет выполнен, если исключение не возникло
            # парсим HTML-код страницы с помощью BeautifulSoup
            
            
            soup = BeautifulSoup(response.text, "html.parser")
            
            # ищем все ссылки на странице
            links = soup.find_all("a")

            list = []
            with open("file.txt", "w") as file:
            # Записываем текст в файл
                file.write('Test\n')



            # выводим все найденные ссылки
            for link in links:
                href = link.get('href')
                if href and not ('support.google.com' in href or 'policies.google.com' in href or '/search?q=' in href or 'www.google.com' in href or '#' in href or 'maps.google.com' in href or '/?sa=' in href	 or '/advanced_search' in href):
                    href = href.replace('/imgres?imgurl=', '')
                    href = href.replace('/url?esrc=s&q=&rct=j&sa=U&url=', '')
                    with open("file.txt", "a") as file:
                    # Записываем текст в файл
                        file.write(href + '\n')
                    list.append(href)

            return list


class Find_Sources_Context(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        '''Конструктор класса'''
        # Сохранение контекста компонента для последующего использования 
        self.findSource = Find_Sources(ctx)
        self.dialog = self.create_link_dialog()
        self.findSource.dialog = self.dialog
    
    def trigger(self, event):
        '''Обработчик события'''
        self.dialog.open()

    def create_link_dialog(self):
        BUTTON_WH = 20
        args= {
            'Name': 'find_source_window',
            'Title': 'Поиск источников текста',
            'Width': 280,
            'Height': 180,
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
            'Type': 'Button',
            'Name': 'find_source_but',
            'Label': 'Поиск',
            'Width': 100,
            'Height': 20,
            'Step': 10,
            'Align': 1,
            'FontHeight': 16,
        }
        ctl = dialog.add_control(args)
        ctl.addActionListener(self.findSource)
        dialog.center(dialog.find_source_but, y=30)

        

        args = {
            'Type': 'Listbox',
            'Name': 'link_list',
            'Width': 260,
            'Height': 120,
            'Step': 10,
            'FontHeight': 14,
        }
        dialog.add_control(args)
        dialog.center(dialog.link_list, y=10)
        dialog.link_list.visible = False

        args = {
            'Type': 'Button',
            'Name': 'follow_link_but',
            'Label': 'Перейти по ссылке',
            'Width': 120,
            'Height': 20,
            'Step': 10,
            'Align': 1,
            'FontHeight': 16,
        }
        ctl = dialog.add_control(args)
        ctl.addActionListener(self.findSource)
        dialog.center(dialog.follow_link_but, y=140)
        dialog.follow_link_but.visible = False

        return dialog

# Регистрация реализации службы
g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation( \
		Find_Sources_Context,                                # Имя класса UNO
		"org.openoffice.comp.pyuno.exp.Find_Sources_Context",# Имя реализации
		("com.sun.star.task.Job",),)                # Список реализованных служб
