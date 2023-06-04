import uno
import unohelper
from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
from com.sun.star.awt.MessageBoxResults import YES
from com.sun.star.awt import Rectangle, Size, Point
from com.sun.star.awt.PosSize import POSSIZE, SIZE
from com.sun.star.beans import PropertyValue, NamedValue
from pathlib import Path
import platform
import getpass
import os
import sys
import re
from typing import Any, Union

CTX = uno.getComponentContext()
SM = CTX.getServiceManager()

def create_instance(name: str, with_context: bool=False, args: Any=None) -> Any:
    if with_context:
        instance = SM.createInstanceWithContext(name, CTX)
    elif args:
        instance = SM.createInstanceWithArguments(name, (args,))
    else:
        instance = SM.createInstance(name)
    return instance

def get_app_config(node_name: str, key: str=''):
    name = 'com.sun.star.configuration.ConfigurationProvider'
    service = 'com.sun.star.configuration.ConfigurationAccess'
    cp = create_instance(name, True)
    node = PropertyValue(Name='nodepath', Value=node_name)
    try:
        ca = cp.createInstanceWithArguments(service, (node,))
        if ca and not key:
            return ca
        if ca and ca.hasByName(key):
            return ca.getPropertyValue(key)
    except Exception as e:
        return ''

NAME = TITLE = get_app_config('org.openoffice.Setup/Product', 'ooName')
DIR = {
    'images': 'images',
    'locales': 'locales',
}

def _set_properties(model, properties):
    if 'X' in properties:
        properties['PositionX'] = properties.pop('X')
    if 'Y' in properties:
        properties['PositionY'] = properties.pop('Y')
    keys = tuple(properties.keys())
    values = tuple(properties.values())
    model.setPropertyValues(keys, values)
    return

class UnoBaseObject(object):

    def __init__(self, obj, path=''):
        self._obj = obj
        self._model = obj.Model

    def __setattr__(self, name, value):
        exists = hasattr(self, name)
        if not exists and not name in ('_obj', '_model'):
            setattr(self._model, name, value)
        else:
            super().__setattr__(name, value)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        pass

    @property
    def obj(self):
        return self._obj

    @property
    def model(self):
        return self._model
    @property
    def m(self):
        return self._model

    @property
    def properties(self):
        return {}
    @properties.setter
    def properties(self, values):
        _set_properties(self.model, values)

    @property
    def name(self):
        return self.model.Name

    @property
    def parent(self):
        return self.obj.Context

    @property
    def tag(self):
        return self.model.Tag
    @tag.setter
    def tag(self, value):
        self.model.Tag = value

    @property
    def visible(self):
        return self.obj.Visible
    @visible.setter
    def visible(self, value):
        self.obj.setVisible(value)

    @property
    def enabled(self):
        return self.model.Enabled
    @enabled.setter
    def enabled(self, value):
        self.model.Enabled = value

    @property
    def step(self):
        return self.model.Step
    @step.setter
    def step(self, value):
        self.model.Step = value

    @property
    def align(self):
        return self.model.Align
    @align.setter
    def align(self, value):
        self.model.Align = value

    @property
    def valign(self):
        return self.model.VerticalAlign
    @valign.setter
    def valign(self, value):
        self.model.VerticalAlign = value

    @property
    def font_weight(self):
        return self.model.FontWeight
    @font_weight.setter
    def font_weight(self, value):
        self.model.FontWeight = value

    @property
    def font_height(self):
        return self.model.FontHeight
    @font_height.setter
    def font_height(self, value):
        self.model.FontHeight = value

    @property
    def font_name(self):
        return self.model.FontName
    @font_name.setter
    def font_name(self, value):
        self.model.FontName = value

    @property
    def font_underline(self):
        return self.model.FontUnderline
    @font_underline.setter
    def font_underline(self, value):
        self.model.FontUnderline = value

    @property
    def text_color(self):
        return self.model.TextColor
    @text_color.setter
    def text_color(self, value):
        self.model.TextColor = value

    @property
    def back_color(self):
        return self.model.BackgroundColor
    @back_color.setter
    def back_color(self, value):
        self.model.BackgroundColor = value

    @property
    def multi_line(self):
        return self.model.MultiLine
    @multi_line.setter
    def multi_line(self, value):
        self.model.MultiLine = value

    @property
    def help_text(self):
        return self.model.HelpText
    @help_text.setter
    def help_text(self, value):
        self.model.HelpText = value

    @property
    def border(self):
        return self.model.Border
    @border.setter
    def border(self, value):
        # ~ Bug for report
        self.model.Border = value

    @property
    def width(self):
        return self._model.Width
    @width.setter
    def width(self, value):
        self.model.Width = value

    @property
    def height(self):
        return self.model.Height
    @height.setter
    def height(self, value):
        self.model.Height = value

    def _get_possize(self, name):
        ps = self.obj.getPosSize()
        return getattr(ps, name)

    def _set_possize(self, name, value):
        ps = self.obj.getPosSize()
        setattr(ps, name, value)
        self.obj.setPosSize(ps.X, ps.Y, ps.Width, ps.Height, POSSIZE)
        return

    @property
    def x(self):
        if hasattr(self.model, 'PositionX'):
            return self.model.PositionX
        return self._get_possize('X')
    @x.setter
    def x(self, value):
        if hasattr(self.model, 'PositionX'):
            self.model.PositionX = value
        else:
            self._set_possize('X', value)

    @property
    def y(self):
        if hasattr(self.model, 'PositionY'):
            return self.model.PositionY
        return self._get_possize('Y')
    @y.setter
    def y(self, value):
        if hasattr(self.model, 'PositionY'):
            self.model.PositionY = value
        else:
            self._set_possize('Y', value)

    @property
    def tab_index(self):
        return self._model.TabIndex
    @tab_index.setter
    def tab_index(self, value):
        self.model.TabIndex = value

    @property
    def tab_stop(self):
        return self._model.Tabstop
    @tab_stop.setter
    def tab_stop(self, value):
        self.model.Tabstop = value

    @property
    def ps(self):
        ps = self.obj.getPosSize()
        return ps
    @ps.setter
    def ps(self, ps):
        self.obj.setPosSize(ps.X, ps.Y, ps.Width, ps.Height, POSSIZE)

    def set_focus(self):
        self.obj.setFocus()
        return

    def ps_from(self, source):
        self.ps = source.ps
        return

    def center(self, horizontal=True, vertical=False):
        p = self.parent.Model
        w = p.Width
        h = p.Height
        if horizontal:
            x = w / 2 - self.width / 2
            self.x = x
        if vertical:
            y = h / 2 - self.height / 2
            self.y = y
        return

    def move(self, origin, x=0, y=5, center=False):
        if x:
            self.x = origin.x + origin.width + x
        else:
            self.x = origin.x
        if y:
            self.y = origin.y + origin.height + y
        else:
            self.y = origin.y

        if center:
            self.center()
        return

# класс компонета label
class UnoLabel(UnoBaseObject):

    def __init__(self, obj):
        super().__init__(obj)

    @property
    def type(self):
        return 'label'

    @property
    def value(self):
        return self.model.Label
    @value.setter
    def value(self, value):
        self.model.Label = value

# класс компонета link
class UnoLabelLink(UnoLabel):

    def __init__(self, obj):
        super().__init__(obj)

    @property
    def type(self):
        return 'link'

# класс компонета button
class UnoButton(UnoBaseObject):

    def __init__(self, obj):
        super().__init__(obj)

    @property
    def type(self):
        return 'button'

    @property
    def value(self):
        return self.model.Label
    @value.setter
    def value(self, value):
        self.model.Label = value

# класс компонета Radio
class UnoRadio(UnoBaseObject):

    def __init__(self, obj):
        super().__init__(obj)

    @property
    def type(self):
        return 'radio'

    @property
    def value(self):
        return self.model.Label
    @value.setter
    def value(self, value):
        self.model.Label = value

# класс компонета checkBox
class UnoCheckBox(UnoBaseObject):

    def __init__(self, obj):
        super().__init__(obj)

    @property
    def type(self):
        return 'checkbox'

    @property
    def value(self):
        return self.model.State
    @value.setter
    def value(self, value):
        self.model.State = value

    @property
    def label(self):
        return self.model.Label
    @label.setter
    def label(self, value):
        self.model.Label = value

    @property
    def tri_state(self):
        return self.model.TriState
    @tri_state.setter
    def tri_state(self, value):
        self.model.TriState = value

# класс компонета text
class UnoText(UnoBaseObject):

    def __init__(self, obj):
        super().__init__(obj)

    @property
    def type(self):
        return 'text'

    @property
    def value(self):
        return self.model.Text
    @value.setter
    def value(self, value):
        self.model.Text = value

    @property
    def echochar(self):
        return chr(self.model.EchoChar)
    @echochar.setter
    def echochar(self, value):
        self.model.EchoChar = ord(value[0])

    def validate(self):
        return

# класс компонета image
class UnoImage(UnoBaseObject):

    def __init__(self, obj):
        super().__init__(obj)

    @property
    def type(self):
        return 'image'

    @property
    def value(self):
        return self.url
    @value.setter
    def value(self, value):
        self.url = value

    @property
    def url(self):
        return self.m.ImageURL
    @url.setter
    def url(self, value):
        self.m.ImageURL = None
        self.m.ImageURL = _P.to_url(value)

# класс компонета listBox
class UnoListBox(UnoBaseObject):

    def __init__(self, obj):
        super().__init__(obj)
        self._path = ''

    def __setattr__(self, name, value):
        if name in ('_path',):
            self.__dict__[name] = value
        else:
            super().__setattr__(name, value)

    @property
    def type(self):
        return 'listbox'

    @property
    def value(self):
        return self.obj.getSelectedItem()

    @property
    def count(self):
        return len(self.data)

    @property
    def data(self):
        return self.model.StringItemList
    @data.setter
    def data(self, values):
        self.model.StringItemList = list(sorted(values))

    @property
    def path(self):
        return self._path
    @path.setter
    def path(self, value):
        self._path = value

    def unselect(self):
        self.obj.selectItem(self.value, False)
        return

    def select(self, pos=0):
        if isinstance(pos, str):
            self.obj.selectItem(pos, True)
        else:
            self.obj.selectItemPos(pos, True)
        return

    def clear(self):
        self.model.removeAllItems()
        return

    def _set_image_url(self, image):
        if _P.exists(image):
            return _P.to_url(image)

        path = _P.join(self._path, DIR['images'], image)
        return _P.to_url(path)

    def insert(self, value, path='', pos=-1, show=True):
        if pos < 0:
            pos = self.count
        if path:
            self.model.insertItem(pos, value, self._set_image_url(path))
        else:
            self.model.insertItemText(pos, value)
        if show:
            self.select(pos)
        return

UNO_CLASSES = {
    'label': UnoLabel,
    'link': UnoLabelLink,
    'button': UnoButton,
    'radio': UnoRadio,
    'checkbox': UnoCheckBox,
    'text': UnoText,
    'image': UnoImage,
    'listbox': UnoListBox,
}

def msgbox(message, title=TITLE, buttons=MSG_BUTTONS.BUTTONS_OK, type_msg='infobox'):
    """ Create message box
        type_msg: infobox, warningbox, errorbox, querybox, messbox
        http://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XMessageBoxFactory.html
    """
    toolkit = create_instance('com.sun.star.awt.Toolkit')
    parent = toolkit.getDesktopWindow()
    box = toolkit.createMessageBox(parent, type_msg, buttons, title, str(message))
    return box.execute()

def question(message, title=TITLE):
    result = msgbox(message, title, MSG_BUTTONS.BUTTONS_YES_NO, 'querybox')
    return result == YES


def warning(message, title=TITLE):
    return msgbox(message, title, type_msg='warningbox')


def errorbox(message, title=TITLE):
    return msgbox(message, title, type_msg='errorbox')

# класс пользовательской формы
class LODialog(object):
    SEPARATION = 5
    MODELS = {
        'label': 'com.sun.star.awt.UnoControlFixedTextModel',
        'link': 'com.sun.star.awt.UnoControlFixedHyperlinkModel',
        'button': 'com.sun.star.awt.UnoControlButtonModel',
        'radio': 'com.sun.star.awt.UnoControlRadioButtonModel',
        'checkbox': 'com.sun.star.awt.UnoControlCheckBoxModel',
        'text': 'com.sun.star.awt.UnoControlEditModel',
        'image': 'com.sun.star.awt.UnoControlImageControlModel',
        'listbox': 'com.sun.star.awt.UnoControlListBoxModel',
        'groupbox': 'com.sun.star.awt.UnoControlGroupBoxModel',
        'combobox': 'com.sun.star.awt.UnoControlComboBoxModel',
    }

    # конструткор класса
    def __init__(self, args):
        self._obj = self._create(args)
        self._model = self.obj.Model
        self._events = None
        self._modal = True
        self._controls = {}
        self._color_on_focus = COLOR_ON_FOCUS
        self._id = ''
        self._path = ''
        self._init_controls()

    # создание формы
    def _create(self, args):
        service = 'com.sun.star.awt.DialogProvider'
        path = args.pop('Path', '')
        if path:
            dp = create_instance(service, True)
            dlg = dp.createDialog(_P.to_url(path))
            return dlg

        if 'Location' in args:
            name = args['Name']
            library = args.get('Library', 'Standard')
            location = args.get('Location', 'application').lower()
            if location == 'user':
                location = 'application'
            url = f'vnd.sun.star.script:{library}.{name}?location={location}'
            if location == 'document':
                dp = create_instance(service, args=docs.active.obj)
            else:
                dp = create_instance(service, True)

            dlg = dp.createDialog(url)
            return dlg

        dlg = create_instance('com.sun.star.awt.UnoControlDialog', True)
        model = create_instance('com.sun.star.awt.UnoControlDialogModel', True)
        toolkit = create_instance('com.sun.star.awt.Toolkit', True)
        _set_properties(model, args)
        dlg.setModel(model)
        dlg.setVisible(False)
        dlg.createPeer(toolkit, None)
        return dlg

    # получение типа компонента
    def _get_type_control(self, name):
        name = name.split('.')[2]
        types = {
            'UnoFixedTextControl': 'label',
            'UnoEditControl': 'text',
            'UnoButtonControl': 'button',
        }
        return types[name]

    # инициализация компонента
    def _init_controls(self):
        for control in self.obj.getControls():
            tipo = self._get_type_control(control.ImplementationName)
            name = control.Model.Name
            control = UNO_CLASSES[tipo](control)
            setattr(self, name, control)
        return

    @property
    def obj(self):
        return self._obj

    @property
    def model(self):
        return self._model

    @property
    def controls(self):
        return self._controls

    @property
    def path(self):
        return self._path
    @property
    def id(self):
        return self._id

    @property
    def height(self):
        return self.model.Height
    @height.setter
    def height(self, value):
        self.model.Height = value

    @property
    def width(self):
        return self.model.Width
    @width.setter
    def width(self, value):
        self.model.Width = value

    @property
    def visible(self):
        return self.obj.Visible
    @visible.setter
    def visible(self, value):
        self.obj.Visible = value

    @property
    def step(self):
        return self.model.Step
    @step.setter
    def step(self, value):
        self.model.Step = value

    @property
    def events(self):
        return self._events
    @events.setter
    def events(self, controllers):
        self._events = controllers(self)
        self._connect_listeners()

    @property
    def color_on_focus(self):
        return self._color_on_focus
    @color_on_focus.setter
    def color_on_focus(self, value):
        self._color_on_focus = get_color(value)

    def _special_properties(self, tipo, args):
        if tipo == 'link' and not 'Label' in args:
            args['Label'] = args['URL']
            return args

        if tipo == 'button':
            if 'ImageURL' in args:
                args['ImageURL'] = self._set_image_url(args['ImageURL'])
            args['FocusOnClick'] = args.get('FocusOnClick', False)
            return args

        return args

    # добавление компонента
    def add_control(self, args):
        tipo = args.pop('Type').lower()

        args = self._special_properties(tipo, args)
        model = self.model.createInstance(self.MODELS[tipo])
        _set_properties(model, args)
        name = args['Name']
        self.model.insertByName(name, model)
        control = self.obj.getControl(name)
        control = UNO_CLASSES[tipo](control)

        if tipo in ('listbox',):
            control.path = self.path

        setattr(self, name, control)
        self._controls[name] = control
        control = self.obj.getControl(name)
        return control

    # центрирования компонента
    def center(self, control, x=0, y=0):
        w = self.width
        h = self.height

        if isinstance(control, tuple):
            wt = self.SEPARATION * -1
            for c in control:
                wt += c.width + self.SEPARATION
            x = w / 2 - wt / 2
            for c in control:
                c.x = x
                x = c.x + c.width + self.SEPARATION
            return

        if x < 0:
            x = w + x - control.width
        elif x == 0:
            x = w / 2 - control.width / 2
        if y < 0:
            y = h + y - control.height
        elif y == 0:
            y = h / 2 - control.height / 2
        control.x = x
        control.y = y
        return

    # открытие формы
    def open(self, modal=True):
        self._modal = modal
        if modal:
            return self.obj.execute()
        else:
            self.visible = True
        return

    # закрытие формы
    def close(self, value=0):
        if self._modal:
            value = self.obj.endDialog(value)
        else:
            self.visible = False
            self.obj.dispose()
        return value

    def set_values(self, data):
        for k, v in data.items():
            self._controls[k].value = v
        return

def create_dialog(args):
    return LODialog(args)


class Paths(object):
    FILE_PICKER = 'com.sun.star.ui.dialogs.FilePicker'
    FOLDER_PICKER = 'com.sun.star.ui.dialogs.FolderPicker'

    def __init__(self, path=''):
        if path.startswith('file://'):
            path = str(Path(uno.fileUrlToSystemPath(path)).resolve())
        self._path = Path(path)

    @classmethod
    def to_url(cls, path):
        if not path.startswith('file://'):
            path = Path(path).as_uri()
        return path

_P = Paths
# ~ https://en.wikipedia.org/wiki/Web_colors
def get_color(value):
    COLORS = {
        'aliceblue': 15792383,
        'antiquewhite': 16444375,
        'aqua': 65535,
        'aquamarine': 8388564,
        'azure': 15794175,
        'beige': 16119260,
        'bisque': 16770244,
        'black': 0,
        'blanchedalmond': 16772045,
        'blue': 255,
        'blueviolet': 9055202,
        'brown': 10824234,
        'burlywood': 14596231,
        'cadetblue': 6266528,
        'chartreuse': 8388352,
        'chocolate': 13789470,
        'coral': 16744272,
        'cornflowerblue': 6591981,
        'cornsilk': 16775388,
        'crimson': 14423100,
        'cyan': 65535,
        'darkblue': 139,
        'darkcyan': 35723,
        'darkgoldenrod': 12092939,
        'darkgray': 11119017,
        'darkgreen': 25600,
        'darkgrey': 11119017,
        'darkkhaki': 12433259,
        'darkmagenta': 9109643,
        'darkolivegreen': 5597999,
        'darkorange': 16747520,
        'darkorchid': 10040012,
        'darkred': 9109504,
        'darksalmon': 15308410,
        'darkseagreen': 9419919,
        'darkslateblue': 4734347,
        'darkslategray': 3100495,
        'darkslategrey': 3100495,
        'darkturquoise': 52945,
        'darkviolet': 9699539,
        'deeppink': 16716947,
        'deepskyblue': 49151,
        'dimgray': 6908265,
        'dimgrey': 6908265,
        'dodgerblue': 2003199,
        'firebrick': 11674146,
        'floralwhite': 16775920,
        'forestgreen': 2263842,
        'fuchsia': 16711935,
        'gainsboro': 14474460,
        'ghostwhite': 16316671,
        'gold': 16766720,
        'goldenrod': 14329120,
        'gray': 8421504,
        'grey': 8421504,
        'green': 32768,
        'greenyellow': 11403055,
        'honeydew': 15794160,
        'hotpink': 16738740,
        'indianred': 13458524,
        'indigo': 4915330,
        'ivory': 16777200,
        'khaki': 15787660,
        'lavender': 15132410,
        'lavenderblush': 16773365,
        'lawngreen': 8190976,
        'lemonchiffon': 16775885,
        'lightblue': 11393254,
        'lightcoral': 15761536,
        'lightcyan': 14745599,
        'lightgoldenrodyellow': 16448210,
        'lightgray': 13882323,
        'lightgreen': 9498256,
        'lightgrey': 13882323,
        'lightpink': 16758465,
        'lightsalmon': 16752762,
        'lightseagreen': 2142890,
        'lightskyblue': 8900346,
        'lightslategray': 7833753,
        'lightslategrey': 7833753,
        'lightsteelblue': 11584734,
        'lightyellow': 16777184,
        'lime': 65280,
        'limegreen': 3329330,
        'linen': 16445670,
        'magenta': 16711935,
        'maroon': 8388608,
        'mediumaquamarine': 6737322,
        'mediumblue': 205,
        'mediumorchid': 12211667,
        'mediumpurple': 9662683,
        'mediumseagreen': 3978097,
        'mediumslateblue': 8087790,
        'mediumspringgreen': 64154,
        'mediumturquoise': 4772300,
        'mediumvioletred': 13047173,
        'midnightblue': 1644912,
        'mintcream': 16121850,
        'mistyrose': 16770273,
        'moccasin': 16770229,
        'navajowhite': 16768685,
        'navy': 128,
        'oldlace': 16643558,
        'olive': 8421376,
        'olivedrab': 7048739,
        'orange': 16753920,
        'orangered': 16729344,
        'orchid': 14315734,
        'palegoldenrod': 15657130,
        'palegreen': 10025880,
        'paleturquoise': 11529966,
        'palevioletred': 14381203,
        'papayawhip': 16773077,
        'peachpuff': 16767673,
        'peru': 13468991,
        'pink': 16761035,
        'plum': 14524637,
        'powderblue': 11591910,
        'purple': 8388736,
        'red': 16711680,
        'rosybrown': 12357519,
        'royalblue': 4286945,
        'saddlebrown': 9127187,
        'salmon': 16416882,
        'sandybrown': 16032864,
        'seagreen': 3050327,
        'seashell': 16774638,
        'sienna': 10506797,
        'silver': 12632256,
        'skyblue': 8900331,
        'slateblue': 6970061,
        'slategray': 7372944,
        'slategrey': 7372944,
        'snow': 16775930,
        'springgreen': 65407,
        'steelblue': 4620980,
        'tan': 13808780,
        'teal': 32896,
        'thistle': 14204888,
        'tomato': 16737095,
        'turquoise': 4251856,
        'violet': 15631086,
        'wheat': 16113331,
        'white': 16777215,
        'whitesmoke': 16119285,
        'yellow': 16776960,
        'yellowgreen': 10145074,
    }

    if isinstance(value, tuple):
        color = (value[0] << 16) + (value[1] << 8) + value[2]
    else:
        if value[0] == '#':
            r, g, b = bytes.fromhex(value[1:])
            color = (r << 16) + (g << 8) + b
        else:
            color = COLORS.get(value.lower(), -1)
    return color

COLOR_ON_FOCUS = get_color('LightYellow')
