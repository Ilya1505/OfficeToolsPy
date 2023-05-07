import uno
import unohelper
from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
from com.sun.star.awt.MessageBoxResults import YES
from com.sun.star.awt import Rectangle, Size, Point
from com.sun.star.awt.PosSize import POSSIZE, SIZE
from com.sun.star.awt import Key, KeyModifier, KeyEvent
from com.sun.star.beans import PropertyValue, NamedValue
from com.sun.star.view.SelectionType import SINGLE, MULTI, RANGE
from com.sun.star.lang import XEventListener
from com.sun.star.awt import XActionListener
from com.sun.star.awt import XMouseListener
from com.sun.star.awt import XMouseMotionListener
from com.sun.star.awt import XFocusListener
from com.sun.star.awt import XKeyListener
from com.sun.star.awt import XItemListener
from com.sun.star.awt import XTabListener
from com.sun.star.awt.grid import XGridDataListener
from com.sun.star.awt.grid import XGridSelectionListener
from pathlib import Path
import tempfile
import json
import shlex
import csv
import zipfile
import platform
import getpass
import os
import sys
import shutil
import subprocess
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

OS = platform.system()
IS_WIN = OS == 'Windows'
IS_MAC = OS == 'Darwin'
USER = getpass.getuser()
PC = platform.node()
DESKTOP = os.environ.get('DESKTOP_SESSION', '')
INFO_DEBUG = f"{sys.version}\n\n{platform.platform()}\n\n" + '\n'.join(sys.path)
NAME = TITLE = get_app_config('org.openoffice.Setup/Product', 'ooName')
VERSION = get_app_config('org.openoffice.Setup/Product','ooSetupVersion')

PYTHON = 'python'
if IS_WIN:
    PYTHON = 'python.exe'

DIR = {
    'images': 'images',
    'locales': 'locales',
}

class EventsListenerBase(unohelper.Base, XEventListener):

    def __init__(self, controller, name, window=None):
        self._controller = controller
        self._name = name
        self._window = window

    @property
    def name(self):
        return self._name

    def disposing(self, event):
        self._controller = None
        if not self._window is None:
            self._window.setMenuBar(None)

class EventsButton(EventsListenerBase, XActionListener):

    def __init__(self, controller, name):
        super().__init__(controller, name)

    def actionPerformed(self, event):
        event_name = f'{self.name}_action'
        if hasattr(self._controller, event_name):
            getattr(self._controller, event_name)(event)
        return

class EventsMouse(EventsListenerBase, XMouseListener, XMouseMotionListener):

    def __init__(self, controller, name):
        super().__init__(controller, name)

    def mousePressed(self, event):
        event_name = '{}_click'.format(self._name)
        if event.ClickCount == 2:
            event_name = '{}_double_click'.format(self._name)
        if hasattr(self._controller, event_name):
            getattr(self._controller, event_name)(event)
        return

    def mouseReleased(self, event):
        pass

    def mouseEntered(self, event):
        pass

    def mouseExited(self, event):
        pass

    # ~ XMouseMotionListener
    def mouseMoved(self, event):
        pass

    def mouseDragged(self, event):
        pass

class EventsGrid(EventsListenerBase, XGridDataListener, XGridSelectionListener):

    def __init__(self, controller, name):
        super().__init__(controller, name)

    def dataChanged(self, event):
        event_name = '{}_data_changed'.format(self.name)
        if hasattr(self._controller, event_name):
            getattr(self._controller, event_name)(event)
        return

    def rowHeadingChanged(self, event):
        pass

    def rowsInserted(self, event):
        pass

    def rowsRemoved(self, evemt):
        pass

    def selectionChanged(self, event):
        event_name = '{}_selection_changed'.format(self.name)
        if hasattr(self._controller, event_name):
            getattr(self._controller, event_name)(event)
        return

class EventsMouseGrid(EventsMouse):
    selected = False

    def mousePressed(self, event):
        super().mousePressed(event)
        # ~ obj = event.Source
        # ~ col = obj.getColumnAtPoint(event.X, event.Y)
        # ~ row = obj.getRowAtPoint(event.X, event.Y)
        # ~ print(col, row)
        # ~ if col == -1 and row == -1:
            # ~ if self.selected:
                # ~ obj.deselectAllRows()
            # ~ else:
                # ~ obj.selectAllRows()
            # ~ self.selected = not self.selected
        return

    def mouseReleased(self, event):
        # ~ obj = event.Source
        # ~ col = obj.getColumnAtPoint(event.X, event.Y)
        # ~ row = obj.getRowAtPoint(event.X, event.Y)
        # ~ if row == -1 and col > -1:
            # ~ gdm = obj.Model.GridDataModel
            # ~ for i in range(gdm.RowCount):
                # ~ gdm.updateRowHeading(i, i + 1)
        return

class EventsMouseLink(EventsMouse):

    def __init__(self, controller, name):
        super().__init__(controller, name)
        self._text_color = 0

    def mouseEntered(self, event):
        model = event.Source.Model
        self._text_color = model.TextColor or 0
        model.TextColor = get_color('blue')
        return

    def mouseExited(self, event):
        model = event.Source.Model
        model.TextColor = self._text_color
        return

class EventsFocus(EventsListenerBase, XFocusListener):
    CONTROLS = (
        'stardiv.Toolkit.UnoControlEditModel',
    )

    def __init__(self, controller, name):
        super().__init__(controller, name)

    def focusGained(self, event):
        service = event.Source.Model.ImplementationName
        # ~ print('Focus enter', service)
        if service in self.CONTROLS:
            obj = event.Source.Model
            obj.BackgroundColor = COLOR_ON_FOCUS
        return

    def focusLost(self, event):
        service = event.Source.Model.ImplementationName
        if service in self.CONTROLS:
            obj = event.Source.Model
            obj.BackgroundColor = -1
        return

class EventsKey(EventsListenerBase, XKeyListener):
    """
        event.KeyChar
        event.KeyCode
        event.KeyFunc
        event.Modifiers
    """

    def __init__(self, controller, name):
        super().__init__(controller, name)

    def keyPressed(self, event):
        pass

    def keyReleased(self, event):
        event_name = '{}_key_released'.format(self._name)
        if hasattr(self._controller, event_name):
            getattr(self._controller, event_name)(event)
        # ~ else:
            # ~ if event.KeyFunc == QUIT and hasattr(self._cls, 'close'):
                # ~ self._cls.close()
        return


class EventsItem(EventsListenerBase, XItemListener):

    def __init__(self, controller, name):
        super().__init__(controller, name)

    def disposing(self, event):
        pass

    def itemStateChanged(self, event):
        event_name = '{}_item_changed'.format(self.name)
        if hasattr(self._controller, event_name):
            getattr(self._controller, event_name)(event)
        return

class EventsTab(EventsListenerBase, XTabListener):

    def __init__(self, controller, name):
        super().__init__(controller, name)

    def activated(self, id):
        event_name = '{}_activated'.format(self.name)
        if hasattr(self._controller, event_name):
            getattr(self._controller, event_name)(id)
        return

def _add_listeners(events, control, name=''):
    listeners = {
        'addActionListener': EventsButton,
        'addMouseListener': EventsMouse,
        'addFocusListener': EventsFocus,
        'addItemListener': EventsItem,
        'addKeyListener': EventsKey,
        'addTabListener': EventsTab,
    }
    if hasattr(control, 'obj'):
        control = control.obj
    # ~ debug(control.ImplementationName)
    is_grid = control.ImplementationName == 'stardiv.Toolkit.GridControl'
    is_link = control.ImplementationName == 'stardiv.Toolkit.UnoFixedHyperlinkControl'
    is_roadmap = control.ImplementationName == 'stardiv.Toolkit.UnoRoadmapControl'
    is_pages = control.ImplementationName == 'stardiv.Toolkit.UnoMultiPageControl'

    for key, value in listeners.items():
        if hasattr(control, key):
            if is_grid and key == 'addMouseListener':
                control.addMouseListener(EventsMouseGrid(events, name))
                continue
            if is_link and key == 'addMouseListener':
                control.addMouseListener(EventsMouseLink(events, name))
                continue
            if is_roadmap and key == 'addItemListener':
                control.addItemListener(EventsItemRoadmap(events, name))
                continue

            getattr(control, key)(listeners[key](events, name))

    if is_grid:
        controllers = EventsGrid(events, name)
        control.addSelectionListener(controllers)
        control.Model.GridDataModel.addGridDataListener(controllers)
    return

class EventsItemRoadmap(EventsItem):

    def itemStateChanged(self, event):
        dialog = event.Source.Context.Model
        dialog.Step = event.ItemId + 1
        return

def _set_properties(model, properties):
    if 'X' in properties:
        properties['PositionX'] = properties.pop('X')
    if 'Y' in properties:
        properties['PositionY'] = properties.pop('Y')
    keys = tuple(properties.keys())
    values = tuple(properties.values())
    model.setPropertyValues(keys, values)
    return

# ~ BorderColor = ?
# ~ FontStyleName = ?
# ~ HelpURL = ?
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


class UnoLabelLink(UnoLabel):

    def __init__(self, obj):
        super().__init__(obj)

    @property
    def type(self):
        return 'link'

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

# ~ https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlEditModel.html
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

class UnoRoadmap(UnoBaseObject):

    def __init__(self, obj):
        super().__init__(obj)
        self._options = ()

    def __setattr__(self, name, value):
        if name in ('_options',):
            self.__dict__[name] = value
        else:
            super().__setattr__(name, value)

    @property
    def options(self):
        return self._options
    @options.setter
    def options(self, values):
        self._options = values
        for i, v in enumerate(values):
            opt = self.model.createInstance()
            opt.ID = i
            opt.Label = v
            self.model.insertByIndex(i, opt)
        return

    @property
    def enabled(self):
        return True
    @enabled.setter
    def enabled(self, value):
        for m in self.model:
            m.Enabled = value
        return

    def set_enabled(self, index, value):
        self.model.getByIndex(index).Enabled = value
        return
    
class UnoTree(UnoBaseObject):

    def __init__(self, obj, ):
        super().__init__(obj)
        self._tdm = None
        self._data = []

    def __setattr__(self, name, value):
        if name in ('_tdm', '_data'):
            self.__dict__[name] = value
        else:
            super().__setattr__(name, value)

    @property
    def selection(self):
        sel = self.obj.Selection
        return sel.DataValue, sel.DisplayValue

    @property
    def parent(self):
        parent = self.obj.Selection.Parent
        if parent is None:
            return ()
        return parent.DataValue, parent.DisplayValue

    def _get_parents(self, node):
        value = (node.DisplayValue,)
        parent = node.Parent
        if parent is None:
            return value
        return self._get_parents(parent) + value

    @property
    def parents(self):
        values = self._get_parents(self.obj.Selection)
        return values

    @property
    def root(self):
        if self._tdm is None:
            return ''
        return self._tdm.Root.DisplayValue
    @root.setter
    def root(self, value):
        self._add_data_model(value)

    def _add_data_model(self, name):
        tdm = create_instance('com.sun.star.awt.tree.MutableTreeDataModel')
        root = tdm.createNode(name, True)
        root.DataValue = 0
        tdm.setRoot(root)
        self.model.DataModel = tdm
        self._tdm = self.model.DataModel
        return

    @property
    def path(self):
        return self.root
    @path.setter
    def path(self, value):
        self.data = _P.walk_dir(value, True)

    @property
    def data(self):
        return self._data
    @data.setter
    def data(self, values):
        self._data = list(values)
        self._add_data()

    def _add_data(self):
        if not self.data:
            return

        parents = {}
        for node in self.data:
            parent = parents.get(node[1], self._tdm.Root)
            child = self._tdm.createNode(node[2], False)
            child.DataValue = node[0]
            parent.appendChild(child)
            parents[node[0]] = child
        self.obj.expandNode(self._tdm.Root)
        return

# ~ https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1awt_1_1grid.html
class UnoGrid(UnoBaseObject):

    def __init__(self, obj):
        super().__init__(obj)
        self._gdm = self.model.GridDataModel
        self._data = []
        self._formats = ()

    def __setattr__(self, name, value):
        if name in ('_gdm', '_data', '_formats'):
            self.__dict__[name] = value
        else:
            super().__setattr__(name, value)

    def __getitem__(self, key):
        value = self._gdm.getCellData(key[0], key[1])
        return value

    def __setitem__(self, key, value):
        self._gdm.updateCellData(key[0], key[1], value)
        return

    @property
    def type(self):
        return 'grid'

    @property
    def columns(self):
        return {}
    @columns.setter
    def columns(self, values):
        # ~ self._columns = values
        #~ https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1grid_1_1XGridColumn.html
        model = create_instance('com.sun.star.awt.grid.DefaultGridColumnModel', True)
        for properties in values:
            column = create_instance('com.sun.star.awt.grid.GridColumn', True)
            for k, v in properties.items():
                setattr(column, k, v)
            model.addColumn(column)
        self.model.ColumnModel = model
        return

    @property
    def data(self):
        return self._data
    @data.setter
    def data(self, values):
        self._data = values
        self.clear()
        headings = tuple(range(1, len(values) + 1))
        self._gdm.addRows(headings, values)
        # ~ rows = range(grid_dm.RowCount)
        # ~ colors = [COLORS['GRAY'] if r % 2 else COLORS['WHITE'] for r in rows]
        # ~ grid.Model.RowBackgroundColors = tuple(colors)
        return

    @property
    def value(self):
        if self.column == -1 or self.row == -1:
            return ''
        return self[self.column, self.row]
    @value.setter
    def value(self, value):
        if self.column > -1 and self.row > -1:
            self[self.column, self.row] = value

    @property
    def row(self):
        return self.obj.CurrentRow

    @property
    def row_count(self):
        return self._gdm.RowCount

    @property
    def column(self):
        return self.obj.CurrentColumn

    @property
    def column(self):
        return self.obj.CurrentColumn

    @property
    def is_valid(self):
        return not (self.row == -1 or self.column == -1)

    @property
    def formats(self):
        return self._formats
    @formats.setter
    def formats(self, values):
        self._formats = values

    def clear(self):
        self._gdm.removeAllRows()
        return

    def _format_columns(self, data):
        row = data
        if self.formats:
            for i, f in enumerate(formats):
                if f:
                    row[i] = f.format(data[i])
        return row

    def add_row(self, data):
        self._data.append(data)
        row = self._format_columns(data)
        self._gdm.addRow(self.row_count + 1, row)
        return

    def set_cell_tooltip(self, col, row, value):
        self._gdm.updateCellToolTip(col, row, value)
        return

    def get_cell_tooltip(self, col, row):
        value = self._gdm.getCellToolTip(col, row)
        return value

    def sort(self, column, asc=True):
        self._gdm.sortByColumn(column, asc)
        self.update_row_heading()
        return

    def update_row_heading(self):
        for i in range(self.row_count):
            self._gdm.updateRowHeading(i, i + 1)
        return

    def remove_row(self, row):
        self._gdm.removeRow(row)
        del self._data[row]
        self.update_row_heading()
        return

class UnoPages(object):

    def __init__(self, obj):
        self._obj = obj
        self._events = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        pass

    @property
    def obj(self):
        return self._obj

    @property
    def model(self):
        return self._obj.Model

    # ~ @property
    # ~ def id(self):
        # ~ return self.m.TabPageID

    @property
    def parent(self):
        return self.obj.Context

    def _set_image_url(self, image):
        if _P.exists(image):
            return _P.to_url(image)

        path = _P.join(self._path, DIR['images'], image)
        return _P.to_url(path)

    def _special_properties(self, tipo, args):
        if tipo == 'link' and not 'Label' in args:
            args['Label'] = args['URL']
            return args

        if tipo == 'button':
            if 'ImageURL' in args:
                args['ImageURL'] = self._set_image_url(args['ImageURL'])
            args['FocusOnClick'] = args.get('FocusOnClick', False)
            return args

        if tipo == 'roadmap':
            args['Height'] = args.get('Height', self.height)
            if 'Title' in args:
                args['Text'] = args.pop('Title')
            return args

        if tipo == 'tree':
            args['SelectionType'] = args.get('SelectionType', SINGLE)
            return args

        if tipo == 'grid':
            args['ShowRowHeader'] = args.get('ShowRowHeader', True)
            return args

        if tipo == 'pages':
            args['Width'] = args.get('Width', self.width)
            args['Height'] = args.get('Height', self.height)

        return args

UNO_CLASSES = {
    'label': UnoLabel,
    'link': UnoLabelLink,
    'button': UnoButton,
    'radio': UnoRadio,
    'checkbox': UnoCheckBox,
    'text': UnoText,
    'image': UnoImage,
    'listbox': UnoListBox,
    'roadmap': UnoRoadmap,
    'tree': UnoTree,
    'grid': UnoGrid,
    'pages': UnoPages,
}

# ~ https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1sheet.html#aa5aa6dbecaeb5e18a476b0a58279c57a
class ValidationType():
    from com.sun.star.sheet.ValidationType \
        import ANY, WHOLE, DECIMAL, DATE, TIME, TEXT_LEN, LIST, CUSTOM
VT = ValidationType

NAME = TITLE = get_app_config('org.openoffice.Setup/Product', 'ooName')

def msgbox(message, title='Анализ текста', buttons=MSG_BUTTONS.BUTTONS_OK, type_msg='infobox'):
    """ Create message box
        type_msg: infobox, warningbox, errorbox, querybox, messbox
        http://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XMessageBoxFactory.html
    """
    toolkit = create_instance('com.sun.star.awt.Toolkit')
    parent = toolkit.getDesktopWindow()
    box = toolkit.createMessageBox(parent, type_msg, buttons, title, str(message))
    return box.execute()


# ~ 'CurrencyField': 'com.sun.star.awt.UnoControlCurrencyFieldModel',
# ~ 'DateField': 'com.sun.star.awt.UnoControlDateFieldModel',
# ~ 'FileControl': 'com.sun.star.awt.UnoControlFileControlModel',
# ~ 'FormattedField': 'com.sun.star.awt.UnoControlFormattedFieldModel',
# ~ 'NumericField': 'com.sun.star.awt.UnoControlNumericFieldModel',
# ~ 'PatternField': 'com.sun.star.awt.UnoControlPatternFieldModel',
# ~ 'ProgressBar': 'com.sun.star.awt.UnoControlProgressBarModel',
# ~ 'ScrollBar': 'com.sun.star.awt.UnoControlScrollBarModel',
# ~ 'SimpleAnimation': 'com.sun.star.awt.UnoControlSimpleAnimationModel',
# ~ 'SpinButton': 'com.sun.star.awt.UnoControlSpinButtonModel',
# ~ 'Throbber': 'com.sun.star.awt.UnoControlThrobberModel',
# ~ 'TimeField': 'com.sun.star.awt.UnoControlTimeFieldModel',


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
        'roadmap': 'com.sun.star.awt.UnoControlRoadmapModel',
        'tree': 'com.sun.star.awt.tree.TreeControlModel',
        'grid': 'com.sun.star.awt.grid.UnoControlGridModel',
        'pages': 'com.sun.star.awt.UnoMultiPageModel',
        'groupbox': 'com.sun.star.awt.UnoControlGroupBoxModel',
        'combobox': 'com.sun.star.awt.UnoControlComboBoxModel',
    }

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
                # ~ uid = docs.active.uid
                # ~ url = f'vnd.sun.star.tdoc:/{uid}/Dialogs/{library}/{name}.xml'
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

    def _get_type_control(self, name):
        name = name.split('.')[2]
        types = {
            'UnoFixedTextControl': 'label',
            'UnoEditControl': 'text',
            'UnoButtonControl': 'button',
        }
        return types[name]

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
    # @id.setter
    # def id(self, value):
    #     self._id = value
    #     self._path = _P.from_id(value)

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

    def _connect_listeners(self):
        for control in self.obj.Controls:
            _add_listeners(self.events, control, control.Model.Name)
        return

    # def _set_image_url(self, image):
    #     if _P.exists(image):
    #         return _P.to_url(image)

    #     path = _P.join(self._path, DIR['images'], image)
    #     return _P.to_url(path)

    def _special_properties(self, tipo, args):
        if tipo == 'link' and not 'Label' in args:
            args['Label'] = args['URL']
            return args

        if tipo == 'button':
            if 'ImageURL' in args:
                args['ImageURL'] = self._set_image_url(args['ImageURL'])
            args['FocusOnClick'] = args.get('FocusOnClick', False)
            return args

        if tipo == 'roadmap':
            args['Height'] = args.get('Height', self.height)
            if 'Title' in args:
                args['Text'] = args.pop('Title')
            return args

        if tipo == 'tree':
            args['SelectionType'] = args.get('SelectionType', SINGLE)
            return args

        if tipo == 'grid':
            args['ShowRowHeader'] = args.get('ShowRowHeader', True)
            return args

        if tipo == 'pages':
            args['Width'] = args.get('Width', self.width)
            args['Height'] = args.get('Height', self.height)

        return args

    def add_control(self, args):
        tipo = args.pop('Type').lower()
        root = args.pop('Root', '')
        sheets = args.pop('Sheets', ())
        columns = args.pop('Columns', ())

        args = self._special_properties(tipo, args)
        model = self.model.createInstance(self.MODELS[tipo])
        _set_properties(model, args)
        name = args['Name']
        self.model.insertByName(name, model)
        control = self.obj.getControl(name)
        _add_listeners(self.events, control, name)
        control = UNO_CLASSES[tipo](control)

        if tipo in ('listbox',):
            control.path = self.path

        if tipo == 'tree' and root:
            control.root = root
        elif tipo == 'grid' and columns:
            control.columns = columns
        elif tipo == 'pages' and sheets:
            control.sheets = sheets
            control.events = self.events

        setattr(self, name, control)
        self._controls[name] = control
        return control

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

    def open(self, modal=True):
        self._modal = modal
        if modal:
            return self.obj.execute()
        else:
            self.visible = True
        return

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

# class classproperty:
#     def __init__(self, method=None):
#         self.fget = method

#     def __get__(self, instance, cls=None):
#         return self.fget(cls)

#     def getter(self, method):
#         self.fget = method
#         return self

class Paths(object):
    FILE_PICKER = 'com.sun.star.ui.dialogs.FilePicker'
    FOLDER_PICKER = 'com.sun.star.ui.dialogs.FolderPicker'

    def __init__(self, path=''):
        if path.startswith('file://'):
            path = str(Path(uno.fileUrlToSystemPath(path)).resolve())
        self._path = Path(path)

#     @property
#     def path(self):
#         return str(self._path.parent)

#     @property
#     def file_name(self):
#         return self._path.name

#     @property
#     def name(self):
#         return self._path.stem

#     @property
#     def ext(self):
#         return self._path.suffix[1:]

#     @property
#     def info(self):
#         return self.path, self.file_name, self.name, self.ext

#     @property
#     def url(self):
#         return self._path.as_uri()

#     @property
#     def size(self):
#         return self._path.stat().st_size

#     @classproperty
#     def home(self):
#         return str(Path.home())

#     @classproperty
#     def documents(self):
#         return self.config()

#     @classproperty
#     def temp_dir(self):
#         return tempfile.gettempdir()

#     @classproperty
#     def python(self):
#         if IS_WIN:
#             path = self.join(self.config('Module'), PYTHON)
#         elif IS_MAC:
#             path = self.join(self.config('Module'), '..', 'Resources', PYTHON)
#         else:
#             path = sys.executable
#         return path

#     @classmethod
#     def dir_tmp(self, only_name=False):
#         dt = tempfile.TemporaryDirectory()
#         if only_name:
#             dt = dt.name
#         return dt

#     @classmethod
#     def tmp(cls, ext=''):
#         tmp = tempfile.NamedTemporaryFile(suffix=ext)
#         return tmp.name

#     @classmethod
#     def save_tmp(cls, data):
#         path_tmp = cls.tmp()
#         cls.save(path_tmp, data)
#         return path_tmp

#     @classmethod
#     def config(cls, name='Work'):
#         """
#             Return de path name in config
#             http://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XPathSettings.html
#         """
#         path = create_instance('com.sun.star.util.PathSettings')
#         return cls.to_system(getattr(path, name))

#     @classmethod
#     def get(cls, init_dir='', filters: str=''):
#         """
#             Options: http://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1ui_1_1dialogs_1_1TemplateDescription.html
#             filters: 'xml' or 'txt,xml'
#         """
#         if not init_dir:
#             init_dir = cls.documents
#         init_dir = cls.to_url(init_dir)
#         file_picker = create_instance(cls.FILE_PICKER)
#         file_picker.setTitle('Select path')
#         file_picker.setDisplayDirectory(init_dir)
#         file_picker.initialize((2,))
#         if filters:
#             filters = [(f.upper(), f'*.{f.lower()}') for f in filters.split(',')]
#             file_picker.setCurrentFilter(filters[0][0])
#             for f in filters:
#                 file_picker.appendFilter(f[0], f[1])

#         path = ''
#         if file_picker.execute():
#             path =  cls.to_system(file_picker.getSelectedFiles()[0])
#         return path

#     @classmethod
#     def get_dir(cls, init_dir=''):
#         folder_picker = create_instance(cls.FOLDER_PICKER)
#         if not init_dir:
#             init_dir = cls.documents
#         init_dir = cls.to_url(init_dir)
#         folder_picker.setTitle('Select directory')
#         folder_picker.setDisplayDirectory(init_dir)

#         path = ''
#         if folder_picker.execute():
#             path = cls.to_system(folder_picker.getDirectory())
#         return path

#     @classmethod
#     def get_file(cls, init_dir: str='', filters: str='', multiple: bool=False):
#         """
#             init_folder: folder default open
#             multiple: True for multiple selected
#             filters: 'xml' or 'xml,txt'
#         """
#         if not init_dir:
#             init_dir = cls.documents
#         init_dir = cls.to_url(init_dir)

#         file_picker = create_instance(cls.FILE_PICKER)
#         file_picker.setTitle('Select file')
#         file_picker.setDisplayDirectory(init_dir)
#         file_picker.setMultiSelectionMode(multiple)

#         if filters:
#             filters = [(f.upper(), f'*.{f.lower()}') for f in filters.split(',')]
#             file_picker.setCurrentFilter(filters[0][0])
#             for f in filters:
#                 file_picker.appendFilter(f[0], f[1])

#         path = ''
#         if file_picker.execute():
#             files = file_picker.getSelectedFiles()
#             path = [cls.to_system(f) for f in files]
#             if not multiple:
#                 path = path[0]
#         return path

#     @classmethod
#     def replace_ext(cls, path, new_ext):
#         p = Paths(path)
#         name = f'{p.name}.{new_ext}'
#         path = cls.join(p.path, name)
#         return path

#     @classmethod
#     def exists(cls, path):
#         result = False
#         if path:
#             path = cls.to_system(path)
#             result = Path(path).exists()
#         return result

#     @classmethod
#     def exists_app(cls, name_app):
#         return bool(shutil.which(name_app))

#     @classmethod
#     def open(cls, path):
#         if IS_WIN:
#             os.startfile(path)
#         else:
#             pid = subprocess.Popen(['xdg-open', path]).pid
#         return

#     @classmethod
#     def is_dir(cls, path):
#         return Path(path).is_dir()

#     @classmethod
#     def is_file(cls, path):
#         return Path(path).is_file()

#     @classmethod
#     def join(cls, *paths):
#         return str(Path(paths[0]).joinpath(*paths[1:]))

#     @classmethod
#     def save(cls, path, data, encoding='utf-8'):
#         result = bool(Path(path).write_text(data, encoding=encoding))
#         return result

#     @classmethod
#     def save_bin(cls, path, data):
#         result = bool(Path(path).write_bytes(data))
#         return result

#     @classmethod
#     def read(cls, path, encoding='utf-8'):
#         data = Path(path).read_text(encoding=encoding)
#         return data

#     @classmethod
#     def read_bin(cls, path):
#         data = Path(path).read_bytes()
#         return data

    @classmethod
    def to_url(cls, path):
        if not path.startswith('file://'):
            path = Path(path).as_uri()
        return path

#     @classmethod
#     def to_system(cls, path):
#         if path.startswith('file://'):
#             path = str(Path(uno.fileUrlToSystemPath(path)).resolve())
#         return path

#     @classmethod
#     def kill(cls, path):
#         result = True
#         p = Path(path)

#         try:
#             if p.is_file():
#                 p.unlink()
#             elif p.is_dir():
#                 shutil.rmtree(path)
#         except OSError as e:
#             result = False

#         return result

#     @classmethod
#     def files(cls, path, pattern='*'):
#         files = [str(p) for p in Path(path).glob(pattern) if p.is_file()]
#         return files

#     @classmethod
#     def dirs(cls, path):
#         dirs = [str(p) for p in Path(path).iterdir() if p.is_dir()]
#         return dirs

#     @classmethod
#     def walk(cls, path, filters=''):
#         paths = []
#         if filters in ('*', '*.*'):
#             filters = ''
#         for folder, _, files in os.walk(path):
#             if filters:
#                 pattern = re.compile(r'\.(?:{})$'.format(filters), re.IGNORECASE)
#                 paths += [cls.join(folder, f) for f in files if pattern.search(f)]
#             else:
#                 paths += [cls.join(folder, f) for f in files]
#         return paths

#     @classmethod
#     def walk_dir(cls, path, tree=False):
#         folders = []
#         if tree:
#             i = 0
#             p = 0
#             parents = {path: 0}
#             for root, dirs, _ in os.walk(path):
#                 for name in dirs:
#                     i += 1
#                     rn = cls.join(root, name)
#                     if not rn in parents:
#                         parents[rn] = i
#                     folders.append((i, parents[root], name))
#         else:
#             for root, dirs, _ in os.walk(path):
#                 folders += [cls.join(root, name) for name in dirs]
#         return folders

#     @classmethod
#     def from_id(cls, id_ext):
#         pip = CTX.getValueByName('/singletons/com.sun.star.deployment.PackageInformationProvider')
#         path = _P.to_system(pip.getPackageLocation(id_ext))
#         return path

#     @classmethod
#     def from_json(cls, path):
#         data = json.loads(cls.read(path))
#         return data

#     @classmethod
#     def to_json(cls, path, data):
#         data = json.dumps(data, indent=4, ensure_ascii=False, sort_keys=True)
#         return cls.save(path, data)

#     @classmethod
#     def from_csv(cls, path, args={}):
#         # ~ See https://docs.python.org/3.7/library/csv.html#csv.reader
#         with open(path) as f:
#             rows = tuple(csv.reader(f, **args))
#         return rows

#     @classmethod
#     def to_csv(cls, path, data, args={}):
#         with open(path, 'w') as f:
#             writer = csv.writer(f, **args)
#             writer.writerows(data)
#         return

#     @classmethod
#     def zip(cls, source, target='', pwd=''):
#         path_zip = target
#         if not isinstance(source, (tuple, list)):
#             path, _, name, _ = _P(source).info
#             start = len(path) + 1
#             if not target:
#                 path_zip = f'{path}/{name}.zip'

#         if isinstance(source, (tuple, list)):
#             files = [(f, f[len(_P(f).path)+1:]) for f in source]
#         elif _P.is_file(source):
#             files = ((source, source[start:]),)
#         else:
#             files = [(f, f[start:]) for f in _P.walk(source)]

#         compression = zipfile.ZIP_DEFLATED
#         with zipfile.ZipFile(path_zip, 'w', compression=compression) as z:
#             for f in files:
#                 z.write(f[0], f[1])
#         return

#     @classmethod
#     def zip_content(cls, path):
#         with zipfile.ZipFile(path) as z:
#             names = z.namelist()
#         return names

#     @classmethod
#     def unzip(cls, source, target='', members=None, pwd=None):
#         path = target
#         if not target:
#             path = _P(source).path
#         with zipfile.ZipFile(source) as z:
#             if not pwd is None:
#                 pwd = pwd.encode()
#             if isinstance(members, str):
#                 members = (members,)
#             z.extractall(path, members=members, pwd=pwd)
#         return True

#     @classmethod
#     def merge_zip(cls, target, zips):
#         try:
#             with zipfile.ZipFile(target, 'w', compression=zipfile.ZIP_DEFLATED) as t:
#                 for path in zips:
#                     with zipfile.ZipFile(path, compression=zipfile.ZIP_DEFLATED) as s:
#                         for name in s.namelist():
#                             t.writestr(name, s.open(name).read())
#         except Exception as e:
#             return False

#         return True

#     @classmethod
#     def image(cls, path):
#         gp = create_instance('com.sun.star.graphic.GraphicProvider')
#         image = gp.queryGraphic((
#             PropertyValue(Name='URL', Value=cls.to_url(path)),
#         ))
#         return image

#     @classmethod
#     def copy(cls, source, target='', name=''):
#         p, f, n, e = _P(source).info
#         if target:
#             p = target
#         if name:
#             e = ''
#             n = name
#         path_new = cls.join(p, f'{n}{e}')
#         shutil.copy(source, path_new)
#         return path_new
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

def run(command, capture=False, split=True):
    if not split:
        return subprocess.check_output(command, shell=True).decode()

    cmd = shlex.split(command)
    result = subprocess.run(cmd, capture_output=capture, text=True, shell=IS_WIN)
    if capture:
        result = result.stdout
    else:
        result = result.returncode
    return result
