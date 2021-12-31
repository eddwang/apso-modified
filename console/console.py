# -*- coding: utf-8 -*-

import uno
import unohelper

def createUnoService(service, ctx=None, args = None):
    smgr = uno.getComponentContext().ServiceManager
    if ctx and args:
        return smgr.createInstanceWithArgumentsAndContext(service, args, ctx)
    elif args:
        return smgr.createInstanceWithArguments(service, args)
    elif ctx:
        return smgr.createInstanceWithContext(service, ctx)
    else:
        return smgr.createInstance(service)


#------------XRAY----------------
from com.sun.star.script.provider import ScriptFrameworkErrorException
def xray(obj):
    '''Macro to call Basic XRay by Bernard Marcelly from Python.'''
    print('Loading Xray...')
    ctx = uno.getComponentContext()
    url = ("vnd.sun.star.script:XRayTool._Main.Xray?language=Basic&location=application")
    mspf = createUnoService("com.sun.star.script.provider.MasterScriptProviderFactory", ctx)
    try:
        script = mspf.createScriptProvider('').getScript(url)
        script.invoke((obj,), (), ())
    except ScriptFrameworkErrorException:
        print('Xray is not installed.\nPlease visit:\n'
              'http://www.openoffice.org/fr/Documentation/Basic')


#------------MRI----------------
def mri(target):
    '''Macro to instantiate Mri introspection tool from hanya.'''
    print('Loading MRI...')
    ctx = uno.getComponentContext()
    mri = ctx.ServiceManager.createInstanceWithContext("mytools.Mri",ctx)
    try:
        mri.inspect(target)
    except AttributeError:
        print('MRI is not installed.\nPlease visit:\nhttp://extensions.services.openoffice.org')


#------------MSGBOX----------------
from com.sun.star.awt.MessageBoxType import (MESSAGEBOX, INFOBOX,
                                             ERRORBOX, WARNINGBOX, QUERYBOX)
def msgbox(message, titre="Message", boxtype='message', boutons=1, frame=None):
    types = {'message': MESSAGEBOX, 'info': INFOBOX, 'error': ERRORBOX,
             'warning': WARNINGBOX, 'query': QUERYBOX}
    if hasattr(builtins, '__console__'):
        win = builtins.__console__.dialog
        tk = win.Toolkit
    else:
        ctx = uno.getComponentContext()
        tk = createUnoService("com.sun.star.awt.Toolkit", ctx)
        if not frame:
            desktop = createUnoService("com.sun.star.frame.Desktop", ctx)
            frame = desktop.ActiveFrame
            if frame.ActiveFrame:
                # top window is a subdocument
                frame = frame.ActiveFrame
        win = frame.ComponentWindow
    box = tk.createMessageBox(win, types[boxtype], boutons, titre, message)
    box.execute()


#------------CONSOLE----------------
import sys
import pdb
import code
import traceback
import threading
import pythonscript
try:
    import queue
    import builtins
except ImportError:
    import Queue as queue
    import __builtin__ as builtins
from com.sun.star.awt import (XKeyHandler, XFocusListener, XWindowListener,
         XTextListener, WindowDescriptor, FontDescriptor, Selection, Rectangle)
from com.sun.star.awt.WindowClass import TOP
from com.sun.star.awt.WindowAttribute import (
                                    MOVEABLE, SIZEABLE, BORDER, CLOSEABLE, SHOW)
from com.sun.star.awt.KeyModifier import SHIFT, MOD1
from com.sun.star.awt.PosSize import SIZE

if sys.version_info < (3, ):
    reload(sys)
    sys.setdefaultencoding(sys.getfilesystemencoding())

EOT = b'\x04'

class UnoScriptImporter(object):
    def __init__(self, ctx):
        self.ctx = ctx
        self.providers = self._load_providers()
        self.nodes = {}

    def _load_providers(self):
        p = {}
        locations = [u'user', u'user:uno_packages', u'share', u'share:uno_packages']
        for location in locations:
            ext = ""
            if location.endswith('uno_packages'):
                ext = ".oxt"
            p[location] = (pythonscript.PythonScriptProvider(self.ctx, location), ext)
        doc = XSCRIPTCONTEXT.getDocument()
        # FIXME: check if fonctional from base subdocuments
        try:
            if doc.ScriptContainer:
                p['document'] = (pythonscript.PythonScriptProvider(self.ctx, doc), "")
        except AttributeError:
            pass
        return p

    def _find_module(self, name, path):
        # print('---_find_module---')
        if path:
            node = path[0]
            if node in self.nodes:
                return self._search_node(self.nodes[node], name)
            else:
                return False
        else:
            for prov in self.providers:
                sp, ext = self.providers[prov]
                if self._search_node(sp, name, ext):
                    self.location = prov
                    return True
        return False

    def _search_node(self, node, name, ext=''):
        # print('---_search_node---')
        for child in node.getChildNodes():
            if child.name == uno.systemPathToFileUrl(name+ext):
                self.nodes[self.fullname] = child
                return True
        return False

    def find_module(self, fullname, path=None):
        # print('_____FIND_MODULE_____')
        # print('fullname: ' + fullname)
        if fullname in ('com',):
            return None
        self.fullname = fullname
        name = fullname.rsplit('.', 1)[-1]
        if self._find_module(name, path):
            return self
        return None

    def _module_from_node(self, node):
        import imp
        mod = imp.new_module(node.name)
        for child in node.getChildNodes():
            try:
                if isinstance(child, pythonscript.FileBrowseNode):
                    setattr(mod, child.name, node.provCtx.getModuleByUrl(child.uri))
                else:
                    setattr(mod, child.name, self._module_from_node(child))
            # import machinery could find a submodule that is
            # not visible from the PythonScriptProvider
            except ImportError:
                print('unexpected error while loading module <{}>'.format(child.name))
                # traceback.print_exc()
        return mod

    def load_module(self, fullname):
        # print('_____LOAD_MODULE_____')
        # print('fullname: ' + fullname)
        # print('self.nodes: ' + str(self.nodes))
        node = self.nodes[fullname]
        if isinstance(node, pythonscript.DirBrowseNode):
            mod = self._module_from_node(node)
            mod.__file__ = '<{}>'.format(self.location)
            mod.__path__ = [fullname]
            mod.__package__ = fullname
        else:
            mod = node.provCtx.getModuleByUrl(node.uri)
        try:
            parent = fullname.rsplit('.', 1)[-2]
            mod.__package__ = parent
        except IndexError:
            pass
        # name = fullname.rsplit('.', 1)[-1]
        mod.__name__ = fullname
        mod.__loader__ = self
        sys.modules[fullname] = mod
        return mod


class ConsoleWindow(object):
    """Interactive console dialog.
    Instantiate with corresponding keyword parameters to override
    default values: ConsoleWindow(BACKGROUND=0x0, FOREGROUND=0xFFFFFF)>"""

    FONT = "DejaVu Sans Mono"
    BACKGROUND = 0xFDF6E3
    FOREGROUND = 0x657B83
    MARGIN = 3
    BUTTON_WIDTH  = 80
    BUTTON_HEIGHT = 26
    EDIT_HEIGHT = 300
    WIDTH = 600
    PS1 = '>>> '
    PS2 = '... '
    NBTAB = 4
    HEIGHT = EDIT_HEIGHT + MARGIN * 3

    def __init__(self, **kwargs):
        self.banner = None
        for key in kwargs:
            setattr(self, key, kwargs[key])
        self.loc.update(pdb=pdb)
        if not 'ctx' in kwargs:
            self.ctx = uno.getComponentContext()
        self.smgr = self.ctx.getServiceManager()
        self.title = "PYTHON console"
        self.edit_name = "console"
        self.product = self.getproduct()
        self.dialog = None
        self.parent = self.getparent()
        self.tk = self.parent.Toolkit
        self.end = 0
        self.more = 0
        self.history = []
        self.historycursor = 0

    def create(self, name, arguments=None):
        """ Create service instance. """
        if arguments:
            return self.smgr.createInstanceWithArgumentsAndContext(name, arguments, self.ctx)
        else:
            return self.smgr.createInstanceWithContext(name, self.ctx)

    def execute(self):
        '''Create dialog and manage context'''
        self._init()
        proc = threading.Thread(target=run_interact, args=(self.inqueue, self.exitevent,
                                                    self.PS1, self.PS2, self.product, self.loc))
        proc.start()
        self.dialog.execute()
        # self.dialog.dispose()

    def __enter__(self):
        # create import hook for macro
        self.importer = UnoScriptImporter(self.ctx)
        sys.meta_path.append(self.importer)
        # redirect output
        self.stdout = sys.stdout
        self.stderr = sys.stderr
        sys.stdout = self
        sys.stderr = self
        return self

    def __exit__(self, type, value, traceback):
        del builtins.__console__
        self.inqueue.put(EOT)
        self.tk.removeKeyHandler(self.keyhandler)
        sys.meta_path.remove(self.importer)
        sys.stdout = self.stdout
        sys.stderr = self.stderr

    def enddialog(self):
        """Terminate dialog"""
        self.dialog.endDialog(0)

    def create_control(self, win, name, type, pos, size, prop_names, prop_values, full_name=False):
        """ Create and insert control. """
        if not full_name:
            type = "com.sun.star.awt.UnoControl" + type + "Model"
        model = self.create("com.sun.star.awt.UnoControlEditModel")
        if prop_names and prop_values:
            model.setPropertyValues(prop_names, prop_values)
        ctrl = self.create(model.DefaultControl)
        ctrl.setModel(model)
        ctrl.createPeer(win.Toolkit, win)
        ctrl.setPosSize(pos[0], pos[1], size[0], size[1], 15)
        return ctrl

    def create_dialog(self, title, size=None, parent=None):
        """ Create modeless dialog. """
        rect = Rectangle()
        rect.Width, rect.Height = size
        ps = self.parent.getPosSize()
        rect.X, rect.Y = (ps.Width - self.WIDTH)/2 -50, (ps.Height - self.HEIGHT)/2 -100

        desc = WindowDescriptor()
        desc.Type = TOP
        desc.WindowServiceName = "dialog"
        desc.ParentIndex = -1
        desc.Bounds = rect
        desc.WindowAttributes = SHOW | MOVEABLE | SIZEABLE | CLOSEABLE | BORDER

        dialog = self.tk.createWindow(desc)
        dialog.setTitle(title)

        self.dialog = dialog

    def create_edit(self, name, pos, size, prop_names=None, prop_values=None):
        """ Create and add new edit control. """
        self.edit = self.create_control(self.dialog, name, "Edit", pos, size,
                                                   prop_names, prop_values)

    class ListenerBase(unohelper.Base):
        def __init__(self, act):
            self.act = act
        def disposing(self, source):
            self.act = None

    class WindowListener(ListenerBase, XWindowListener):
        def windowResized(self, e):
            size = e.Source.Size
            margin = self.act.MARGIN
            e.Source.Windows[0].setPosSize(0, 0, size.Width - margin*2,
                                                 size.Height - margin*3, SIZE)

    class FocusListener(ListenerBase, XFocusListener):
        def focusGained(self, ev):
            self.act.tk.addKeyHandler(self.act.keyhandler)
        def focusLost(self, ev):
            self.act.tk.removeKeyHandler(self.act.keyhandler)

    class TextListener(ListenerBase, XTextListener):
        def textChanged(self, ev):
            self.act.end = len(ev.Source.Text) +1 # +1 for trailing new line character

    class KeyHandler(ListenerBase, XKeyHandler):
        def keyPressed(self, ev):
            try:
                return getattr(self.act, "onkey_"+str(ev.KeyCode))(ev.Modifiers)
            except AttributeError:
                return 0
        def keyReleased(self, ev):
            return 0

    def _init(self):
        margin = self.MARGIN
        font = FontDescriptor()
        font.Name = self.FONT
        self.create_dialog(self.title, size=(self.WIDTH, self.HEIGHT))
        self.dialog.addWindowListener(self.WindowListener(self))
        self.create_edit(self.edit_name, pos=(margin, margin * 2),
            size=(self.WIDTH - margin * 2, self.EDIT_HEIGHT),
            prop_names=("AutoVScroll", "BackgroundColor", "FontDescriptor",
                        "HideInactiveSelection", "MultiLine", "TextColor"),
            prop_values=(True, self.BACKGROUND, font, True, True, self.FOREGROUND))

        self.keyhandler = self.KeyHandler(self)
        self.edit.addTextListener(self.TextListener(self))
        self.edit.addFocusListener(self.FocusListener(self))
        self.edit.setFocus()

    def getparent(self):
        '''Returns parent frame'''
        desktop = self.create("com.sun.star.frame.Desktop")
        return desktop.ActiveFrame.ContainerWindow

    def getproduct(self):
        cp = self.create("com.sun.star.configuration.ConfigurationProvider")
        node = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        node.Name = "nodepath"
        node.Value = "/org.openoffice.Setup/Product"
        reader = cp.createInstanceWithArguments(
                "com.sun.star.configuration.ConfigurationAccess", (node,))
        return '{} {}'.format(reader.ooName, reader.ooSetupVersion)

    def onkey_514(self, modifiers):  # C
        '''Catch ctrl+C keyboard entry'''
        if (modifiers & MOD1):
            sel = self.edit.Selection
            if sel.Min == sel.Max:
                self._keyboardinterrupt()
                return 1
        return 0

    def onkey_515(self, modifiers):  # D
        '''Catch ctrl+D keyboard entry'''
        if (modifiers & MOD1):
            try:
                self.enddialog()
                return 1
            except:
                traceback.print_exc()
                self._write(self.PS1)
                return 1
        return 0

    def onkey_537(self, modifiers):  # Z
        '''Catch ctrl+Z keyboard entry'''
        if (modifiers & MOD1):
            try:
                self.enddialog()
                return 1
            except:
                traceback.print_exc()
                self._write(self.PS1)
                return 1
        return 0

    def onkey_1024(self, modifiers):  # DOWN
        '''Catch DOWN keyboard entry'''
        try:
            line = self._readline()
            if self.historycursor < len(self.history)-1:
                self.historycursor +=1
                self._write(self.history[self.historycursor], (self.end-len(line), self.end))
            else:
                self.historycursor = len(self.history)
                self._write("", (self.end-len(line), self.end))
            self.gotoendofinput()
            return 1
        except Exception as e:
            self.inqueue.put(traceback.format_exc())

    def onkey_1025(self, modifiers):  # UP
        '''Catch UP keyboard entry'''
        try:
            line = self._readline()
            if self.historycursor > 0:
                self.historycursor -= 1
                self._write(self.history[self.historycursor], (self.end-len(line), self.end))
            self.gotoendofinput()
            return 1
        except Exception as e:
            self.inqueue.put(traceback.format_exc())

    def onkey_1028(self, modifiers):  # HOME
        '''Catch HOME keyboard entry'''
        if not (modifiers & SHIFT):
            self.gotostartofinput()
            return 1
        return 0

    def onkey_1029(self, modifiers):  # END
        '''Catch END keyboard entry'''
        if not (modifiers & SHIFT):
            self.gotoendofinput()
            return 1
        return 0

    def onkey_1280(self, modifiers):  # RETURN
        '''Catch RETURN keyboard entry'''
        try:
            line = self._readline()
            cmd = line.rstrip('\n')
            if cmd:
                if not self.history or cmd != self.history[-1]:
                    self.history.append(cmd)
                self.historycursor = len(self.history)
            if self.edit.Selection.Max >= (self.end-len(self.prompt+line)):
                if cmd in ("clear", "clear()"):
                    self.clear()
                else:
                    self._write("\n")
                    self.prompt = ""
                    self.inqueue.put(line)
            self.gotoendofinput()
            return 1
        except:
            self.inqueue.put(traceback.format_exc())

    def onkey_1281(self, modifiers):  # ESCAPE
        '''Catch ESCAPE keyboard entry'''
        return 1

    def onkey_1282(self, modifiers):  # TAB
        '''Catch TAB keyboard entry'''
        if not modifiers:
            self._write(' ' * self.NBTAB)
            return 1
        return 0

    def _keyboardinterrupt(self):
        '''Send KeyboardInterror exception.
        This exception will not allways work as expected'''
        self.inqueue.put(KeyboardInterrupt)

    def gotoendofinput(self):
        '''Send visible cursor to end of input line'''
        self.edit.setSelection(Selection(self.end, self.end))

    def gotostartofinput(self):
        '''Send visible cursor to start of input line'''
        line = self._readline()
        pos = self.end - len(line)
        self.edit.setSelection(Selection(pos, pos))

    def clear(self):
        '''Clear edit area'''
        self.edit.Text = self.prompt

    def write(self, data):
        '''Implements sys.stdout/stderr readline method'''
        try:
            if self.exitevent.is_set():
                self.enddialog()
                return
            self.prompt = data.split("\n")[-1]
            self._write(data)
        except:
            msgbox("Error on ConsoleWindow.write:\n" + traceback.format_exc())

    def _write(self, data, sel=None):
        '''Append data to edit control text'''
        if not sel:
            sel = (self.end, self.end)
        self.edit.insertText(Selection(*sel), data)

    def _readline(self):
        '''Returns input text'''
        lines = self.edit.Text.rsplit("\n", 1)
        line = lines[-1] + '\n'
        return line[len(self.prompt):]


class InteractiveConsole(code.InteractiveConsole):
    def __init__(self, inqueue, exitevent, ps1, ps2, product, loc=None):
        code.InteractiveConsole.__init__(self, loc)
        self.inqueue = inqueue
        self.exitevent = exitevent
        self.ps1, self.ps2 = ps1, ps2
        self.product = product
        self.keep_prompting = True
        self.stdin = sys.stdin
        sys.stdin = self

    def readline(self):
        '''Implements sys.stdin readline method'''
        rl = self.inqueue.get()
        if rl == KeyboardInterrupt:
            raise KeyboardInterrupt()
        if rl == EOT:
            self.keep_prompting = False
        return rl

    def interact(self, banner=None):
        '''Overwrite default interact method'''
        cprt = 'Type "help", "copyright", "credits" or "license" for more information.'
        if banner is None:
            self.write("PYTHON console [{}]\n{}\n{}\n".format(self.product, sys.version, cprt))
        elif banner:
            self.write("%s\n" % str(banner))
        more = 0
        try:
            while self.keep_prompting:
                try:
                    if more:
                        prompt = self.ps2
                    else:
                        prompt = self.ps1
                    line = self.raw_input(prompt)
                    more = self.push(line)
                except KeyboardInterrupt:
                    self.write("\nKeyboardInterrupt\n")
                    self.resetbuffer()
                    more = 0
                except SystemExit:
                    self.exitevent.set()
                except:
                    print('INTERACT\n')
                    traceback.print_exc()
                    self.resetbuffer()
                    more = 0
        finally:
            sys.stdin = self.stdin


def run_interact(*args):
    '''Start the"python interpreter" engine'''
    iconsole = InteractiveConsole(*args)
    iconsole.interact()


def console(event=None):
    '''Start the"python interpreter" gui'''
    if hasattr(builtins, '__console__'):
        __console__.dialog.toFront()
        return
    # test = "locals passed"
    inqueue = queue.Queue()
    exitevent = threading.Event()
    loc = locals()
    loc.update(globals())
    try:
        with ConsoleWindow(inqueue=inqueue, exitevent=exitevent, loc=loc) as console_:
            builtins.__console__ = console_
            console_.execute()
    except:
        msgbox(traceback.format_exc())


g_exportedScripts = console,
