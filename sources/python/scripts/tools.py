# -*- coding: utf-8 -*-


def console(*args, **kwargs):
    '''
    Launch the "python interpreter" gui.

    Positional arguments are only intended for arguments automatically
    passed by the program gui.

    Keyword arguments are:
    - 'loc': for passing caller's locales and/or globals to the console context
    - any constructor constant (BACKGROUND, FOREGROUND...) to tweak the console aspect.

    Examples:
    - console()  # defaut constructor)
    - console(loc=locals())
    - console(BACKGROUND=0x0, FOREGROUND=0xFFFFFF)

    More infos: https://extensions.libreoffice.org/en/extensions/show/apso-alternative-script-organizer-for-python.
    '''

    # we need to load apso before import statement
    ctx = XSCRIPTCONTEXT.getComponentContext()
    ctx.ServiceManager.createInstance("apso.python.script.organizer.impl")
    # now we can use apso_utils library
    from apso_utils import console
    from pathlib import Path
    from uno import fileUrlToSystemPath
    import sys
    import os

    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    __file__ = os.path.join(str(Path.home()),"python") if doc.Location == "" else doc.Location
    try:
        os.chdir(fileUrlToSystemPath(os.path.dirname(__file__)))
    except:
        os.chdir(os.path.dirname(__file__))
    sys.path.append(".")
    kwargs.setdefault('loc', {})
    kwargs['loc'].setdefault('XSCRIPTCONTEXT', XSCRIPTCONTEXT)
    kwargs['loc'].setdefault('__file__', __file__)
    kwargs['loc'].setdefault('os', os)
    kwargs['loc'].setdefault('sys', sys)
    kwargs.setdefault('BACKGROUND',0x0)
    kwargs.setdefault('FOREGROUND',0xFFFFFF)
    kwargs.setdefault('WIDTH',1000)
    console(**kwargs)

g_exportedScripts = console,
