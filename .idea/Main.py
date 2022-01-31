
import imp


try:
    fp, pathname, description = imp.find_module("gtr")
    _mod = imp.load_module("gtr", fp, pathname, description)
    _mod._init_()
except Exception as exc:
    print(exc)
    input();


    # компилировать в exe командой pyinstaller .idea/Main.py --hidden-import 'gtr'
    # модуль gtr должен включать все библиотеки, используемые остальными модулями