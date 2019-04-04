"""A command module for Dragonfly, for generic editing help.

-----------------------------------------------------------------------------
This is a heavily modified version of the _multiedit-en.py script at:
http://dragonfly-modules.googlecode.com/svn/trunk/command-modules/documentation/mod-_multiedit.html  # @IgnorePep8
Licensed under the LGPL, see http://www.gnu.org/licenses/

"""
from time import sleep
from natlink import setMicState
from dragonfly import (
    Alternative, AppContext,
    Choice, CompoundRule, Config, Dictation,
    FocusWindow, Function, Grammar,
    Integer, IntegerRef, Item, Key,
    MappingRule, Mimic, Mouse,
    Pause, Repetition, RuleRef,
    Section, Text, Window,
    get_engine
)

import win32con

from dragonfly.actions.keyboard import Typeable, keyboard
from dragonfly.actions.typeables import typeables

if not 'Control_R' in typeables:
    keycode = win32con.VK_RCONTROL
    typeables["Control_R"] = Typeable(code=keycode, name="Control_R")
if not 'semicolon' in typeables:
    typeables["semicolon"] = keyboard.get_typeable(char=';')

import lib.config
config = lib.config.get_config()

import lib.sound as sound
from lib.format import (
    camel_case_count,
    pascal_case_count,
    snake_case_count,
    squash_count,
    expand_count,
    uppercase_count,
    lowercase_count,
    format_text,
    FormatTypes as ft,
)

release = Key("shift:up, ctrl:up, alt:up, win:up")

################################################################################
## Disabling dictation mode

def noop(text=None):
    pass

class DictationOffRule(MappingRule):
    mapping = {
        "<text>": Function(noop),
    }
    extras = [
        Dictation("text"),
    ]
    defaults = {
    }

dictationOffGrammar = Grammar("Dictation off")
dictationOffGrammar.add_rule(DictationOffRule())
dictationOffGrammar.load()
dictationOffGrammar.disable()

dictation_enabled = True

def dictation_off():
    global dictation_enabled
    if dictation_enabled:
        print "Dictation now OFF."
        dictation_enabled = False
        dictationOffGrammar.enable()
    else:
        print "Dictation already off."

def dictation_on():
    global dictation_enabled
    if not dictation_enabled:
        print "Dictation now ON."
        dictation_enabled = True
        dictationOffGrammar.disable()
    else:
        print "Dictation already on."

################################################################################

def cancel_and_sleep(text=None, text2=None):
    """random mumbling go to sleep'" => Microphone sleep."""
    #print("* Dictation canceled. Going to sleep. *")
    #sound.play(sound.SND_DING)
    setMicState("sleeping")

def reload_natlink():
    """Reloads Natlink and custom Python modules."""
    win = Window.get_foreground()
    FocusWindow(executable="natspeak",
        title="Messages from Python Macros").execute()
    Pause("10").execute()
    Key("a-f, r").execute()
    Pause("10").execute()
    win.set_foreground()

def copy_command():
    # Add Command Prompt, putty, ...?
    context = AppContext(executable="console")
    window = Window.get_foreground()
    if context.matches(window.executable, window.title, window.handle):
        return
    release.execute()
    Key("c-c/3").execute()

def paste_command():
    # Add Command Prompt, putty, ...?
    context = AppContext(executable="console")
    window = Window.get_foreground()
    if context.matches(window.executable, window.title, window.handle):
        return
    release.execute()
    Key("c-v/3").execute()

################################################################################

formatMap = {
    "camel": ft.camelCase,
    "pascal": ft.pascalCase,
    "studly": ft.pascalCase,
    "snake": ft.snakeCase,
    "uppercase": ft.upperCase,
    "lowercase": ft.lowerCase,
    "squash": ft.squash,
    "lowercase squash": [ft.squash, ft.lowerCase],
    "uppercase squash": [ft.squash, ft.upperCase],
    "squash lowercase": [ft.squash, ft.lowerCase],
    "squash uppercase": [ft.squash, ft.upperCase],
    "dashify": ft.dashify,
    "lowercase dashify": [ft.dashify, ft.lowerCase],
    "uppercase dashify": [ft.dashify, ft.upperCase],
    "dashify lowercase": [ft.dashify, ft.lowerCase],
    "dashify uppercase": [ft.dashify, ft.upperCase],
    "dotify": ft.dotify,
    "lowercase dotify": [ft.dotify, ft.lowerCase],
    "uppercase dotify": [ft.dotify, ft.upperCase],
    "dotify lowercase": [ft.dotify, ft.lowerCase],
    "dotify uppercase": [ft.dotify, ft.upperCase],
    "say": ft.spokenForm,
    "environment variable": [ft.snakeCase, ft.upperCase],
}

abbreviationMap = {
    "administrator": "admin",
    "administrators": "admins",
    "application": "app",
    "applications": "apps",
    "argument": "arg",
    "arguments": "args",
    "attribute": "attr",
    "attributes": "attrs",
    "(authenticate|authentication)": "auth",
    "binary": "bin",
    "button": "btn",
    "class": "cls",
    "command": "cmd",
    "(config|configuration)": "cfg",
    "context": "ctx",
    "control": "ctrl",
    "database": "db",
    "(define|definition)": "def",
    "description": "desc",
    "(develop|development)": "dev",
    "(dictionary|dictation)": "dict",
    "(direction|directory)": "dir",
    "dynamic": "dyn",
    "example": "ex",
    "execute": "exec",
    "exception": "exc",
    "expression": "exp",
    "(extension|extend)": "ext",
    "function": "func",
    "framework": "fw",
    "(initialize|initializer)": "init",
    "instance": "inst",
    "integer": "int",
    "iterate": "iter",
    "java archive": "jar",
    "javascript": "js",
    "keyword": "kw",
    "keyword arguments": "kwargs",
    "language": "lng",
    "library": "lib",
    "length": "len",
    "number": "num",
    "object": "obj",
    "okay": "ok",
    "package": "pkg",
    "parameter": "param",
    "parameters": "params",
    "pixel": "px",
    "position": "pos",
    "point": "pt",
    "previous": "prev",
    "property": "prop",
    "python": "py",
    "query string": "qs",
    "reference": "ref",
    "references": "refs",
    "(represent|representation)": "repr",
    "regular (expression|expressions)": "regex",
    "request": "req",
    "revision": "rev",
    "ruby": "rb",
    "session aidee": "sid",  # "session id" didn't work for some reason.
    "source": "src",
    "(special|specify|specific|specification)": "spec",
    "standard": "std",
    "standard in": "stdin",
    "standard out": "stdout",
    "string": "str",
    "(synchronize|synchronous)": "sync",
    "system": "sys",
    "utility": "util",
    "utilities": "utils",
    "temporary": "tmp",
    "text": "txt",
    "value": "val",
    "window": "win",
}

# For use with "say"-command. Words that are commands in the generic edit
# grammar were treated as separate commands and could not be written with the
# "say"-command. This overrides that behavior.
# Other words that won't work for one reason or another, can also be added to
# this list.
reservedWord = {
    "up": "up",
    "down": "down",
    "left": "left",
    "right": "right",
    "home": "home",
    "end": "end",
    "space": "space",
    "tab": "tab",
    "backspace": "backspace",
    "delete": "delete",
    "enter": "enter",
    "paste": "paste",
    "copy": "copy",
    "cut": "cut",
    "undo": "undo",
    "release": "release",
    "page up": "page up",
    "page down": "page down",
    "say": "say",
    "select": "select",
    "select all": "select all",
    "abbreviate": "abbreviate",
    "uppercase": "uppercase",
    "lowercase": "lowercase",
    "expand": "expand",
    "squash": "squash",
    "dash": "dash",
    "underscore": "underscore",
    "dot": "dot",
    "period": "period",
    "minus": "minus",
    "semi-colon": "semi-colon",
    "hyphen": "hyphen",
    "triple": "triple"
}

################################################################################

grammarCfg = Config("multi edit")
grammarCfg.cmd = Section("Language section")
grammarCfg.cmd.map = Item(
    {
        # Navigation keys.
        "up [<n>]": Key("up:%(n)d"),
        "up [<n>] slow": Key("up/15:%(n)d"),
        "down [<n>]": Key("down:%(n)d"),
        "down [<n>] slow": Key("down/15:%(n)d"),
        "left [<n>]": Key("left:%(n)d"),
        "left [<n>] slow": Key("left/15:%(n)d"),
        "right [<n>]": Key("right:%(n)d"),
        "right [<n>] slow": Key("right/15:%(n)d"),
        "page up [<n>]": Key("pgup:%(n)d"),
        "page down [<n>]": Key("pgdown:%(n)d"),
        "up <n> (page|pages)": Key("pgup:%(n)d"),
        "down <n> (page|pages)": Key("pgdown:%(n)d"),
        "left <n> (word|words)": Key("c-left/3:%(n)d/10"),
        "right <n> (word|words)": Key("c-right/3:%(n)d/10"),
        "home|homer": Key("home"),
        "end": Key("end"),
        "doc home": Key("c-home/3"),
        "doc end": Key("c-end/3"),
        "doc save": Key("c-s"),
        "doc open": Key("c-o"),
        "doc save as": Key("alt:down/3,f/3,a/3,alt:up/3"),
        "doc close [<n>]": Key("c-w:%(n)d"),
        "doc close all": Key("cs-w:%(n)d"),
        "doc next [<n>]": Key("c-tab:%(n)d"),
        "doc create": Key("c-n"),
        "doc new tab": Key("c-t"),
        "doc (previous|back) [<n>]": Key("cs-tab:%(n)d"),
        "doc (search|find)": Key("c-f"),
        "doc print": Key("c-p"),
        "doc format": Key("a-e, v, a"),
        "select line [<n>]": release + Key("home, home, s-down:%(n)d"),
        "go block start [<n>]": Key("sa-[:%(n)d"),
        "show parameters": Key("cs-space"),
        "integer": Text("int"),
        "variable": Text("var"),
        "uno [<n>]": Key("f1:%(n)d"),
        "doss [<n>]": Key("f2:%(n)d"),
        "trez [<n>]": Key("f3:%(n)d"),
        "quatro [<n>]": Key("f4:%(n)d"),
        "sinko [<n>]": Key("f5:%(n)d"),
        "see ettay [<n>]": Key("f7:%(n)d"),
        "occo [<n>]": Key("f8:%(n)d"),
        "noo evvay [<n>]": Key("f9:%(n)d"),
        "dee ez [<n>]": Key("f10:%(n)d"),
        "onsay [<n>]": Key("f11:%(n)d"),
        "dossay [<n>]": Key("f12:%(n)d"),
        # Functional keys.
        "space": release + Key("space"),
        "space [<n>]": release + Key("space:%(n)d"),
        "(slap|go) [<n>]": release + Key("enter:%(n)d"),
        "tab [<n>]": Key("tab:%(n)d"),
        "delete [<n>]": Key("del/3:%(n)d"),
        "delete [this] line": Key("home, home, s-end, del, del"),  # @IgnorePep8
        "backspace [<n>]": release + Key("backspace:%(n)d"),
        "application key": release + Key("apps/3"),
        "win key": release + Key("win/3"),
        "paste [that]": Function(paste_command),
        "copy [that]": Function(copy_command),
        "cut [that]": release + Key("c-x/3"),
        "select all": release + Key("c-a/3"),
        "undo": release + Key("c-z/3"),
        "undo <n> [times]": release + Key("c-z/3:%(n)d"),
        "redo": release + Key("c-y/3"),
        "redo <n> [times]": release + Key("c-y/3:%(n)d"),
        "[(hold|press)] (alt|meta)": Key("alt:down/3"),
        "release (alt|meta)": Key("alt:up"),
        "[(hold|press)] shift": Key("shift:down/3"),
        "release shift": Key("shift:up"),
        "[(hold|press)] control": Key("ctrl:down/3"),
        "release control": Key("ctrl:up"),
        "(hold|press) (hyper|windows)": Key("win:down/3"),
        "release (hyper|windows)": Key("win:up"),
        "[press] (hyper|windows) first": Key("win:down/3, 1/3, win:up/3"),
        "[press] (hyper|windows) <m>": Key("win:down/3, %(m)d/3, win:up/3"),
        "release [all]": release,
        "git commit minus M": Text("git commit -m "),
        # Closures.
        "angle brackets": Key("langle, rangle, left/3"),
        "brackets": Key("lbracket, rbracket, left/3"),
        "braces": Key("lbrace, rbrace, left/3"),
        "parens": Key("lparen, rparen, left/3"),
        "quotes": Key("dquote/3, dquote/3, left/3"),
        "single quotes": Key("squote, squote, left/3"),
        # To release keyboard capture by VirtualBox.
        "press right control": Key("Control_R"),
        # Formatting <n> words to the left of the cursor.
        "camel case <n> [words]": Function(camel_case_count),
        "pascal case <n> [words]": Function(pascal_case_count),
        "snake case <n> [words]": Function(snake_case_count),
        "squash <n> [words]": Function(squash_count),
        "expand <n> [words]": Function(expand_count),
        "uppercase <n> [words]": Function(uppercase_count),
        "lowercase <n> [words]": Function(lowercase_count),
        # Format dictated words. See the formatMap for all available types.
        # Ex: "camel case my new variable" -> "myNewVariable"
        # Ex: "snake case my new variable" -> "my_new_variable"
        # Ex: "uppercase squash my new hyphen variable" -> "MYNEW-VARIABLE"
        "<formatType> <text>": Function(format_text),
        # For writing words that would otherwise be characters or commands.
        # Ex: "period", tab", "left", "right", "home".
        "Simon says <reservedWord>": Text("%(reservedWord)s"),
        # Abbreviate words commonly used in programming.
        # Ex: arguments -> args, parameterers -> params.
        "abbreviate <abbreviation>": Text("%(abbreviation)s"),
        # Text corrections.
        "(add|fix) missing space": Key("c-left/3, space, c-right/3"),
        "(delete|remove) (double|extra) (space|whitespace)": Key("c-left/3, backspace, c-right/3"),  # @IgnorePep8
        "(delete|remove) (double|extra) (type|char|character)": Key("c-left/3, del, c-right/3"),  # @IgnorePep8
        # Microphone sleep/cancel started dictation.
        "[<text>] (go to sleep|cancel and sleep) [<text2>]": Function(cancel_and_sleep),  # @IgnorePep8
        # Reload Natlink.
        "reload Natlink": Function(reload_natlink),
        # Ego
        "alpha": Text("a"),
        "bravo": Text("b"),
        "charlie": Text("c"),
        "delta": Text("d"),
        "echo": Text("e"),
        "foxtrot": Text("f"),
        "golf": Text("g"),
        "hotel": Text("h"),
        "(india|indigo)": Text("i"),
        "juliet": Text("j"),
        "kilo": Text("k"),
        "lima": Text("l"),
        "mike": Text("m"),
        "november": Text("n"),
        "oscar": Text("o"),
        "(Papa|pappa|pepper|popper)": Text("p"),
        "quebec": Text("q"),
        "romeo": Text("r"),
        "sierra": Text("s"),
        "tango": Text("t"),
        "uniform": Text("u"),
        "victor": Text("v"),
        "whiskey": Text("w"),
        "x-ray": Text("x"),
        "yankee": Text("y"),
        "zulu": Text("z"),
        "big alpha": Text("A"),
        "big bravo": Text("B"),
        "big charlie": Text("C"),
        "big delta": Text("D"),
        "big echo": Text("E"),
        "big foxtrot": Text("F"),
        "big golf": Text("G"),
        "big hotel": Text("H"),
        "big india": Text("I"),
        "big juliet": Text("J"),
        "big kilo": Text("K"),
        "big lima": Text("L"),
        "big mike": Text("M"),
        "big november": Text("N"),
        "big oscar": Text("O"),
        "big (Papa|pappa|pepper|popper)": Text("P"),
        "big quebec": Text("Q"),
        "big romeo": Text("R"),
        "big sierra": Text("S"),
        "big tango": Text("T"),
        "big uniform": Text("U"),
        "big victor": Text("V"),
        "big whiskey": Text("W"),
        "big x-ray": Text("X"),
        "big yankee": Text("Y"),
        "big zulu": Text("Z"),
        # Ego
        "quote [<n>]": release + Key("dquote:%(n)d"),
        "dot": Text("."),
        "backslish": Text("\\"),
        "slish": Text("/"),
        "lape": Text("("),
        "rape": Text(")"),
        "lace": Text("{"),
        "race": Text("}"),
        "(lack|bra)": Text("["),
        "(rack|ket)": Text("]"),
        "lang": Text("<"),
        "rang": Text(">"),
        "pipe": Text("|"),
        "eke": Text("="),
        "bang": Text("!"),
        "plus": Text("+"),
        "minus": Text("-"),
        "mull": Text("*"),
        "(underscore|bar) [<n>]": Key("underscore/2:%(n)d"),
        "(sem|semi|rock|semicolon)": Text(";"),
        "(coal|colon)": Text(":"),
        "(comma|cam) [<n>]": Key("comma/2:%(n)d"),
        "(dot|period) [<n>]": Key("dot/2:%(n)d"),
        "(dash|hyphen) [<n>]": Key("hyphen/2:%(n)d"),
        "hash [<n>]": Key("hash/2:%(n)d"),
        #"": Text(""),
        # Ego
        "quit [<n>]": release + Key("escape:%(n)d"),
        "switch [<n>]": release + Key("alt:down/10, tab:%(n)d/10, alt:up"),
        "hut": release + Key("alt:down/5, d, alt:up"),
        "context menu": release + Key("apps"),
        # Ego
        "hash include stud Io": Text("#include <stdio.h>"),
        "hash include stud lib": Text("#include <stdlib.h>"),
        "change directory": Text("cd "),
        "change directory drive": Text("cd /D "),
        "directory list": Text("dir "),
        "strike [<n>]": release + Key("backspace:%(n)d"),
        "kill [<n>]": release + Key("delete:%(n)d"),
        "pig [<n>]": release + Key("pgdown:%(n)d"),
        "pug [<n>]": release + Key("pgup:%(n)d"),
        "comment": Text("// "),
        # Ego
        "dictation on": Function(dictation_on),
        "dictation off": Function(dictation_off),
        "say <text>": Text("%(text)s"),
        # Ego
        "show dragon tip": Text(""),
        "move dragon tip": Mouse("[15, 8]/200, left:down, [0.66, 0.98], left:up, <-50,0>"),
    },
    namespace={
        "Key": Key,
        "Text": Text,
    }
)

class KeystrokeRule(MappingRule):
    exported = False
    mapping = grammarCfg.cmd.map
    extras = [
        IntegerRef("n", 1, 100),
        IntegerRef("m", 0, 100),
        Dictation("text"),
        Dictation("text2"),
        Choice("formatType", formatMap),
        Choice("abbreviation", abbreviationMap),
        Choice("reservedWord", reservedWord),
    ]
    defaults = {
        "n": 1,
        "m": 0,
    }

alternatives = []
alternatives.append(RuleRef(rule=KeystrokeRule()))
single_action = Alternative(alternatives)

sequence = Repetition(single_action, min=1, max=16, name="sequence")

class RepeatRule(CompoundRule):
    # Here we define this rule's spoken-form and special elements.
    spec = "<sequence> [[[and] repeat [that]] <n> times]"
    extras = [
        sequence,  # Sequence of actions defined above.
        IntegerRef("n", 1, 100),  # Times to repeat the sequence.
    ]
    defaults = {
        "n": 1,  # Default repeat count.
    }

    def _process_recognition(self, node, extras):  # @UnusedVariable
        sequence = extras["sequence"]  # A sequence of actions.
        count = extras["n"]  # An integer repeat count.
        for i in range(count):  # @UnusedVariable
            for action in sequence:
                action.execute()
        release.execute()

elite_context = AppContext(executable="elitedangerous")

grammar = Grammar("Generic edit", context=~elite_context)
grammar.add_rule(RepeatRule())  # Add the top-level rule.
grammar.load()  # Load the grammar.

def unload():
    """Unload function which will be called at unload time."""
    global grammar
    if grammar:
        grammar.unload()
    grammar = None

################################################################################

elite_mode = "flight"
elite_driving_speed = 0

def elite_key(n=1, name=None):
    k = Key("{}:down/5, {}:up/10".format(name, name))
    for i in range(n):
        k.execute()

def elite_long_key(n=1, pre=None, name=None):
    k = Key("alt:down, {}:down, {}:down/5, {}:up, {}:up, alt:up/10".format(pre, name, name, pre))
    for i in range(n):
        k.execute()

def elite_longer_key(n=1, pre=None, name=None):
    k = Key("shift:down, alt:down, {}:down/5, {}:down/5, {}:up, {}:up, alt:up, shift:up/10".format(pre, name, name, pre))
    for i in range(n):
        k.execute()

def elite_timed_key(n=1, name=None, factor=12):
    k = Key("{}:down/{}, {}:up/10".format(name, 5+n*factor, name))
    k.execute()

def elite_timed_key_pair(n=1, name=None, factor=12, second=None):
    k = Key("{}:down/5, {}:down/{}, {}:up/10, {}:up/10".format(name, second, 5+n*factor, second, name))
    k.execute()

def elite_switch_mode(name):
    global elite_mode
    elite_mode = name
    get_engine().speak(name + " mode")

def elite_galaxy_map():
    elite_longer_key(pre="f1", name="u")

def elite_next_tab():
    elite_longer_key(pre="f2", name="i")

def elite_select():
    elite_key(name="b")

def elite_galaxy_route():
    elite_galaxy_map()
    sleep(2)
    elite_next_tab()
    sleep(0.1)
    elite_select()

def elite_galaxy_bookmarks():
    elite_galaxy_map()
    sleep(2)
    elite_next_tab()
    elite_next_tab()

def elite_navigate(name=None, n=1):
    global elite_mode
    if name == "left":
        if elite_mode == "flight":
            Function(elite_key, name="left", n=n).execute()
        elif elite_mode == "driving":
            Function(elite_timed_key, name="a", factor=6, n=n).execute()
    elif name == "right":
        if elite_mode == "flight":
            Function(elite_key, name="right", n=n).execute()
        elif elite_mode == "driving":
            Function(elite_timed_key, name="d", factor=6, n=n).execute()

def elite_relative_speed(name=None, n=1):
    global elite_mode
    global elite_driving_speed
    if name == "faster":
        if elite_mode == "flight":
            Function(elite_long_key, pre="f5", name="q").execute()
        elif elite_mode == "driving":
            for i in range(n):
                elite_driving_speed = elite_driving_speed + 1
                Key("e:down/5, e:up/10").execute()

    elif name == "slower":
        if elite_mode == "flight":
            Function(elite_long_key, pre="f5", name="w").execute()
        elif elite_mode == "driving":
            for i in range(n):
                elite_driving_speed = elite_driving_speed - 1
                Key("q:down/5, q:up/10").execute()

def elite_all_stop():
    global elite_mode
    global elite_driving_speed
    if elite_mode == "flight":
        Function(elite_long_key, pre="f5", name="u").execute()
    elif elite_mode == "driving":
        while elite_driving_speed > 0:
            elite_driving_speed = elite_driving_speed - 1
            Key("q:down/5, q:up/10").execute()
        while elite_driving_speed < 0:
            elite_driving_speed = elite_driving_speed + 1
            Key("e:down/5, e:up/10").execute()

def elite_what_driving_speed():
    get_engine().speak("driving speed is {}".format(elite_driving_speed))

def elite_reset_driving_speed():
    global elite_driving_speed
    elite_driving_speed = 0
    get_engine().speak("driving speed is now {}".format(elite_driving_speed))

def elite_target(name):
    global elite_mode
    global elite_driving_speed
    if name == "ahead":
        if elite_mode == "flight":
            Function(elite_long_key, pre="f7", name="w").execute()
        elif elite_mode == "driving":
            Mouse("right:down/5, right:up/10").execute()

def elite_what_mode():
    global elite_mode
    get_engine().speak(elite_mode + " mode")

elite_galaxy_release = Key("w:up, a:up, s:up, r:up, x:up, z:up, f:up, d:up/10")

class EliteRule(MappingRule):
    mapping = {
        # Modes
        "driving mode": Function(elite_switch_mode, name="driving"),
        "(flight|normal) mode": Function(elite_switch_mode, name="flight"),
        "reset driving speed": Function(elite_reset_driving_speed),
        "what mode": Function(elite_what_mode),
        "what driving speed": Function(elite_what_driving_speed),

        # Mouse controls
        "reset mouse": Function(elite_long_key, pre="9", name="q"),

        # Flight rotation
        "sky <n>": Function(elite_timed_key, name="b"),
        "sea <n>": Function(elite_timed_key, name="np5"),
        "loll <n>": Function(elite_timed_key, name="np4", factor=3),
        "roll <n>": Function(elite_timed_key, name="np6", factor=3),
        "port <n>": Function(elite_timed_key, name="a"),
        "starboard <n>": Function(elite_timed_key, name="c"),
        "sky": Key("np5:up/10, b:down/5"),
        "sea": Key("b:up/10, np5:down/5"),
        "loll": Key("np6:up/10, np4:down/5"),
        "roll": Key("np4:up/10, np6:down/5"),
        "port": Key("c:up/10, a:down/5"),
        "starboard": Key("a:up/10, c:down/5"),
        "lock": Key("np4, np5, np6, np7, np8, np9, a, b, c, np4:up, np5:up, np6:up, np7:up, np8:up, np9:up, a:up, b:up, c:up/10"),

        # Flight thrust
        "thrust up": Key("f:up/10, r:down/5"),
        "thrust down": Key("r:up/10, f:down/5"),
        "thrust left": Key("e:up/10, q:down/5"),
        "thrust right": Key("q:up/10, e:down/5"),
        "thrust forward": Key("npdec:up/10, np0:down/5"),
        "thrust (back|backward)": Key("np0:up/10, npdec:down/5"),
        "nix": Key("f, r, q, e, npdec, np0, f:up/10, r:up/10, q:up/10, e:up/10, np0:up/10, npdec:up/10"),
        "thrust up <n>": Function(elite_timed_key, name="r"),
        "thrust down <n>": Function(elite_timed_key, name="f"),
        "thrust left <n>": Function(elite_timed_key, name="q"),
        "thrust right <n>": Function(elite_timed_key, name="e"),
        "thrust forward <n>": Function(elite_timed_key, name="np0"),
        "thrust (back|backward) <n>": Function(elite_timed_key, name="npdec"),

        # Flight throttle
        "faster [<n>]": Function(elite_relative_speed, name="faster"),
        "slower [<n>]": Function(elite_relative_speed, name="slower"),
        "full reverse": Function(elite_long_key, pre="f5", name="e"),
        "three quarter reverse": Function(elite_long_key, pre="f5", name="r"),
        "half reverse": Function(elite_long_key, pre="f5", name="t"),
        "quarter reverse": Function(elite_long_key, pre="f5", name="y"),
        "all stop": Function(elite_all_stop),
        "quarter impulse": Function(elite_long_key, pre="f5", name="i"),
        "half impulse": Function(elite_long_key, pre="f5", name="o"),
        "(three quarter impulse|cruising speed)": Function(elite_long_key, pre="f5", name="p"),
        "(full impulse|engage|[all] ahead full)": Function(elite_long_key, pre="f5", name="a"),

        # Flight miscellaneous
        "boost [engines]": Function(elite_long_key, pre="f6", name="w"),
        "[charge the] hyperdrive": Function(elite_long_key, pre="f6", name="e"),
        "(supercruise|cruise)": Function(elite_long_key, pre="f6", name="r"),
        "(hyperspace|jump)": Function(elite_long_key, pre="f6", name="t"),
        "toggle orbit lines": Function(elite_long_key, pre="f6", name="y"),

        # Targeting
        "target ahead": Function(elite_target, name="ahead"),
        "[target] next (target|ship) [<n>]": Function(elite_long_key, pre="f7", name="w"),
        "[target] previous (target|ship) [<n>]": Function(elite_long_key, pre="f7", name="e"),
        "[target] highest (threat|ship) [<n>]": Function(elite_long_key, pre="f7", name="r"),
        "[target] next (hostile|enemy|foe) [<n>]": Function(elite_long_key, pre="f7", name="t"),
        "[target] previous (hostile|enemy|foe [<n>])": Function(elite_long_key, pre="f7", name="y"),
        "[target] wingman one": Function(elite_long_key, pre="f7", name="u"),
        "[target] wingman two": Function(elite_long_key, pre="f7", name="i"),
        "[target] wingman three": Function(elite_long_key, pre="f7", name="o"),
        "[target] (wingman|wingman's) target": Function(elite_long_key, pre="f7", name="a"),
        "wingman nav lock": Function(elite_long_key, pre="f7", name="s"),
        "[target] next subsystem [<n>]": Function(elite_long_key, pre="f7", name="d"),
        "[target] previous subsystem [<n>]": Function(elite_long_key, pre="f7", name="f"),
        "[target] (next system [in route]|destination)": Function(elite_long_key, pre="f7", name="g"),

        # Weapons
        "fire 1": Mouse("left:down/5"),
        "fire 2": Mouse("right:down/5"),
        "fire 3": Mouse("middle:down/5"),
        "commence honking": Mouse("left:down/550, left:up/10"),
        "(cease fire|stop firing|stop shooting)": Mouse("left:up, middle:up, right:up/10"),
        "fox 1": Mouse("left:down/5, left:up/10"),
        "fox 2": Mouse("right:down/5, right:up/10"),
        "fox 3": Mouse("middle:down/5, middle:up/10"),
        "(next fire group|next weapon|weapon) [<n>]": Function(elite_long_key, pre="f8", name="e"),
        "(previous fire group|previous weapon) [<n>]": Function(elite_long_key, pre="f8", name="r"),
        "(deploy|retract|stow) (hardpoints|weapons)": Function(elite_long_key, pre="f8", name="t"),

        # Cooling
        "(silent running|run silent)": Function(elite_long_key, pre="f9", name="q"),
        "(use|deploy|launch|fire) heatsink": Function(elite_long_key, pre="f9", name="w"),

        # Miscellaneous
        "[toggle] [ship] lights": Function(elite_long_key, pre="f11", name="q"),
        "[increase] sensor zoom [in]": Function(elite_long_key, pre="f11", name="w"),
        "[decrease] sensor zoom [out]": Function(elite_long_key, pre="f11", name="e"),
        "power to engines": Key("alt:down/5, f11:down/5, u:down/5, u:up/10, u:down/5, u:up/10, r:down/5, r:up/10, r:down/5, r:up/10, r:down/5, r:up/10, f11:up/10, alt:up/10"),
        "power to weapons": Key("alt:down/5, f11:down/5, u:down/5, u:up/10, u:down/5, u:up/10, t:down/5, t:up/10, t:down/5, t:up/10, t:down/5, t:up/10, f11:up/10, alt:up/10"),
        "power to (systems|shields)": Key("alt:down/5, f11:down/5, u:down/5, u:up/10, y:down/5, y:up/10, y:down/5, y:up/10, y:down/5, y:up/10, f11:up/10, alt:up/10"),
        "power to (shields and engines|engines and shields)": Key("alt:down/5, f11:down/5, u:down/5, u:up/10, y:down/5, y:up/10, r:down/5, r:up/10, y:down/5, y:up/10, r:down/5, r:up/10, f11:up/10, alt:up/10"),
        "balance power": Function(elite_long_key, pre="f11", name="u"),
        "reset HMD": Function(elite_long_key, pre="f11", name="i"),
        "cargo scoop": Function(elite_long_key, pre="f11", name="o"),
        "jettison all cargo": Function(elite_long_key, pre="f11", name="p"),
        "landing gear": Function(elite_long_key, pre="f11", name="a"),
        "(use|deploy|launch|fire) shield cell": Function(elite_long_key, pre="f11", name="d"),
        "(use|deploy|launch|fire) (chaff|chef|chav) [launcher]": Function(elite_long_key, pre="f11", name="f"),
        "charge (ECM|countermeasures)": Function(elite_long_key, pre="f11", name="g"),
        "weapon colour": Function(elite_long_key, pre="f11", name="h"),
        "engine colour": Function(elite_long_key, pre="f11", name="j"),
        "night vision": Function(elite_long_key, pre="f11", name="k"),

        # Mode switches
        "(left panel|target panel|Nav)": Function(elite_key, name="1"),
        "(right panel|systems|inventory|ship)": Function(elite_key, name="4"),
        "(top panel|comms)": Function(elite_key, name="2"),
        "(bottom panel|role panel)": Function(elite_key, name="3"),
        "galaxy map": Function(elite_galaxy_map),
        "galaxy route": Function(elite_galaxy_route),
        "galaxy (bookmarks|bookmark)": Function(elite_galaxy_bookmarks),
        "system map": Function(elite_longer_key, pre="f1", name="i"),
        "show CQC score screen": Function(elite_longer_key, pre="f1", name="o"),
        "headlook": Function(elite_longer_key, pre="f1", name="p"),
        "game menu": Function(elite_longer_key, pre="f1", name="a"),
        "open discovery": Function(elite_longer_key, pre="f1", name="d"),
        "[switch|change] [hud|hut|hudd|weapon] mode": Function(elite_longer_key, pre="f1", name="f"),
        "request docking": Key("1:down/5, 1:up/50, e:down/5, e:up/50, e:down/5, e:up/50, right:down/5, right:up/50, space:down/5, space:up/50, left:down/5, left:up/50, q:down/5, q:up/50, q:down/5, q:up/50, backspace:down/5, backspace:up/50"),

        # Interface mode
        "up [<n>]": Function(elite_key, name="up"),
        "down [<n>]": Function(elite_key, name="down"),
        "right [<n>]": Function(elite_navigate, name="right"),
        "left [<n>]": Function(elite_navigate, name="left"),
        "slap [<n>]": Function(elite_key, name="enter"),
        "quit [<n>]": Function(elite_key, name="escape"),
        "(space|go|select|make it so) [<n>]": Function(elite_select),
        "back [<n>]": Function(elite_key, name="backspace"),
        "expand|collapse": Function(elite_longer_key, pre="f2", name="u"),
        "[next] tab [<n>]": Function(elite_next_tab),
        "(previous tab|shift tab|bat) [<n>]": Function(elite_longer_key, pre="f2", name="h"),
        "next page [<n>]": Function(elite_longer_key, pre="f2", name="j"),
        "previous page [<n>]": Function(elite_longer_key, pre="f2", name="k"),

        "(meta down|alt key down)": Key("alt:down/5"),
        "(meta up|alt key up)": Key("alt:up/10"),
        "shift key down": Key("alt:down/5"),
        "shift key up": Key("shift:up/10"),
        "control key down": Key("ctrl:down/5"),
        "control key up": Key("ctrl:up/10"),

        # Headlook mode

        # Galaxy map
        "[walk] north <n>": Function(elite_timed_key, name="w", factor=3),
        "[walk] south <n>": Function(elite_timed_key, name="s", factor=3),
        "[walk] east <n>": Function(elite_timed_key, name="d", factor=3),
        "[walk] west <n>": Function(elite_timed_key, name="a", factor=3),
        "walk north east <n>": Function(elite_timed_key_pair, name="w", second="d", factor=3),
        "walk north west <n>": Function(elite_timed_key_pair, name="w", second="a", factor=3),
        "walk south east <n>": Function(elite_timed_key_pair, name="s", second="d", factor=3),
        "walk south west <n>": Function(elite_timed_key_pair, name="s", second="a", factor=3),
        "walk north": elite_galaxy_release + Key("w:down/5"),
        "walk south": elite_galaxy_release + Key("s:down/5"),
        "walk east": elite_galaxy_release + Key("d:down/5"),
        "walk west": elite_galaxy_release + Key("a:down/5"),
        "walk north east": elite_galaxy_release + Key("w:down/5, d:down/5"),
        "walk south east": elite_galaxy_release + Key("s:down/5, d:down/5"),
        "walk north west": elite_galaxy_release + Key("w:down/5, a:down/5"),
        "walk south west": elite_galaxy_release + Key("s:down/5, a:down/5"),
        "wait": elite_galaxy_release,
        "spin left <n>": Function(elite_timed_key, name="q", factor=60),
        "spin right <n>": Function(elite_timed_key, name="e", factor=60),
        "walk up <n>": Function(elite_timed_key, name="r", factor=6),
        "walk down <n>": Function(elite_timed_key, name="f", factor=6),
        "walk up": elite_galaxy_release + Key("r:down/5"),
        "walk down": elite_galaxy_release + Key("f:down/5"),
        "zoom out <n>": Function(elite_timed_key, name="x", factor=6),
        "zoom in <n>": Function(elite_timed_key, name="z", factor=6),
        "zoom out": elite_galaxy_release + Key("x:down/5"),
        "zoom in": elite_galaxy_release + Key("z:down/5"),

        # Driving
        #"drive left <n>": Function(elite_timed_key, name="a", factor=6),
        #"drive right <n>": Function(elite_timed_key, name="d", factor=6),
        #"drive faster": Function(elite_key, name="e"),
        #"drive slower": Function(elite_key, name="q"),
        "drive target [ahead]": Mouse("right:down/5, right:up/10"),
        "drive fox 1": Mouse("left:down/5, left:up/10"),
        "drive fire 1": Mouse("left:down/5"),
        "drive cease fire": Mouse("left:up/10"),
        "drive lights": Function(elite_key, name="l"),
        "drive turret": Function(elite_key, name="u"),
        "drive [weapon] mode": Function(elite_key, name="m"),
        "drive [next] weapon": Function(elite_key, name="n"),
        "drive [cargo] scoop": Function(elite_key, name="home"),
        "drive jump <n>": Function(elite_timed_key, name="space", factor=30),

        # Camera Suite
        "camera [suite]": Function(elite_long_key, pre="3", name="q"),
        "next camera": Function(elite_long_key, pre="3", name="w"),
        "previous camera": Function(elite_long_key, pre="3", name="e"),
        "free camera": Function(elite_long_key, pre="3", name="r"),
        "camera one": Function(elite_long_key, pre="3", name="a"),
        "camera two": Function(elite_long_key, pre="3", name="s"),
        "camera three": Function(elite_long_key, pre="3", name="d"),
        "camera four": Function(elite_long_key, pre="3", name="f"),
        "camera five": Function(elite_long_key, pre="3", name="g"),
        "camera six": Function(elite_long_key, pre="3", name="h"),
        "camera seven": Function(elite_long_key, pre="3", name="j"),
        "camera eight": Function(elite_long_key, pre="3", name="k"),
        "camera nine": Function(elite_long_key, pre="3", name="l"),
        "toggle hud": Function(elite_long_key, pre="4", name="q"),
        "(show|hide) interface": Key("ctrl:down/5, ctrl:up/10"),
        "increase camera speed": Function(elite_long_key, pre="4", name="w"),
        "decrease camera speed": Function(elite_long_key, pre="4", name="e"),

        # Playlist
        "playlist (play|pause)": Function(elite_long_key, pre="6", name="q"),
        "playlist skip forward": Function(elite_long_key, pre="6", name="w"),
        "playlist skip backward": Function(elite_long_key, pre="6", name="e"),
        "playlist clear queue": Function(elite_long_key, pre="6", name="r"),

        # Full-spectrum system scanner
        "(FSS|scan system|system scanner|full spectrum system scanner)": Function(elite_longer_key, pre="7", name="q"),

        # Ego
        "alpha": Text("a"),
        "bravo": Text("b"),
        "charlie": Text("c"),
        "delta": Text("d"),
        "echo": Text("e"),
        "foxtrot": Text("f"),
        "golf": Text("g"),
        "hotel": Text("h"),
        "(india|indigo)": Text("i"),
        "juliet": Text("j"),
        "kilo": Text("k"),
        "lima": Text("l"),
        "mike": Text("m"),
        "november": Text("n"),
        "oscar": Text("o"),
        "(Papa|pappa|pepper|popper)": Text("p"),
        "quebec": Text("q"),
        "romeo": Text("r"),
        "sierra": Text("s"),
        "tango": Text("t"),
        "uniform": Text("u"),
        "victor": Text("v"),
        "whiskey": Text("w"),
        "x-ray": Text("x"),
        "yankee": Text("y"),
        "zulu": Text("z"),
        "big alpha": Text("A"),
        "big bravo": Text("B"),
        "big charlie": Text("C"),
        "big delta": Text("D"),
        "big echo": Text("E"),
        "big foxtrot": Text("F"),
        "big golf": Text("G"),
        "big hotel": Text("H"),
        "big india": Text("I"),
        "big juliet": Text("J"),
        "big kilo": Text("K"),
        "big lima": Text("L"),
        "big mike": Text("M"),
        "big november": Text("N"),
        "big oscar": Text("O"),
        "big (Papa|pappa|pepper|popper)": Text("P"),
        "big quebec": Text("Q"),
        "big romeo": Text("R"),
        "big sierra": Text("S"),
        "big tango": Text("T"),
        "big uniform": Text("U"),
        "big victor": Text("V"),
        "big whiskey": Text("W"),
        "big x-ray": Text("X"),
        "big yankee": Text("Y"),
        "big zulu": Text("Z"),
        "(dash|hyphen|minus)": Text("-"),
        "space": Text(" "),
        "strike": Key("backspace"),
        "strike <n>": Key("backspace:%(n)d"),
    }
    extras = [
        Dictation("text"),
        IntegerRef("n", 1, 100),
    ]
    defaults = {
    }

elite_alternatives = []
elite_alternatives.append(RuleRef(rule=EliteRule()))
elite_single_action = Alternative(elite_alternatives)

elite_sequence = Repetition(elite_single_action, min=1, max=16, name="elite_sequence")

class EliteRepeatRule(CompoundRule):
    # Here we define this rule's spoken-form and special elements.
    spec = "<elite_sequence> [[[and] repeat [that]] <n> times]"
    extras = [
        elite_sequence,
        IntegerRef("n", 1, 100),
    ]
    defaults = {
        "n": 1,
    }

    def _process_recognition(self, node, extras):
        elite_sequence = extras["elite_sequence"]
        count = extras["n"]
        for i in range(count):
            for action in elite_sequence:
                action.execute()

try:
    if eliteGrammar:
        eliteGrammar.unload()
except NameError:
    pass

eliteGrammar = Grammar("Elite Dangerous grammar", context=elite_context)
eliteGrammar.add_rule(EliteRepeatRule())
eliteGrammar.load()
