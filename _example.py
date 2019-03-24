from dragonfly import Grammar, CompoundRule, Key, Text, Mouse

# Voice command rule combining spoken form and recognition processing.
class ExampleRule(CompoundRule):
    spec = "do something computer"                  # Spoken form of command.
    def _process_recognition(self, node, extras):   # Callback when command is spoken.
        print "I am become death, destroyer of worlds."

class ClickRule(CompoundRule):
    spec = "touch"
    def _process_recognition(self, node, extras):
        action = Mouse("left")
        action.execute()

# Create a grammar which contains and loads the command rule.
grammar = Grammar("example grammar")                # Create a grammar to contain the command rule.
grammar.add_rule(ExampleRule())                     # Add the command rule to the grammar.
grammar.add_rule(ClickRule())
grammar.load()
