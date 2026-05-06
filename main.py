from terminaltexteffects.effects.effect_blackhole import Blackhole
from terminaltexteffects.effects.effect_wipe import Wipe
from terminaltexteffects.effects.effect_middleout import MiddleOut
import openpyxl
from pyfiglet import figlet_format

def terminal_effect_print(effect):
    with effect.terminal_output() as terminal:
        for frame in effect:
            terminal.print(frame)

def introduction():
    title = Blackhole(figlet_format("STREAK - IT", font='ansi_shadow'))
    terminal_effect_print(effect=title)

    subtitle = "Streak it for me baby. Put in as much as you want, I can take it all"
    terminal_effect_print(Wipe(subtitle))

    separator = "============================================================================="
    terminal_effect_print(MiddleOut(separator))



def main():
    introduction()
    


if __name__ == "__main__":
    main()
