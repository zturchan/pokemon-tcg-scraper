#-------------------------------------------------------------------------------
# Name:        Pokemon TCG Scraper
# Purpose:     Scrapes the POKEMON TCG API for all cards and renders them in
#              a desired Excel format.
#
# Author:      Zak Turchansky
#
# Created:     15-06-2019
# Copyright:   (c) Zak Turchansky 2019
# Licence:     LICENSE.txt
#-------------------------------------------------------------------------------

from pokemontcgsdk import Card
from openpyxl import Workbook
from openpyxl.styles import Font
from os import remove
from pathlib import Path


def main():
    print("Downloading all cards... Please be patient")

    fields_remain_same = ['id',
                          'name',
                          'national_pokedex_number',
                          'image_url',
                          'image_url_hi_res',
                          'subtype',
                          'supertype',
                          'hp',
                          'number',
                          'artist',
                          'rarity',
                          'series',
                          'set',
                          'set_code',
                          'converted_retreat_cost',
                          'evolves_from']
    cards = Card.all()
    create_workbook(cards, fields_remain_same)

def create_workbook(cards, default_fields):
    if Path("pkmn_output.xlsx").is_file():
        remove("pkmn_output.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.append(create_header_row(default_fields))
    for cell in ws["1:1"]:
        cell.font = Font(bold=True)
    for card in cards:
        row = []
        for field in default_fields:
            value = getattr(card, field)
            if isinstance(value, list) and len(value) == 1:
                row.append(value[0])
            elif not isinstance(value, list):
                row.append(value)
            else:
                raise Exception("Unexpected data type found: " + field +
                                ". Value is " + str(value))
        parse_ancient_trait(row, card)
        parse_types(row, card)
        parse_ability(row, card)
        parse_text(row, card)
        parse_attacks(row, card)
        parse_weakness(row, card)
        parse_resistance(row, card)
        ws.append(row)

    wb.save("pkmn_output.xlsx")

def parse_ancient_trait(row, card):
    if card.ancient_trait != None:
        row.extend([card.ancient_trait["name"],card.ancient_trait["text"]])
    else:
        row.extend(["",""])

def parse_types(row, card):
    for i in range(2):
        if(card.types == None or len(card.types) <= i):
            row.append("")
        else:
            row.append(card.types[i])

def parse_ability(row, card):
    if card.ability != None:
        row.extend([card.ability["name"],card.ability["text"]])
    else:
        row.extend(["",""])

def parse_text(row, card):
    # Get Rules text (e.g. Mega Evolution text) if applicable
    if card.text != None:
        row.append(" ".join(card.text))
    else:
        row.append("")

def parse_weakness(row, card):
    if card.weaknesses != None:
        row.extend([card.weaknesses[0]["type"], card.weaknesses[0]["value"]])
    else:
        row.extend(["",""])

def parse_resistance(row, card):
    if card.resistances != None:
        row.extend([card.resistances[0]["type"], card.resistances[0]["value"]])
    else:
        row.extend(["",""])

def parse_attacks(row, card):
    for i in range(3):
        if(card.attacks == None or len(card.attacks) <= i):
            row.extend(["","","","",""])
        else:
            parse_attack(row, card.attacks[i])

def parse_attack(row, attack):
    if attack != None:
        damage = attack["damage"] if "damage" in attack else ""
        text = attack["text"] if "text" in attack else ""
        row.extend([''.join(attack["cost"]),
                    attack["name"],
                    text,
                    damage,
                    attack["convertedEnergyCost"]])
    else:
        row.extend(["","","","",""])

def create_header_row(default_fields):
    row = default_fields.copy()
    row.extend(["ancient_trait_name",
                "ancient_trait_text",
                "type1",
                "type2",
                "ability_name",
                "ability_text",
                "text",
                "attack1_cost",
                "attack1_name",
                "attack1_text",
                "attack1_damage",
                "attack1_converted_energy_cost",
                "attack2_cost",
                "attack2_name",
                "attack2_text",
                "attack2_damage",
                "attack2_converted_energy_cost",
                "attack3_cost",
                "attack3_name",
                "attack3_text",
                "attack3_damage",
                "attack3_converted_energy_cost",
                "weakness_type",
                "weakness_value",
                "resistance_type",
                "resistance_value"])
    return row

if __name__ == '__main__':
    main()
