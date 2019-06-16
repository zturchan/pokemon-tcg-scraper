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
                          'ancient_trait',
                          'hp',
                          'number',
                          'artist',
                          'rarity',
                          'series',
                          'set',
                          'set_code',
                          'converted_retreat_cost',
                          'types',
                          'evolves_from']
    #cards = Card.all()
    cards = Card.where(set='generations')
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
        ws.append(row)


    wb.save("pkmn_output.xlsx")

def create_header_row(default_fields):
    row = default_fields.copy()
    return row

if __name__ == '__main__':
    main()
